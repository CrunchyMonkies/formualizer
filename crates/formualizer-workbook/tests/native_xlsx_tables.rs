#![cfg(feature = "xlsx-recalc")]
//! The native XLSX loader registers worksheet tables (ListObjects) with the
//! engine, so structured references evaluate instead of returning `#NAME?`
//! (psu3d0/formualizer#571).
use formualizer_common::ExcelErrorKind;
use formualizer_workbook::{
    LiteralValue, LoadStrategy, NativeXlsxAdapter, SpreadsheetReader, Workbook, WorkbookConfig,
};
use std::io::{Cursor, Write};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn s(cell: &str, text: &str) -> String {
    format!("<c r='{cell}' t='inlineStr'><is><t>{text}</t></is></c>")
}
fn n(cell: &str, value: f64) -> String {
    format!("<c r='{cell}'><v>{value}</v></c>")
}
fn f(cell: &str, formula: &str) -> String {
    format!("<c r='{cell}'><f>{formula}</f></c>")
}
fn sheet(rows: &[Vec<String>], tables: &[&str]) -> String {
    let data: String = rows
        .iter()
        .enumerate()
        .map(|(i, cells)| format!("<row r='{}'>{}</row>", i + 1, cells.concat()))
        .collect();
    let parts = if tables.is_empty() {
        String::new()
    } else {
        let parts: String = tables
            .iter()
            .map(|id| format!("<tablePart r:id='{id}'/>"))
            .collect();
        format!("<tableParts count='{}'>{parts}</tableParts>", tables.len())
    };
    format!(
        "<worksheet xmlns='{MAIN}' xmlns:r='{OFFICE}'><sheetData>{data}</sheetData>{parts}</worksheet>"
    )
}

/// Two worksheets: `Data` owns the tables, `Report` references them.
/// `tables` is `(rId, target, xml)` for `Data`'s table relationships; a
/// `None` xml leaves the target part out of the package.
fn workbook(data: &str, report: &str, tables: &[(&str, &str, Option<&str>)]) -> Vec<u8> {
    let mut parts = vec![
        (
            "_rels/.rels".to_owned(),
            format!(
                "<Relationships xmlns='{RELS}'><Relationship Id='r1' Type='{OFFICE}/officeDocument' Target='xl/workbook.xml'/></Relationships>"
            ),
        ),
        (
            "xl/workbook.xml".to_owned(),
            format!(
                "<workbook xmlns='{MAIN}' xmlns:r='{OFFICE}'><sheets><sheet name='Data' sheetId='1' r:id='r1'/><sheet name='Report' sheetId='2' r:id='r2'/></sheets></workbook>"
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            format!(
                "<Relationships xmlns='{RELS}'><Relationship Id='r1' Type='{OFFICE}/worksheet' Target='worksheets/sheet1.xml'/><Relationship Id='r2' Type='{OFFICE}/worksheet' Target='worksheets/sheet2.xml'/></Relationships>"
            ),
        ),
        ("xl/worksheets/sheet1.xml".to_owned(), data.to_owned()),
        ("xl/worksheets/sheet2.xml".to_owned(), report.to_owned()),
    ];
    if !tables.is_empty() {
        let rels: String = tables
            .iter()
            .map(|(id, target, _)| {
                format!(
                    "<Relationship Id='{id}' Type='{OFFICE}/table' Target='../tables/{target}'/>"
                )
            })
            .collect();
        parts.push((
            "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
            format!("<Relationships xmlns='{RELS}'>{rels}</Relationships>"),
        ));
        for (_, target, xml) in tables {
            if let Some(xml) = xml {
                parts.push((format!("xl/tables/{target}"), (*xml).to_owned()));
            }
        }
    }
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (path, data) in parts {
        zip.start_file(path, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// `Sales` (A1:B5): header row, three body rows, a totals row.
fn sales_table() -> String {
    format!(
        "<table xmlns='{MAIN}' id='1' name='Table1' displayName='Sales' ref='A1:B5' totalsRowCount='1'><autoFilter ref='A1:B4'/><tableColumns count='2'><tableColumn id='1' name='Qty' totalsRowFunction='sum'/><tableColumn id='2' name='Price'/></tableColumns><tableStyleInfo name='TableStyleMedium2' showRowStripes='1'/></table>"
    )
}
/// `Raw` (D1:D3): no header row, three body rows.
fn raw_table() -> String {
    format!(
        "<table xmlns='{MAIN}' id='2' name='Raw' displayName='Raw' ref='D1:D3' headerRowCount='0'><tableColumns count='1'><tableColumn id='1' name='Val'/></tableColumns></table>"
    )
}
fn data_sheet(tables: &[&str]) -> String {
    sheet(
        &[
            vec![s("A1", "Qty"), s("B1", "Price"), n("D1", 7.0)],
            vec![n("A2", 1.0), n("B2", 10.0), n("D2", 8.0)],
            vec![n("A3", 2.0), n("B3", 20.0), n("D3", 9.0)],
            vec![n("A4", 3.0), n("B4", 30.0)],
            vec![
                f("A5", "SUBTOTAL(109,Sales[Qty])"),
                f("F5", "SUM(Sales[Price])"),
            ],
        ],
        tables,
    )
}
fn report_sheet() -> String {
    sheet(
        &[vec![
            f("A1", "SUM(Sales[Qty])"),
            f("B1", "ROWS(Sales[#Data])"),
            f("C1", "ROWS(Sales[#All])"),
            f("D1", "COUNTA(Sales[#Headers])"),
            f("E1", "SUM(Raw[Val])"),
            f("F1", "ROWS(Raw[Val])"),
            f("G1", "SUM(Sales[#Totals])"),
            f("H1", "SUM(Raw[#All])"),
        ]],
        &[],
    )
}

fn load(bytes: Vec<u8>) -> Workbook {
    Workbook::from_reader(
        NativeXlsxAdapter::open_bytes(bytes).unwrap(),
        LoadStrategy::EagerAll,
        WorkbookConfig::ephemeral(),
    )
    .unwrap()
}
fn number(wb: &mut Workbook, sheet: &str, row: u32, col: u32) -> f64 {
    match wb.evaluate_cell(sheet, row, col).unwrap() {
        LiteralValue::Number(value) => value,
        LiteralValue::Int(value) => value as f64,
        other => panic!("{sheet}!R{row}C{col} = {other:?}"),
    }
}
fn is_name_error(wb: &mut Workbook, sheet: &str, row: u32, col: u32) -> bool {
    matches!(
        wb.evaluate_cell(sheet, row, col).unwrap(),
        LiteralValue::Error(e) if e.kind == ExcelErrorKind::Name
    )
}

#[test]
fn native_loader_registers_tables_for_structured_references() {
    let sales = sales_table();
    let raw = raw_table();
    let bytes = workbook(
        &data_sheet(&["rId1", "rId2"]),
        &report_sheet(),
        &[
            ("rId1", "table1.xml", Some(&sales)),
            ("rId2", "table2.xml", Some(&raw)),
        ],
    );
    let mut wb = load(bytes);
    // The engine knows both tables by their display names.
    let mut names: Vec<_> = wb.tables().into_iter().map(|t| t.name).collect();
    names.sort();
    assert_eq!(names, ["Raw", "Sales"]);

    // Same-sheet whole-column references, including in the totals row.
    assert_eq!(number(&mut wb, "Data", 5, 1), 6.0);
    assert_eq!(number(&mut wb, "Data", 5, 6), 60.0);
    // Cross-sheet references: whole column, ROWS, #All, #Headers, #Totals.
    // (A bare `ROWS(Sales)` is not covered: the engine does not resolve bare
    // table names on any load path; source recalculation lowers them first.
    // Combined items such as `Sales[[#Totals],[Qty]]` are likewise `#N/IMPL`
    // in the engine itself, independent of how the table was registered.)
    assert_eq!(number(&mut wb, "Report", 1, 1), 6.0);
    assert_eq!(number(&mut wb, "Report", 1, 2), 3.0);
    assert_eq!(number(&mut wb, "Report", 1, 3), 5.0);
    assert_eq!(number(&mut wb, "Report", 1, 4), 2.0);
    assert_eq!(number(&mut wb, "Report", 1, 7), 6.0);
    // Header-less table: every row of the range is body.
    assert_eq!(number(&mut wb, "Report", 1, 5), 24.0);
    assert_eq!(number(&mut wb, "Report", 1, 6), 3.0);
    assert_eq!(number(&mut wb, "Report", 1, 8), 24.0);
}

#[test]
fn malformed_or_missing_table_parts_are_skipped_without_failing_the_load() {
    let raw = raw_table();
    for broken in [
        // tableColumns disagrees with the ref width.
        Some(format!(
            "<table xmlns='{MAIN}' id='1' name='Sales' displayName='Sales' ref='A1:B5'><tableColumns count='1'><tableColumn id='1' name='Qty'/></tableColumns></table>"
        )),
        // Unparseable range.
        Some(format!(
            "<table xmlns='{MAIN}' id='1' name='Sales' displayName='Sales' ref='nope'><tableColumns count='2'><tableColumn id='1' name='Qty'/><tableColumn id='2' name='Price'/></tableColumns></table>"
        )),
        // Not XML at all.
        Some("<table".to_owned()),
        // Relationship target missing from the package.
        None,
    ] {
        let bytes = workbook(
            &data_sheet(&["rId1", "rId2"]),
            &report_sheet(),
            &[
                ("rId1", "table1.xml", broken.as_deref()),
                ("rId2", "table2.xml", Some(&raw)),
            ],
        );
        let mut wb = load(bytes);
        assert!(is_name_error(&mut wb, "Report", 1, 1), "{broken:?}");
        assert_eq!(number(&mut wb, "Report", 1, 5), 24.0, "{broken:?}");
    }
}

#[test]
fn excel_written_table_workbook_evaluates_whole_column_references() {
    let bytes = include_bytes!("fixtures/excel_tables.xlsx").to_vec();
    let mut wb = load(bytes);
    assert_eq!(
        wb.tables().into_iter().map(|t| t.name).collect::<Vec<_>>(),
        ["Sales"]
    );
    // Totals row: SUBTOTAL(109,Sales[Qty]) over Qty 10/20/3.
    assert_eq!(number(&mut wb, "Data", 5, 2), 33.0);
    // Defined name next to the table, and a cross-sheet #Headers reference.
    assert_eq!(number(&mut wb, "Summary", 1, 4), 12.75);
    assert_eq!(number(&mut wb, "Summary", 1, 5), 5.0);
    // Calculated-column cells use `[#This Row]`, which the engine does not
    // evaluate yet; whatever they return, it must no longer be `#NAME?`.
    for (sheet, row, col) in [("Data", 2, 4), ("Summary", 1, 1)] {
        assert!(
            !is_name_error(&mut wb, sheet, row, col),
            "{sheet}!R{row}C{col}"
        );
    }
}

#[test]
fn reader_surface_still_opens_table_workbooks() {
    let bytes = include_bytes!("fixtures/excel_tables.xlsx").to_vec();
    let mut adapter = NativeXlsxAdapter::open_bytes(bytes).unwrap();
    assert_eq!(adapter.sheet_names().unwrap(), ["Data", "Summary"]);
    let data = adapter.read_sheet("Data").unwrap();
    assert!(data.cells.contains_key(&(5, 2)));
}
