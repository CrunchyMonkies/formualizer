//! `Workbook::to_xlsx_bytes` writes the engine's defined names and tables, so
//! a saved `=SUM(Prices)` or `=SUM(Sales[Total])` reloads with its value
//! instead of `#NAME?`.
//!
//! Tables reload through Umya here. The native reader registers tables only
//! with `fix/native-xlsx-tables`; the native checks below compare table
//! results whenever that reader registered the tables, so they run end to
//! end on a branch that has it and are skipped (names still checked) on one
//! that does not.
#![cfg(all(feature = "umya", feature = "xlsx-recalc"))]
use formualizer_workbook::traits::NamedRangeScope;
use formualizer_workbook::{
    LiteralValue, LoadStrategy, NativeXlsxAdapter, RangeAddress, SpreadsheetReader, Workbook,
    WorkbookConfig,
};
use std::io::Read;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Reader {
    Native,
    Umya,
}

fn load(bytes: Vec<u8>, reader: Reader) -> Workbook {
    let config = WorkbookConfig::interactive();
    let mut wb = match reader {
        Reader::Native => Workbook::from_reader(
            NativeXlsxAdapter::open_bytes(bytes).unwrap(),
            LoadStrategy::EagerAll,
            config,
        )
        .unwrap(),
        Reader::Umya => Workbook::from_reader(
            formualizer_workbook::UmyaAdapter::open_bytes(bytes).unwrap(),
            LoadStrategy::EagerAll,
            config,
        )
        .unwrap(),
    };
    wb.evaluate_all().unwrap();
    wb
}

fn part(bytes: &[u8], name: &str) -> String {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut out = String::new();
    archive
        .by_name(name)
        .unwrap_or_else(|e| panic!("{name}: {e}"))
        .read_to_string(&mut out)
        .unwrap();
    out
}

fn n(value: f64) -> Option<LiteralValue> {
    Some(LiteralValue::Number(value))
}

fn text(value: &str) -> Option<LiteralValue> {
    Some(LiteralValue::Text(value.to_string()))
}

/// Every cell value of every sheet, by address.
fn values(wb: &Workbook) -> Vec<(String, u32, u32, Option<LiteralValue>)> {
    let mut out = Vec::new();
    for sheet in wb.sheet_names() {
        let Some((rows, cols)) = wb.sheet_dimensions(&sheet) else {
            continue;
        };
        for row in 1..=rows {
            for col in 1..=cols {
                let value = wb.get_value(&sheet, row, col);
                if value.is_some() {
                    out.push((sheet.clone(), row, col, value));
                }
            }
        }
    }
    out
}

/// Normalize integral numbers so a reader's Int and Number compare equal.
fn normalized(v: Vec<(String, u32, u32, Option<LiteralValue>)>) -> Vec<String> {
    v.into_iter()
        .map(|(s, r, c, value)| {
            let value = match value {
                Some(LiteralValue::Int(i)) => Some(LiteralValue::Number(i as f64)),
                other => other,
            };
            format!("{s}!R{r}C{c}={value:?}")
        })
        .collect()
}

fn has_name_error(wb: &Workbook) -> bool {
    values(wb).into_iter().any(|(_, _, _, v)| {
        matches!(v, Some(LiteralValue::Error(e)) if e.kind == formualizer_common::ExcelErrorKind::Name)
    })
}

/// Save `wb`, reload it through Umya and (where it registered the tables)
/// the native reader, and check every value matches. Returns the package.
fn assert_roundtrip(wb: &Workbook) -> Vec<u8> {
    let expected = normalized(values(wb));
    assert!(
        !has_name_error(wb),
        "source workbook has #NAME?: {expected:#?}"
    );
    let bytes = wb.to_xlsx_bytes().unwrap();
    for reader in [Reader::Umya, Reader::Native] {
        let reloaded = load(bytes.clone(), reader);
        if reader == Reader::Native && reloaded.tables().len() != wb.tables().len() {
            // The native reader on this branch does not register tables.
            continue;
        }
        assert_eq!(
            reloaded.tables(),
            wb.tables(),
            "{reader:?}: tables after reload"
        );
        assert_eq!(
            normalized(values(&reloaded)),
            expected,
            "{reader:?}: values after reload"
        );
    }
    bytes
}

fn range(sheet: &str, r1: u32, c1: u32, r2: u32, c2: u32) -> RangeAddress {
    RangeAddress::new(sheet, r1, c1, r2, c2).unwrap()
}

#[test]
fn workbook_and_sheet_scoped_names_round_trip() {
    let mut wb = Workbook::new_with_config(WorkbookConfig::interactive());
    wb.add_sheet("Data").unwrap();
    wb.add_sheet("Other Sheet").unwrap();
    for (row, value) in [(1, 2.0), (2, 3.0), (3, 5.0)] {
        wb.set_value("Data", row, 1, LiteralValue::Number(value))
            .unwrap();
        wb.set_value("Other Sheet", row, 1, LiteralValue::Number(value * 10.0))
            .unwrap();
    }
    wb.define_named_range(
        "Prices",
        &range("Data", 1, 1, 3, 1),
        NamedRangeScope::Workbook,
    )
    .unwrap();
    wb.define_named_range(
        "Rate",
        &range("Data", 2, 1, 2, 1),
        NamedRangeScope::Workbook,
    )
    .unwrap();
    // Same name, scoped to each sheet.
    wb.define_named_range("Local", &range("Data", 1, 1, 1, 1), NamedRangeScope::Sheet)
        .unwrap();
    wb.define_named_range(
        "Local",
        &range("Other Sheet", 1, 1, 3, 1),
        NamedRangeScope::Sheet,
    )
    .unwrap();
    wb.set_formula("Data", 1, 2, "=SUM(Prices)").unwrap();
    wb.set_formula("Data", 2, 2, "=Rate*2").unwrap();
    wb.set_formula("Data", 3, 2, "=SUM(Local)").unwrap();
    wb.set_formula("Other Sheet", 1, 2, "=SUM(Local)").unwrap();
    wb.set_formula("Other Sheet", 2, 2, "=SUM(Prices)+Rate")
        .unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(wb.get_value("Data", 1, 2), n(10.0));
    assert_eq!(wb.get_value("Data", 3, 2), n(2.0));
    assert_eq!(wb.get_value("Other Sheet", 1, 2), n(100.0));

    let bytes = assert_roundtrip(&wb);
    let workbook = part(&bytes, "xl/workbook.xml");
    assert!(
        workbook.contains(r#"<definedName name="Prices">Data!$A$1:$A$3</definedName>"#),
        "{workbook}"
    );
    assert!(
        workbook.contains(r#"<definedName name="Local" localSheetId="1">Data!$A$1</definedName>"#),
        "{workbook}"
    );
    assert!(
        workbook.contains(
            r#"<definedName name="Local" localSheetId="2">'Other Sheet'!$A$1:$A$3</definedName>"#
        ),
        "{workbook}"
    );
    // `definedNames` follows `sheets` and precedes `calcPr` (CT_Workbook).
    let (sheets, names, calc) = (
        workbook.find("</sheets>").unwrap(),
        workbook.find("<definedNames>").unwrap(),
        workbook.find("<calcPr").unwrap(),
    );
    assert!(sheets < names && names < calc, "{workbook}");

    for reader in [Reader::Native, Reader::Umya] {
        let reloaded = load(bytes.clone(), reader);
        assert!(reloaded.has_name("Prices", None), "{reader:?}");
        assert!(reloaded.has_name("Local", Some("Data")), "{reader:?}");
        assert!(!reloaded.has_name("Local", None), "{reader:?}");
    }
}

/// `Sales` (header + totals) on `Data`; read from `Summary` by structured
/// references.
fn sales_workbook() -> Workbook {
    let mut wb = Workbook::new_with_config(WorkbookConfig::interactive());
    wb.add_sheet("Data").unwrap();
    wb.add_sheet("Summary").unwrap();
    for (col, header) in ["Item", "Qty", "Price", "Total"].iter().enumerate() {
        wb.set_value(
            "Data",
            1,
            col as u32 + 1,
            LiteralValue::Text(header.to_string()),
        )
        .unwrap();
    }
    for (row, (item, qty, price)) in [("Apple", 3.0, 2.0), ("Pear", 7.0, 1.5), ("Fig", 2.0, 4.0)]
        .into_iter()
        .enumerate()
    {
        let row = row as u32 + 2;
        wb.set_value("Data", row, 1, LiteralValue::Text(item.into()))
            .unwrap();
        wb.set_value("Data", row, 2, LiteralValue::Number(qty))
            .unwrap();
        wb.set_value("Data", row, 3, LiteralValue::Number(price))
            .unwrap();
        wb.set_formula("Data", row, 4, &format!("=B{row}*C{row}"))
            .unwrap();
    }
    wb.set_value("Data", 5, 1, LiteralValue::Text("Total".into()))
        .unwrap();
    wb.set_formula("Data", 5, 2, "=SUBTOTAL(109,Sales[Qty])")
        .unwrap();
    wb.set_formula("Data", 5, 4, "=SUBTOTAL(109,Sales[Total])")
        .unwrap();
    wb.define_table(
        "Sales",
        "Data",
        (1, 1, 5, 4),
        vec!["Item".into(), "Qty".into(), "Price".into(), "Total".into()],
        true,
        true,
    )
    .unwrap();
    wb.set_formula("Summary", 1, 1, "=SUM(Sales[Total])")
        .unwrap();
    wb.set_formula("Summary", 2, 1, "=Sales[[#Totals],[Total]]")
        .unwrap();
    wb.set_formula("Summary", 3, 1, "=COUNTA(Sales[#Headers])")
        .unwrap();
    wb.set_formula("Summary", 4, 1, "=AVERAGE(Sales[Price])")
        .unwrap();
    wb.set_formula("Summary", 5, 1, "=ROWS(Sales[Item])")
        .unwrap();
    wb.evaluate_all().unwrap();
    wb
}

#[test]
fn table_with_header_and_totals_round_trips() {
    let wb = sales_workbook();
    assert_eq!(wb.get_value("Summary", 1, 1), n(24.5));
    // `[#Totals]` evaluates with the this-row structured-reference engine
    // work; without it the value is `#N/IMPL!` both before and after saving.
    let totals = wb.get_value("Summary", 2, 1);
    assert!(
        totals == n(24.5)
            || matches!(&totals, Some(LiteralValue::Error(e)) if e.kind == formualizer_common::ExcelErrorKind::NImpl),
        "{totals:?}"
    );
    assert_eq!(wb.get_value("Summary", 3, 1), n(4.0));
    assert_eq!(wb.get_value("Summary", 4, 1), n(2.5));
    assert_eq!(wb.get_value("Summary", 5, 1), n(3.0));

    let bytes = assert_roundtrip(&wb);
    let table = part(&bytes, "xl/tables/table1.xml");
    assert!(
        table
            .contains(r#"id="1" name="Sales" displayName="Sales" ref="A1:D5" totalsRowCount="1" totalsRowShown="1">"#),
        "{table}"
    );
    assert!(table.contains(r#"<autoFilter ref="A1:D4"/>"#), "{table}");
    assert!(
        table.contains(r#"<tableColumns count="4"><tableColumn id="1" name="Item"/>"#),
        "{table}"
    );
    assert!(!table.contains("headerRowCount"), "{table}");
    // Sheets: Sheet1 (the default), Data, Summary.
    let sheet = part(&bytes, "xl/worksheets/sheet2.xml");
    assert!(
        sheet.contains(r#"<tableParts count="1"><tablePart r:id="#),
        "{sheet}"
    );
    let rels = part(&bytes, "xl/worksheets/_rels/sheet2.xml.rels");
    assert!(rels.contains(r#"Target="../tables/table1.xml""#), "{rels}");
    let types = part(&bytes, "[Content_Types].xml");
    assert!(
        types.contains(
            r#"<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>"#
        ),
        "{types}"
    );
}

#[test]
fn headerless_table_round_trips() {
    let mut wb = Workbook::new_with_config(WorkbookConfig::interactive());
    wb.add_sheet("S").unwrap();
    for (row, value) in [(2, 4.0), (3, 6.0), (4, 8.0)] {
        wb.set_value("S", row, 2, LiteralValue::Number(value))
            .unwrap();
        wb.set_value("S", row, 3, LiteralValue::Number(value / 2.0))
            .unwrap();
    }
    wb.define_table(
        "Raw",
        "S",
        (2, 2, 4, 3),
        vec!["Column1".into(), "Column2".into()],
        false,
        false,
    )
    .unwrap();
    wb.set_formula("S", 6, 1, "=SUM(Raw[Column1])").unwrap();
    wb.set_formula("S", 7, 1, "=SUM(Raw[Column2])").unwrap();
    wb.set_formula("S", 8, 1, "=COUNT(Raw[Column2])").unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(wb.get_value("S", 6, 1), n(18.0));
    assert_eq!(wb.get_value("S", 8, 1), n(3.0));

    let bytes = assert_roundtrip(&wb);
    let table = part(&bytes, "xl/tables/table1.xml");
    assert!(
        table.contains(r#"ref="B2:C4" headerRowCount="0" totalsRowShown="0">"#),
        "{table}"
    );
    assert!(!table.contains("autoFilter"), "{table}");
}

#[test]
fn tables_on_two_sheets_round_trip() {
    let mut wb = sales_workbook();
    wb.add_sheet("Costs").unwrap();
    wb.set_value("Costs", 3, 2, LiteralValue::Text("Kind".into()))
        .unwrap();
    wb.set_value("Costs", 3, 3, LiteralValue::Text("Amount".into()))
        .unwrap();
    for (row, (kind, amount)) in [("Rent", 100.0), ("Power", 40.0)].into_iter().enumerate() {
        let row = row as u32 + 4;
        wb.set_value("Costs", row, 2, LiteralValue::Text(kind.into()))
            .unwrap();
        wb.set_value("Costs", row, 3, LiteralValue::Number(amount))
            .unwrap();
    }
    wb.define_table(
        "Costs",
        "Costs",
        (3, 2, 5, 3),
        vec!["Kind".into(), "Amount".into()],
        true,
        false,
    )
    .unwrap();
    wb.set_formula("Summary", 6, 1, "=SUM(Sales[Total])-SUM(Costs[Amount])")
        .unwrap();
    wb.set_formula("Summary", 7, 1, "=COUNTA(Costs[Kind])&\" kinds\"")
        .unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(wb.get_value("Summary", 6, 1), n(24.5 - 140.0));
    assert_eq!(wb.get_value("Summary", 7, 1), text("2 kinds"));

    let bytes = assert_roundtrip(&wb);
    // Tables are written in name order: Costs, then Sales.
    let costs = part(&bytes, "xl/tables/table1.xml");
    assert!(
        costs
            .contains(r#"id="1" name="Costs" displayName="Costs" ref="B3:C5" totalsRowShown="0">"#),
        "{costs}"
    );
    assert!(costs.contains(r#"<autoFilter ref="B3:C5"/>"#), "{costs}");
    let sales = part(&bytes, "xl/tables/table2.xml");
    assert!(sales.contains(r#"id="2" name="Sales""#), "{sales}");
    let costs_sheet = part(&bytes, "xl/worksheets/sheet4.xml");
    assert!(
        costs_sheet.contains("<tableParts count=\"1\">"),
        "{costs_sheet}"
    );
    let rels = part(&bytes, "xl/worksheets/_rels/sheet4.xml.rels");
    assert!(rels.contains(r#"Target="../tables/table1.xml""#), "{rels}");
}

#[test]
fn excel_tables_fixture_round_trips() {
    let source = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/excel_tables.xlsx"
    ))
    .unwrap();
    // LibreOffice wrote this file with shared structured-reference formulas,
    // which Umya 2.3.2 cannot tokenize (its bracket scanner never advances),
    // so the source loads through the native reader only.
    let mut wb = load(source, Reader::Native);
    if wb.tables().is_empty() {
        // Without the native table loader, register `Sales` (part name
        // `Table1`) as the file declares it.
        wb.define_table(
            "Sales",
            "Data",
            (1, 1, 5, 5),
            ["Item", "Qty", "Unit Price", "Total", "Running"]
                .map(String::from)
                .to_vec(),
            true,
            true,
        )
        .unwrap();
        wb.evaluate_all().unwrap();
    }
    assert!(wb.has_name("Prices", None));
    let bytes = assert_roundtrip(&wb);
    let workbook = part(&bytes, "xl/workbook.xml");
    assert!(
        workbook.contains(r#"<definedName name="Prices">Data!$C$2:$C$4</definedName>"#),
        "{workbook}"
    );
    let table = part(&bytes, "xl/tables/table1.xml");
    assert!(
        table.contains(
            r#"name="Sales" displayName="Sales" ref="A1:E5" totalsRowCount="1" totalsRowShown="1">"#
        ),
        "{table}"
    );
    let data = part(&bytes, "xl/worksheets/sheet1.xml");
    assert!(
        data.contains("<f>Sales[[#This Row],[Qty]] * Sales[[#This Row],[Unit Price]]</f>"),
        "{data}"
    );
}
