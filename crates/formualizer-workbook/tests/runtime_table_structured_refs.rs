//! Structured-reference forms Excel uses for calculated columns and totals:
//! `#This Row` (`Table1[@Qty]`, `Table1[[#This Row],[Qty]]`, spans), section
//! and column combinations (`Table1[[#Totals],[Total]]`), unqualified
//! references inside a table, and a bare table name.
use formualizer_common::{ExcelErrorKind, LiteralValue};
use formualizer_workbook::Workbook;

const HEADERS: [&str; 4] = ["Item", "Qty", "Unit Price", "Total"];

/// `Table1` at A1:D5 on sheet `S`: header row 1, data rows 2..=4 with
/// Qty = 2, 3, 4 and Unit Price = 10, 20, 30, totals row 5 (Total = 999).
fn workbook() -> Workbook {
    let mut wb = Workbook::new();
    wb.add_sheet("S").unwrap();
    wb.add_sheet("Other").unwrap();
    for (c, h) in HEADERS.iter().enumerate() {
        wb.set_value("S", 1, c as u32 + 1, LiteralValue::Text((*h).into()))
            .unwrap();
    }
    for r in 2..=4u32 {
        wb.set_value("S", r, 1, LiteralValue::Text(format!("item{r}")))
            .unwrap();
        wb.set_value("S", r, 2, LiteralValue::Number(f64::from(r)))
            .unwrap();
        wb.set_value("S", r, 3, LiteralValue::Number(f64::from(r - 1) * 10.0))
            .unwrap();
    }
    wb.set_value("S", 5, 4, LiteralValue::Number(999.0))
        .unwrap();
    wb.define_table(
        "Table1",
        "S",
        (1, 1, 5, 4),
        HEADERS.iter().map(|h| h.to_string()).collect(),
        true,
        true,
    )
    .unwrap();
    wb
}

fn num(wb: &Workbook, sheet: &str, row: u32, col: u32) -> f64 {
    match wb.get_value(sheet, row, col) {
        Some(LiteralValue::Number(n)) => n,
        Some(LiteralValue::Int(n)) => n as f64,
        other => panic!("{sheet}!R{row}C{col}: expected a number, got {other:?}"),
    }
}

fn error_kind(wb: &Workbook, sheet: &str, row: u32, col: u32) -> ExcelErrorKind {
    match wb.get_value(sheet, row, col) {
        Some(LiteralValue::Error(e)) => e.kind,
        other => panic!("{sheet}!R{row}C{col}: expected an error, got {other:?}"),
    }
}

/// Fill the Total column (D2:D4) with one formula text, like an Excel
/// calculated column, and return the per-row results.
fn fill_down(wb: &mut Workbook, formula: &str) -> Vec<f64> {
    for r in 2..=4 {
        wb.set_formula("S", r, 4, formula).unwrap();
    }
    wb.evaluate_all().unwrap();
    (2..=4).map(|r| num(wb, "S", r, 4)).collect()
}

#[test]
fn this_row_spellings_fill_down_with_per_row_results() {
    for formula in [
        "=Table1[[#This Row],[Qty]]*Table1[[#This Row],[Unit Price]]",
        "=Table1[@Qty]*Table1[@[Unit Price]]",
        "=Table1[@[Qty]]*Table1[@[Unit Price]]",
        "=Table1[[@],[Qty]]*Table1[[@],[Unit Price]]",
        "=[@Qty]*[@[Unit Price]]",
        "=PRODUCT(Table1[@[Qty]:[Unit Price]])",
        "=PRODUCT(Table1[[#This Row],[Qty]:[Unit Price]])",
        "=PRODUCT([@[Qty]:[Unit Price]])",
    ] {
        let mut wb = workbook();
        assert_eq!(
            fill_down(&mut wb, formula),
            vec![20.0, 60.0, 120.0],
            "{formula}"
        );
    }
}

#[test]
fn whole_this_row_selects_every_column_of_the_row() {
    let mut wb = workbook();
    wb.set_formula("S", 3, 6, "=COUNT(Table1[@])").unwrap();
    wb.set_formula("S", 3, 7, "=COUNTA(Table1[[#This Row]])")
        .unwrap();
    wb.evaluate_all().unwrap();
    // Item is text, Qty and Unit Price are numbers, Total is empty.
    assert_eq!(num(&wb, "S", 3, 6), 2.0);
    assert_eq!(num(&wb, "S", 3, 7), 3.0);
}

#[test]
fn editing_an_input_recomputes_only_its_rows_formula() {
    let mut wb = workbook();
    assert_eq!(
        fill_down(&mut wb, "=Table1[@Qty]*Table1[@[Unit Price]]"),
        vec![20.0, 60.0, 120.0]
    );
    wb.set_value("S", 3, 2, LiteralValue::Number(7.0)).unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(
        (2..=4).map(|r| num(&wb, "S", r, 4)).collect::<Vec<_>>(),
        vec![20.0, 140.0, 120.0]
    );
    wb.set_value("S", 4, 3, LiteralValue::Number(1.0)).unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(num(&wb, "S", 4, 4), 4.0);

    // A this-row reference beside the table follows its row as well.
    wb.set_formula("S", 2, 6, "=Table1[[#This Row],[Qty]]+1")
        .unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(num(&wb, "S", 2, 6), 3.0);
    wb.set_value("S", 2, 2, LiteralValue::Number(10.0)).unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(num(&wb, "S", 2, 6), 11.0);
}

#[test]
fn this_row_outside_the_data_rows_is_value_error() {
    let mut wb = workbook();
    let cells = [
        ("S", 1, 6),     // beside the header row
        ("S", 5, 6),     // beside the totals row
        ("S", 5, 1),     // in the totals row
        ("S", 9, 2),     // below the table
        ("Other", 3, 1), // a data row number, but another sheet
    ];
    for (sheet, row, col) in cells {
        wb.set_formula(sheet, row, col, "=Table1[@Qty]").unwrap();
    }
    wb.set_formula("S", 10, 2, "=IFERROR(Table1[[#This Row],[Qty]],-1)")
        .unwrap();
    wb.evaluate_all().unwrap();
    for (sheet, row, col) in cells {
        assert_eq!(
            error_kind(&wb, sheet, row, col),
            ExcelErrorKind::Value,
            "{sheet}!R{row}C{col}"
        );
    }
    assert_eq!(num(&wb, "S", 10, 2), -1.0);
}

#[test]
fn section_and_column_combinations_select_excel_areas() {
    let mut wb = workbook();
    let cases = [
        ("=SUM(Table1[[Qty]:[Unit Price]])", 69.0),
        ("=SUM(Table1[[#Totals],[Total]])", 999.0),
        ("=COUNTA(Table1[[#Headers],[Qty]])", 1.0),
        ("=COUNTA(Table1[[#Headers],[#Data],[Qty]])", 4.0),
        ("=SUM(Table1[[#Data],[#Totals],[Total]])", 999.0),
        ("=COUNTA(Table1[[#Data],[#Totals],[Item]])", 3.0),
        ("=ROWS(Table1[[#Data],[#Totals],[Qty]])", 4.0),
        ("=COLUMNS(Table1[[#Headers],[Qty]:[Total]])", 3.0),
        ("=ROWS(Table1[[#All],[Qty]])", 5.0),
        ("=SUM(Table1)", 69.0),
        ("=ROWS(Table1)", 3.0),
        ("=COLUMNS(Table1)", 4.0),
    ];
    for (i, (formula, _)) in cases.iter().enumerate() {
        wb.set_formula("Other", i as u32 + 1, 1, formula).unwrap();
    }
    wb.evaluate_all().unwrap();
    for (i, (formula, expected)) in cases.iter().enumerate() {
        assert_eq!(num(&wb, "Other", i as u32 + 1, 1), *expected, "{formula}");
    }
}

#[test]
fn bare_table_name_tracks_edits() {
    let mut wb = workbook();
    wb.set_formula("Other", 1, 1, "=SUM(Table1)").unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(num(&wb, "Other", 1, 1), 69.0);
    wb.set_value("S", 2, 3, LiteralValue::Number(110.0))
        .unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(num(&wb, "Other", 1, 1), 169.0);
}

#[test]
fn unqualified_references_inside_a_table_bind_to_that_table() {
    let mut wb = workbook();
    wb.set_formula("S", 3, 4, "=[@Qty]+[@[Unit Price]]")
        .unwrap();
    // This-row in the totals row has no data row.
    wb.set_formula("S", 5, 2, "=[@Qty]").unwrap();
    // A section/column combination names the containing table.
    wb.set_formula("S", 5, 3, "=SUM([[#Data],[Unit Price]])")
        .unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(num(&wb, "S", 3, 4), 23.0);
    assert_eq!(error_kind(&wb, "S", 5, 2), ExcelErrorKind::Value);
    assert_eq!(num(&wb, "S", 5, 3), 60.0);
    assert_eq!(
        wb.get_formula("S", 5, 3).as_deref(),
        Some("=SUM(Table1[[#Data],[Unit Price]])")
    );
}

#[test]
fn this_row_formulas_keep_their_structured_text() {
    let mut wb = workbook();
    for formula in [
        "=Table1[[#This Row],[Qty]]*2",
        "=Table1[@Qty]*2",
        "=SUM(Table1[@[Qty]:[Unit Price]])",
    ] {
        wb.set_formula("S", 3, 4, formula).unwrap();
        wb.evaluate_all().unwrap();
        let text = wb.get_formula("S", 3, 4).unwrap();
        assert!(text.contains("Table1["), "{formula} printed as {text}");
        assert!(!text.contains('$'), "{formula} printed as {text}");
    }
}

#[test]
fn rows_added_by_resizing_get_this_row_results() {
    let mut wb = workbook();
    // Move the totals row down: the table becomes A1:D6 with data rows 2..=5.
    wb.set_value("S", 5, 4, LiteralValue::Empty).unwrap();
    wb.set_value("S", 5, 2, LiteralValue::Number(5.0)).unwrap();
    wb.set_value("S", 5, 3, LiteralValue::Number(40.0)).unwrap();
    wb.resize_table("Table1", 6).unwrap();
    for r in 2..=5 {
        wb.set_formula("S", r, 4, "=Table1[@Qty]*Table1[@[Unit Price]]")
            .unwrap();
    }
    wb.set_formula("Other", 1, 1, "=SUM(Table1[Total])")
        .unwrap();
    wb.evaluate_all().unwrap();
    assert_eq!(num(&wb, "S", 5, 4), 200.0);
    assert_eq!(num(&wb, "Other", 1, 1), 400.0);
}
