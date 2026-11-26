use crate::engine::{Engine, EvalConfig};
use crate::reference::{CellRef, Coord, RangeRef};
use formualizer_common::{ExcelErrorKind, LiteralValue};

#[test]
fn structured_ref_table_column_tracks_cell_edits_via_table_vertex() {
    let ctx = crate::test_workbook::TestWorkbook::new();
    let mut engine: Engine<_> = Engine::new(ctx, EvalConfig::default());

    engine.add_sheet("Sheet1").unwrap();

    // Table region A1:B3 (header + 2 data rows)
    // Headers: Region, Amount
    // Make the non-selected column numeric so we catch over-wide selection bugs.
    engine
        .set_cell_value("Sheet1", 2, 1, LiteralValue::Number(5.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 2, 2, LiteralValue::Number(10.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 3, 1, LiteralValue::Number(7.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 3, 2, LiteralValue::Number(20.0))
        .unwrap();

    let sheet_id = engine.sheet_id("Sheet1").unwrap();
    let start = CellRef::new(sheet_id, Coord::from_excel(1, 1, true, true));
    let end = CellRef::new(sheet_id, Coord::from_excel(3, 2, true, true));
    let range = RangeRef::new(start, end);
    engine
        .define_table(
            "Sales",
            range,
            true,
            vec!["Region".into(), "Amount".into()],
            false,
        )
        .unwrap();

    let ast = formualizer_parse::parser::parse("=SUM(Sales[Amount])").unwrap();
    engine.set_cell_formula("Sheet1", 1, 4, ast).unwrap();

    let v = engine
        .evaluate_cell("Sheet1", 1, 4)
        .unwrap()
        .expect("computed value");
    assert_eq!(v, LiteralValue::Number(30.0));

    // Edit a precedent cell inside the table and ensure the table-dependent formula is dirtied.
    engine
        .set_cell_value("Sheet1", 2, 2, LiteralValue::Number(100.0))
        .unwrap();
    let v2 = engine
        .evaluate_cell("Sheet1", 1, 4)
        .unwrap()
        .expect("computed value");
    assert_eq!(v2, LiteralValue::Number(120.0));
}

#[test]
fn table_resize_without_cell_edits_rebuilds_dependencies_and_journal_replay() {
    use crate::engine::ChangeEvent;
    use crate::engine::graph::editor::undo_engine::UndoEngine;
    let mut engine = Engine::new(
        crate::test_workbook::TestWorkbook::new(),
        EvalConfig::default(),
    );
    engine.add_sheet("Sheet1").unwrap();
    for row in 2..=4 {
        engine
            .set_cell_value("Sheet1", row, 1, LiteralValue::Number(row as f64))
            .unwrap();
    }
    let sheet = engine.sheet_id("Sheet1").unwrap();
    engine
        .define_table(
            "Orders",
            RangeRef::new(
                CellRef::new(sheet, Coord::from_excel(1, 1, true, true)),
                CellRef::new(sheet, Coord::from_excel(2, 1, true, true)),
            ),
            true,
            vec!["seats".into()],
            false,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            1,
            3,
            formualizer_parse::parser::parse("=SUM(Orders[seats])").unwrap(),
        )
        .unwrap();
    assert_eq!(
        engine.evaluate_cell("Sheet1", 1, 3).unwrap(),
        Some(LiteralValue::Number(2.0))
    );
    engine.resize_table("Orders", 4).unwrap();
    assert_eq!(
        engine.evaluate_cell("Sheet1", 1, 3).unwrap(),
        Some(LiteralValue::Number(9.0))
    );

    let (_, journal) = engine
        .action_atomic_journal("shrink", |tx| {
            tx.resize_table("Orders", 2)?;
            Ok(())
        })
        .unwrap();
    assert!(
        journal
            .graph
            .events
            .iter()
            .any(|event| matches!(event, ChangeEvent::TableResized { .. }))
    );
    assert_eq!(
        engine.evaluate_cell("Sheet1", 1, 3).unwrap(),
        Some(LiteralValue::Number(2.0))
    );
    let mut undo = UndoEngine::new();
    undo.push_action(journal);
    engine.undo_action(&mut undo).unwrap();
    assert_eq!(engine.table_metadata("Orders").unwrap().end_row, 4);
    assert_eq!(
        engine.evaluate_cell("Sheet1", 1, 3).unwrap(),
        Some(LiteralValue::Number(9.0))
    );
    engine.redo_action(&mut undo).unwrap();
    assert_eq!(engine.table_metadata("Orders").unwrap().end_row, 2);
    assert_eq!(
        engine.evaluate_cell("Sheet1", 1, 3).unwrap(),
        Some(LiteralValue::Number(2.0))
    );
    assert!(engine.resize_table("Orders", 0).is_err());
    assert!(engine.resize_table("Orders", 1).is_err());
    assert!(engine.resize_table("Orders", 1_048_577).is_err());
}

#[test]
fn structured_ref_this_row_column_rewrites_to_concrete_cell() {
    let ctx = crate::test_workbook::TestWorkbook::new();
    let mut engine: Engine<_> = Engine::new(ctx, EvalConfig::default());

    engine.add_sheet("Sheet1").unwrap();

    // Table region A1:C3 (header + 2 data rows)
    // Headers: Region, Amount, Double
    engine
        .set_cell_value("Sheet1", 2, 1, LiteralValue::Text("N".into()))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 2, 2, LiteralValue::Number(10.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 3, 1, LiteralValue::Text("S".into()))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 3, 2, LiteralValue::Number(20.0))
        .unwrap();

    let sheet_id = engine.sheet_id("Sheet1").unwrap();
    let start = CellRef::new(sheet_id, Coord::from_excel(1, 1, true, true));
    let end = CellRef::new(sheet_id, Coord::from_excel(3, 3, true, true));
    let range = RangeRef::new(start, end);
    engine
        .define_table(
            "Sales",
            range,
            true,
            vec!["Region".into(), "Amount".into(), "Double".into()],
            false,
        )
        .unwrap();

    // Formula inside the table: Double = [@Amount] * 2
    let f2 = formualizer_parse::parser::parse("=[@Amount]*2").unwrap();
    engine.set_cell_formula("Sheet1", 2, 3, f2).unwrap();
    let f3 = formualizer_parse::parser::parse("=[@[Amount]]*2").unwrap();
    engine.set_cell_formula("Sheet1", 3, 3, f3).unwrap();

    let v2 = engine
        .evaluate_cell("Sheet1", 2, 3)
        .unwrap()
        .expect("computed value");
    assert_eq!(v2, LiteralValue::Number(20.0));
    let v3 = engine
        .evaluate_cell("Sheet1", 3, 3)
        .unwrap()
        .expect("computed value");
    assert_eq!(v3, LiteralValue::Number(40.0));

    // Editing Amount should dirty and recompute the this-row dependents.
    engine
        .set_cell_value("Sheet1", 2, 2, LiteralValue::Number(100.0))
        .unwrap();
    let v2b = engine
        .evaluate_cell("Sheet1", 2, 3)
        .unwrap()
        .expect("computed value");
    assert_eq!(v2b, LiteralValue::Number(200.0));
}

#[test]
fn structured_ref_bracket_table_shorthand_selects_data_body() {
    let ctx = crate::test_workbook::TestWorkbook::new();
    let mut engine: Engine<_> = Engine::new(ctx, EvalConfig::default());

    engine.add_sheet("Sheet1").unwrap();

    // Table region A1:B3 (header + 2 data rows)
    // Headers: Region, Amount
    engine
        .set_cell_value("Sheet1", 1, 2, LiteralValue::Number(100.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 2, 1, LiteralValue::Number(5.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 2, 2, LiteralValue::Number(10.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 3, 1, LiteralValue::Number(7.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 3, 2, LiteralValue::Number(20.0))
        .unwrap();

    let sheet_id = engine.sheet_id("Sheet1").unwrap();
    let start = CellRef::new(sheet_id, Coord::from_excel(1, 1, true, true));
    let end = CellRef::new(sheet_id, Coord::from_excel(3, 2, true, true));
    let range = RangeRef::new(start, end);
    engine
        .define_table(
            "Sales",
            range,
            true,
            vec!["Region".into(), "Amount".into()],
            false,
        )
        .unwrap();

    let ast = formualizer_parse::parser::parse("=SUM([Sales])").unwrap();
    engine.set_cell_formula("Sheet1", 1, 4, ast).unwrap();

    // Header cell (B1=100) should not be included; only data body should contribute.
    let v = engine
        .evaluate_cell("Sheet1", 1, 4)
        .unwrap()
        .expect("computed value");
    assert_eq!(v, LiteralValue::Number(42.0));
}

#[test]
fn structured_ref_table_name_resolution_is_case_insensitive_by_default() {
    let ctx = crate::test_workbook::TestWorkbook::new();
    let mut engine: Engine<_> = Engine::new(ctx, EvalConfig::default());

    engine.add_sheet("Sheet1").unwrap();

    // Table region A1:B3 (header + 2 data rows)
    engine
        .set_cell_value("Sheet1", 2, 2, LiteralValue::Number(10.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 3, 2, LiteralValue::Number(20.0))
        .unwrap();

    let sheet_id = engine.sheet_id("Sheet1").unwrap();
    let start = CellRef::new(sheet_id, Coord::from_excel(1, 1, true, true));
    let end = CellRef::new(sheet_id, Coord::from_excel(3, 2, true, true));
    let range = RangeRef::new(start, end);
    engine
        .define_table(
            "Sales",
            range,
            true,
            vec!["Region".into(), "Amount".into()],
            false,
        )
        .unwrap();

    // Reference the table using different casing.
    let ast = formualizer_parse::parser::parse("=SUM(sales[Amount])").unwrap();
    engine.set_cell_formula("Sheet1", 1, 4, ast).unwrap();
    let v = engine
        .evaluate_cell("Sheet1", 1, 4)
        .unwrap()
        .expect("computed value");
    assert_eq!(v, LiteralValue::Number(30.0));
}

#[test]
fn structured_ref_unicode_table_and_column_resolution_is_case_insensitive_by_default() {
    let ctx = crate::test_workbook::TestWorkbook::new();
    let mut engine: Engine<_> = Engine::new(ctx, EvalConfig::default());

    engine.add_sheet("Sheet1").unwrap();

    engine
        .set_cell_value("Sheet1", 2, 2, LiteralValue::Number(10.0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", 3, 2, LiteralValue::Number(20.0))
        .unwrap();

    let sheet_id = engine.sheet_id("Sheet1").unwrap();
    let start = CellRef::new(sheet_id, Coord::from_excel(1, 1, true, true));
    let end = CellRef::new(sheet_id, Coord::from_excel(3, 2, true, true));
    let range = RangeRef::new(start, end);
    engine
        .define_table(
            "Продажи",
            range,
            true,
            vec!["Регион".into(), "Сумма".into()],
            false,
        )
        .unwrap();

    let ast = formualizer_parse::parser::parse("=SUM(продажи[сумма])").unwrap();
    engine.set_cell_formula("Sheet1", 1, 4, ast).unwrap();
    let v = engine
        .evaluate_cell("Sheet1", 1, 4)
        .unwrap()
        .expect("computed value");
    assert_eq!(v, LiteralValue::Number(30.0));
}

#[test]
fn table_definition_rejects_case_insensitive_collisions() {
    let ctx = crate::test_workbook::TestWorkbook::new();
    let mut engine: Engine<_> = Engine::new(ctx, EvalConfig::default());
    engine.add_sheet("Sheet1").unwrap();

    let sheet_id = engine.sheet_id("Sheet1").unwrap();
    let start = CellRef::new(sheet_id, Coord::from_excel(1, 1, true, true));
    let end = CellRef::new(sheet_id, Coord::from_excel(1, 1, true, true));
    let range = RangeRef::new(start, end);

    engine
        .define_table("Sales", range, true, vec!["A".into()], false)
        .unwrap();

    let err = engine
        .define_table("sales", range, true, vec!["A".into()], false)
        .expect_err("expected collision error");
    assert_eq!(err.kind, ExcelErrorKind::Name);
}

#[test]
fn empty_table_rectangles_obstruct_spills_before_dependents_are_published() {
    for (header, totals) in [(false, false), (true, false), (false, true)] {
        let mut engine = Engine::new(
            crate::test_workbook::TestWorkbook::new(),
            EvalConfig::default(),
        );
        let id = engine.add_sheet("Sheet1").unwrap();
        engine
            .define_table(
                "Table1",
                RangeRef::new(
                    CellRef::new(id, Coord::from_excel(2, 2, true, true)),
                    CellRef::new(id, Coord::from_excel(4, 2, true, true)),
                ),
                header,
                vec!["Qty".into()],
                totals,
            )
            .unwrap();
        engine
            .set_cell_formula(
                "Sheet1",
                1,
                2,
                formualizer_parse::parser::parse("=SEQUENCE(3)").unwrap(),
            )
            .unwrap();
        engine
            .set_cell_formula(
                "Sheet1",
                1,
                4,
                formualizer_parse::parser::parse("=SUM(B1)").unwrap(),
            )
            .unwrap();
        engine.evaluate_all().unwrap();
        for col in [2, 4] {
            assert!(
                matches!(engine.get_cell_value("Sheet1",1,col),Some(LiteralValue::Error(e)) if e.kind==ExcelErrorKind::Spill)
            );
        }
    }
}

/// `Orders` at A1:C5 on Sheet1: header row 1, data rows 2..=4 (Qty = 2, 3, 4
/// in A, Price = 10, 20, 30 in B), totals row 5. C2:C4 hold `formula`.
fn calculated_column(formula: &str) -> Engine<crate::test_workbook::TestWorkbook> {
    let mut engine = Engine::new(
        crate::test_workbook::TestWorkbook::new(),
        EvalConfig::default(),
    );
    engine.add_sheet("Sheet1").unwrap();
    for row in 2..=4u32 {
        engine
            .set_cell_value("Sheet1", row, 1, LiteralValue::Number(f64::from(row)))
            .unwrap();
        engine
            .set_cell_value(
                "Sheet1",
                row,
                2,
                LiteralValue::Number(f64::from(row - 1) * 10.0),
            )
            .unwrap();
    }
    let sheet = engine.sheet_id("Sheet1").unwrap();
    engine
        .define_table(
            "Orders",
            RangeRef::new(
                CellRef::new(sheet, Coord::from_excel(1, 1, true, true)),
                CellRef::new(sheet, Coord::from_excel(5, 3, true, true)),
            ),
            true,
            vec!["Qty".into(), "Price".into(), "Total".into()],
            true,
        )
        .unwrap();
    for row in 2..=4 {
        engine
            .set_cell_formula(
                "Sheet1",
                row,
                3,
                formualizer_parse::parser::parse(formula).unwrap(),
            )
            .unwrap();
    }
    engine.evaluate_all().unwrap();
    engine
}

fn column_c(
    engine: &Engine<crate::test_workbook::TestWorkbook>,
    rows: std::ops::RangeInclusive<u32>,
) -> Vec<Option<LiteralValue>> {
    rows.map(|r| engine.get_cell_value("Sheet1", r, 3))
        .collect()
}

#[test]
fn qualified_this_row_column_stays_symbolic_and_dirties_only_its_row() {
    let n = |v: f64| Some(LiteralValue::Number(v));
    for formula in [
        "=Orders[[#This Row],[Qty]]*Orders[[#This Row],[Price]]",
        "=Orders[@Qty]*Orders[@Price]",
        "=PRODUCT(Orders[@[Qty]:[Price]])",
    ] {
        let mut engine = calculated_column(formula);
        assert_eq!(
            column_c(&engine, 2..=4),
            vec![n(20.0), n(60.0), n(120.0)],
            "{formula}"
        );

        // The stored formula keeps its structured reference.
        let (ast, _) = engine.get_cell("Sheet1", 3, 3).unwrap();
        let text = formualizer_parse::pretty::canonical_formula(&ast.unwrap());
        assert!(text.contains("Orders["), "{formula} stored as {text}");

        // Editing Qty in row 3 recomputes only C3.
        engine
            .set_cell_value("Sheet1", 3, 1, LiteralValue::Number(7.0))
            .unwrap();
        let result = engine.evaluate_all().unwrap();
        assert_eq!(result.computed_vertices, 1, "{formula}");
        assert_eq!(
            column_c(&engine, 2..=4),
            vec![n(20.0), n(140.0), n(120.0)],
            "{formula}"
        );
    }
}

#[test]
fn this_row_formula_moved_by_a_row_insert_reads_its_new_row() {
    let n = |v: f64| Some(LiteralValue::Number(v));
    let mut engine = calculated_column("=Orders[@Qty]*Orders[@Price]");
    // Insert a sheet row before row 3: old row 3 (Qty 3, Price 20) moves to
    // row 4. (Table extents do not follow row inserts yet; row 4 is a data
    // row either way.)
    engine.insert_rows("Sheet1", 3, 1).unwrap();
    engine.evaluate_all().unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", 2, 3), n(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", 4, 3), n(60.0));
    engine
        .set_cell_value("Sheet1", 4, 2, LiteralValue::Number(1.0))
        .unwrap();
    engine.evaluate_all().unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", 4, 3), n(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", 2, 3), n(20.0));
}

#[test]
fn bare_table_name_is_the_data_body_and_tracks_edits() {
    let mut engine = calculated_column("=1");
    for (col, formula) in [
        (5, "=SUM(Orders)"),
        (6, "=ROWS(Orders)"),
        (7, "=COLUMNS(Orders)"),
    ] {
        engine
            .set_cell_formula(
                "Sheet1",
                1,
                col,
                formualizer_parse::parser::parse(formula).unwrap(),
            )
            .unwrap();
    }
    engine.evaluate_all().unwrap();
    // Qty 9 + Price 60 + Total 3 (three `=1` cells).
    assert_eq!(
        engine.get_cell_value("Sheet1", 1, 5),
        Some(LiteralValue::Number(72.0))
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", 1, 6),
        Some(LiteralValue::Number(3.0))
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", 1, 7),
        Some(LiteralValue::Number(3.0))
    );
    engine
        .set_cell_value("Sheet1", 2, 1, LiteralValue::Number(102.0))
        .unwrap();
    engine.evaluate_all().unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", 1, 5),
        Some(LiteralValue::Number(172.0))
    );
}

#[test]
fn this_row_outside_the_data_rows_evaluates_to_value_error() {
    let mut engine = calculated_column("=1");
    for (row, col) in [(1, 5), (5, 5), (5, 1), (9, 1)] {
        engine
            .set_cell_formula(
                "Sheet1",
                row,
                col,
                formualizer_parse::parser::parse("=Orders[@Qty]").unwrap(),
            )
            .unwrap();
    }
    engine.evaluate_all().unwrap();
    for (row, col) in [(1, 5), (5, 5), (5, 1), (9, 1)] {
        match engine.get_cell_value("Sheet1", row, col) {
            Some(LiteralValue::Error(e)) => {
                assert_eq!(e.kind, ExcelErrorKind::Value, "R{row}C{col}")
            }
            other => panic!("R{row}C{col}: {other:?}"),
        }
    }
}
