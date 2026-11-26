//! `Workbook::to_xlsx_bytes` stores dynamic-array anchors as Excel does
//! (`cm` + `<f t="array" ref>` + XLDAPR metadata), so a saved workbook
//! reloads with the same results instead of `#SPILL!`.
#![cfg(all(feature = "umya", feature = "xlsx-recalc"))]
use formualizer_common::ExcelErrorKind;
use formualizer_workbook::{
    LiteralValue, LoadStrategy, NativeXlsxAdapter, SpreadsheetReader, Workbook, WorkbookConfig,
};
use std::io::Read;

fn load(bytes: Vec<u8>, backend: usize) -> Workbook {
    let config = WorkbookConfig::interactive();
    match backend {
        0 => Workbook::from_reader(
            NativeXlsxAdapter::open_bytes(bytes).unwrap(),
            LoadStrategy::EagerAll,
            config,
        )
        .unwrap(),
        1 => Workbook::from_reader(
            formualizer_workbook::UmyaAdapter::open_bytes(bytes).unwrap(),
            LoadStrategy::EagerAll,
            config,
        )
        .unwrap(),
        #[cfg(feature = "umya3")]
        2 => Workbook::from_reader(
            formualizer_workbook::Umya3Adapter::open_bytes(bytes).unwrap(),
            LoadStrategy::EagerAll,
            config,
        )
        .unwrap(),
        _ => unreachable!(),
    }
}

fn backends() -> Vec<usize> {
    let mut out = vec![0, 1];
    if cfg!(feature = "umya3") {
        out.push(2);
    }
    out
}

fn n(value: f64) -> Option<LiteralValue> {
    Some(LiteralValue::Number(value))
}

fn is_spill(value: Option<LiteralValue>) -> bool {
    matches!(value, Some(LiteralValue::Error(e)) if e.kind == ExcelErrorKind::Spill)
}

fn part(bytes: &[u8], name: &str) -> String {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut out = String::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_string(&mut out)
        .unwrap();
    out
}

/// One sheet `S`: inputs in column H, then `formulas` at (row, col).
fn workbook(formulas: &[(u32, u32, &str)]) -> Workbook {
    let mut wb = Workbook::new_with_config(WorkbookConfig::interactive());
    wb.add_sheet("S").unwrap();
    for (row, value) in [(1, 5.0), (2, 1.0), (3, 7.0), (4, 3.0), (5, 9.0)] {
        wb.set_value("S", row, 8, LiteralValue::Number(value))
            .unwrap();
    }
    for &(row, col, formula) in formulas {
        wb.set_formula("S", row, col, formula).unwrap();
    }
    wb.evaluate_all().unwrap();
    wb
}

/// Every cell of `S!A1:F8` (values only).
fn grid(wb: &Workbook) -> Vec<Option<LiteralValue>> {
    (1..=8)
        .flat_map(|r| (1..=6).map(move |c| wb.get_value("S", r, c)))
        .collect()
}

/// Save, reload through every reader, recalculate, and compare to the
/// original's results.
fn assert_roundtrip(wb: &Workbook) -> Vec<u8> {
    let expected = grid(wb);
    let bytes = wb.to_xlsx_bytes().unwrap();
    for backend in backends() {
        let mut loaded = load(bytes.clone(), backend);
        loaded.evaluate_all().unwrap();
        assert_eq!(grid(&loaded), expected, "backend {backend}");
    }
    bytes
}

#[test]
fn sequence_spill_reloads_without_spill_error() {
    let wb = workbook(&[(1, 1, "=SEQUENCE(3)")]);
    assert_eq!(wb.get_value("S", 3, 1), n(3.0));
    let bytes = assert_roundtrip(&wb);

    let sheet = part(&bytes, "xl/worksheets/sheet2.xml");
    assert!(
        sheet.contains(
            r#"<c r="A1" cm="1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(3)</f><v>1</v></c>"#
        ),
        "{sheet}"
    );
    // Children carry cached values only.
    assert!(sheet.contains(r#"<c r="A2"><v>2</v></c>"#), "{sheet}");
    assert!(part(&bytes, "xl/metadata.xml").contains("fDynamic=\"1\""));
    assert!(part(&bytes, "xl/_rels/workbook.xml.rels").contains("Target=\"metadata.xml\""));
    assert!(part(&bytes, "[Content_Types].xml").contains("PartName=\"/xl/metadata.xml\""));
}

#[test]
fn reloaded_spill_is_dynamic_and_resizes() {
    let wb = workbook(&[(1, 1, "=SEQUENCE(H4)")]);
    let bytes = wb.to_xlsx_bytes().unwrap();
    for backend in backends() {
        let mut loaded = load(bytes.clone(), backend);
        loaded.evaluate_all().unwrap();
        assert_eq!(loaded.get_value("S", 3, 1), n(3.0), "backend {backend}");
        loaded
            .set_value("S", 4, 8, LiteralValue::Number(5.0))
            .unwrap();
        loaded.evaluate_all().unwrap();
        assert_eq!(
            loaded.get_value("S", 5, 1),
            n(5.0),
            "grow, backend {backend}"
        );
        loaded
            .set_value("S", 4, 8, LiteralValue::Number(2.0))
            .unwrap();
        loaded.evaluate_all().unwrap();
        assert_eq!(loaded.get_value("S", 2, 1), n(2.0), "backend {backend}");
        assert_eq!(
            loaded.get_value("S", 3, 1),
            None,
            "shrink, backend {backend}"
        );
    }
}

#[test]
fn filter_spill_roundtrips() {
    let wb = workbook(&[(1, 1, "=FILTER(H1:H5, H1:H5 > 4)")]);
    assert_eq!(wb.get_value("S", 3, 1), n(9.0));
    let bytes = assert_roundtrip(&wb);
    let sheet = part(&bytes, "xl/worksheets/sheet2.xml");
    assert!(
        sheet.contains(r#"ref="A1:A3">_xlfn._xlws.FILTER(H1:H5, H1:H5 &gt; 4)</f>"#),
        "{sheet}"
    );
}

#[test]
fn two_dimensional_spill_roundtrips() {
    let wb = workbook(&[(2, 2, "=SEQUENCE(2, 3)")]);
    assert_eq!(wb.get_value("S", 3, 4), n(6.0));
    let bytes = assert_roundtrip(&wb);
    let sheet = part(&bytes, "xl/worksheets/sheet2.xml");
    assert!(sheet.contains(r#"<f t="array" ref="B2:D3">"#), "{sheet}");
}

#[test]
fn formula_reading_a_spill_child_roundtrips() {
    let wb = workbook(&[
        (1, 1, "=SEQUENCE(3)"),
        (1, 3, "=A2 * 10"),
        (2, 3, "=A3 + A1"),
    ]);
    assert_eq!(wb.get_value("S", 1, 3), n(20.0));
    assert_roundtrip(&wb);
}

#[test]
fn spill_operator_roundtrips_as_anchorarray() {
    let wb = workbook(&[
        (1, 1, "=SEQUENCE(H4)"),
        (1, 3, "=SUM(A1#)"),
        (2, 3, "=ROWS(A1#)"),
    ]);
    assert_eq!(wb.get_value("S", 1, 3), n(6.0));
    let bytes = assert_roundtrip(&wb);
    let sheet = part(&bytes, "xl/worksheets/sheet2.xml");
    assert!(
        sheet.contains("<f>SUM(_xlfn.ANCHORARRAY(A1))</f>"),
        "{sheet}"
    );

    // The reference follows the reloaded spill when it resizes.
    for backend in backends() {
        let mut loaded = load(bytes.clone(), backend);
        loaded
            .set_value("S", 4, 8, LiteralValue::Number(4.0))
            .unwrap();
        loaded.evaluate_all().unwrap();
        assert_eq!(loaded.get_value("S", 1, 3), n(10.0), "backend {backend}");
        assert_eq!(loaded.get_value("S", 2, 3), n(4.0), "backend {backend}");
    }
}

#[test]
fn blocked_spill_keeps_its_blocker_and_spills_once_cleared() {
    let mut wb = workbook(&[(1, 1, "=SEQUENCE(3)")]);
    wb.set_value("S", 3, 1, LiteralValue::Number(99.0)).unwrap();
    wb.evaluate_all().unwrap();
    assert!(is_spill(wb.get_value("S", 1, 1)));
    let bytes = assert_roundtrip(&wb);
    let sheet = part(&bytes, "xl/worksheets/sheet2.xml");
    // Excel's form for a blocked spill: the anchor cell alone.
    assert!(
        sheet.contains(r#" cm="1"><f t="array" ref="A1">"#),
        "{sheet}"
    );
    assert!(sheet.contains(r#"<c r="A3"><v>99</v></c>"#), "{sheet}");

    for backend in backends() {
        let mut loaded = load(bytes.clone(), backend);
        loaded.evaluate_all().unwrap();
        assert!(is_spill(loaded.get_value("S", 1, 1)), "backend {backend}");
        assert_eq!(loaded.get_value("S", 3, 1), n(99.0), "backend {backend}");
        loaded.set_value("S", 3, 1, LiteralValue::Empty).unwrap();
        loaded.evaluate_all().unwrap();
        assert_eq!(loaded.get_value("S", 3, 1), n(3.0), "backend {backend}");
    }
}

#[test]
fn ordinary_formulas_and_scalar_results_are_not_marked() {
    let wb = workbook(&[(1, 1, "=SUM(H1:H5)"), (2, 1, "=XLOOKUP(7, H1:H5, H1:H5)")]);
    let bytes = assert_roundtrip(&wb);
    let sheet = part(&bytes, "xl/worksheets/sheet2.xml");
    assert!(!sheet.contains("cm="), "{sheet}");
    assert!(
        sheet.contains("<f>_xlfn.XLOOKUP(7, H1:H5, H1:H5)</f>"),
        "{sheet}"
    );
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    assert!(archive.by_name("xl/metadata.xml").is_err());
}

#[test]
fn spills_on_several_sheets_share_one_metadata_record() {
    let mut wb = workbook(&[(1, 1, "=SEQUENCE(2)")]);
    wb.add_sheet("T").unwrap();
    wb.set_formula("T", 1, 1, "=SEQUENCE(1, 2)").unwrap();
    wb.set_formula("T", 2, 1, "=SUM(S!A1#)").unwrap();
    wb.evaluate_all().unwrap();
    let bytes = wb.to_xlsx_bytes().unwrap();
    for backend in backends() {
        let mut loaded = load(bytes.clone(), backend);
        loaded.evaluate_all().unwrap();
        assert_eq!(loaded.get_value("S", 2, 1), n(2.0), "backend {backend}");
        assert_eq!(loaded.get_value("T", 1, 2), n(2.0), "backend {backend}");
        assert_eq!(loaded.get_value("T", 2, 1), n(3.0), "backend {backend}");
    }
}

/// The repro from the bug report: save to a file and open it by path.
#[test]
fn saved_file_reopened_by_path_matches() {
    let wb = workbook(&[(1, 1, "=SEQUENCE(3)")]);
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("spill.xlsx");
    std::fs::write(&path, wb.to_xlsx_bytes().unwrap()).unwrap();
    let reader = NativeXlsxAdapter::open_path(&path).unwrap();
    let mut wb2 = Workbook::from_reader(
        reader,
        LoadStrategy::EagerAll,
        WorkbookConfig::interactive(),
    )
    .unwrap();
    wb2.evaluate_all().unwrap();
    for row in 1..=3 {
        assert_eq!(wb2.get_value("S", row, 1), n(row as f64));
    }
}

/// Saved output is accepted by the strict source-preserving recalculator.
#[test]
fn strict_recalculation_accepts_saved_spills() {
    let wb = workbook(&[(1, 1, "=SEQUENCE(H4)"), (1, 3, "=SUM(A1#)")]);
    let bytes = wb.to_xlsx_bytes().unwrap();
    let result = formualizer_workbook::recalculate_xlsx_bytes(&bytes, Default::default()).unwrap();
    let mut loaded = load(result.bytes, 0);
    loaded.evaluate_all().unwrap();
    assert_eq!(loaded.get_value("S", 3, 1), n(3.0));
    assert_eq!(loaded.get_value("S", 1, 3), n(6.0));
}
