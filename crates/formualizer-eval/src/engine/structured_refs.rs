//! Geometry of structured (table) references.
//!
//! One place decides which cells a table specifier selects, so the ingest
//! lowering of this-row references and the evaluator's resolution of named
//! table references agree. Coordinates are 0-based and inclusive.
//!
//! Excel's rules, which this module follows:
//! * the row part is one of: nothing (the data body), `#All`, `#Data`,
//!   `#Headers`, `#Totals`, `#Headers`+`#Data`, `#Data`+`#Totals`, or
//!   `#This Row` alone; any other mix is `#REF!`;
//! * at most one column or column span (`[A]:[B]`) may follow; no column
//!   selects every column of the table;
//! * `#This Row` selects the formula cell's own row and is `#VALUE!` when that
//!   row is not a data row of the table (header, totals, outside the table, or
//!   another sheet).

use crate::SheetId;
use crate::reference::CellRef;
use formualizer_common::{ExcelError, ExcelErrorKind};
use formualizer_parse::parser::{ReferenceType, SpecialItem, TableRowSpecifier, TableSpecifier};

/// The placement and shape of a table, independent of how it is stored.
#[derive(Clone, Copy, Debug)]
pub(crate) struct TableGeometry<'a> {
    pub(crate) sheet_id: SheetId,
    pub(crate) start_row: u32,
    pub(crate) start_col: u32,
    pub(crate) end_row: u32,
    pub(crate) end_col: u32,
    pub(crate) header_row: bool,
    pub(crate) totals_row: bool,
    /// A table whose only body row is the blank placeholder Excel keeps for an
    /// empty table: it has no data rows.
    pub(crate) empty_body: bool,
    pub(crate) headers: &'a [String],
}

impl TableGeometry<'_> {
    /// First and last data rows, or `None` when the table has no data rows.
    pub(crate) fn data_rows(&self) -> Option<(u32, u32)> {
        if self.empty_body {
            return None;
        }
        let first = self.start_row + u32::from(self.header_row);
        let last = self.end_row.checked_sub(u32::from(self.totals_row))?;
        (first <= last).then_some((first, last))
    }

    fn col_index(&self, name: &str) -> Option<u32> {
        let key = name.trim().to_lowercase();
        self.headers
            .iter()
            .position(|h| h.to_lowercase() == key)
            .and_then(|i| u32::try_from(i).ok())
    }
}

/// The cells a structured reference selects.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum TableArea {
    Rect {
        start_row: u32,
        start_col: u32,
        end_row: u32,
        end_col: u32,
    },
    /// A section the table does not have (e.g. `#Totals` without a totals
    /// row, or the data body of an empty table).
    Empty,
}

/// Whether a specifier selects the formula cell's own row.
pub(crate) fn specifier_has_this_row(spec: &TableSpecifier) -> bool {
    match spec {
        TableSpecifier::SpecialItem(SpecialItem::ThisRow)
        | TableSpecifier::Row(TableRowSpecifier::Current) => true,
        TableSpecifier::Combination(parts) => parts.iter().any(|p| specifier_has_this_row(p)),
        _ => false,
    }
}

fn ref_error(message: impl Into<String>) -> ExcelError {
    ExcelError::new(ExcelErrorKind::Ref).with_message(message.into())
}

fn value_error(message: impl Into<String>) -> ExcelError {
    ExcelError::new(ExcelErrorKind::Value).with_message(message.into())
}

#[derive(Default)]
struct Parts {
    all: bool,
    headers: bool,
    data: bool,
    totals: bool,
    this_row: bool,
    columns: Option<(u32, u32)>,
}

fn set_flag(flag: &mut bool, spec: &TableSpecifier) -> Result<(), ExcelError> {
    if std::mem::replace(flag, true) {
        return Err(ref_error(format!(
            "Structured reference repeats the item {spec}"
        )));
    }
    Ok(())
}

fn collect_parts(
    geom: &TableGeometry<'_>,
    spec: &TableSpecifier,
    parts: &mut Parts,
    nested: bool,
) -> Result<(), ExcelError> {
    match spec {
        TableSpecifier::All
        | TableSpecifier::SpecialItem(SpecialItem::All)
        | TableSpecifier::Row(TableRowSpecifier::All) => set_flag(&mut parts.all, spec),
        TableSpecifier::Data
        | TableSpecifier::SpecialItem(SpecialItem::Data)
        | TableSpecifier::Row(TableRowSpecifier::Data) => set_flag(&mut parts.data, spec),
        TableSpecifier::Headers
        | TableSpecifier::SpecialItem(SpecialItem::Headers)
        | TableSpecifier::Row(TableRowSpecifier::Headers) => set_flag(&mut parts.headers, spec),
        TableSpecifier::Totals
        | TableSpecifier::SpecialItem(SpecialItem::Totals)
        | TableSpecifier::Row(TableRowSpecifier::Totals) => set_flag(&mut parts.totals, spec),
        TableSpecifier::SpecialItem(SpecialItem::ThisRow)
        | TableSpecifier::Row(TableRowSpecifier::Current) => set_flag(&mut parts.this_row, spec),
        TableSpecifier::Row(TableRowSpecifier::Index(_)) => {
            Err(ExcelError::new(ExcelErrorKind::NImpl)
                .with_message("Row-index structured references are not supported".to_string()))
        }
        TableSpecifier::Column(name) => {
            let i = geom
                .col_index(name)
                .ok_or_else(|| ref_error(format!("Unknown table column: {name}")))?;
            set_columns(parts, (i, i))
        }
        TableSpecifier::ColumnRange(first, last) => {
            let a = geom
                .col_index(first)
                .ok_or_else(|| ref_error(format!("Unknown table column: {first}")))?;
            let b = geom
                .col_index(last)
                .ok_or_else(|| ref_error(format!("Unknown table column: {last}")))?;
            set_columns(parts, (a.min(b), a.max(b)))
        }
        TableSpecifier::Combination(items) => {
            if nested {
                return Err(ref_error("Nested structured-reference combination"));
            }
            for item in items {
                collect_parts(geom, item, parts, true)?;
            }
            Ok(())
        }
    }
}

fn set_columns(parts: &mut Parts, cols: (u32, u32)) -> Result<(), ExcelError> {
    if parts.columns.replace(cols).is_some() {
        return Err(ref_error(
            "Structured reference selects more than one column or column span",
        ));
    }
    Ok(())
}

/// The rows a structured reference selects, given an optional formula cell
/// for `#This Row`.
fn select_rows(
    geom: &TableGeometry<'_>,
    parts: &Parts,
    current: Option<CellRef>,
) -> Result<Option<(u32, u32)>, ExcelError> {
    let row_items = [
        parts.all,
        parts.headers,
        parts.data,
        parts.totals,
        parts.this_row,
    ]
    .iter()
    .filter(|b| **b)
    .count();

    if parts.this_row {
        if row_items > 1 {
            return Err(ref_error(
                "#This Row cannot be combined with another row item",
            ));
        }
        let Some(cell) = current else {
            return Err(value_error(
                "#This Row needs the formula cell's row, which is not known here",
            ));
        };
        let row = cell.coord.row();
        return match geom.data_rows() {
            Some((first, last))
                if cell.sheet_id == geom.sheet_id && first <= row && row <= last =>
            {
                Ok(Some((row, row)))
            }
            _ => Err(value_error("#This Row used outside the table's data rows")),
        };
    }
    if parts.all {
        if row_items > 1 {
            return Err(ref_error("#All cannot be combined with another row item"));
        }
        return Ok(Some((geom.start_row, geom.end_row)));
    }
    if parts.headers && parts.totals {
        return Err(ref_error("#Headers and #Totals cannot be combined"));
    }

    let headers = (parts.headers && geom.header_row).then_some((geom.start_row, geom.start_row));
    let totals = (parts.totals && geom.totals_row).then_some((geom.end_row, geom.end_row));
    // No row item means the data body.
    let data = if parts.data || row_items == 0 {
        geom.data_rows()
    } else {
        None
    };
    Ok([headers, data, totals]
        .into_iter()
        .flatten()
        .reduce(|(a0, a1), (b0, b1)| (a0.min(b0), a1.max(b1))))
}

/// The cells `spec` selects in the table. `current` is the formula cell, used
/// only by `#This Row`.
pub(crate) fn select_table_area(
    geom: &TableGeometry<'_>,
    spec: &TableSpecifier,
    current: Option<CellRef>,
) -> Result<TableArea, ExcelError> {
    let mut parts = Parts::default();
    collect_parts(geom, spec, &mut parts, false)?;
    let Some((start_row, end_row)) = select_rows(geom, &parts, current)? else {
        return Ok(TableArea::Empty);
    };
    let (start_col, end_col) = match parts.columns {
        Some((a, b)) => (geom.start_col + a, geom.start_col + b),
        None => (geom.start_col, geom.end_col),
    };
    Ok(TableArea::Rect {
        start_row,
        start_col,
        end_row,
        end_col,
    })
}

/// Lower a this-row structured reference written in `cell` to the absolute
/// A1 reference it denotes there. The result has no sheet qualifier: a
/// this-row reference only resolves on its own sheet, which is the formula's.
///
/// `Err` is the Excel error the reference evaluates to in that cell.
pub(crate) fn lower_this_row_reference(
    geom: &TableGeometry<'_>,
    spec: &TableSpecifier,
    cell: CellRef,
) -> Result<ReferenceType, ExcelError> {
    match select_table_area(geom, spec, Some(cell))? {
        TableArea::Rect {
            start_row,
            start_col,
            end_row,
            end_col,
        } if start_row == end_row && start_col == end_col => Ok(ReferenceType::Cell {
            sheet: None,
            row: start_row + 1,
            col: start_col + 1,
            row_abs: true,
            col_abs: true,
        }),
        TableArea::Rect {
            start_row,
            start_col,
            end_row,
            end_col,
        } => Ok(ReferenceType::Range {
            sheet: None,
            start_row: Some(start_row + 1),
            start_col: Some(start_col + 1),
            end_row: Some(end_row + 1),
            end_col: Some(end_col + 1),
            start_row_abs: true,
            start_col_abs: true,
            end_row_abs: true,
            end_col_abs: true,
        }),
        TableArea::Empty => Err(value_error("#This Row used outside the table's data rows")),
    }
}

/// Whether `ast` contains a qualified this-row reference, which dependency
/// planning reads from a lowered copy of the formula.
pub(crate) fn ast_has_named_this_row(ast: &formualizer_parse::parser::ASTNode) -> bool {
    use formualizer_parse::parser::ASTNodeType;
    match &ast.node_type {
        ASTNodeType::Reference {
            reference: ReferenceType::Table(t),
            ..
        } => !t.name.is_empty() && t.specifier.as_ref().is_some_and(specifier_has_this_row),
        ASTNodeType::Reference { .. } | ASTNodeType::Literal(_) | ASTNodeType::Omitted => false,
        ASTNodeType::UnaryOp { expr, .. } => ast_has_named_this_row(expr),
        ASTNodeType::BinaryOp { left, right, .. } => {
            ast_has_named_this_row(left) || ast_has_named_this_row(right)
        }
        ASTNodeType::Function { args, .. } => args.iter().any(ast_has_named_this_row),
        ASTNodeType::Call { callee, args } => {
            ast_has_named_this_row(callee) || args.iter().any(ast_has_named_this_row)
        }
        ASTNodeType::Array(rows) => rows.iter().flatten().any(ast_has_named_this_row),
    }
}

/// A table definition that can describe its geometry.
pub(crate) trait TableShape {
    fn table_name(&self) -> &str;
    fn geometry(&self) -> TableGeometry<'_>;
}

impl<T: TableShape + ?Sized> TableShape for &T {
    fn table_name(&self) -> &str {
        (**self).table_name()
    }
    fn geometry(&self) -> TableGeometry<'_> {
        (**self).geometry()
    }
}

/// What [`rewrite_table_reference`] did to a reference.
#[derive(Debug)]
pub(crate) enum StructuredRewrite {
    Unchanged,
    /// The reference was replaced in place.
    Rewritten,
    /// The reference denotes this error in the formula cell; the caller
    /// replaces the reference node with the error literal.
    Error(ExcelError),
}

/// Bind a structured reference to the formula cell at ingest.
///
/// * An unqualified this-row reference (`[@Qty]`) becomes the absolute A1
///   cell or row span it denotes in `cell` (`#VALUE!` outside the table's
///   data rows).
/// * A qualified this-row reference (`Table1[@Qty]`,
///   `Table1[[#This Row],[Qty]]`) is lowered the same way only when
///   `lower_named` is set. The stored formula keeps it symbolic (it is resolved
///   against the formula cell at evaluation); the lowered copy is what
///   dependency planning reads, so a calculated column depends on its own
///   row's cells rather than on the whole table (which contains it).
/// * An unqualified reference without `#This Row` (`[[#Totals],[Qty]]`) is
///   bound to the table that contains `cell`.
///
/// `Err` is an ingest failure: an unqualified reference outside any table.
pub(crate) fn rewrite_table_reference<T: TableShape>(
    reference: &mut ReferenceType,
    cell: CellRef,
    lower_named: bool,
    named: impl FnOnce(&str) -> Option<T>,
    containing: impl FnOnce(CellRef) -> Option<T>,
) -> Result<StructuredRewrite, ExcelError> {
    let ReferenceType::Table(tref) = reference else {
        return Ok(StructuredRewrite::Unchanged);
    };
    let Some(spec) = tref.specifier.as_ref() else {
        return Ok(StructuredRewrite::Unchanged);
    };
    let this_row = specifier_has_this_row(spec);
    let table = if tref.name.is_empty() {
        let Some(table) = containing(cell) else {
            return Err(ExcelError::new(ExcelErrorKind::Name).with_message(
                "Unqualified structured reference used outside a table".to_string(),
            ));
        };
        if !this_row {
            tref.name = table.table_name().to_string();
            return Ok(StructuredRewrite::Rewritten);
        }
        table
    } else {
        if !this_row || !lower_named {
            return Ok(StructuredRewrite::Unchanged);
        }
        // An unknown table is left for dependency planning to report.
        let Some(table) = named(&tref.name) else {
            return Ok(StructuredRewrite::Unchanged);
        };
        table
    };
    match lower_this_row_reference(&table.geometry(), spec, cell) {
        Ok(lowered) => {
            *reference = lowered;
            Ok(StructuredRewrite::Rewritten)
        }
        Err(error) => Ok(StructuredRewrite::Error(error)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::reference::Coord;
    use formualizer_parse::parser::{ASTNodeType, parse};

    fn headers() -> Vec<String> {
        ["Item", "Qty", "Unit Price", "Total"]
            .iter()
            .map(|s| s.to_string())
            .collect()
    }

    /// Table at B2:E7 (0-based rows 1..=6, cols 1..=4): header row 1, data
    /// rows 2..=5, totals row 6.
    fn geom(h: &[String]) -> TableGeometry<'_> {
        TableGeometry {
            sheet_id: 0,
            start_row: 1,
            start_col: 1,
            end_row: 6,
            end_col: 4,
            header_row: true,
            totals_row: true,
            empty_body: false,
            headers: h,
        }
    }

    fn spec(formula: &str) -> TableSpecifier {
        let ast = parse(formula).unwrap();
        match ast.node_type {
            ASTNodeType::Reference {
                reference: ReferenceType::Table(t),
                ..
            } => t.specifier.unwrap(),
            other => panic!("{other:?}"),
        }
    }

    fn rect(start_row: u32, start_col: u32, end_row: u32, end_col: u32) -> TableArea {
        TableArea::Rect {
            start_row,
            start_col,
            end_row,
            end_col,
        }
    }

    fn cell(row0: u32, col0: u32) -> CellRef {
        CellRef::new(0, Coord::new(row0, col0, true, true))
    }

    #[test]
    fn sections_and_combinations_select_excel_areas() {
        let h = headers();
        let g = geom(&h);
        for (formula, expected) in [
            ("=T[Qty]", rect(2, 2, 5, 2)),
            ("=T[[Qty]:[Unit Price]]", rect(2, 2, 5, 3)),
            ("=T[[#Totals],[Total]]", rect(6, 4, 6, 4)),
            ("=T[[#Headers],[Qty]]", rect(1, 2, 1, 2)),
            ("=T[[#Headers],[#Data],[Qty]]", rect(1, 2, 5, 2)),
            ("=T[[#Data],[#Totals],[Qty]]", rect(2, 2, 6, 2)),
            ("=T[[#All],[Qty]:[Total]]", rect(1, 2, 6, 4)),
            ("=T[#Data]", rect(2, 1, 5, 4)),
            ("=T[[#Headers],[Qty]:[Unit Price]]", rect(1, 2, 1, 3)),
        ] {
            assert_eq!(
                select_table_area(&g, &spec(formula), None),
                Ok(expected),
                "{formula}"
            );
        }
        for formula in [
            "=T[[#Headers],[#Totals],[Qty]]",
            "=T[[#All],[#Data]]",
            "=T[[#This Row],[#Data],[Qty]]",
            "=T[[Qty],[Total]]",
            "=T[[#Totals],[Nope]]",
        ] {
            assert_eq!(
                select_table_area(&g, &spec(formula), Some(cell(3, 4))).map_err(|e| e.kind),
                Err(ExcelErrorKind::Ref),
                "{formula}"
            );
        }
    }

    #[test]
    fn missing_sections_are_empty() {
        let h = headers();
        let mut g = geom(&h);
        g.totals_row = false;
        g.header_row = false;
        assert_eq!(
            select_table_area(&g, &spec("=T[[#Totals],[Total]]"), None),
            Ok(TableArea::Empty)
        );
        assert_eq!(
            select_table_area(&g, &spec("=T[[#Headers],[#Data],[Qty]]"), None),
            Ok(rect(1, 2, 6, 2))
        );
        g.empty_body = true;
        assert_eq!(
            select_table_area(&g, &spec("=T[Qty]"), None),
            Ok(TableArea::Empty)
        );
    }

    #[test]
    fn this_row_lowers_to_the_formula_rows_cells() {
        let h = headers();
        let g = geom(&h);
        let at = cell(3, 4); // E4, a data row
        for formula in [
            "=T[[#This Row],[Qty]]",
            "=T[@Qty]",
            "=T[@[Qty]]",
            "=T[[@],[Qty]]",
            "=[@Qty]",
        ] {
            assert_eq!(
                lower_this_row_reference(&g, &spec(formula), at),
                Ok(ReferenceType::Cell {
                    sheet: None,
                    row: 4,
                    col: 3,
                    row_abs: true,
                    col_abs: true,
                }),
                "{formula}"
            );
        }
        let span = lower_this_row_reference(&g, &spec("=T[@[Qty]:[Unit Price]]"), at).unwrap();
        assert_eq!(span.to_string(), "$C$4:$D$4");
        let whole = lower_this_row_reference(&g, &spec("=T[@]"), at).unwrap();
        assert_eq!(whole.to_string(), "$B$4:$E$4");
    }

    #[test]
    fn this_row_outside_data_rows_is_value_error() {
        let h = headers();
        let g = geom(&h);
        let qty = spec("=T[@Qty]");
        for at in [
            cell(1, 4),                                    // header row
            cell(6, 4),                                    // totals row
            cell(10, 4),                                   // below the table
            cell(0, 0),                                    // above and left
            CellRef::new(1, Coord::new(3, 4, true, true)), // another sheet
        ] {
            assert_eq!(
                lower_this_row_reference(&g, &qty, at).map_err(|e| e.kind),
                Err(ExcelErrorKind::Value),
                "{at:?}"
            );
        }
        // A this-row reference beside the table, on a data row, is valid in
        // Excel (implicit row match), and selects that row.
        assert!(lower_this_row_reference(&g, &qty, cell(3, 8)).is_ok());
        assert_eq!(
            select_table_area(&g, &qty, None).map_err(|e| e.kind),
            Err(ExcelErrorKind::Value)
        );
    }
}
