//! Excel storage form for packages written by `Workbook::to_xlsx_bytes`.
//!
//! Umya writes every formula as a plain `<f>` and has no API for array
//! formula references or cell metadata. A dynamic-array anchor saved that way
//! reloads as an ordinary formula whose saved spill children are constants
//! occupying its spill range, so it evaluates to `#SPILL!`. This module puts
//! the saved package into the form Excel itself writes:
//!
//! * each anchor becomes `<c … cm="1"><f t="array" ref="A1:A3">…</f>`, the
//!   `cm` selecting an XLDAPR (`fDynamic="1"`) cell-metadata record in
//!   `xl/metadata.xml`, so readers treat the cached children inside `ref` as
//!   generated spill output rather than as values blocking the spill;
//! * formula text uses the stored spellings: `_xlfn.`/`_xlfn._xlws.` function
//!   prefixes, `_xlfn.ANCHORARRAY(A1)` for `A1#` and `_xlfn.SINGLE(x)` for `@x`.
use quick_xml::{
    Reader, Writer,
    events::{BytesStart, Event},
};
use std::collections::{BTreeMap, HashMap};
use std::io::{self, Cursor, Read, Write};

/// A dynamic-array anchor and the `ref` to publish for it (1-based,
/// inclusive). A blocked or one-cell result publishes the anchor cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DynamicArrayAnchor {
    pub sheet: String,
    pub row: u32,
    pub col: u32,
    pub last_row: u32,
    pub last_col: u32,
}

const WORKBOOK: &str = "xl/workbook.xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES: &str = "[Content_Types].xml";
const METADATA_PART: &str = "xl/metadata.xml";
const SHEET_METADATA_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
const SHEET_METADATA_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
/// The one-based `c/@cm` of the XLDAPR record in [`METADATA_XML`].
const DYNAMIC_CM: &str = "1";
/// One XLDAPR metadata type, one `fDynamic=1 fCollapsed=0` future block and
/// one cell block selecting it: the record Excel writes for spilling formulas
/// (and the layout the strict XLSX reader's canonical part uses).
const METADATA_XML: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
    "xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">",
    "<metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" ",
    "minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" ",
    "merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" ",
    "clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/></metadataTypes>",
    "<futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst>",
    "<ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\">",
    "<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/>",
    "</ext></extLst></bk></futureMetadata>",
    "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata>",
    "</metadata>"
);

/// Functions Excel stores as `_xlfn._xlws.NAME`.
const XLWS_FUNCTIONS: &[&str] = &["FILTER", "SORT"];
/// Functions newer than the original OOXML function table, which Excel
/// stores as `_xlfn.NAME` (sorted, for binary search).
const XLFN_FUNCTIONS: &[&str] = &[
    "ACOT",
    "ACOTH",
    "AGGREGATE",
    "ANCHORARRAY",
    "ARABIC",
    "ARRAYTOTEXT",
    "BASE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.DIST.RANGE",
    "BINOM.INV",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "BYCOL",
    "BYROW",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "COMBINA",
    "CONCAT",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COT",
    "COTH",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "DROP",
    "ECMA.CEILING",
    "ENCODEURL",
    "ERF.PRECISE",
    "ERFC.PRECISE",
    "EXPAND",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FIELDVALUE",
    "FILTERXML",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "FORMULATEXT",
    "GAMMA",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "GAUSS",
    "GROUPBY",
    "HSTACK",
    "HYPGEOM.DIST",
    "IFNA",
    "IFS",
    "IMAGE",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMSEC",
    "IMSECH",
    "IMSINH",
    "IMTAN",
    "ISFORMULA",
    "ISO.CEILING",
    "ISOMITTED",
    "ISOWEEKNUM",
    "LAMBDA",
    "LET",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MAKEARRAY",
    "MAP",
    "MAXIFS",
    "MINIFS",
    "MODE.MULT",
    "MODE.SNGL",
    "MUNIT",
    "NEGBINOM.DIST",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "NUMBERVALUE",
    "PDURATION",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTOF",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "PERMUTATIONA",
    "PHI",
    "PIVOTBY",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "RANDARRAY",
    "RANK.AVG",
    "RANK.EQ",
    "REDUCE",
    "REGEXEXTRACT",
    "REGEXREPLACE",
    "REGEXTEST",
    "RRI",
    "SCAN",
    "SEC",
    "SECH",
    "SEQUENCE",
    "SHEET",
    "SHEETS",
    "SINGLE",
    "SKEW.P",
    "SORTBY",
    "STDEV.P",
    "STDEV.S",
    "STOCKHISTORY",
    "SWITCH",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTJOIN",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "TRIMRANGE",
    "UNICHAR",
    "UNICODE",
    "UNIQUE",
    "VALUETOTEXT",
    "VAR.P",
    "VAR.S",
    "VSTACK",
    "WEBSERVICE",
    "WEIBULL.DIST",
    "WORKDAY.INTL",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
    "XOR",
    "Z.TEST",
];

/// The stored prefix of a function name, if Excel needs one.
fn function_prefix(name: &str) -> Option<&'static str> {
    if name.starts_with('_') {
        // Already prefixed (`_xlfn.`, `_xlws.`, `_xll.`): keep as written.
        return None;
    }
    let upper = name.to_ascii_uppercase();
    if XLWS_FUNCTIONS.contains(&upper.as_str()) {
        Some("_xlfn._xlws.")
    } else if XLFN_FUNCTIONS.binary_search(&upper.as_str()).is_ok() {
        Some("_xlfn.")
    } else {
        None
    }
}

/// The formula text as Excel stores it (see the module docs). The leading
/// `=` is kept when present. Text the tokenizer cannot read is returned
/// unchanged: the engine accepted it, and a guessed rewrite could corrupt it.
pub(crate) fn storage_formula(formula: &str) -> String {
    use formualizer_parse::{TokenSubType, TokenType};
    let normalized;
    let (source, had_equals) = match formula.strip_prefix('=') {
        Some(_) => (formula, true),
        None => {
            normalized = format!("={formula}");
            (normalized.as_str(), false)
        }
    };
    let Ok(tokens) = formualizer_parse::TokenStream::new(source) else {
        return formula.to_string();
    };
    let views: Vec<_> = (0..tokens.len()).filter_map(|i| tokens.get(i)).collect();
    let is_spill_postfix = |i: usize| {
        views
            .get(i)
            .is_some_and(|t| t.span.token_type == TokenType::OpPostfix && t.value == "#")
    };

    // Output pieces; `unit_start[i]` is the first piece of the operand that
    // piece `i` completes (itself, or the matching opener of a close).
    let mut pieces: Vec<String> = Vec::with_capacity(views.len() + 1);
    let mut unit_start: Vec<usize> = Vec::with_capacity(views.len() + 1);
    let mut openers: Vec<usize> = Vec::new();
    // Depths at which an `@` waits for its operand to complete.
    let mut pending_single: Vec<usize> = Vec::new();
    let mut cursor = 0usize;
    for (i, token) in views.iter().enumerate() {
        let (start, end) = (token.span.start, token.span.end);
        if start > cursor {
            // Untokenized text (the leading `=`): keep verbatim.
            pieces.push(source[cursor..start].to_string());
            unit_start.push(pieces.len() - 1);
        }
        cursor = end.max(cursor);
        let text = &source[start..end];
        let index = pieces.len();
        let mut completes_unit = false;
        match (token.span.token_type, token.span.subtype) {
            (TokenType::Func, TokenSubType::Open) => {
                let name = text.strip_suffix('(').unwrap_or(text);
                let piece = match function_prefix(name) {
                    Some(prefix) => format!("{prefix}{text}"),
                    None => text.to_string(),
                };
                pieces.push(piece);
                unit_start.push(index);
                openers.push(index);
            }
            (TokenType::Paren, TokenSubType::Open) | (TokenType::Array, TokenSubType::Open) => {
                pieces.push(text.to_string());
                unit_start.push(index);
                openers.push(index);
            }
            (TokenType::Func | TokenType::Paren | TokenType::Array, TokenSubType::Close) => {
                let opener = openers.pop().unwrap_or(index);
                pieces.push(text.to_string());
                unit_start.push(opener);
                completes_unit = true;
            }
            (TokenType::Operand, _) => {
                pieces.push(text.to_string());
                unit_start.push(index);
                completes_unit = true;
            }
            (TokenType::OpPostfix, _) if text == "#" => {
                // Wrap the operand just completed: `A1#` → ANCHORARRAY(A1).
                let start = unit_start.last().copied().unwrap_or(index);
                pieces.insert(start, "_xlfn.ANCHORARRAY(".to_string());
                unit_start.insert(start, start);
                pieces.push(")".to_string());
                unit_start.push(start);
                completes_unit = true;
            }
            (TokenType::OpPrefix, _) if text == "@" => {
                pieces.push("_xlfn.SINGLE(".to_string());
                unit_start.push(index);
                pending_single.push(openers.len());
            }
            _ => {
                pieces.push(text.to_string());
                unit_start.push(index);
            }
        }
        // Close each `@` whose operand is now complete (after any `#`).
        if completes_unit && !is_spill_postfix(i + 1) {
            while pending_single.last() == Some(&openers.len()) {
                pending_single.pop();
                pieces.push(")".to_string());
                unit_start.push(pieces.len() - 1);
            }
        }
    }
    if cursor < source.len() {
        pieces.push(source[cursor..].to_string());
    }
    if !pending_single.is_empty() {
        // An `@` without an operand: leave the text alone.
        return formula.to_string();
    }
    let rewritten = pieces.concat();
    if had_equals {
        rewritten
    } else {
        rewritten
            .strip_prefix('=')
            .map(str::to_string)
            .unwrap_or(rewritten)
    }
}

/// A1-style cell reference (1-based).
fn a1(row: u32, col: u32) -> String {
    format!(
        "{}{row}",
        formualizer_common::coord::col_letters_from_1based(col).unwrap_or_default()
    )
}

fn invalid(message: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, message.into())
}

fn attribute(start: &BytesStart<'_>, key: &[u8]) -> io::Result<Option<String>> {
    use crate::xml_compat::AttributeExt as _;
    for a in start.attributes() {
        let a = a.map_err(io::Error::other)?;
        if a.key.as_ref() == key {
            return a
                .decoded_without_normalization(Reader::from_str("").decoder())
                .map(|v| Some(v.into_owned()))
                .map_err(io::Error::other);
        }
    }
    Ok(None)
}

/// `start` with `key="value"` set (replacing an existing `key`).
fn with_attributes(
    start: &BytesStart<'_>,
    set: &[(&str, &str)],
) -> io::Result<BytesStart<'static>> {
    let mut out = BytesStart::new(String::from_utf8_lossy(start.name().as_ref()).into_owned());
    for a in start.attributes() {
        let a = a.map_err(io::Error::other)?;
        if !set.iter().any(|(k, _)| a.key.as_ref() == k.as_bytes()) {
            out.push_attribute(a);
        }
    }
    for (k, v) in set {
        out.push_attribute((*k, *v));
    }
    Ok(out)
}

/// Resolve each worksheet name to its package part through the workbook
/// relationships.
fn sheet_parts(workbook: &[u8], rels: &[u8]) -> io::Result<HashMap<String, String>> {
    let mut targets = HashMap::new();
    let mut reader = Reader::from_reader(rels);
    loop {
        match reader.read_event().map_err(io::Error::other)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"Relationship" => {
                if let (Some(id), Some(target)) = (attribute(&e, b"Id")?, attribute(&e, b"Target")?)
                {
                    let part = match target.strip_prefix('/') {
                        Some(absolute) => absolute.to_string(),
                        None => format!("xl/{target}"),
                    };
                    targets.insert(id, part);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let mut parts = HashMap::new();
    let mut reader = Reader::from_reader(workbook);
    loop {
        match reader.read_event().map_err(io::Error::other)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheet" => {
                let name = attribute(&e, b"name")?;
                let id = attribute(&e, b"r:id")?;
                if let (Some(name), Some(part)) = (name, id.and_then(|id| targets.get(&id))) {
                    parts.insert(name, part.clone());
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(parts)
}

/// Mark the anchors of one worksheet: `cm` on the cell, `t="array"` and
/// `ref` on its formula. Every anchor must be found with a formula.
fn mark_sheet(xml: &[u8], refs: &BTreeMap<String, String>) -> io::Result<Vec<u8>> {
    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + refs.len() * 48));
    // The `ref` of the anchor cell being copied, until its formula is marked.
    let mut current: Option<&String> = None;
    let mut marked = 0usize;
    loop {
        let event = reader.read_event().map_err(io::Error::other)?;
        match event {
            Event::Start(e) if e.local_name().as_ref() == b"c" => {
                current = match attribute(&e, b"r")? {
                    Some(r) => refs.get(&r),
                    None => None,
                };
                if current.is_some() {
                    let e = with_attributes(&e, &[("cm", DYNAMIC_CM)])?;
                    writer.write_event(Event::Start(e))?;
                } else {
                    writer.write_event(Event::Start(e))?;
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"f" && current.is_some() => {
                let reference = current.take().expect("checked");
                let e = with_attributes(&e, &[("t", "array"), ("ref", reference)])?;
                writer.write_event(Event::Start(e))?;
                marked += 1;
            }
            Event::End(e) if e.local_name().as_ref() == b"c" => {
                current = None;
                writer.write_event(Event::End(e))?;
            }
            Event::Eof => break,
            other => writer.write_event(other)?,
        }
    }
    if marked != refs.len() {
        return Err(invalid(format!(
            "dynamic array export marked {marked} of {} anchors",
            refs.len()
        )));
    }
    Ok(writer.into_inner())
}

/// Insert `fragment` before the closing `</{root}>` tag of `xml`.
fn insert_before_close(xml: &str, root: &str, fragment: &str) -> io::Result<String> {
    let close = format!("</{root}>");
    let at = xml
        .rfind(&close)
        .ok_or_else(|| invalid(format!("missing {close}")))?;
    let mut out = String::with_capacity(xml.len() + fragment.len());
    out.push_str(&xml[..at]);
    out.push_str(fragment);
    out.push_str(&xml[at..]);
    Ok(out)
}

/// An unused `rIdN` relationship id.
fn fresh_relationship_id(rels: &str) -> String {
    let max = rels
        .split("Id=\"rId")
        .skip(1)
        .filter_map(|rest| {
            let digits: String = rest.chars().take_while(char::is_ascii_digit).collect();
            digits.parse::<u64>().ok()
        })
        .max()
        .unwrap_or(0);
    format!("rId{}", max + 1)
}

/// Rewrite a package saved by Umya so `anchors` are stored as dynamic-array
/// formulas: their worksheet cells are marked and `xl/metadata.xml` is added
/// with its workbook relationship and content-type override.
pub(crate) fn write_dynamic_arrays(
    bytes: &[u8],
    anchors: &[DynamicArrayAnchor],
) -> io::Result<Vec<u8>> {
    use zip::write::SimpleFileOptions;
    if anchors.is_empty() {
        return Ok(bytes.to_vec());
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).map_err(io::Error::other)?;
    let mut read = |name: &str| -> io::Result<Vec<u8>> {
        let mut entry = archive.by_name(name).map_err(io::Error::other)?;
        let mut out = Vec::new();
        entry.read_to_end(&mut out)?;
        Ok(out)
    };
    if archive_has(&mut read, METADATA_PART) {
        // Umya never writes one; refuse rather than clobber foreign metadata.
        return Err(invalid("package already has a sheet metadata part"));
    }
    let rels = String::from_utf8(read(WORKBOOK_RELS)?).map_err(io::Error::other)?;
    let parts = sheet_parts(&read(WORKBOOK)?, rels.as_bytes())?;

    let mut by_part: BTreeMap<String, BTreeMap<String, String>> = BTreeMap::new();
    for anchor in anchors {
        let part = parts
            .get(&anchor.sheet)
            .ok_or_else(|| invalid(format!("no worksheet part for '{}'", anchor.sheet)))?;
        let first = a1(anchor.row, anchor.col);
        let reference = if (anchor.row, anchor.col) == (anchor.last_row, anchor.last_col) {
            first.clone()
        } else {
            format!("{first}:{}", a1(anchor.last_row, anchor.last_col))
        };
        by_part
            .entry(part.clone())
            .or_default()
            .insert(first, reference);
    }
    let mut replaced: HashMap<String, Vec<u8>> = HashMap::new();
    for (part, refs) in &by_part {
        replaced.insert(part.clone(), mark_sheet(&read(part)?, refs)?);
    }
    let id = fresh_relationship_id(&rels);
    let rels = insert_before_close(
        &rels,
        "Relationships",
        &format!(
            "<Relationship Id=\"{id}\" Type=\"{SHEET_METADATA_RELATIONSHIP}\" Target=\"metadata.xml\"/>"
        ),
    )?;
    replaced.insert(WORKBOOK_RELS.to_string(), rels.into_bytes());
    let types = String::from_utf8(read(CONTENT_TYPES)?).map_err(io::Error::other)?;
    let types = insert_before_close(
        &types,
        "Types",
        &format!(
            "<Override PartName=\"/{METADATA_PART}\" ContentType=\"{SHEET_METADATA_CONTENT_TYPE}\"/>"
        ),
    )?;
    replaced.insert(CONTENT_TYPES.to_string(), types.into_bytes());

    let mut out = Vec::with_capacity(bytes.len() + 2048);
    {
        let mut writer = zip::ZipWriter::new(Cursor::new(&mut out));
        for i in 0..archive.len() {
            let entry = archive.by_index(i).map_err(io::Error::other)?;
            let name = entry.name().map_err(io::Error::other)?.into_owned();
            match replaced.remove(&name) {
                Some(data) => {
                    let options = SimpleFileOptions::default()
                        .compression_method(entry.compression())
                        .last_modified_time(entry.last_modified().unwrap_or_default());
                    drop(entry);
                    writer.start_file(name, options).map_err(io::Error::other)?;
                    writer.write_all(&data)?;
                }
                None => writer.raw_copy_file(entry).map_err(io::Error::other)?,
            }
        }
        writer
            .start_file(
                METADATA_PART,
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated),
            )
            .map_err(io::Error::other)?;
        writer.write_all(METADATA_XML.as_bytes())?;
        writer.finish().map_err(io::Error::other)?;
    }
    if !replaced.is_empty() {
        return Err(invalid("dynamic array export lost a rewritten part"));
    }
    Ok(out)
}

fn archive_has(read: &mut impl FnMut(&str) -> io::Result<Vec<u8>>, name: &str) -> bool {
    read(name).is_ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn function_table_is_sorted() {
        assert!(XLFN_FUNCTIONS.windows(2).all(|w| w[0] < w[1]));
    }

    #[test]
    fn storage_formula_uses_excel_spellings() {
        for (input, expected) in [
            ("=SEQUENCE(3)", "=_xlfn.SEQUENCE(3)"),
            ("=sequence(2, 3)", "=_xlfn.sequence(2, 3)"),
            (
                "=FILTER(A1:A3, A1:A3 > 1)",
                "=_xlfn._xlws.FILTER(A1:A3, A1:A3 > 1)",
            ),
            ("=SORT(UNIQUE(A:A))", "=_xlfn._xlws.SORT(_xlfn.UNIQUE(A:A))"),
            ("=_xlfn.XLOOKUP(1, A:A, B:B)", "=_xlfn.XLOOKUP(1, A:A, B:B)"),
            ("=SUM(A1:A3) & \"SEQUENCE(\"", "=SUM(A1:A3) & \"SEQUENCE(\""),
            ("=SUM(A1#)", "=SUM(_xlfn.ANCHORARRAY(A1))"),
            ("='My S'!B2# * 2", "=_xlfn.ANCHORARRAY('My S'!B2) * 2"),
            ("=@A1:A3", "=_xlfn.SINGLE(A1:A3)"),
            ("=@A1#", "=_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))"),
            ("=@INDEX(A:A, 2) + 1", "=_xlfn.SINGLE(INDEX(A:A, 2)) + 1"),
            ("=Table1[@Col] + 1", "=Table1[@Col] + 1"),
            ("=A1 * 10%", "=A1 * 10%"),
            ("SEQUENCE(3)", "_xlfn.SEQUENCE(3)"),
        ] {
            assert_eq!(storage_formula(input), expected, "{input}");
        }
    }

    #[test]
    fn sheet_marking_adds_array_reference_and_metadata() {
        let xml = br#"<worksheet><sheetData><row r="1"><c r="A1" s="2"><f>_xlfn.SEQUENCE(3)</f><v>1</v></c><c r="B1"><f>A1</f><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>"#;
        let refs = BTreeMap::from([("A1".to_string(), "A1:A3".to_string())]);
        let out = String::from_utf8(mark_sheet(xml, &refs).unwrap()).unwrap();
        assert!(
            out.contains(r#"<c r="A1" s="2" cm="1"><f t="array" ref="A1:A3">_xlfn.SEQUENCE(3)</f><v>1</v></c>"#),
            "{out}"
        );
        assert!(out.contains(r#"<c r="B1"><f>A1</f>"#), "{out}");
        assert!(out.contains(r#"<c r="A2"><v>2</v></c>"#), "{out}");

        let missing = BTreeMap::from([("C9".to_string(), "C9".to_string())]);
        assert!(mark_sheet(xml, &missing).is_err());
    }

    #[test]
    fn relationship_ids_do_not_collide() {
        assert_eq!(
            fresh_relationship_id(r#"<Relationship Id="rId3"/><Relationship Id="rId12"/>"#),
            "rId13"
        );
        assert_eq!(fresh_relationship_id("<Relationships/>"), "rId1");
    }
}
