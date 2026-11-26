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
            (TokenType::Operand, subtype) => {
                let piece = if subtype == TokenSubType::Range {
                    storage_structured_reference(text)
                } else {
                    text.to_string()
                };
                pieces.push(piece);
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

/// A structured reference as Excel stores it: the `@` this-row shorthand is
/// UI syntax only, so `T[@]`, `T[@Col]`, `T[@[Col]]` and the engine's
/// canonical `T[[@],[Col]]` become `T[#This Row]` / `T[[#This Row],[Col]]`.
/// A reference without `[@` is returned unchanged.
fn storage_structured_reference(text: &str) -> String {
    const THIS_ROW: &str = "#This Row";
    if !text.contains("[@") {
        return text.to_string();
    }
    let chars: Vec<char> = text.chars().collect();
    let mut out = String::with_capacity(text.len() + 16);
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '\'' {
            // `'` escapes the next character of a column name.
            out.push(c);
            if let Some(&next) = chars.get(i + 1) {
                out.push(next);
            }
            i += 2;
            continue;
        }
        if c == '[' && chars.get(i + 1) == Some(&'@') {
            match chars.get(i + 2) {
                // `[@]`
                Some(']') => {
                    out.push_str(&format!("[{THIS_ROW}]"));
                    i += 3;
                }
                // `[@[Col]…]`: the bracketed columns follow the item.
                Some('[') => {
                    out.push_str(&format!("[[{THIS_ROW}],"));
                    i += 2;
                }
                // `[@Col]`: a bare column name up to the unescaped `]`.
                _ => {
                    let mut j = i + 2;
                    let mut column = String::new();
                    while j < chars.len() && chars[j] != ']' {
                        if chars[j] == '\'' && j + 1 < chars.len() {
                            column.push(chars[j]);
                            j += 1;
                        }
                        column.push(chars[j]);
                        j += 1;
                    }
                    if j >= chars.len() {
                        return text.to_string();
                    }
                    out.push_str(&format!("[[{THIS_ROW}],[{column}]]"));
                    i = j + 1;
                }
            }
            continue;
        }
        out.push(c);
        i += 1;
    }
    out
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
    replaced.insert(METADATA_PART.to_string(), METADATA_XML.as_bytes().to_vec());
    rewrite_archive(&mut archive, replaced, bytes.len())
}

fn archive_has(read: &mut impl FnMut(&str) -> io::Result<Vec<u8>>, name: &str) -> bool {
    read(name).is_ok()
}

/// Copy `archive`, replacing the entries named in `parts` and appending the
/// remaining `parts` as new deflated entries (in name order).
fn rewrite_archive(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    mut parts: HashMap<String, Vec<u8>>,
    size_hint: usize,
) -> io::Result<Vec<u8>> {
    use zip::write::SimpleFileOptions;
    let mut out = Vec::with_capacity(size_hint + 4096);
    let mut writer = zip::ZipWriter::new(Cursor::new(&mut out));
    for i in 0..archive.len() {
        let entry = archive.by_index(i).map_err(io::Error::other)?;
        let name = entry.name().map_err(io::Error::other)?.into_owned();
        match parts.remove(&name) {
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
    let added: BTreeMap<String, Vec<u8>> = parts.into_iter().collect();
    for (name, data) in added {
        writer
            .start_file(
                name,
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated),
            )
            .map_err(io::Error::other)?;
        writer.write_all(&data)?;
    }
    writer.finish().map_err(io::Error::other)?;
    Ok(out)
}

/// A defined name to write into `xl/workbook.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DefinedNameExport {
    pub name: String,
    /// Zero-based index of the scope sheet in workbook order; `None` for
    /// workbook scope.
    pub local_sheet: Option<usize>,
    /// Formula text without a leading `=`, in storage form.
    pub formula: String,
}

/// A table (ListObject) to write as an `xl/tables/tableN.xml` part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct TableExport {
    pub sheet: String,
    pub name: String,
    /// Full extent (1-based, inclusive), including header and totals rows.
    pub first_row: u32,
    pub first_col: u32,
    pub last_row: u32,
    pub last_col: u32,
    pub header_row: bool,
    pub totals_row: bool,
    pub headers: Vec<String>,
}

const TABLE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const RELATIONSHIPS_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIPS_NS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
const SHEET_MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn escape_xml(text: &str) -> String {
    quick_xml::escape::escape(text).into_owned()
}

/// The `<definedNames>` children for `names`.
fn defined_names_xml(names: &[DefinedNameExport]) -> String {
    let mut out = String::new();
    for name in names {
        out.push_str("<definedName name=\"");
        out.push_str(&escape_xml(&name.name));
        out.push('"');
        if let Some(sheet) = name.local_sheet {
            out.push_str(&format!(" localSheetId=\"{sheet}\""));
        }
        out.push('>');
        out.push_str(&quick_xml::escape::partial_escape(&name.formula));
        out.push_str("</definedName>");
    }
    out
}

/// Add `names` to a workbook part: into an existing `<definedNames>`, or as
/// a new one directly after `</sheets>` (its place in `CT_Workbook`, since
/// Umya writes no `functionGroups` or `externalReferences`).
fn add_defined_names(workbook: &str, names: &[DefinedNameExport]) -> io::Result<String> {
    let entries = defined_names_xml(names);
    if workbook.contains("</definedNames>") {
        return insert_before_close(workbook, "definedNames", &entries);
    }
    let fragment = format!("<definedNames>{entries}</definedNames>");
    if let Some(at) = workbook.find("</sheets>") {
        let at = at + "</sheets>".len();
        return Ok(format!("{}{fragment}{}", &workbook[..at], &workbook[at..]));
    }
    Err(invalid("workbook part has no </sheets>"))
}

/// A1 range text for an inclusive 1-based rectangle.
fn a1_range(first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> String {
    format!("{}:{}", a1(first_row, first_col), a1(last_row, last_col))
}

/// The `xl/tables/tableN.xml` part for `table`.
fn table_xml(table: &TableExport, id: u32) -> String {
    let reference = a1_range(
        table.first_row,
        table.first_col,
        table.last_row,
        table.last_col,
    );
    let name = escape_xml(&table.name);
    let mut out = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    out.push_str(&format!(
        "<table xmlns=\"{SHEET_MAIN_NS}\" id=\"{id}\" name=\"{name}\" displayName=\"{name}\" ref=\"{reference}\""
    ));
    if !table.header_row {
        out.push_str(" headerRowCount=\"0\"");
    }
    if table.totals_row {
        // Umya's reader takes the totals row from `totalsRowShown`, so state
        // the spec default explicitly.
        out.push_str(" totalsRowCount=\"1\" totalsRowShown=\"1\"");
    } else {
        out.push_str(" totalsRowShown=\"0\"");
    }
    out.push('>');
    if table.header_row {
        // The filter covers the header and data rows, never the totals row.
        let last = table.last_row - u32::from(table.totals_row);
        out.push_str(&format!(
            "<autoFilter ref=\"{}\"/>",
            a1_range(table.first_row, table.first_col, last, table.last_col)
        ));
    }
    out.push_str(&format!("<tableColumns count=\"{}\">", table.headers.len()));
    for (i, header) in table.headers.iter().enumerate() {
        out.push_str(&format!(
            "<tableColumn id=\"{}\" name=\"{}\"/>",
            i + 1,
            escape_xml(header)
        ));
    }
    out.push_str("</tableColumns>");
    out.push_str(
        "<tableStyleInfo name=\"TableStyleMedium2\" showFirstColumn=\"0\" \
         showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>",
    );
    out.push_str("</table>");
    out
}

/// The relationships part of a package part (`xl/worksheets/sheet1.xml` →
/// `xl/worksheets/_rels/sheet1.xml.rels`).
fn rels_part_for(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

/// Add `<tableParts>` referencing `ids` to a worksheet part. It goes after
/// every other child but a trailing worksheet `<extLst>`.
fn add_table_parts(sheet: &str, ids: &[String]) -> io::Result<String> {
    if sheet.contains("<tableParts") {
        return Err(invalid("worksheet already has table parts"));
    }
    let root_end = sheet
        .find("<worksheet")
        .and_then(|at| sheet[at..].find('>').map(|end| at + end))
        .ok_or_else(|| invalid("worksheet part has no root"))?;
    // Declare `r:` locally when the root does not.
    let declare = if sheet[..root_end].contains(&format!("xmlns:r=\"{RELATIONSHIPS_NS}\"")) {
        String::new()
    } else {
        format!(" xmlns:r=\"{RELATIONSHIPS_NS}\"")
    };
    let mut fragment = format!("<tableParts count=\"{}\"{declare}>", ids.len());
    for id in ids {
        fragment.push_str(&format!("<tablePart r:id=\"{id}\"/>"));
    }
    fragment.push_str("</tableParts>");
    let sheet_data_end = sheet.find("</sheetData>").unwrap_or(0);
    match sheet.rfind("<extLst") {
        Some(at) if at > sheet_data_end => {
            Ok(format!("{}{fragment}{}", &sheet[..at], &sheet[at..]))
        }
        _ => insert_before_close(sheet, "worksheet", &fragment),
    }
}

/// The largest `N` among entries named `{prefix}N{suffix}`.
fn max_numbered_part(archive: &zip::ZipArchive<Cursor<&[u8]>>, prefix: &str, suffix: &str) -> u32 {
    archive
        .file_names()
        .filter_map(|name| {
            name.ok()?
                .strip_prefix(prefix)?
                .strip_suffix(suffix)?
                .parse()
                .ok()
        })
        .max()
        .unwrap_or(0)
}

/// Rewrite a package saved by Umya so it carries the engine's defined names
/// (`<definedNames>` in `xl/workbook.xml`) and tables (`xl/tables/tableN.xml`
/// parts, each referenced from its worksheet's `<tableParts>` and `.rels`
/// and given a content-type override). Table formulas stay ordinary cell
/// formulas; Excel does not need `calculatedColumnFormula` to open them.
pub(crate) fn write_names_and_tables(
    bytes: &[u8],
    names: &[DefinedNameExport],
    tables: &[TableExport],
) -> io::Result<Vec<u8>> {
    if names.is_empty() && tables.is_empty() {
        return Ok(bytes.to_vec());
    }
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).map_err(io::Error::other)?;
    let read =
        |archive: &mut zip::ZipArchive<Cursor<&[u8]>>, name: &str| -> io::Result<Option<String>> {
            let mut entry = match archive.by_name(name) {
                Ok(entry) => entry,
                Err(zip::result::ZipError::FileNotFound) => return Ok(None),
                Err(e) => return Err(io::Error::other(e)),
            };
            let mut out = String::new();
            entry.read_to_string(&mut out)?;
            Ok(Some(out))
        };
    let required = |part: Option<String>, name: &str| {
        part.ok_or_else(|| invalid(format!("package has no {name}")))
    };
    let mut parts: HashMap<String, Vec<u8>> = HashMap::new();
    let workbook = required(read(&mut archive, WORKBOOK)?, WORKBOOK)?;

    if !tables.is_empty() {
        let rels = required(read(&mut archive, WORKBOOK_RELS)?, WORKBOOK_RELS)?;
        let sheet_parts = sheet_parts(workbook.as_bytes(), rels.as_bytes())?;
        let mut types = required(read(&mut archive, CONTENT_TYPES)?, CONTENT_TYPES)?;
        let first = max_numbered_part(&archive, "xl/tables/table", ".xml") + 1;
        let mut by_part: BTreeMap<&str, Vec<String>> = BTreeMap::new();
        let mut sheet_rels: HashMap<String, String> = HashMap::new();
        let mut type_overrides = String::new();
        for (next, table) in (first..).zip(tables) {
            let part = sheet_parts
                .get(&table.sheet)
                .ok_or_else(|| invalid(format!("no worksheet part for '{}'", table.sheet)))?;
            let rels_part = rels_part_for(part);
            let rels_xml = match sheet_rels.get(&rels_part) {
                Some(xml) => xml.clone(),
                None => read(&mut archive, &rels_part)?.unwrap_or_else(|| {
                    format!(
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
                         <Relationships xmlns=\"{PACKAGE_RELATIONSHIPS_NS}\"></Relationships>"
                    )
                }),
            };
            let rid = fresh_relationship_id(&rels_xml);
            let rels_xml = insert_before_close(
                &rels_xml,
                "Relationships",
                &format!(
                    "<Relationship Id=\"{rid}\" Type=\"{TABLE_RELATIONSHIP}\" Target=\"../tables/table{next}.xml\"/>"
                ),
            )?;
            sheet_rels.insert(rels_part, rels_xml);
            by_part.entry(part.as_str()).or_default().push(rid);
            parts.insert(
                format!("xl/tables/table{next}.xml"),
                table_xml(table, next).into_bytes(),
            );
            type_overrides.push_str(&format!(
                "<Override PartName=\"/xl/tables/table{next}.xml\" ContentType=\"{TABLE_CONTENT_TYPE}\"/>"
            ));
        }
        for (part, ids) in by_part {
            let sheet = required(read(&mut archive, part)?, part)?;
            parts.insert(
                part.to_string(),
                add_table_parts(&sheet, &ids)?.into_bytes(),
            );
        }
        for (part, xml) in sheet_rels {
            parts.insert(part, xml.into_bytes());
        }
        if !types.contains("Extension=\"rels\"") {
            type_overrides.push_str(
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
            );
        }
        types = insert_before_close(&types, "Types", &type_overrides)?;
        parts.insert(CONTENT_TYPES.to_string(), types.into_bytes());
    }

    if !names.is_empty() {
        parts.insert(
            WORKBOOK.to_string(),
            add_defined_names(&workbook, names)?.into_bytes(),
        );
    }
    rewrite_archive(&mut archive, parts, bytes.len())
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
            ("=Table1[@Col] + 1", "=Table1[[#This Row],[Col]] + 1"),
            (
                "=Sales[[@],[Qty]] * Sales[[@],[Unit Price]]",
                "=Sales[[#This Row],[Qty]] * Sales[[#This Row],[Unit Price]]",
            ),
            ("=Sales[@[Unit Price]]", "=Sales[[#This Row],[Unit Price]]"),
            (
                "=SUM(Sales[@[Qty]:[Total]])",
                "=SUM(Sales[[#This Row],[Qty]:[Total]])",
            ),
            ("=Sales[@]", "=Sales[#This Row]"),
            ("=Sales[@'[x']]", "=Sales[[#This Row],['[x']]]"),
            ("=Sales[[#Totals],[Total]]", "=Sales[[#Totals],[Total]]"),
            ("=\"[@x]\"", "=\"[@x]\""),
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
