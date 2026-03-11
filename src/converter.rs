/// Excel (.xlsx) → custom XML converter.
///
/// An xlsx file is a ZIP archive containing XML documents.  We parse:
///   xl/workbook.xml            – sheet list
///   xl/_rels/workbook.xml.rels – sheet-id → file-name mapping
///   xl/sharedStrings.xml       – shared string table
///   xl/styles.xml              – fonts, fills, borders, cellXfs
///   xl/worksheets/sheet*.xml   – cell data, merges, col/row dims
///   xl/drawings/drawing*.xml   – image anchors (optional)
///   xl/media/*                 – raw image bytes (optional)
///
/// The resulting XML format matches what the JavaScript frontend expects.
use std::collections::HashMap;
use std::fs;
use std::io::{BufReader, Read};
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use base64::{engine::general_purpose::STANDARD as B64, Engine};
use quick_xml::events::Event;
use quick_xml::Reader;

// ── Public entry-points ───────────────────────────────────────────────────────

/// Compute the MD5 hex digest of a file (for cache validation).
pub fn get_file_hash(path: &Path) -> Result<String> {
    let bytes = fs::read(path).with_context(|| format!("reading {}", path.display()))?;
    let digest = md5::compute(&bytes);
    Ok(format!("{:x}", digest))
}

/// Return `true` when the XML cache is present and its stored hash still
/// matches the current xlsx file.
pub fn is_cache_valid(xlsx_path: &Path, xml_path: &Path) -> bool {
    if !xml_path.exists() {
        return false;
    }
    let Ok(xml_bytes) = fs::read(xml_path) else {
        return false;
    };
    let Ok(xml_str) = std::str::from_utf8(&xml_bytes) else {
        return false;
    };
    // Extract the hash attribute from the root <workbook> element
    let cached_hash = extract_root_attr(xml_str, "hash").unwrap_or_default();
    let Ok(current_hash) = get_file_hash(xlsx_path) else {
        return false;
    };
    cached_hash == current_hash
}

/// Convert an xlsx file to our custom XML format, writing the result to
/// `xml_path`.  Creates parent directories as needed.
pub fn convert_excel_to_xml(xlsx_path: &Path, xml_path: &Path) -> Result<()> {
    let file_hash = get_file_hash(xlsx_path)?;
    let filename = xlsx_path
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("unknown.xlsx")
        .to_string();
    let converted = chrono::Local::now().format("%Y-%m-%dT%H:%M:%S").to_string();

    // Load the entire ZIP into a name→bytes map for easy random access.
    let zip_bytes = fs::read(xlsx_path)?;
    let entries = read_zip_entries(&zip_bytes)?;

    // Parse components
    let shared_strings = parse_shared_strings(&entries);
    let styles = parse_styles(&entries);
    let sheet_entries = parse_workbook(&entries);

    // Build XML
    let mut xml = String::with_capacity(65536);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    xml.push_str(&format!(
        "<workbook filename=\"{}\" converted=\"{}\" hash=\"{}\">\n",
        xml_escape(&filename),
        xml_escape(&converted),
        xml_escape(&file_hash)
    ));

    for (sheet_name, sheet_path) in &sheet_entries {
        let xml_rel_path = &sheet_path[..]; // e.g. "xl/worksheets/sheet1.xml"
        if let Some(ws_bytes) = entries.get(xml_rel_path) {
            let ws_data =
                WorksheetData::parse(ws_bytes, sheet_name, &shared_strings, &styles, &entries)?;
            ws_data.write_xml(&mut xml);
        }
    }

    xml.push_str("</workbook>\n");

    if let Some(parent) = xml_path.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::write(xml_path, xml.as_bytes())?;
    Ok(())
}

// ── ZIP helpers ───────────────────────────────────────────────────────────────

fn read_zip_entries(bytes: &[u8]) -> Result<HashMap<String, Vec<u8>>> {
    let cursor = std::io::Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(cursor)?;
    let mut map = HashMap::new();
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i)?;
        let name = entry.name().to_string();
        let mut data = Vec::new();
        entry.read_to_end(&mut data)?;
        map.insert(name, data);
    }
    Ok(map)
}

// ── Shared strings ────────────────────────────────────────────────────────────

fn parse_shared_strings(entries: &HashMap<String, Vec<u8>>) -> Vec<String> {
    let Some(data) = entries.get("xl/sharedStrings.xml") else {
        return vec![];
    };
    let mut strings = Vec::new();
    let mut reader = Reader::from_reader(data.as_slice());
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut inside_si = false;
    let mut current = String::new();
    let mut inside_t = false;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"si" => {
                    inside_si = true;
                    current.clear();
                }
                b"t" if inside_si => {
                    inside_t = true;
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"t" && inside_si {
                    // empty <t/> contributes nothing
                }
            }
            Ok(Event::Text(ref e)) => {
                if inside_t {
                    if let Ok(s) = e.unescape() {
                        current.push_str(&s);
                    }
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"t" => inside_t = false,
                b"si" => {
                    strings.push(current.clone());
                    inside_si = false;
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    strings
}

// ── Styles ────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Default)]
struct StyleFont {
    bold: bool,
    italic: bool,
    underline: bool,
    size: f64,
    name: String,
    color: String,
}

#[derive(Debug, Clone, Default)]
struct StyleFill {
    bg_color: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct StyleBorderSide {
    style: String,
}

#[derive(Debug, Clone, Default)]
struct StyleBorder {
    left: StyleBorderSide,
    right: StyleBorderSide,
    top: StyleBorderSide,
    bottom: StyleBorderSide,
}

#[derive(Debug, Clone, Default)]
struct StyleAlign {
    horizontal: String,
    vertical: String,
    wrap_text: bool,
    indent: u32,
}

#[derive(Debug, Clone, Default)]
struct CellXf {
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    num_fmt_id: u32,
    alignment: StyleAlign,
}

#[derive(Debug, Default)]
struct Styles {
    fonts: Vec<StyleFont>,
    fills: Vec<StyleFill>,
    borders: Vec<StyleBorder>,
    cell_xfs: Vec<CellXf>,
    num_fmts: HashMap<u32, String>,
}

impl Styles {
    fn get_xf(&self, idx: usize) -> Option<&CellXf> {
        self.cell_xfs.get(idx)
    }

    fn font(&self, idx: usize) -> &StyleFont {
        static DEFAULT: StyleFont = StyleFont {
            bold: false,
            italic: false,
            underline: false,
            size: 11.0,
            name: String::new(),
            color: String::new(),
        };
        self.fonts.get(idx).unwrap_or(&DEFAULT)
    }

    fn fill(&self, idx: usize) -> &StyleFill {
        static DEFAULT: StyleFill = StyleFill { bg_color: None };
        self.fills.get(idx).unwrap_or(&DEFAULT)
    }

    fn border(&self, idx: usize) -> &StyleBorder {
        static DEFAULT: StyleBorder = StyleBorder {
            left: StyleBorderSide { style: String::new() },
            right: StyleBorderSide { style: String::new() },
            top: StyleBorderSide { style: String::new() },
            bottom: StyleBorderSide { style: String::new() },
        };
        self.borders.get(idx).unwrap_or(&DEFAULT)
    }

    fn num_fmt(&self, id: u32) -> &str {
        if let Some(s) = self.num_fmts.get(&id) {
            return s.as_str();
        }
        // Built-in number format IDs
        match id {
            0 => "General",
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
            9 => "0%",
            10 => "0.00%",
            11 => "0.00E+00",
            12 => "# ?/?",
            13 => "# ??/??",
            14 => "m/d/yyyy",
            15 => "d-mmm-yy",
            16 => "d-mmm",
            17 => "mmm-yy",
            18 => "h:mm AM/PM",
            19 => "h:mm:ss AM/PM",
            20 => "h:mm",
            21 => "h:mm:ss",
            22 => "m/d/yyyy h:mm",
            37 => "#,##0 ;(#,##0)",
            38 => "#,##0 ;[Red](#,##0)",
            39 => "#,##0.00;(#,##0.00)",
            40 => "#,##0.00;[Red](#,##0.00)",
            45 => "mm:ss",
            46 => "[h]:mm:ss",
            47 => "mmss.0",
            48 => "##0.0E+0",
            49 => "@",
            _ => "General",
        }
    }
}

// Possible states while parsing styles.xml
#[derive(Debug, PartialEq, Clone, Copy)]
enum StyleSection {
    None,
    NumFmts,
    Fonts,
    Fills,
    Borders,
    CellXfs,
}
#[derive(Debug, PartialEq, Clone, Copy)]
enum BorderSideState {
    None,
    Left,
    Right,
    Top,
    Bottom,
}

fn parse_styles(entries: &HashMap<String, Vec<u8>>) -> Styles {
    let Some(data) = entries.get("xl/styles.xml") else {
        return Styles::default();
    };
    let mut styles = Styles::default();
    let mut reader = Reader::from_reader(data.as_slice());
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut section = StyleSection::None;
    let mut border_side = BorderSideState::None;

    // Working copies during parsing
    let mut cur_font = StyleFont { size: 11.0, ..Default::default() };
    let mut cur_fill = StyleFill::default();
    let mut cur_border = StyleBorder::default();
    let mut cur_xf = CellXf::default();
    // Track depth inside a <font>/<fill>/<border> element
    let mut depth: u32 = 0;

    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf);
        // Track whether the current event is an empty (self-closing) element.
        // Self-closing elements like <xf ... /> do NOT fire a matching End event,
        // so we must push accumulated data immediately for those.
        let is_empty_ev = matches!(&event, Ok(Event::Empty(_)));
        match event {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                macro_rules! attr {
                    ($name:expr) => {
                        e.try_get_attribute($name)
                            .ok()
                            .flatten()
                            .and_then(|a| a.unescape_value().ok())
                            .map(|c| c.into_owned())
                    };
                }

                match e.local_name().as_ref() {
                    // ── Section starts ──
                    b"numFmts" => section = StyleSection::NumFmts,
                    b"fonts" => section = StyleSection::Fonts,
                    b"fills" => section = StyleSection::Fills,
                    b"borders" => section = StyleSection::Borders,
                    b"cellXfs" => section = StyleSection::CellXfs,

                    // ── numFmt ──
                    b"numFmt" if section == StyleSection::NumFmts => {
                        if let (Some(id_s), Some(code)) =
                            (attr!("numFmtId"), attr!("formatCode"))
                        {
                            if let Ok(id) = id_s.parse::<u32>() {
                                styles.num_fmts.insert(id, code);
                            }
                        }
                    }

                    // ── Font sub-elements ──
                    b"b" if section == StyleSection::Fonts => {
                        cur_font.bold = attr!("val").map_or(true, |v| v != "0");
                    }
                    b"i" if section == StyleSection::Fonts => {
                        cur_font.italic = attr!("val").map_or(true, |v| v != "0");
                    }
                    b"u" if section == StyleSection::Fonts => {
                        let val = attr!("val").unwrap_or_else(|| "single".to_string());
                        cur_font.underline = val != "none";
                    }
                    b"sz" if section == StyleSection::Fonts => {
                        if let Some(v) = attr!("val") {
                            cur_font.size = v.parse::<f64>().unwrap_or(11.0);
                        }
                    }
                    b"name" if section == StyleSection::Fonts => {
                        if let Some(v) = attr!("val") {
                            cur_font.name = v;
                        }
                    }
                    b"color" if section == StyleSection::Fonts => {
                        if let Some(c) = resolve_color(
                            attr!("rgb").as_deref(),
                            attr!("theme").as_deref().and_then(|s| s.parse().ok()),
                            attr!("indexed").as_deref().and_then(|s| s.parse().ok()),
                        ) {
                            cur_font.color = c;
                        }
                    }

                    // ── Fill sub-elements ──
                    b"fgColor" if section == StyleSection::Fills => {
                        cur_fill.bg_color = resolve_color(
                            attr!("rgb").as_deref(),
                            attr!("theme").as_deref().and_then(|s| s.parse().ok()),
                            attr!("indexed").as_deref().and_then(|s| s.parse().ok()),
                        )
                        .filter(|c| c != "FF000000" && c != "00000000" && c != "FFFFFFFF");
                    }

                    // ── Border side states ──
                    b"left" if section == StyleSection::Borders => {
                        border_side = BorderSideState::Left;
                        if let Some(s) = attr!("style") {
                            cur_border.left.style = s;
                        }
                    }
                    b"right" if section == StyleSection::Borders => {
                        border_side = BorderSideState::Right;
                        if let Some(s) = attr!("style") {
                            cur_border.right.style = s;
                        }
                    }
                    b"top" if section == StyleSection::Borders => {
                        border_side = BorderSideState::Top;
                        if let Some(s) = attr!("style") {
                            cur_border.top.style = s;
                        }
                    }
                    b"bottom" if section == StyleSection::Borders => {
                        border_side = BorderSideState::Bottom;
                        if let Some(s) = attr!("style") {
                            cur_border.bottom.style = s;
                        }
                    }
                    b"diagonal" if section == StyleSection::Borders => {
                        border_side = BorderSideState::None;
                    }

                    // ── CellXf ──
                    b"xf" if section == StyleSection::CellXfs => {
                        cur_xf = CellXf::default();
                        if let Some(v) = attr!("fontId").and_then(|s| s.parse().ok()) {
                            cur_xf.font_id = v;
                        }
                        if let Some(v) = attr!("fillId").and_then(|s| s.parse().ok()) {
                            cur_xf.fill_id = v;
                        }
                        if let Some(v) = attr!("borderId").and_then(|s| s.parse().ok()) {
                            cur_xf.border_id = v;
                        }
                        if let Some(v) = attr!("numFmtId").and_then(|s| s.parse().ok()) {
                            cur_xf.num_fmt_id = v;
                        }
                    }
                    b"alignment" if section == StyleSection::CellXfs => {
                        cur_xf.alignment.horizontal =
                            attr!("horizontal").unwrap_or_default();
                        cur_xf.alignment.vertical =
                            attr!("vertical").unwrap_or_default();
                        cur_xf.alignment.wrap_text = attr!("wrapText")
                            .map_or(false, |v| v == "1" || v == "true");
                        cur_xf.alignment.indent = attr!("indent")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(0);
                    }

                    // ── Depth tracking for font/fill/border sub-elements ──
                    b"font" if section == StyleSection::Fonts => {
                        cur_font = StyleFont { size: 11.0, ..Default::default() };
                        depth = 1;
                    }
                    b"fill" if section == StyleSection::Fills => {
                        cur_fill = StyleFill::default();
                        depth = 1;
                    }
                    b"border" if section == StyleSection::Borders => {
                        cur_border = StyleBorder::default();
                        border_side = BorderSideState::None;
                        depth = 1;
                    }

                    _ => {}
                }

                // Self-closing <xf .../> elements have no matching End event,
                // so we must push the accumulated CellXf immediately.
                if is_empty_ev && section == StyleSection::CellXfs
                    && e.local_name().as_ref() == b"xf"
                {
                    styles.cell_xfs.push(cur_xf.clone());
                    cur_xf = CellXf::default();
                }
            }

            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"numFmts" | b"cellXfs" => section = StyleSection::None,
                b"fonts" => section = StyleSection::None,
                b"fills" => section = StyleSection::None,
                b"borders" => section = StyleSection::None,

                b"font" if section == StyleSection::Fonts => {
                    styles.fonts.push(cur_font.clone());
                    cur_font = StyleFont { size: 11.0, ..Default::default() };
                }
                b"fill" if section == StyleSection::Fills => {
                    styles.fills.push(cur_fill.clone());
                    cur_fill = StyleFill::default();
                }
                b"border" if section == StyleSection::Borders => {
                    styles.borders.push(cur_border.clone());
                    cur_border = StyleBorder::default();
                    border_side = BorderSideState::None;
                }
                b"xf" if section == StyleSection::CellXfs => {
                    styles.cell_xfs.push(cur_xf.clone());
                    cur_xf = CellXf::default();
                }

                b"left" | b"right" | b"top" | b"bottom"
                    if section == StyleSection::Borders =>
                {
                    border_side = BorderSideState::None;
                }

                _ => {}
            },

            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    styles
}

// ── Workbook / sheet list ─────────────────────────────────────────────────────

/// Returns (sheet_name, relative_xml_path) in workbook order.
fn parse_workbook(entries: &HashMap<String, Vec<u8>>) -> Vec<(String, String)> {
    // 1. Read workbook.xml to get ordered sheet names & r:id
    let sheet_ids: Vec<(String, String)> = entries
        .get("xl/workbook.xml")
        .map(|d| extract_sheets(d))
        .unwrap_or_default();

    // 2. Read rels to map r:id → file path
    let rels = entries
        .get("xl/_rels/workbook.xml.rels")
        .map(|d| parse_rels(d))
        .unwrap_or_default();

    sheet_ids
        .into_iter()
        .filter_map(|(name, rid)| {
            let target = rels.get(&rid)?;
            let path = if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("xl/{}", target)
            };
            Some((name, path))
        })
        .collect()
}

fn extract_sheets(data: &[u8]) -> Vec<(String, String)> {
    let mut out = Vec::new();
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut in_sheets = false;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"sheets" {
                    in_sheets = true;
                }
            }
            Ok(Event::Empty(ref e)) if in_sheets && e.local_name().as_ref() == b"sheet" => {
                let name = e
                    .try_get_attribute("name")
                    .ok()
                    .flatten()
                    .and_then(|a| a.unescape_value().ok())
                    .map(|c| c.into_owned())
                    .unwrap_or_default();
                let rid = e
                    .try_get_attribute("id")
                    .ok()
                    .flatten()
                    .or_else(|| {
                        // r:id is stored as attribute with namespace prefix
                        // quick-xml may see it as "r:id"
                        e.attributes()
                            .filter_map(|a| a.ok())
                            .find(|a| a.key.as_ref().ends_with(b":id") || a.key.as_ref() == b"id")
                            .map(|a| a)
                    })
                    .and_then(|a| a.unescape_value().ok())
                    .map(|c| c.into_owned())
                    .unwrap_or_default();
                if !name.is_empty() && !rid.is_empty() {
                    out.push((name, rid));
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheets" => {
                in_sheets = false;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    out
}

fn parse_rels(data: &[u8]) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"Relationship" => {
                let id = e
                    .try_get_attribute("Id")
                    .ok()
                    .flatten()
                    .and_then(|a| a.unescape_value().ok())
                    .map(|c| c.into_owned())
                    .unwrap_or_default();
                let target = e
                    .try_get_attribute("Target")
                    .ok()
                    .flatten()
                    .and_then(|a| a.unescape_value().ok())
                    .map(|c| c.into_owned())
                    .unwrap_or_default();
                if !id.is_empty() && !target.is_empty() {
                    map.insert(id, target);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    map
}

// ── Worksheet parsing ─────────────────────────────────────────────────────────

#[derive(Debug, Default)]
struct WorksheetData {
    name: String,
    max_row: u32,
    max_col: u32,
    col_widths: Vec<ColWidth>,
    row_heights: Vec<RowHeight>,
    merges: Vec<Merge>,
    cells: Vec<Cell>,
    images: Vec<Image>,
}

#[derive(Debug)]
struct ColWidth {
    index: u32,
    width: f64,
    hidden: bool,
}

#[derive(Debug)]
struct RowHeight {
    index: u32,
    height: f64,
    hidden: bool,
}

#[derive(Debug)]
struct Merge {
    ref_str: String,
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
}

#[derive(Debug, Default, Clone)]
struct CellFormat {
    bold: bool,
    italic: bool,
    underline: bool,
    font_size: f64,
    font_name: String,
    font_color: String,
    bg_color: Option<String>,
    halign: String,
    valign: String,
    wrap_text: bool,
    indent: u32,
    border_top: String,
    border_bottom: String,
    border_left: String,
    border_right: String,
    num_fmt: String,
}

#[derive(Debug)]
struct Cell {
    row: u32,
    col: u32,
    value: String,
    cell_type: String,
    rowspan: u32,
    colspan: u32,
    format: Option<CellFormat>,
}

#[derive(Debug)]
struct Image {
    row: u32,
    col: u32,
    to_row: Option<u32>,
    to_col: Option<u32>,
    data_b64: String,
    fmt: String,
    width: Option<f64>,
    height: Option<f64>,
}

// Worksheet parser state
#[derive(Debug, PartialEq, Clone, Copy)]
enum WsSection {
    None,
    Cols,
    SheetData,
    MergeCells,
}

impl WorksheetData {
    fn parse(
        data: &[u8],
        sheet_name: &str,
        shared_strings: &[String],
        styles: &Styles,
        entries: &HashMap<String, Vec<u8>>,
    ) -> Result<Self> {
        let mut ws = WorksheetData {
            name: sheet_name.to_string(),
            ..Default::default()
        };

        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();

        let mut section = WsSection::None;
        let mut cur_row: u32 = 0;
        let mut cur_col: u32 = 0;
        let mut cur_style: Option<usize> = None;
        let mut cur_type = String::new();
        let mut in_v = false;
        let mut in_t = false;  // inside <is><t>
        let mut cur_value = String::new();
        let mut drawing_rid: Option<String> = None;

        // Build merge lookup: (row, col) → (start_row, start_col, end_row, end_col)
        let mut merge_map: HashMap<(u32, u32), (u32, u32, u32, u32)> = HashMap::new();

        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    macro_rules! attr {
                        ($name:expr) => {
                            e.try_get_attribute($name)
                                .ok()
                                .flatten()
                                .and_then(|a| a.unescape_value().ok())
                                .map(|c| c.into_owned())
                        };
                    }

                    match e.local_name().as_ref() {
                        b"dimension" => {
                            // <dimension ref="A1:D7"/>
                            if let Some(ref_str) = attr!("ref") {
                                if let Some(colon) = ref_str.find(':') {
                                    let end = &ref_str[colon + 1..];
                                    let (r, c) = cell_ref_to_row_col(end);
                                    ws.max_row = r;
                                    ws.max_col = c;
                                }
                            }
                        }

                        b"cols" => section = WsSection::Cols,

                        b"col" if section == WsSection::Cols => {
                            let min: u32 = attr!("min").and_then(|s| s.parse().ok()).unwrap_or(1);
                            let max: u32 = attr!("max").and_then(|s| s.parse().ok()).unwrap_or(min);
                            let width: f64 =
                                attr!("width").and_then(|s| s.parse().ok()).unwrap_or(8.43);
                            let hidden = attr!("hidden").map_or(false, |v| v == "1");
                            for idx in min..=max {
                                ws.col_widths.push(ColWidth {
                                    index: idx,
                                    width,
                                    hidden,
                                });
                            }
                        }

                        b"sheetData" => section = WsSection::SheetData,

                        b"row" if section == WsSection::SheetData => {
                            cur_row = attr!("r").and_then(|s| s.parse().ok()).unwrap_or(0);
                            let ht: f64 = attr!("ht").and_then(|s| s.parse().ok()).unwrap_or(0.0);
                            let hidden = attr!("hidden").map_or(false, |v| v == "1");
                            if ht > 0.0 || hidden {
                                let custom =
                                    attr!("customHeight").map_or(false, |v| v == "1");
                                if custom || hidden {
                                    ws.row_heights.push(RowHeight {
                                        index: cur_row,
                                        height: if ht > 0.0 { ht } else { 15.0 },
                                        hidden,
                                    });
                                    // Update max_row from actual rows
                                    if cur_row > ws.max_row {
                                        ws.max_row = cur_row;
                                    }
                                }
                            }
                        }

                        b"c" if section == WsSection::SheetData => {
                            // <c r="A1" s="1" t="s">
                            if let Some(r) = attr!("r") {
                                let (row, col) = cell_ref_to_row_col(&r);
                                cur_row = row;
                                cur_col = col;
                                if row > ws.max_row {
                                    ws.max_row = row;
                                }
                                if col > ws.max_col {
                                    ws.max_col = col;
                                }
                            }
                            cur_style = attr!("s").and_then(|s| s.parse().ok());
                            cur_type = attr!("t").unwrap_or_else(|| "n".to_string());
                            cur_value.clear();
                            in_v = false;
                            in_t = false;
                        }

                        b"v" if section == WsSection::SheetData => {
                            in_v = true;
                            cur_value.clear();
                        }
                        b"t" if section == WsSection::SheetData => {
                            in_t = true;
                            cur_value.clear();
                        }

                        b"mergeCells" => section = WsSection::MergeCells,

                        b"mergeCell" if section == WsSection::MergeCells => {
                            if let Some(ref_str) = attr!("ref") {
                                if let Some(m) = parse_merge_ref(&ref_str) {
                                    let (sr, sc, er, ec) = m;
                                    ws.merges.push(Merge {
                                        ref_str: ref_str.clone(),
                                        start_row: sr,
                                        start_col: sc,
                                        end_row: er,
                                        end_col: ec,
                                    });
                                    for r in sr..=er {
                                        for c in sc..=ec {
                                            merge_map.insert((r, c), (sr, sc, er, ec));
                                        }
                                    }
                                }
                            }
                        }

                        b"drawing" => {
                            // <drawing r:id="rId2"/>
                            drawing_rid = e
                                .attributes()
                                .filter_map(|a| a.ok())
                                .find(|a| {
                                    let k = a.key.as_ref();
                                    k.ends_with(b":id") || k == b"id"
                                })
                                .and_then(|a| a.unescape_value().ok())
                                .map(|c| c.into_owned());
                        }

                        _ => {}
                    }
                }

                Ok(Event::Text(ref e)) => {
                    if in_v || in_t {
                        if let Ok(s) = e.unescape() {
                            cur_value.push_str(&s);
                        }
                    }
                }

                Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                    b"cols" => section = WsSection::None,
                    b"sheetData" => section = WsSection::None,
                    b"mergeCells" => section = WsSection::None,

                    b"v" => in_v = false,
                    b"t" => in_t = false,

                    b"c" if section == WsSection::SheetData => {
                        // Resolve shared string
                        let value = if cur_type == "s" {
                            let idx: usize = cur_value.trim().parse().unwrap_or(0);
                            shared_strings
                                .get(idx)
                                .cloned()
                                .unwrap_or_else(|| cur_value.clone())
                        } else if cur_type == "b" {
                            if cur_value.trim() == "1" {
                                "TRUE".to_string()
                            } else {
                                "FALSE".to_string()
                            }
                        } else if cur_type == "inlineStr" {
                            cur_value.clone()
                        } else {
                            cur_value.trim().to_string()
                        };

                        // Determine cell_type for our XML ("s", "n", "b")
                        let xml_type = match cur_type.as_str() {
                            "s" | "inlineStr" | "str" | "e" => "s",
                            "b" => "b",
                            _ => "n",
                        };

                        // Look up style
                        let fmt = cur_style.and_then(|s| styles.get_xf(s)).map(|xf| {
                            let font = styles.font(xf.font_id);
                            let fill = styles.fill(xf.fill_id);
                            let border = styles.border(xf.border_id);
                            let num_fmt = styles.num_fmt(xf.num_fmt_id).to_string();

                            let font_color = if font.color.is_empty() {
                                "FF000000".to_string()
                            } else {
                                font.color.clone()
                            };
                            let font_size = if font.size > 0.0 { font.size } else { 11.0 };
                            let font_name = if font.name.is_empty() {
                                "Calibri".to_string()
                            } else {
                                font.name.clone()
                            };

                            CellFormat {
                                bold: font.bold,
                                italic: font.italic,
                                underline: font.underline,
                                font_size,
                                font_name,
                                font_color,
                                bg_color: fill.bg_color.clone(),
                                halign: xf.alignment.horizontal.clone(),
                                valign: xf.alignment.vertical.clone(),
                                wrap_text: xf.alignment.wrap_text,
                                indent: xf.alignment.indent,
                                border_top: border.top.style.clone(),
                                border_bottom: border.bottom.style.clone(),
                                border_left: border.left.style.clone(),
                                border_right: border.right.style.clone(),
                                num_fmt,
                            }
                        });

                        // Determine merge spans
                        let (rowspan, colspan) =
                            if let Some((sr, sc, er, ec)) = merge_map.get(&(cur_row, cur_col)) {
                                let rs = er - sr + 1;
                                let cs = ec - sc + 1;
                                (rs, cs)
                            } else {
                                (1, 1)
                            };

                        // Skip non-origin merge cells
                        let is_merge_non_origin =
                            if let Some((sr, sc, _er, _ec)) = merge_map.get(&(cur_row, cur_col)) {
                                cur_row != *sr || cur_col != *sc
                            } else {
                                false
                            };

                        if !is_merge_non_origin {
                            // Emit cell if it has content or style
                            if !value.is_empty() || fmt.is_some() || rowspan > 1 || colspan > 1 {
                                ws.cells.push(Cell {
                                    row: cur_row,
                                    col: cur_col,
                                    value,
                                    cell_type: xml_type.to_string(),
                                    rowspan,
                                    colspan,
                                    format: fmt,
                                });
                            }
                        }
                    }

                    _ => {}
                },

                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }

        // Post-process: apply merge spans to origin cells now that merge_map is fully populated.
        // (In xlsx, <mergeCells> appears AFTER <sheetData>, so merges weren't known during cell parsing.)
        for cell in ws.cells.iter_mut() {
            if let Some((sr, sc, er, ec)) = merge_map.get(&(cell.row, cell.col)) {
                if cell.row == *sr && cell.col == *sc {
                    cell.rowspan = er - sr + 1;
                    cell.colspan = ec - sc + 1;
                }
            }
        }

        // Parse images if a drawing reference was found
        if let Some(rid) = drawing_rid {
            // Find the sheet's rels file to resolve the drawing path
            // We need to know the sheet file name to find its rels
            // Images are parsed from drawings
            if let Some(imgs) = parse_images_for_drawing(&rid, entries, &ws.name) {
                ws.images = imgs;
            }
        }

        Ok(ws)
    }

    fn write_xml(&self, out: &mut String) {
        out.push_str(&format!(
            "  <sheet name=\"{}\">\n",
            xml_escape(&self.name)
        ));

        // dimensions
        out.push_str(&format!(
            "    <dimensions maxRow=\"{}\" maxCol=\"{}\"/>\n",
            self.max_row, self.max_col
        ));

        // colWidths
        out.push_str("    <colWidths>\n");
        for cw in &self.col_widths {
            out.push_str(&format!(
                "      <col index=\"{}\" width=\"{}\" hidden=\"{}\"/>\n",
                cw.index, cw.width, cw.hidden
            ));
        }
        out.push_str("    </colWidths>\n");

        // rowHeights
        out.push_str("    <rowHeights>\n");
        for rh in &self.row_heights {
            out.push_str(&format!(
                "      <row index=\"{}\" height=\"{}\" hidden=\"{}\"/>\n",
                rh.index, rh.height, rh.hidden
            ));
        }
        out.push_str("    </rowHeights>\n");

        // merges
        out.push_str("    <merges>\n");
        for m in &self.merges {
            out.push_str(&format!(
                "      <merge ref=\"{}\" startRow=\"{}\" startCol=\"{}\" endRow=\"{}\" endCol=\"{}\"/>\n",
                xml_escape(&m.ref_str),
                m.start_row,
                m.start_col,
                m.end_row,
                m.end_col
            ));
        }
        out.push_str("    </merges>\n");

        // rows / cells
        out.push_str("    <rows>\n");

        // Group cells by row
        let mut rows_map: std::collections::BTreeMap<u32, Vec<&Cell>> =
            std::collections::BTreeMap::new();
        for cell in &self.cells {
            rows_map.entry(cell.row).or_default().push(cell);
        }

        for (row_idx, cells) in &rows_map {
            out.push_str(&format!("      <row index=\"{}\">\n", row_idx));
            // Sort cells by column
            let mut sorted = cells.clone();
            sorted.sort_by_key(|c| c.col);
            for cell in sorted {
                self.write_cell_xml(out, cell);
            }
            out.push_str("      </row>\n");
        }
        out.push_str("    </rows>\n");

        // images
        out.push_str("    <images>\n");
        for img in &self.images {
            let mut attrs = format!(
                "row=\"{}\" col=\"{}\"",
                img.row, img.col
            );
            if let Some(tr) = img.to_row {
                attrs.push_str(&format!(" toRow=\"{}\"", tr));
            }
            if let Some(tc) = img.to_col {
                attrs.push_str(&format!(" toCol=\"{}\"", tc));
            }
            attrs.push_str(&format!(" format=\"{}\"", xml_escape(&img.fmt)));
            if let Some(w) = img.width {
                attrs.push_str(&format!(" width=\"{}\"", w));
            }
            if let Some(h) = img.height {
                attrs.push_str(&format!(" height=\"{}\"", h));
            }
            attrs.push_str(&format!(" data=\"{}\"", img.data_b64));
            out.push_str(&format!("      <image {}/>\n", attrs));
        }
        out.push_str("    </images>\n");

        out.push_str("  </sheet>\n");
    }

    fn write_cell_xml(&self, out: &mut String, cell: &Cell) {
        let mut attrs = format!(
            "row=\"{}\" col=\"{}\" value=\"{}\" type=\"{}\"",
            cell.row,
            cell.col,
            xml_escape(&cell.value),
            xml_escape(&cell.cell_type)
        );
        if cell.rowspan > 1 {
            attrs.push_str(&format!(" rowspan=\"{}\"", cell.rowspan));
        }
        if cell.colspan > 1 {
            attrs.push_str(&format!(" colspan=\"{}\"", cell.colspan));
        }

        if let Some(ref fmt) = cell.format {
            out.push_str(&format!("        <cell {}>\n", attrs));
            let border_top = if fmt.border_top.is_empty() { "none" } else { &fmt.border_top };
            let border_bottom = if fmt.border_bottom.is_empty() { "none" } else { &fmt.border_bottom };
            let border_left = if fmt.border_left.is_empty() { "none" } else { &fmt.border_left };
            let border_right = if fmt.border_right.is_empty() { "none" } else { &fmt.border_right };
            let bg = fmt
                .bg_color
                .as_deref()
                .map(|c| format!(" bgColor=\"{}\"", c))
                .unwrap_or_default();
            out.push_str(&format!(
                "          <format bold=\"{}\" italic=\"{}\" underline=\"{}\" \
                 fontSize=\"{}\" fontName=\"{}\" fontColor=\"{}\"{}  \
                 halign=\"{}\" valign=\"{}\" wrapText=\"{}\" indent=\"{}\" \
                 borderTop=\"{}\" borderBottom=\"{}\" borderLeft=\"{}\" borderRight=\"{}\" \
                 numFmt=\"{}\"/>\n",
                fmt.bold,
                fmt.italic,
                fmt.underline,
                fmt.font_size,
                xml_escape(&fmt.font_name),
                xml_escape(&fmt.font_color),
                bg,
                xml_escape(&fmt.halign),
                xml_escape(&fmt.valign),
                fmt.wrap_text,
                fmt.indent,
                border_top,
                border_bottom,
                border_left,
                border_right,
                xml_escape(&fmt.num_fmt),
            ));
            out.push_str("        </cell>\n");
        } else {
            out.push_str(&format!("        <cell {}/>\n", attrs));
        }
    }
}

// ── Image parsing (from drawing XML) ─────────────────────────────────────────

fn parse_images_for_drawing(
    _drawing_rid: &str,
    entries: &HashMap<String, Vec<u8>>,
    _sheet_name: &str,
) -> Option<Vec<Image>> {
    // We need to find the drawing file.  The rels file for a given worksheet
    // maps rIds to drawing files.  Since we don't always know the worksheet path
    // at this point, we scan all drawing files in the ZIP.
    let mut images = Vec::new();

    // Find all drawing XML files
    let drawing_keys: Vec<String> = entries
        .keys()
        .filter(|k| k.starts_with("xl/drawings/drawing") && k.ends_with(".xml"))
        .cloned()
        .collect();

    for drawing_key in drawing_keys {
        let Some(drawing_data) = entries.get(&drawing_key) else {
            continue;
        };
        // Find the rels for this drawing
        let rels_key = drawing_key
            .replacen("xl/drawings/", "xl/drawings/_rels/", 1)
            + ".rels";
        let drawing_rels = entries
            .get(&rels_key)
            .map(|d| parse_rels(d))
            .unwrap_or_default();

        parse_drawing_images(drawing_data, &drawing_rels, entries, &mut images);
    }

    if images.is_empty() {
        None
    } else {
        Some(images)
    }
}

fn parse_drawing_images(
    data: &[u8],
    rels: &HashMap<String, String>,
    entries: &HashMap<String, Vec<u8>>,
    images: &mut Vec<Image>,
) {
    let mut reader = Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    // State tracking for anchor
    let mut from_col: u32 = 0;
    let mut from_row: u32 = 0;
    let mut to_col: Option<u32> = None;
    let mut to_row: Option<u32> = None;
    let mut cur_val_target = String::new();
    let mut cur_val_text = String::new();
    let mut in_from = false;
    let mut in_to = false;
    let mut in_col = false;
    let mut in_row = false;
    let mut embed_rid: Option<String> = None;
    let mut cur_cx: Option<f64> = None;
    let mut cur_cy: Option<f64> = None;

    macro_rules! push_image {
        () => {
            if let Some(rid) = embed_rid.take() {
                if let Some(target) = rels.get(&rid) {
                    // Resolve path: target is relative to xl/drawings/
                    let media_path = if target.starts_with('/') {
                        target.trim_start_matches('/').to_string()
                    } else {
                        format!("xl/media/{}", target.trim_start_matches("../media/"))
                    };
                    if let Some(img_data) = entries.get(&media_path) {
                        let fmt = media_path
                            .rsplit('.')
                            .next()
                            .unwrap_or("png")
                            .to_lowercase();
                        let data_b64 = B64.encode(img_data);
                        images.push(Image {
                            row: from_row + 1,
                            col: from_col + 1,
                            to_row: to_row.map(|r| r + 1),
                            to_col: to_col.map(|c| c + 1),
                            data_b64,
                            fmt,
                            width: cur_cx.map(|cx| cx / 914400.0 * 96.0),
                            height: cur_cy.map(|cy| cy / 914400.0 * 96.0),
                        });
                    }
                }
            }
            from_col = 0;
            from_row = 0;
            to_col = None;
            to_row = None;
            cur_cx = None;
            cur_cy = None;
            in_from = false;
            in_to = false;
        };
    }

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                macro_rules! attr {
                    ($name:expr) => {
                        e.try_get_attribute($name)
                            .ok()
                            .flatten()
                            .and_then(|a| a.unescape_value().ok())
                            .map(|c| c.into_owned())
                    };
                }
                match e.local_name().as_ref() {
                    b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                        push_image!();
                    }
                    b"from" => {
                        in_from = true;
                        in_to = false;
                    }
                    b"to" => {
                        in_to = true;
                        in_from = false;
                    }
                    b"col" => in_col = true,
                    b"row" => in_row = true,
                    b"blip" => {
                        // <a:blip r:embed="rId1"/>
                        embed_rid = e
                            .attributes()
                            .filter_map(|a| a.ok())
                            .find(|a| {
                                let k = a.key.as_ref();
                                k.ends_with(b":embed") || k == b"embed"
                            })
                            .and_then(|a| a.unescape_value().ok())
                            .map(|c| c.into_owned());
                    }
                    b"ext" => {
                        // <a:ext cx="1234" cy="567"/>  (EMU units)
                        if let Some(v) = attr!("cx") {
                            cur_cx = v.parse::<f64>().ok();
                        }
                        if let Some(v) = attr!("cy") {
                            cur_cy = v.parse::<f64>().ok();
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if let Ok(s) = e.unescape() {
                    cur_val_text = s.trim().to_string();
                }
                if in_col {
                    let n: u32 = cur_val_text.parse().unwrap_or(0);
                    if in_from {
                        from_col = n;
                    } else if in_to {
                        to_col = Some(n);
                    }
                }
                if in_row {
                    let n: u32 = cur_val_text.parse().unwrap_or(0);
                    if in_from {
                        from_row = n;
                    } else if in_to {
                        to_row = Some(n);
                    }
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"col" => in_col = false,
                b"row" => in_row = false,
                b"from" => in_from = false,
                b"to" => in_to = false,
                b"pic" | b"sp" => {
                    // End of a picture element – emit the image
                    push_image!();
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => {
                push_image!();
                break;
            }
            _ => {}
        }
    }
}

// ── Color resolution ──────────────────────────────────────────────────────────

fn resolve_color(rgb: Option<&str>, theme: Option<u32>, indexed: Option<u32>) -> Option<String> {
    if let Some(rgb) = rgb {
        let s = rgb.to_uppercase();
        if s.len() == 8 && s != "00000000" {
            return Some(s);
        }
        if s.len() == 6 {
            return Some(format!("FF{}", s));
        }
    }
    if let Some(t) = theme {
        let defaults = [
            "FFFFFFFF", "FF000000", "FFE7E6E6", "FF44546A", "FF4472C4",
            "FFED7D31", "FFA5A5A5", "FFFFC000", "FF4472C4", "FFED7D31",
        ];
        return defaults.get(t as usize).map(|s| s.to_string());
    }
    if let Some(idx) = indexed {
        let map: &[(u32, &str)] = &[
            (0, "FF000000"),
            (1, "FFFFFFFF"),
            (2, "FFFF0000"),
            (3, "FF00FF00"),
            (4, "FF0000FF"),
            (5, "FFFFFF00"),
            (6, "FFFF00FF"),
            (7, "FF00FFFF"),
            (8, "FF000000"),
            (9, "FFFFFFFF"),
            (64, "FF000000"),
            (65, "FFFFFFFF"),
        ];
        return map.iter().find(|(k, _)| *k == idx).map(|(_, v)| v.to_string());
    }
    None
}

// ── Cell reference utilities ──────────────────────────────────────────────────

pub fn cell_ref_to_row_col(cell_ref: &str) -> (u32, u32) {
    let mut col_str = String::new();
    let mut row_str = String::new();
    for c in cell_ref.chars() {
        if c.is_ascii_alphabetic() {
            col_str.push(c.to_ascii_uppercase());
        } else if c.is_ascii_digit() {
            row_str.push(c);
        }
    }
    let col = col_str.chars().fold(0u32, |acc, c| {
        acc * 26 + (c as u32 - 'A' as u32 + 1)
    });
    let row = row_str.parse::<u32>().unwrap_or(0);
    (row, col)
}

pub fn col_num_to_str(mut n: u32) -> String {
    let mut s = String::new();
    while n > 0 {
        let r = (n - 1) % 26;
        s.insert(0, char::from_u32('A' as u32 + r).unwrap_or('A'));
        n = (n - 1) / 26;
    }
    s
}

fn parse_merge_ref(ref_str: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = ref_str.split(':').collect();
    if parts.len() != 2 {
        return None;
    }
    let (r1, c1) = cell_ref_to_row_col(parts[0]);
    let (r2, c2) = cell_ref_to_row_col(parts[1]);
    if r1 > 0 && c1 > 0 {
        Some((r1, c1, r2, c2))
    } else {
        None
    }
}

// ── XML utilities ─────────────────────────────────────────────────────────────

/// Escape special XML characters.
pub fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Extract the value of `attr_name` from the root element of an XML string
/// without fully parsing the document.
fn extract_root_attr(xml: &str, attr_name: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if let Ok(Some(attr)) = e.try_get_attribute(attr_name) {
                    return attr.unescape_value().ok().map(|c| c.into_owned());
                }
                return None; // Only check root element
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
    }
}
