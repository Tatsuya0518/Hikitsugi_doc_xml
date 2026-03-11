/// Custom XML → Excel (.xlsx) exporter.
///
/// Reads our custom XML format and writes an xlsx file using rust_xlsxwriter,
/// restoring values, formatting, merged cells, and embedded images.
use std::path::Path;

use anyhow::{Context, Result};
use base64::{engine::general_purpose::STANDARD as B64, Engine};
use quick_xml::events::Event;
use quick_xml::Reader;
use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, FormatUnderline, Image, Workbook, Worksheet};

use crate::converter::col_num_to_str;

// ── Public entry-points ───────────────────────────────────────────────────────

pub fn export_xml_to_excel(xml_path: &Path, output_path: &Path) -> Result<()> {
    let bytes = export_xml_to_excel_bytes(xml_path)?;
    std::fs::write(output_path, &bytes)
        .with_context(|| format!("writing {}", output_path.display()))
}

/// Returns xlsx bytes in memory (for HTTP streaming without a temp file).
pub fn export_xml_to_excel_bytes(xml_path: &Path) -> Result<Vec<u8>> {
    let xml_bytes = std::fs::read(xml_path)
        .with_context(|| format!("reading {}", xml_path.display()))?;
    let xml_str = std::str::from_utf8(&xml_bytes)?;
    build_xlsx(xml_str)
}

// ── XML → workbook builder ────────────────────────────────────────────────────

#[derive(Clone, Default)]
struct CellFmt {
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

struct MergeOp {
    r1: u32, c1: u32, r2: u32, c2: u32,
    value: String,
    fmt: Option<CellFmt>,
}

struct CellOp {
    row: u32, col: u32,
    value: String,
    cell_type: String,
    fmt: Option<CellFmt>,
}

struct SheetOps {
    name: String,
    col_ops: Vec<(u32, f64, bool)>,
    row_ops: Vec<(u32, f64, bool)>,
    merges: Vec<MergeOp>,
    cells: Vec<CellOp>,
    image_ops: Vec<(u32, u32, Vec<u8>)>,
}

#[derive(Debug, PartialEq, Clone, Copy)]
enum Sec { None, ColWidths, RowHeights, Merges, Rows, RowsRow, Images }

fn build_xlsx(xml_str: &str) -> Result<Vec<u8>> {
    let mut reader = Reader::from_str(xml_str);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut sec = Sec::None;

    // Pending cell currently being assembled
    #[derive(Default, Clone)]
    struct PC {
        row: u32, col: u32,
        value: String,
        cell_type: String,
        rowspan: u32, colspan: u32,
        fmt: Option<CellFmt>,
    }
    let mut pc: Option<PC> = None;

    let mut sheets: Vec<SheetOps> = Vec::new();
    let mut cur: Option<SheetOps> = None;

    macro_rules! cur_sheet {
        () => { cur.as_mut().expect("sheet") }
    }

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                macro_rules! attr {
                    ($n:expr) => {
                        e.try_get_attribute($n).ok().flatten()
                            .and_then(|a| a.unescape_value().ok())
                            .map(|c| c.into_owned())
                    };
                }

                match e.local_name().as_ref() {
                    // ── Sheet start ──────────────────────────────────────────
                    b"sheet" if cur.is_none() && attr!("ref").is_none() => {
                        let name = attr!("name").unwrap_or_else(|| "Sheet".into());
                        cur = Some(SheetOps {
                            name, col_ops: vec![], row_ops: vec![],
                            merges: vec![], cells: vec![], image_ops: vec![],
                        });
                        sec = Sec::None;
                    }

                    // ── Section headers ──────────────────────────────────────
                    b"colWidths"  => sec = Sec::ColWidths,
                    b"rowHeights" => sec = Sec::RowHeights,
                    b"merges"    => sec = Sec::Merges,
                    b"rows"      => sec = Sec::Rows,
                    b"images"    => sec = Sec::Images,

                    // ── Column widths ────────────────────────────────────────
                    b"col" if sec == Sec::ColWidths => {
                        let idx: u32 = attr!("index").and_then(|s| s.parse().ok()).unwrap_or(1);
                        let w: f64   = attr!("width").and_then(|s| s.parse().ok()).unwrap_or(8.43);
                        let hid = attr!("hidden").map_or(false, |v| v == "true");
                        cur_sheet!().col_ops.push((idx.saturating_sub(1), w, hid));
                    }

                    // ── Row heights ──────────────────────────────────────────
                    b"row" if sec == Sec::RowHeights => {
                        let idx: u32 = attr!("index").and_then(|s| s.parse().ok()).unwrap_or(1);
                        let h: f64   = attr!("height").and_then(|s| s.parse().ok()).unwrap_or(15.0);
                        let hid = attr!("hidden").map_or(false, |v| v == "true");
                        cur_sheet!().row_ops.push((idx.saturating_sub(1), h, hid));
                    }

                    // ── Row inside <rows> ────────────────────────────────────
                    b"row" if sec == Sec::Rows => {
                        sec = Sec::RowsRow;
                    }

                    // ── Merge ranges ─────────────────────────────────────────
                    b"merge" if sec == Sec::Merges => {
                        let r1: u32 = attr!("startRow").and_then(|s| s.parse().ok()).unwrap_or(1);
                        let c1: u32 = attr!("startCol").and_then(|s| s.parse().ok()).unwrap_or(1);
                        let r2: u32 = attr!("endRow").and_then(|s| s.parse().ok()).unwrap_or(r1);
                        let c2: u32 = attr!("endCol").and_then(|s| s.parse().ok()).unwrap_or(c1);
                        cur_sheet!().merges.push(MergeOp {
                            r1, c1, r2, c2, value: String::new(), fmt: None,
                        });
                    }

                    // ── Cell ─────────────────────────────────────────────────
                    b"cell" if sec == Sec::RowsRow => {
                        let row: u32    = attr!("row").and_then(|s| s.parse().ok()).unwrap_or(1);
                        let col: u32    = attr!("col").and_then(|s| s.parse().ok()).unwrap_or(1);
                        let value       = attr!("value").unwrap_or_default();
                        let cell_type   = attr!("type").unwrap_or_else(|| "s".into());
                        let rowspan: u32 = attr!("rowspan").and_then(|s| s.parse().ok()).unwrap_or(1);
                        let colspan: u32 = attr!("colspan").and_then(|s| s.parse().ok()).unwrap_or(1);
                        pc = Some(PC { row, col, value, cell_type, rowspan, colspan, fmt: None });
                    }

                    // ── Format (nested inside <cell>) ─────────────────────────
                    b"format" => {
                        if let Some(ref mut p) = pc {
                            let font_size = attr!("fontSize")
                                .and_then(|s| s.parse::<f64>().ok()).unwrap_or(11.0);
                            p.fmt = Some(CellFmt {
                                bold:          attr!("bold").map_or(false, |v| v == "true"),
                                italic:        attr!("italic").map_or(false, |v| v == "true"),
                                underline:     attr!("underline").map_or(false, |v| v == "true"),
                                font_size,
                                font_name:     attr!("fontName").unwrap_or_else(|| "Calibri".into()),
                                font_color:    attr!("fontColor").unwrap_or_else(|| "FF000000".into()),
                                bg_color:      attr!("bgColor").filter(|s| !s.is_empty()),
                                halign:        attr!("halign").unwrap_or_default(),
                                valign:        attr!("valign").unwrap_or_default(),
                                wrap_text:     attr!("wrapText").map_or(false, |v| v == "true"),
                                indent:        attr!("indent").and_then(|s| s.parse().ok()).unwrap_or(0),
                                border_top:    attr!("borderTop").unwrap_or_else(|| "none".into()),
                                border_bottom: attr!("borderBottom").unwrap_or_else(|| "none".into()),
                                border_left:   attr!("borderLeft").unwrap_or_else(|| "none".into()),
                                border_right:  attr!("borderRight").unwrap_or_else(|| "none".into()),
                                num_fmt:       attr!("numFmt").unwrap_or_else(|| "General".into()),
                            });
                        }
                    }

                    // ── Images ───────────────────────────────────────────────
                    b"image" if sec == Sec::Images => {
                        let row: u32 = attr!("row").and_then(|s| s.parse().ok()).unwrap_or(1);
                        let col: u32 = attr!("col").and_then(|s| s.parse().ok()).unwrap_or(1);
                        let b64 = attr!("data").unwrap_or_default();
                        if !b64.is_empty() {
                            match B64.decode(&b64) {
                                Ok(bytes) => {
                                    cur_sheet!().image_ops.push((
                                        row.saturating_sub(1),
                                        col.saturating_sub(1),
                                        bytes,
                                    ));
                                }
                                Err(e) => {
                                    // Log decode failure and continue; image is skipped
                                    eprintln!(
                                        "warning: image base64 decode failed at row={} col={}: {}",
                                        row, col, e
                                    );
                                }
                            }
                        }
                    }

                    _ => {}
                }
            }

            Ok(Event::End(ref e)) => {
                match e.local_name().as_ref() {
                    b"cell" => {
                        if let (Some(p), Some(ref mut sh)) = (pc.take(), cur.as_mut()) {
                            if p.rowspan > 1 || p.colspan > 1 {
                                // Fill in value/format for the matching merge record
                                for m in sh.merges.iter_mut() {
                                    if m.r1 == p.row && m.c1 == p.col {
                                        m.value = p.value.clone();
                                        m.fmt = p.fmt.clone();
                                        break;
                                    }
                                }
                            } else {
                                sh.cells.push(CellOp {
                                    row: p.row, col: p.col,
                                    value: p.value.clone(),
                                    cell_type: p.cell_type.clone(),
                                    fmt: p.fmt.clone(),
                                });
                            }
                        }
                    }
                    b"row" if sec == Sec::RowsRow => sec = Sec::Rows,
                    b"colWidths" | b"rowHeights" | b"merges" | b"rows" | b"images" => {
                        sec = Sec::None;
                    }
                    b"sheet" => {
                        if let Some(sh) = cur.take() {
                            sheets.push(sh);
                        }
                        sec = Sec::None;
                    }
                    _ => {}
                }
            }

            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    // ── Apply operations to workbook ─────────────────────────────────────────
    let mut workbook = Workbook::new();

    for sh in sheets {
        let ws = workbook.add_worksheet();
        let _ = ws.set_name(&sh.name);

        for (col_0, width, hidden) in sh.col_ops {
            let _ = ws.set_column_width(col_0 as u16, width);
            if hidden { let _ = ws.set_column_hidden(col_0 as u16); }
        }
        for (row_0, height, hidden) in sh.row_ops {
            let _ = ws.set_row_height(row_0, height);
            if hidden { let _ = ws.set_row_hidden(row_0); }
        }

        // Plain (non-merged) cells
        for op in &sh.cells {
            let r = op.row.saturating_sub(1);
            let c = op.col.saturating_sub(1) as u16;
            let fmt = op.fmt.as_ref().map(mk_fmt).unwrap_or_default();
            write_cell(ws, r, c, &op.value, &op.cell_type, &fmt);
        }

        // Merged cells
        for m in &sh.merges {
            let r1 = m.r1.saturating_sub(1);
            let c1 = m.c1.saturating_sub(1) as u16;
            let r2 = m.r2.saturating_sub(1);
            let c2 = m.c2.saturating_sub(1) as u16;
            let fmt = m.fmt.as_ref().map(mk_fmt).unwrap_or_default();
            if r1 == r2 && c1 == c2 {
                write_cell(ws, r1, c1, &m.value, "s", &fmt);
            } else if m.value.is_empty() {
                let _ = ws.merge_range(r1, c1, r2, c2, "", &fmt);
            } else if let Ok(_n) = m.value.parse::<f64>() {
                let _ = ws.merge_range(r1, c1, r2, c2, &m.value, &fmt);
            } else {
                let _ = ws.merge_range(r1, c1, r2, c2, m.value.as_str(), &fmt);
            }
        }

        // Images
        for (row_0, col_0, img_bytes) in &sh.image_ops {
            if let Ok(img) = Image::new_from_buffer(img_bytes) {
                let _ = ws.insert_image(*row_0, *col_0 as u16, &img);
            }
        }
    }

    workbook.save_to_buffer()
        .map_err(|e| anyhow::anyhow!("xlsx save error: {}", e))
}

// ── Format builder ────────────────────────────────────────────────────────────

fn mk_fmt(cf: &CellFmt) -> Format {
    let mut fmt = Format::new();
    if cf.bold    { fmt = fmt.set_bold(); }
    if cf.italic  { fmt = fmt.set_italic(); }
    if cf.underline { fmt = fmt.set_underline(FormatUnderline::Single); }
    if cf.font_size > 0.0 { fmt = fmt.set_font_size(cf.font_size); }
    if !cf.font_name.is_empty() { fmt = fmt.set_font_name(&cf.font_name); }
    if let Some(c) = argb_u32(&cf.font_color) { fmt = fmt.set_font_color(Color::RGB(c)); }
    if let Some(ref bg) = cf.bg_color {
        if let Some(c) = argb_u32(bg) { fmt = fmt.set_background_color(Color::RGB(c)); }
    }
    if let Some(h) = halign(&cf.halign) { fmt = fmt.set_align(h); }
    if let Some(v) = valign(&cf.valign) { fmt = fmt.set_align(v); }
    if cf.wrap_text { fmt = fmt.set_text_wrap(); }
    if cf.indent > 0 { fmt = fmt.set_indent(cf.indent as u8); }
    if let Some(b) = border_style(&cf.border_top)    { fmt = fmt.set_border_top(b); }
    if let Some(b) = border_style(&cf.border_bottom) { fmt = fmt.set_border_bottom(b); }
    if let Some(b) = border_style(&cf.border_left)   { fmt = fmt.set_border_left(b); }
    if let Some(b) = border_style(&cf.border_right)  { fmt = fmt.set_border_right(b); }
    if !cf.num_fmt.is_empty() && cf.num_fmt != "General" {
        fmt = fmt.set_num_format(&cf.num_fmt);
    }
    fmt
}

fn write_cell(ws: &mut Worksheet, row: u32, col: u16, value: &str, ctype: &str, fmt: &Format) {
    if value.is_empty() {
        let _ = ws.write_blank(row, col, fmt);
        return;
    }
    match ctype {
        "n" => {
            if let Ok(n) = value.parse::<f64>() {
                let _ = ws.write_number_with_format(row, col, n, fmt);
            } else {
                let _ = ws.write_string_with_format(row, col, value, fmt);
            }
        }
        "b" => {
            let b = value.to_lowercase() == "true" || value == "1" || value.to_uppercase() == "TRUE";
            let _ = ws.write_boolean_with_format(row, col, b, fmt);
        }
        _ => {
            let _ = ws.write_string_with_format(row, col, value, fmt);
        }
    }
}

fn argb_u32(argb: &str) -> Option<u32> {
    let s = argb.trim_start_matches('#');
    let s = if s.len() == 8 { &s[2..] } else { s };
    if s.len() == 6 { u32::from_str_radix(s, 16).ok() } else { None }
}

fn halign(s: &str) -> Option<FormatAlign> {
    match s {
        "center"           => Some(FormatAlign::Center),
        "left"             => Some(FormatAlign::Left),
        "right"            => Some(FormatAlign::Right),
        "justify"          => Some(FormatAlign::Justify),
        "fill"             => Some(FormatAlign::Fill),
        "centerContinuous" => Some(FormatAlign::CenterAcross),
        _                  => None,
    }
}

fn valign(s: &str) -> Option<FormatAlign> {
    match s {
        "center"  => Some(FormatAlign::VerticalCenter),
        "top"     => Some(FormatAlign::Top),
        "bottom"  => Some(FormatAlign::Bottom),
        "justify" => Some(FormatAlign::VerticalJustify),
        _         => None,
    }
}

fn border_style(s: &str) -> Option<FormatBorder> {
    match s {
        "thin"               => Some(FormatBorder::Thin),
        "medium"             => Some(FormatBorder::Medium),
        "thick"              => Some(FormatBorder::Thick),
        "dashed"             => Some(FormatBorder::Dashed),
        "dotted"             => Some(FormatBorder::Dotted),
        "double"             => Some(FormatBorder::Double),
        "hair"               => Some(FormatBorder::Hair),
        "mediumDashed"       => Some(FormatBorder::MediumDashed),
        "dashDot"            => Some(FormatBorder::DashDot),
        "mediumDashDot"      => Some(FormatBorder::MediumDashDot),
        "dashDotDot"         => Some(FormatBorder::DashDotDot),
        "mediumDashDotDot"   => Some(FormatBorder::MediumDashDotDot),
        "slantDashDot"       => Some(FormatBorder::SlantDashDot),
        _                    => None,
    }
}
