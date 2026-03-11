import os
import hashlib
import base64
from datetime import datetime
import xml.etree.ElementTree as ET
from xml.dom import minidom

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string


def get_file_hash(path: str) -> str:
    """Calculate MD5 hash of a file for cache validation."""
    md5 = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            md5.update(chunk)
    return md5.hexdigest()


def is_cache_valid(xlsx_path: str, xml_path: str) -> bool:
    """Check whether the XML cache is still valid for the given xlsx file."""
    if not os.path.exists(xml_path):
        return False
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        cached_hash = root.get("hash", "")
        current_hash = get_file_hash(xlsx_path)
        return cached_hash == current_hash
    except Exception:
        return False


def _extract_color(color_obj) -> str | None:
    """Return an 8-char AARRGGBB hex string from an openpyxl Color, or None."""
    if color_obj is None:
        return None
    try:
        ctype = color_obj.type
        if ctype == "rgb":
            rgb = color_obj.rgb
            if rgb and rgb not in ("00000000", "000000"):
                if len(rgb) == 6:
                    return "FF" + rgb
                return rgb
        elif ctype == "theme":
            theme_defaults = {
                0: "FFFFFFFF", 1: "FF000000", 2: "FFE7E6E6", 3: "FF44546A",
                4: "FF4472C4", 5: "FFED7D31", 6: "FFA5A5A5", 7: "FFFFC000",
                8: "FF4472C4", 9: "FFED7D31", 10: "FF4472C4",
            }
            return theme_defaults.get(color_obj.theme, "FF000000")
        elif ctype == "indexed":
            indexed_defaults = {
                0: "FF000000", 1: "FFFFFFFF", 2: "FFFF0000", 3: "FF00FF00",
                4: "FF0000FF", 5: "FFFFFF00", 6: "FFFF00FF", 7: "FF00FFFF",
                8: "FF000000", 9: "FFFFFFFF", 64: "FF000000", 65: "FFFFFFFF",
            }
            return indexed_defaults.get(color_obj.indexed, None)
    except Exception:
        pass
    return None


def _border_style(side) -> str:
    if side is None:
        return "none"
    return side.border_style or "none"


def _safe(value) -> str:
    if value is None:
        return ""
    return str(value)


def convert_excel_to_xml(xlsx_path: str, xml_path: str) -> str:
    """
    Convert an Excel (.xlsx) file to a custom XML format that preserves
    cell values, formatting, merged cells, and embedded images.
    """
    file_hash = get_file_hash(xlsx_path)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    root = ET.Element("workbook")
    root.set("filename", os.path.basename(xlsx_path))
    root.set("converted", datetime.now().isoformat())
    root.set("hash", file_hash)

    for ws in wb.worksheets:
        sheet_elem = ET.SubElement(root, "sheet")
        sheet_elem.set("name", ws.title)

        # ── Dimensions ──────────────────────────────────────────────────────────
        dim = ET.SubElement(sheet_elem, "dimensions")
        dim.set("maxRow", str(ws.max_row or 0))
        dim.set("maxCol", str(ws.max_column or 0))

        # ── Column widths ────────────────────────────────────────────────────────
        cols_elem = ET.SubElement(sheet_elem, "colWidths")
        for col_letter, col_dim in ws.column_dimensions.items():
            col_el = ET.SubElement(cols_elem, "col")
            col_el.set("index", str(column_index_from_string(col_letter)))
            col_el.set("width", str(col_dim.width or 8.43))
            col_el.set("hidden", "true" if col_dim.hidden else "false")

        # ── Row heights ──────────────────────────────────────────────────────────
        rows_h_elem = ET.SubElement(sheet_elem, "rowHeights")
        for row_idx, row_dim in ws.row_dimensions.items():
            row_el = ET.SubElement(rows_h_elem, "row")
            row_el.set("index", str(row_idx))
            row_el.set("height", str(row_dim.height or 15))
            row_el.set("hidden", "true" if row_dim.hidden else "false")

        # ── Merged cells ─────────────────────────────────────────────────────────
        merges_elem = ET.SubElement(sheet_elem, "merges")
        merge_info: dict[tuple[int, int], tuple[int, int, int, int]] = {}

        for mr in ws.merged_cells.ranges:
            merge_el = ET.SubElement(merges_elem, "merge")
            merge_el.set("ref", str(mr))
            merge_el.set("startRow", str(mr.min_row))
            merge_el.set("startCol", str(mr.min_col))
            merge_el.set("endRow", str(mr.max_row))
            merge_el.set("endCol", str(mr.max_col))
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    merge_info[(r, c)] = (mr.min_row, mr.min_col,
                                          mr.max_row, mr.max_col)

        # ── Rows & cells ─────────────────────────────────────────────────────────
        rows_elem = ET.SubElement(sheet_elem, "rows")
        max_r = ws.max_row or 0
        max_c = ws.max_column or 0

        for row_cells in ws.iter_rows(min_row=1, max_row=max_r,
                                      min_col=1, max_col=max_c):
            row_idx = row_cells[0].row

            # Skip rows with no content and no merge involvement
            has_content = any(
                cell.value is not None
                or cell.has_style
                or (cell.row, cell.column) in merge_info
                for cell in row_cells
            )
            if not has_content:
                continue

            row_elem = ET.SubElement(rows_elem, "row")
            row_elem.set("index", str(row_idx))

            for cell in row_cells:
                pos = (cell.row, cell.column)
                info = merge_info.get(pos)

                # Skip non-origin cells within a merge
                if info and not (cell.row == info[0] and cell.column == info[1]):
                    continue
                # Also skip cells that have no value, no style, and are not
                # merge origins
                if (cell.value is None and not cell.has_style and info is None):
                    continue

                cell_elem = ET.SubElement(row_elem, "cell")
                cell_elem.set("row", str(cell.row))
                cell_elem.set("col", str(cell.column))
                cell_elem.set("value", _safe(cell.value))
                cell_elem.set("type", cell.data_type or "s")

                # Merge span attributes for origin cells
                if info and cell.row == info[0] and cell.column == info[1]:
                    rowspan = info[2] - info[0] + 1
                    colspan = info[3] - info[1] + 1
                    if rowspan > 1:
                        cell_elem.set("rowspan", str(rowspan))
                    if colspan > 1:
                        cell_elem.set("colspan", str(colspan))

                # Style / formatting
                if cell.has_style:
                    fmt = ET.SubElement(cell_elem, "format")

                    f = cell.font
                    if f:
                        fmt.set("bold", "true" if f.bold else "false")
                        fmt.set("italic", "true" if f.italic else "false")
                        fmt.set("underline",
                                "true" if (f.underline and f.underline != "none")
                                else "false")
                        fmt.set("fontSize", str(f.size) if f.size else "11")
                        fmt.set("fontName", f.name or "Calibri")
                        fc = _extract_color(f.color)
                        fmt.set("fontColor", fc or "FF000000")

                    fill = cell.fill
                    if fill and fill.fill_type not in (None, "none"):
                        fgc = _extract_color(fill.fgColor) if fill.fgColor else None
                        if fgc and fgc not in ("FF000000", "00000000", "FFFFFFFF"):
                            fmt.set("bgColor", fgc)

                    align = cell.alignment
                    if align:
                        fmt.set("halign", align.horizontal or "")
                        fmt.set("valign", align.vertical or "")
                        fmt.set("wrapText", "true" if align.wrap_text else "false")
                        fmt.set("indent", str(align.indent or 0))

                    border = cell.border
                    if border:
                        fmt.set("borderTop",    _border_style(border.top))
                        fmt.set("borderBottom", _border_style(border.bottom))
                        fmt.set("borderLeft",   _border_style(border.left))
                        fmt.set("borderRight",  _border_style(border.right))

                    fmt.set("numFmt", cell.number_format or "General")

        # ── Embedded images ──────────────────────────────────────────────────────
        images_elem = ET.SubElement(sheet_elem, "images")
        try:
            for image in ws._images:
                img_elem = ET.SubElement(images_elem, "image")

                # Anchor position
                try:
                    anchor = image.anchor
                    if hasattr(anchor, "_from"):
                        img_elem.set("row", str(anchor._from.row + 1))
                        img_elem.set("col", str(anchor._from.col + 1))
                        if hasattr(anchor, "_to"):
                            img_elem.set("toRow", str(anchor._to.row + 1))
                            img_elem.set("toCol", str(anchor._to.col + 1))
                    elif hasattr(anchor, "row"):
                        img_elem.set("row", str(anchor.row + 1))
                        img_elem.set("col", str(anchor.col + 1))
                    else:
                        img_elem.set("row", "1")
                        img_elem.set("col", "1")
                except Exception:
                    img_elem.set("row", "1")
                    img_elem.set("col", "1")

                # Raw image bytes → base64
                try:
                    img_data = image._data()
                    img_elem.set("data",
                                 base64.b64encode(img_data).decode("utf-8"))
                    fmt_name = (getattr(image, "format", None) or "png").lower()
                    img_elem.set("format", fmt_name)
                except Exception as exc:
                    img_elem.set("error", str(exc))

                # Pixel dimensions
                try:
                    if getattr(image, "width", None):
                        img_elem.set("width", str(image.width))
                    if getattr(image, "height", None):
                        img_elem.set("height", str(image.height))
                except Exception:
                    pass
        except Exception:
            pass  # sheets without images raise AttributeError on ._images

    # ── Serialize to pretty-printed XML ─────────────────────────────────────────
    xml_str = ET.tostring(root, encoding="unicode")
    dom = minidom.parseString(xml_str)
    pretty = dom.toprettyxml(indent="  ", encoding="UTF-8").decode("utf-8")

    os.makedirs(os.path.dirname(xml_path), exist_ok=True)
    with open(xml_path, "w", encoding="utf-8") as fh:
        fh.write(pretty)

    return xml_path
