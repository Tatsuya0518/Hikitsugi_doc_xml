import io
import os
import tempfile
import xml.etree.ElementTree as ET

from flask import Flask, render_template, request, jsonify, send_file, Response

from excel_converter import convert_excel_to_xml, is_cache_valid, get_file_hash
from xml_exporter import export_xml_to_excel

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data", "watch")
CACHE_DIR = os.path.join(BASE_DIR, "cache")


# ── Helpers ───────────────────────────────────────────────────────────────────

def _cache_path(xlsx_path: str) -> str:
    """Return the XML cache path that corresponds to an xlsx file."""
    rel = os.path.relpath(xlsx_path, DATA_DIR)
    xml_rel = os.path.splitext(rel)[0] + ".xml"
    return os.path.join(CACHE_DIR, xml_rel)


def _ensure_converted(xlsx_path: str) -> str:
    """Convert xlsx → XML if the cache is missing or stale, then return the path."""
    xml_path = _cache_path(xlsx_path)
    if not is_cache_valid(xlsx_path, xml_path):
        convert_excel_to_xml(xlsx_path, xml_path)
    return xml_path


def _search_in_xml(xml_path: str, query_lower: str) -> list[dict]:
    """Return a list of cell-level matches for *query_lower* inside *xml_path*."""
    from openpyxl.utils import get_column_letter

    matches = []
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        for sheet in root.findall("sheet"):
            sheet_name = sheet.get("name", "")
            for cell in sheet.findall(".//cell"):
                value = cell.get("value", "")
                if value and query_lower in value.lower():
                    row = cell.get("row", "")
                    col = cell.get("col", "")
                    try:
                        cell_ref = f"{get_column_letter(int(col))}{row}"
                    except Exception:
                        cell_ref = f"({row},{col})"
                    matches.append({
                        "sheet": sheet_name,
                        "cell": cell_ref,
                        "value": value[:300],
                    })
    except Exception:
        pass
    return matches


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/structure")
def api_structure():
    """Return the full team → code → document hierarchy as JSON."""
    teams = []
    if os.path.exists(DATA_DIR):
        for team_name in sorted(os.listdir(DATA_DIR)):
            team_path = os.path.join(DATA_DIR, team_name)
            if not os.path.isdir(team_path) or team_name.startswith("."):
                continue
            codes = []
            for code_name in sorted(os.listdir(team_path)):
                code_path = os.path.join(team_path, code_name)
                if not os.path.isdir(code_path) or code_name.startswith("."):
                    continue
                docs = [
                    f for f in sorted(os.listdir(code_path))
                    if f.endswith(".xlsx") and not f.startswith("~$")
                ]
                if docs:
                    codes.append({"name": code_name, "documents": docs})
            if codes:
                teams.append({"name": team_name, "codes": codes})
    return jsonify({"teams": teams})


@app.route("/api/document/<team>/<code>/<filename>")
def api_get_document(team, code, filename):
    """Convert (if needed) and return the document XML."""
    xlsx_path = os.path.join(DATA_DIR, team, code, filename)
    if not os.path.exists(xlsx_path):
        return jsonify({"error": "ファイルが見つかりません"}), 404
    try:
        xml_path = _ensure_converted(xlsx_path)
        with open(xml_path, "r", encoding="utf-8") as fh:
            content = fh.read()
        return Response(content, mimetype="application/xml")
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/document/<team>/<code>/<filename>/save", methods=["POST"])
def api_save_document(team, code, filename):
    """Persist an edited XML back to the cache (does NOT overwrite the xlsx)."""
    xml_path = _cache_path(os.path.join(DATA_DIR, team, code, filename))
    try:
        xml_content = request.get_data(as_text=True)
        # Validate XML before saving
        ET.fromstring(xml_content)
        os.makedirs(os.path.dirname(xml_path), exist_ok=True)
        with open(xml_path, "w", encoding="utf-8") as fh:
            fh.write(xml_content)
        return jsonify({"success": True})
    except ET.ParseError as exc:
        return jsonify({"error": f"XMLの形式が不正です: {exc}"}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/document/<team>/<code>/<filename>/export")
def api_export_document(team, code, filename):
    """Export the (possibly edited) XML cache back to an Excel file for download."""
    xlsx_path = os.path.join(DATA_DIR, team, code, filename)
    xml_path = _cache_path(xlsx_path)

    if not os.path.exists(xml_path):
        if not os.path.exists(xlsx_path):
            return jsonify({"error": "ファイルが見つかりません"}), 404
        _ensure_converted(xlsx_path)

    try:
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        export_xml_to_excel(xml_path, tmp.name)
        return send_file(
            tmp.name,
            as_attachment=True,
            download_name=filename,
            mimetype=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/search")
def api_search():
    """Full-text search across all converted documents."""
    query = request.args.get("q", "").strip()
    if not query:
        return jsonify({"results": [], "query": ""})

    query_lower = query.lower()
    results = []

    if os.path.exists(DATA_DIR):
        for team_name in sorted(os.listdir(DATA_DIR)):
            team_path = os.path.join(DATA_DIR, team_name)
            if not os.path.isdir(team_path) or team_name.startswith("."):
                continue
            for code_name in sorted(os.listdir(team_path)):
                code_path = os.path.join(team_path, code_name)
                if not os.path.isdir(code_path) or code_name.startswith("."):
                    continue
                for doc_name in sorted(os.listdir(code_path)):
                    if not doc_name.endswith(".xlsx") or doc_name.startswith("~$"):
                        continue
                    xlsx_path = os.path.join(code_path, doc_name)
                    try:
                        xml_path = _ensure_converted(xlsx_path)
                        for match in _search_in_xml(xml_path, query_lower):
                            results.append({
                                "team": team_name,
                                "code": code_name,
                                "filename": doc_name,
                                **match,
                            })
                    except Exception:
                        pass

    return jsonify({"results": results, "query": query})


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(CACHE_DIR, exist_ok=True)
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug, host="0.0.0.0", port=5000)
