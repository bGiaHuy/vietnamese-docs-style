"""Independent renderer (LibreOffice / optional Word COM) and PyMuPDF geometry inspector."""

from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

import pymupdf

from eval_harness.ooxml_reader import PT_PER_CM, nfc

_SIG_LINE_RE = re.compile(
    r"^(nơi nhận|lưu:|-\s*như trên|-\s*lưu|kt\.|tm\.|t/m|tl\.|q\.|"
    r"phó\s+|trưởng\s+|giám đốc|cục trưởng|chủ tịch|hiệu trưởng|"
    r"chánh văn phòng|bí thư|phó bí thư|thừa lệnh|kính gửi).*$",
    re.IGNORECASE,
)
_PAGE_NUM_RE = re.compile(r"^[-–—\s]*(trang\s*)?\d+(\s*/\s*\d+)?[-–—\s]*$", re.IGNORECASE)


def find_soffice() -> str | None:
    """Locate LibreOffice headless binary."""
    hit = shutil.which("soffice") or shutil.which("libreoffice")
    if hit:
        return hit
    win_candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.COM"),
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for cand in win_candidates:
        if cand.is_file():
            return str(cand)
    return None


def is_word_com_available() -> bool:
    """Check whether Microsoft Word COM automation is registered on Windows."""
    try:
        import winreg
    except ImportError:
        return False
    for subkey in (r"Word.Application", r"SOFTWARE\Classes\Word.Application"):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey):
                return True
        except OSError:
            pass
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"):
            return True
    except OSError:
        return False


def render_to_pdf(
    docx_path: str | Path,
    out_dir: str | Path,
    *,
    backend: str = "libreoffice",
    timeout_sec: int = 90,
) -> dict[str, Any]:
    """Render DOCX to PDF independently of repo render scripts."""
    src = Path(docx_path).resolve()
    out = Path(out_dir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    if backend == "auto":
        backend = "libreoffice" if find_soffice() else ("word_com" if is_word_com_available() else "none")

    if backend == "word_com":
        if not is_word_com_available():
            return {"status": "unavailable", "renderer_backend": "word_com", "reason": "Microsoft Word COM is not installed"}
        try:
            import win32com.client  # type: ignore
            pdf_out = out / f"{src.stem}.pdf"
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                doc = word.Documents.Open(str(src), ReadOnly=True)
                doc.ExportAsFixedFormat(str(pdf_out), 17)
                doc.Close(False)
            finally:
                word.Quit()
            return {"status": "rendered", "renderer_backend": "word_com", "pdf_path": str(pdf_out)}
        except Exception as exc:
            return {"status": "error", "renderer_backend": "word_com", "reason": str(exc)}

    if backend != "libreoffice":
        return {"status": "unavailable", "renderer_backend": backend, "reason": f"Unsupported or unavailable renderer {backend!r}"}

    soffice = find_soffice()
    if not soffice:
        return {"status": "unavailable", "renderer_backend": "libreoffice", "reason": "soffice binary not found"}

    with tempfile.TemporaryDirectory(prefix="eval-lo-prof-") as prof_dir:
        prof_uri = Path(prof_dir).resolve().as_uri()
        cmd = [
            soffice,
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation={prof_uri}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out),
            str(src),
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout_sec, check=False)
        except subprocess.TimeoutExpired:
            return {"status": "error", "renderer_backend": "libreoffice", "reason": f"soffice timed out after {timeout_sec}s"}

        expected_pdf = out / f"{src.stem}.pdf"
        if proc.returncode != 0 or not expected_pdf.is_file():
            return {
                "status": "error",
                "renderer_backend": "libreoffice",
                "reason": (proc.stderr or proc.stdout or f"exit={proc.returncode}").strip()[:500],
            }
        return {"status": "rendered", "renderer_backend": "libreoffice", "pdf_path": str(expected_pdf)}


def _tokenize_words(text: str) -> set[str]:
    clean = re.sub(r"[^\w]+", " ", nfc(text).lower())
    return {tok for tok in clean.split() if len(tok) >= 2}


def _collect_table_candidates(pdf_doc: pymupdf.Document) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for p_idx, page in enumerate(pdf_doc, start=1):
        try:
            found = page.find_tables()
            tbl_objs = list(getattr(found, "tables", []) or [])
        except Exception:
            tbl_objs = []
        for tb in tbl_objs:
            try:
                rows_txt = tb.extract() or []
            except Exception:
                rows_txt = []
            flat_txt = " ".join(str(c or "") for r in rows_txt for c in r)
            candidates.append({
                "page": p_idx,
                "bbox": tuple(float(v) for v in tb.bbox),
                "tokens": _tokenize_words(flat_txt),
                "text_nfc": nfc(flat_txt),
                "rows": len(rows_txt),
            })
    return candidates


def _inside_any_bbox(word_box: tuple[float, float, float, float], boxes: list[tuple[float, float, float, float]]) -> bool:
    wx0, wy0, wx1, wy1 = word_box
    cx = (wx0 + wx1) / 2.0
    cy = (wy0 + wy1) / 2.0
    return any(bx0 - 2.0 <= cx <= bx1 + 2.0 and by0 - 2.0 <= cy <= by1 + 2.0 for bx0, by0, bx1, by1 in boxes)


def _fallback_word_bbox(
    pdf_doc: pymupdf.Document,
    row_texts: list[list[str]],
    start_page: int,
    start_y: float,
    *,
    excluded_bboxes_by_page: dict[int, list[tuple[float, float, float, float]]] | None = None,
) -> tuple[int, tuple[float, float, float, float]] | None:
    """Locate bounding box of a borderless or partial-border table via clustered page words."""
    probe_phrases = [
        nfc(cell).strip().lower()
        for row in row_texts[:3]
        for cell in row
        if nfc(cell).strip()
    ]
    probe_tokens = {tok for ph in probe_phrases for tok in _tokenize_words(ph)}
    if not probe_tokens:
        return None

    best_score: tuple[int, int] = (0, 0)
    best_hit: tuple[int, tuple[float, float, float, float]] | None = None
    for p_idx in range(start_page, len(pdf_doc) + 1):
        page = pdf_doc[p_idx - 1]
        words = page.get_text("words") or []
        excl = (excluded_bboxes_by_page or {}).get(p_idx, [])
        matched: list[tuple[float, float, float, float, set[str]]] = []
        for x0, y0, x1, y1, w_txt, *_ in words:
            if p_idx == start_page and y0 < start_y - 2.0:
                continue
            if excl and _inside_any_bbox((x0, y0, x1, y1), excl):
                continue
            hit_toks = _tokenize_words(w_txt) & probe_tokens
            if hit_toks:
                matched.append((x0, y0, x1, y1, hit_toks))
        if not matched:
            continue

        matched.sort(key=lambda item: (item[1], item[0]))
        clusters: list[list[tuple[float, float, float, float, set[str]]]] = []
        cur_cluster: list[tuple[float, float, float, float, set[str]]] = [matched[0]]
        cur_max_y1 = matched[0][3]
        for item in matched[1:]:
            if item[1] - cur_max_y1 > 28.0:
                clusters.append(cur_cluster)
                cur_cluster = [item]
                cur_max_y1 = item[3]
            else:
                cur_cluster.append(item)
                cur_max_y1 = max(cur_max_y1, item[3])
        clusters.append(cur_cluster)

        for cl in clusters:
            if len(cl) < min(2, len(probe_tokens)):
                continue
            uniq_toks: set[str] = set()
            for _, _, _, _, toks in cl:
                uniq_toks |= toks
            score = (len(uniq_toks), len(cl))
            if score > best_score:
                best_score = score
                bbox = (
                    min(m[0] for m in cl),
                    min(m[1] for m in cl),
                    max(m[2] for m in cl),
                    max(m[3] for m in cl),
                )
                best_hit = (p_idx, bbox)
                if len(uniq_toks) >= len(probe_tokens):
                    return best_hit
    return best_hit


def _has_visible_borders(borders: dict[str, Any] | None) -> bool:
    if not borders:
        return False
    return any((b or {}).get("val") not in (None, "none", "nil") for b in borders.values())


def _match_tables_in_pdf(pdf_doc: pymupdf.Document, docx_info: dict[str, Any]) -> list[dict[str, Any]]:
    candidates = _collect_table_candidates(pdf_doc)
    excl_by_page: dict[int, list[tuple[float, float, float, float]]] = {}
    for cand in candidates:
        excl_by_page.setdefault(cand["page"], []).append(cand["bbox"])
    sections = {s["section_index"]: s for s in docx_info.get("sections", [])}
    results: list[dict[str, Any]] = []
    used_indices: set[int] = set()
    cursor_page = 1
    cursor_y = 0.0

    for tbl in docx_info.get("tables", []):
        row_texts = tbl.get("row_texts_nfc", [])
        all_tokens = _tokenize_words(" ".join(c for r in row_texts for c in r))
        head_tokens = _tokenize_words(" ".join(c for r in row_texts[:2] for c in r))
        is_borderless_layout = bool(tbl.get("is_layout_table")) and not _has_visible_borders(tbl.get("borders"))
        best_idx: int | None = None
        best_score = -1.0

        for idx, cand in enumerate(candidates):
            if idx in used_indices:
                continue
            if cand["page"] < cursor_page or (cand["page"] == cursor_page and cand["bbox"][1] < cursor_y - 4.0):
                continue
            if tbl["rows"] <= 2 and cand["rows"] > 4:
                continue
            if is_borderless_layout and abs(cand["rows"] - tbl["rows"]) > 2:
                continue
            if not all_tokens:
                overlap = 1.0 if cand["rows"] == tbl["rows"] else 0.0
            else:
                inter_all = len(all_tokens & cand["tokens"])
                inter_head = len(head_tokens & cand["tokens"])
                if inter_all == 0:
                    continue
                if is_borderless_layout and head_tokens and inter_head < min(2, len(head_tokens)):
                    continue
                overlap = (inter_all / max(1, len(all_tokens))) + 0.5 * (inter_head / max(1, len(head_tokens)))
            if overlap > best_score:
                best_score = overlap
                best_idx = idx

        sec = sections.get(tbl["section_index"], next(iter(sections.values())))
        printable_right_cm = sec["page_width_cm"] - sec["margins_cm"]["right"]

        min_req = 0.55 if is_borderless_layout else 0.35
        if best_idx is not None and best_score >= min_req:
            used_indices.add(best_idx)
            first = candidates[best_idx]
            spanned_pages = [first["page"]]
            # Detect continuation slices of multi-page tables on subsequent pages.
            tail_tokens = _tokenize_words(" ".join(c for r in row_texts[-3:] for c in r))
            if len(row_texts) > 6 and first["rows"] < tbl["rows"]:
                last_page = first["page"]
                for idx2 in range(best_idx + 1, len(candidates)):
                    c2 = candidates[idx2]
                    if c2["page"] == last_page + 1 and (c2["tokens"] & all_tokens):
                        used_indices.add(idx2)
                        spanned_pages.append(c2["page"])
                        last_page = c2["page"]
                        if tail_tokens and (c2["tokens"] & tail_tokens):
                            break

            x0, y0, x1, y1 = first["bbox"]
            cursor_page = spanned_pages[-1]
            cursor_y = y1 if len(spanned_pages) == 1 else 0.0
            x0_cm = round(x0 / PT_PER_CM, 3)
            x1_cm = round(x1 / PT_PER_CM, 3)
            w_cm = round((x1 - x0) / PT_PER_CM, 3)
            results.append({
                "ordinal": tbl["ordinal"],
                "matched": True,
                "match_method": "find_tables",
                "pdf_pages": spanned_pages,
                "bbox_pt": [round(x0, 2), round(y0, 2), round(x1, 2), round(y1, 2)],
                "x0_cm": x0_cm,
                "x1_cm": x1_cm,
                "width_cm": w_cm,
                "center_cm": round((x0_cm + x1_cm) / 2.0, 3),
                "overflow_right_mm": round(max(0.0, (x1_cm - printable_right_cm) * 10.0), 2),
            })
        else:
            fb = _fallback_word_bbox(
                pdf_doc,
                row_texts,
                cursor_page,
                cursor_y,
                excluded_bboxes_by_page=excl_by_page if tbl.get("is_layout_table") else None,
            )
            if fb is not None:
                p_num, (x0, y0, x1, y1) = fb
                cursor_page = p_num
                cursor_y = y1
                x0_cm = round(x0 / PT_PER_CM, 3)
                x1_cm = round(x1 / PT_PER_CM, 3)
                results.append({
                    "ordinal": tbl["ordinal"],
                    "matched": True,
                    "match_method": "word_bbox_fallback",
                    "pdf_pages": [p_num],
                    "bbox_pt": [round(x0, 2), round(y0, 2), round(x1, 2), round(y1, 2)],
                    "x0_cm": x0_cm,
                    "x1_cm": x1_cm,
                    "width_cm": round((x1 - x0) / PT_PER_CM, 3),
                    "center_cm": round((x0_cm + x1_cm) / 2.0, 3),
                    "overflow_right_mm": round(max(0.0, (x1_cm - printable_right_cm) * 10.0), 2),
                })
            else:
                results.append({"ordinal": tbl["ordinal"], "matched": False})

    return results


def _collect_vector_drawing_rects(pdf_doc: pymupdf.Document) -> list[dict[str, Any]]:
    """Collect large standalone vector shape boxes (e.g. wps:wsp / wpg:wgp rendered as vector paths)."""
    vec_rects: list[dict[str, Any]] = []
    for p_idx, page in enumerate(pdf_doc, start=1):
        seen_bboxes: set[tuple[int, int, int, int]] = set()
        page_w = float(page.rect.width)
        page_h = float(page.rect.height)
        for dr in page.get_drawings():
            rect = dr.get("rect")
            if rect is None:
                continue
            w_pt = float(rect.width)
            h_pt = float(rect.height)
            # Ignore small primitives and full-page frames
            if w_pt < 45.0 or h_pt < 45.0:
                continue
            if w_pt > page_w * 0.96 and h_pt > page_h * 0.96:
                continue
            bbox = (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))
            key = tuple(round(v) for v in bbox)
            if key in seen_bboxes:
                continue
            seen_bboxes.add(key)
            vec_rects.append({"page": p_idx, "bbox": bbox})
    vec_rects.sort(key=lambda r: (r["page"], r["bbox"][1], r["bbox"][0]))
    return vec_rects


def _match_images_in_pdf(pdf_doc: pymupdf.Document, docx_info: dict[str, Any]) -> list[dict[str, Any]]:
    sections = {s["section_index"]: s for s in docx_info.get("sections", [])}
    hf_img_cnt = sum(v.get("image_count", 0) for v in docx_info.get("headers_footers", {}).values())
    img_rects: list[dict[str, Any]] = []

    for p_idx, page in enumerate(pdf_doc, start=1):
        seen_bboxes: set[tuple[int, int, int, int]] = set()
        page_h = float(page.rect.height)
        for info in page.get_image_info(xrefs=True):
            bbox = info.get("bbox")
            if not bbox or len(bbox) != 4:
                continue
            key = tuple(round(v) for v in bbox)
            if key in seen_bboxes:
                continue
            seen_bboxes.add(key)
            # Exclude probable header/footer logos when headers/footers contain images.
            if hf_img_cnt > 0 and (bbox[3] < 45.0 or bbox[1] > page_h - 45.0):
                continue
            img_rects.append({"page": p_idx, "bbox": tuple(float(v) for v in bbox)})

    img_rects.sort(key=lambda r: (r["page"], r["bbox"][1], r["bbox"][0]))
    docx_imgs = docx_info.get("images", [])
    vec_rects = _collect_vector_drawing_rects(pdf_doc) if len(img_rects) < len(docx_imgs) else []
    results: list[dict[str, Any]] = []
    used_vec: set[int] = set()
    for idx, img in enumerate(docx_imgs):
        sec = sections.get(img["section_index"], next(iter(sections.values())))
        printable_right_cm = sec["page_width_cm"] - sec["margins_cm"]["right"]
        rec: dict[str, Any] | None = img_rects[idx] if idx < len(img_rects) else None
        if rec is None and vec_rects:
            want_w = img.get("width_cm") or 0.0
            want_h = img.get("height_cm") or 0.0
            best_vi: int | None = None
            best_diff = 1e9
            for vi, vr in enumerate(vec_rects):
                if vi in used_vec:
                    continue
                vw = (vr["bbox"][2] - vr["bbox"][0]) / PT_PER_CM
                vh = (vr["bbox"][3] - vr["bbox"][1]) / PT_PER_CM
                diff = abs(vw - want_w) + abs(vh - want_h)
                if diff < best_diff:
                    best_diff = diff
                    best_vi = vi
            if best_vi is not None and best_diff <= 1.5:
                used_vec.add(best_vi)
                rec = vec_rects[best_vi]
        if rec is None:
            results.append({"ordinal": img["ordinal"], "matched": False})
            continue
        x0, y0, x1, y1 = rec["bbox"]
        x0_cm = round(x0 / PT_PER_CM, 3)
        x1_cm = round(x1 / PT_PER_CM, 3)
        results.append({
            "ordinal": img["ordinal"],
            "matched": True,
            "page": rec["page"],
            "bbox_pt": [round(x0, 2), round(y0, 2), round(x1, 2), round(y1, 2)],
            "x0_cm": x0_cm,
            "x1_cm": x1_cm,
            "width_cm": round((x1 - x0) / PT_PER_CM, 3),
            "height_cm": round((y1 - y0) / PT_PER_CM, 3),
            "center_cm": round((x0_cm + x1_cm) / 2.0, 3),
            "overflow_right_mm": round(max(0.0, (x1_cm - printable_right_cm) * 10.0), 2),
        })
    return results


def detect_orphan_signature_pdf(pdf_path: str | Path) -> bool:
    """Detect via PDF text alone whether the last page contains only signature/recipient boilerplate."""
    with pymupdf.open(str(pdf_path)) as pdf_doc:
        if len(pdf_doc) < 2:
            return False
        raw_lines = (pdf_doc[-1].get_text("text") or "").splitlines()

    meaningful = []
    for ln in raw_lines:
        s = nfc(ln).strip()
        if not s or _PAGE_NUM_RE.match(s):
            continue
        meaningful.append(s)

    if not meaningful:
        return False

    non_sig_prose = 0
    saw_sig_marker = False
    for line in meaningful:
        if _SIG_LINE_RE.match(line):
            saw_sig_marker = True
            continue
        words = line.split()
        # Short signer name (<=5 words without terminal sentence punctuation) is valid signature block.
        if len(words) <= 5 and not line.endswith("."):
            continue
        non_sig_prose += 1

    return bool(saw_sig_marker and non_sig_prose == 0)


def detect_pagination_ripple(
    before_pdf: str | Path,
    after_pdf: str | Path,
    probe_texts: list[str] | None = None,
) -> bool:
    """Check whether page count or probe paragraph page locations shifted between renders."""
    with pymupdf.open(str(before_pdf)) as bdoc, pymupdf.open(str(after_pdf)) as adoc:
        if len(bdoc) != len(adoc):
            return True
        if probe_texts:
            for probe in probe_texts:
                p_norm = " ".join(nfc(probe).split())
                if not p_norm:
                    continue
                b_pages = [i for i, pg in enumerate(bdoc, start=1) if p_norm in " ".join(nfc(pg.get_text("text")).split())]
                a_pages = [i for i, pg in enumerate(adoc, start=1) if p_norm in " ".join(nfc(pg.get_text("text")).split())]
                if b_pages != a_pages:
                    return True
    return False


def inspect_pdf_layout(pdf_path: str | Path, docx_info: dict[str, Any]) -> dict[str, Any]:
    """Extract page count, page dimensions, table bboxes, image bboxes, and orphan signature check."""
    with pymupdf.open(str(pdf_path)) as pdf_doc:
        pages_meta = [
            {
                "page": idx,
                "width_cm": round(float(pg.rect.width) / PT_PER_CM, 3),
                "height_cm": round(float(pg.rect.height) / PT_PER_CM, 3),
            }
            for idx, pg in enumerate(pdf_doc, start=1)
        ]
        tables = _match_tables_in_pdf(pdf_doc, docx_info)
        images = _match_images_in_pdf(pdf_doc, docx_info)

    return {
        "page_count": len(pages_meta),
        "pages": pages_meta,
        "tables": tables,
        "images": images,
        "orphan_signature": detect_orphan_signature_pdf(pdf_path),
    }


def render_pages_png(pdf_path: str | Path, out_dir: str | Path, dpi: int = 150) -> list[str]:
    """Render PDF pages to PNG files for blind human visual evaluation."""
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    scale = dpi / 72.0
    mat = pymupdf.Matrix(scale, scale)
    paths: list[str] = []
    with pymupdf.open(str(pdf_path)) as pdf_doc:
        for idx, page in enumerate(pdf_doc, start=1):
            pix = page.get_pixmap(matrix=mat)
            png_path = out / f"page-{idx}.png"
            pix.save(str(png_path))
            paths.append(str(png_path))
    return paths


def render_and_inspect(
    docx_path: str | Path,
    docx_info: dict[str, Any],
    out_dir: str | Path,
    *,
    backend: str = "libreoffice",
) -> dict[str, Any]:
    """Render DOCX to PDF and inspect rendered geometry in a single call."""
    ren = render_to_pdf(docx_path, out_dir, backend=backend)
    if ren.get("status") != "rendered":
        return {**ren, "layout": None}
    layout = inspect_pdf_layout(ren["pdf_path"], docx_info)
    return {**ren, "layout": layout}
