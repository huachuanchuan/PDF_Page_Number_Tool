from __future__ import annotations

import hashlib
import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pymupdf

POSITION_OPTIONS = {
    "左下角": "left-bottom",
    "下部中间": "bottom-center",
    "右下角": "right-bottom",
}

FONT_OPTIONS: Dict[str, Dict[str, object]] = {
    "微软雅黑 (Microsoft YaHei)": {
        "builtin": "helv",
        "files": (
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/msyh.ttf",
            "/mnt/c/Windows/Fonts/msyh.ttc",
            "/mnt/c/Windows/Fonts/msyh.ttf",
        ),
    },
    "宋体 (SimSun)": {
        "builtin": "helv",
        "files": (
            "C:/Windows/Fonts/simsun.ttc",
            "/mnt/c/Windows/Fonts/simsun.ttc",
        ),
    },
    "黑体 (SimHei)": {
        "builtin": "helv",
        "files": (
            "C:/Windows/Fonts/simhei.ttf",
            "/mnt/c/Windows/Fonts/simhei.ttf",
        ),
    },
    "楷体 (KaiTi)": {
        "builtin": "helv",
        "files": (
            "C:/Windows/Fonts/simkai.ttf",
            "/mnt/c/Windows/Fonts/simkai.ttf",
        ),
    },
    "仿宋 (FangSong)": {
        "builtin": "helv",
        "files": (
            "C:/Windows/Fonts/simfang.ttf",
            "/mnt/c/Windows/Fonts/simfang.ttf",
        ),
    },
    "Arial": {
        "builtin": "helv",
        "files": (
            "C:/Windows/Fonts/arial.ttf",
            "/mnt/c/Windows/Fonts/arial.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ),
    },
    "Times New Roman": {
        "builtin": "tiro",
        "files": (
            "C:/Windows/Fonts/times.ttf",
            "/mnt/c/Windows/Fonts/times.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
        ),
    },
    "Courier New": {
        "builtin": "cour",
        "files": (
            "C:/Windows/Fonts/cour.ttf",
            "/mnt/c/Windows/Fonts/cour.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
        ),
    },
    "Helvetica (内置)": {"builtin": "helv", "files": ()},
    "Times (内置)": {"builtin": "tiro", "files": ()},
    "Courier (内置)": {"builtin": "cour", "files": ()},
}

DEFAULT_PAGE_NUMBER_FONT = "微软雅黑 (Microsoft YaHei)"
DEFAULT_TOC_TITLE_FONT = "微软雅黑 (Microsoft YaHei)"
DEFAULT_TOC_BODY_FONT = "宋体 (SimSun)"

VALID_ROMAN_RE = re.compile(
    r"^M{0,4}(CM|CD|D?C{0,3})"
    r"(XC|XL|L?X{0,3})"
    r"(IX|IV|V?I{0,3})$",
    re.IGNORECASE,
)


@dataclass
class MergeSource:
    path: Path
    page_count: int


def normalize_path(path_text: str) -> Path:
    return Path(path_text.strip().strip('"').strip("'")).expanduser()


def validate_pdf_file(path_like: str | Path) -> Path:
    path = normalize_path(str(path_like))
    if not path.exists():
        raise FileNotFoundError(f"找不到该文件: {path}")
    if path.suffix.lower() != ".pdf":
        raise ValueError(f"不是 PDF 文件: {path}")
    return path


def get_pdf_page_count(path: Path) -> int:
    doc = pymupdf.open(path)
    try:
        return int(doc.page_count)
    finally:
        doc.close()


def build_default_uncompressed_path(reference_pdf: Path) -> Path:
    return reference_pdf.with_name(f"{reference_pdf.stem}_完整版.pdf")


def build_default_compressed_path(uncompressed_pdf: Path) -> Path:
    return uncompressed_pdf.with_name(f"{uncompressed_pdf.stem}_压缩版.pdf")


def hex_to_mupdf_color(hex_color: str) -> Tuple[float, float, float]:
    color_text = hex_color.strip().lstrip("#")
    if len(color_text) != 6:
        raise ValueError("颜色格式无效，请使用 #RRGGBB。")

    red = int(color_text[0:2], 16) / 255
    green = int(color_text[2:4], 16) / 255
    blue = int(color_text[4:6], 16) / 255
    return red, green, blue


def make_font_alias(prefix: str, choice: str) -> str:
    checksum = sum(ord(ch) for ch in choice) % 10000
    return f"{prefix}_{checksum}"


def resolve_font_resource(choice: str, alias_prefix: str) -> Tuple[str, Optional[str]]:
    default_option = FONT_OPTIONS["Helvetica (内置)"]
    option = FONT_OPTIONS.get(choice, default_option)
    raw_candidates = option.get("files", ())
    candidates = raw_candidates if isinstance(raw_candidates, (tuple, list)) else ()

    for candidate in candidates:
        candidate_path = Path(str(candidate))
        if candidate_path.exists():
            return make_font_alias(alias_prefix, choice), str(candidate_path)

    return str(option.get("builtin", "helv")), None


def estimate_text_width(text: str, font_size: int) -> float:
    width = 0.0
    for char in text:
        width += font_size * (1.0 if ord(char) > 127 else 0.56)
    return width


def measure_text_width(text: str, font_size: int, font_name: str) -> float:
    try:
        return float(pymupdf.get_text_length(text, fontname=font_name, fontsize=font_size))
    except Exception:
        return estimate_text_width(text, font_size)


def truncate_text_to_width(text: str, max_width: float, font_size: int, font_name: str) -> str:
    if measure_text_width(text, font_size, font_name) <= max_width:
        return text

    ellipsis = "..."
    trimmed = text
    while trimmed:
        candidate = f"{trimmed}{ellipsis}"
        if measure_text_width(candidate, font_size, font_name) <= max_width:
            return candidate
        trimmed = trimmed[:-1]
    return ellipsis


def draw_dotted_leader(
    page: Any,
    start_x: float,
    end_x: float,
    baseline_y: float,
    font_size: int,
) -> None:
    if end_x <= start_x:
        return

    step = max(1.55, font_size * 0.15)
    radius = max(0.42, font_size * 0.062)
    center_y = baseline_y - (font_size * 0.28)

    shape = page.new_shape()
    x = start_x
    while x <= end_x:
        shape.draw_circle(pymupdf.Point(x, center_y), radius)
        x += step

    shape.finish(color=(0, 0, 0), fill=(0, 0, 0), width=0.1)
    shape.commit()


def draw_toc_entry_line(
    page: Any,
    y: float,
    item_text: str,
    page_text: str,
    left_x: float,
    right_x: float,
    font_size: int,
    font_name: str,
) -> None:
    clean_item = item_text.strip()
    clean_page = page_text.strip()
    if not clean_item or not clean_page:
        raise ValueError("目录项和页码都不能为空。")

    page_width = measure_text_width(clean_page, font_size, font_name)
    page_x = right_x - page_width
    if page_x <= left_x + (font_size * 3):
        raise ValueError("目录排版空间不足，请减小字号后重试。")

    pre_dot_gap = max(1.2, font_size * 0.16)
    post_dot_gap = max(0.2, font_size * 0.03)
    reserve_for_dots = font_size * 12
    max_item_width = page_x - left_x - pre_dot_gap - post_dot_gap - reserve_for_dots
    max_item_width = max(max_item_width, font_size * 4)

    item_display = truncate_text_to_width(clean_item, max_item_width, font_size, font_name)
    item_width = measure_text_width(item_display, font_size, font_name)

    page.insert_text(
        pymupdf.Point(left_x, y),
        item_display,
        fontsize=font_size,
        fontname=font_name,
        color=(0, 0, 0),
    )

    dots_start = left_x + item_width + pre_dot_gap
    dots_end = page_x - post_dot_gap
    draw_dotted_leader(
        page=page,
        start_x=dots_start,
        end_x=dots_end,
        baseline_y=y,
        font_size=font_size,
    )

    # 页码右侧对齐到固定边界：不同位数也在同一列结束。
    page.insert_text(
        pymupdf.Point(page_x, y),
        clean_page,
        fontsize=font_size,
        fontname=font_name,
        color=(0, 0, 0),
    )


def roman_to_int(value_text: str) -> Optional[int]:
    clean = value_text.strip().upper()
    if not clean or VALID_ROMAN_RE.fullmatch(clean) is None:
        return None

    values = {
        "I": 1,
        "V": 5,
        "X": 10,
        "L": 50,
        "C": 100,
        "D": 500,
        "M": 1000,
    }
    total = 0
    prev = 0
    for ch in reversed(clean):
        current = values[ch]
        if current < prev:
            total -= current
        else:
            total += current
            prev = current

    return total if total > 0 else None


def parse_page_number(page_text: str) -> Optional[int]:
    clean = page_text.strip()
    if not clean:
        return None

    digit_matches = re.findall(r"\d+", clean)
    for token in reversed(digit_matches):
        value = int(token)
        if value > 0:
            return value

    direct_roman = roman_to_int(clean)
    if direct_roman is not None:
        return direct_roman

    alpha_tokens = re.findall(r"[A-Za-z]+", clean)
    for token in reversed(alpha_tokens):
        roman_value = roman_to_int(token)
        if roman_value is not None:
            return roman_value

    return None


def build_toc_bookmarks(
    entries: Sequence[Tuple[str, str]],
    toc_page_count: int,
    content_page_count: int,
    toc_title: str,
) -> List[List[Any]]:
    title = toc_title.strip() or "目录"
    bookmarks: List[List[Any]] = [[1, title, 1]]

    for item_text, page_text in entries:
        label = item_text.strip()
        if not label:
            continue

        page_no = parse_page_number(page_text)
        if page_no is None:
            continue

        if content_page_count > 0:
            page_no = min(page_no, content_page_count)

        target_page = max(1, toc_page_count + page_no)
        bookmarks.append([2, label, target_page])

    return bookmarks


def copy_bookmarks_between_docs(source_doc: Any, target_doc: Any) -> None:
    raw_toc = source_doc.get_toc()
    if not raw_toc:
        return

    max_page = max(1, int(target_doc.page_count))
    normalized_toc: List[List[Any]] = []
    for item in raw_toc:
        if len(item) < 3:
            continue

        level_raw, title_raw, page_raw = item[0], item[1], item[2]
        try:
            level = int(level_raw)
        except (TypeError, ValueError):
            level = 1

        title = str(title_raw).strip() if title_raw is not None else ""
        if not title:
            continue

        try:
            page_no = int(page_raw)
        except (TypeError, ValueError):
            page_no = 1

        page_no = max(1, min(page_no, max_page))
        normalized_toc.append([max(1, level), title, page_no])

    if normalized_toc:
        target_doc.set_toc(normalized_toc)


def resolve_text_point(
    page: Any,
    text: str,
    font_size: int,
    position: str,
    font_name: str,
    margin: int = 40,
) -> pymupdf.Point:
    page_width = page.rect.width
    page_height = page.rect.height
    text_width = measure_text_width(text, font_size, font_name)

    if position == "left-bottom":
        x = margin
    elif position == "bottom-center":
        x = (page_width - text_width) / 2
    else:
        x = page_width - text_width - margin

    x = max(margin / 2, x)
    y = page_height - margin
    return pymupdf.Point(x, y)


def merge_pdfs(pdf_paths: Sequence[Path], output_path: Path) -> Path:
    if not pdf_paths:
        raise ValueError("请至少选择一个 PDF 文件用于合并。")

    doc_out = pymupdf.open()
    try:
        for pdf_path in pdf_paths:
            path = validate_pdf_file(pdf_path)
            doc_in = pymupdf.open(path)
            try:
                if doc_in.page_count > 0:
                    doc_out.insert_pdf(doc_in, from_page=0, to_page=doc_in.page_count - 1)
            finally:
                doc_in.close()

        if output_path.exists():
            output_path.unlink()
        doc_out.save(output_path)
        return output_path
    finally:
        doc_out.close()


def add_page_numbers(
    input_pdf: Path,
    output_pdf: Path,
    font_size: int = 18,
    color_hex: str = "#000000",
    position: str = "right-bottom",
    font_choice: str = DEFAULT_PAGE_NUMBER_FONT,
) -> Path:
    source = validate_pdf_file(input_pdf)
    if source.resolve() == output_pdf.resolve():
        raise ValueError("页码输出不能覆盖输入文件。")

    color = hex_to_mupdf_color(color_hex)
    font_name, font_file = resolve_font_resource(font_choice, "page_num")

    doc = pymupdf.open(source)
    try:
        for page_index in range(doc.page_count):
            page = doc.load_page(page_index)
            if font_file:
                page.insert_font(fontname=font_name, fontfile=font_file)

            text = str(page_index + 1)
            insert_point = resolve_text_point(page, text, font_size, position, font_name)
            page.insert_text(
                insert_point,
                text,
                fontsize=font_size,
                fontname=font_name,
                color=color,
            )

        if output_pdf.exists():
            output_pdf.unlink()
        doc.save(output_pdf)
        return output_pdf
    finally:
        doc.close()


def prepend_toc_pages(
    numbered_pdf: Path,
    entries: Sequence[Tuple[str, str]],
    title: str,
    output_pdf: Path,
    title_font_choice: str = DEFAULT_TOC_TITLE_FONT,
    title_font_size: int = 24,
    body_font_choice: str = DEFAULT_TOC_BODY_FONT,
    body_font_size: int = 14,
) -> Path:
    source = validate_pdf_file(numbered_pdf)
    if not entries:
        raise ValueError("请至少添加一条目录项。")
    if source.resolve() == output_pdf.resolve():
        raise ValueError("目录输出不能覆盖已编号文件。")

    content_doc = pymupdf.open(source)
    toc_doc = pymupdf.open()
    try:
        if content_doc.page_count > 0:
            first_rect = content_doc.load_page(0).rect
            page_width = first_rect.width
            page_height = first_rect.height
        else:
            page_width = 595.0
            page_height = 842.0

        title_font_name, title_font_file = resolve_font_resource(title_font_choice, "toc_title")
        body_font_name, body_font_file = resolve_font_resource(body_font_choice, "toc_body")

        left_margin = 56
        right_margin = page_width - 56
        title_y = 102
        first_line_y = 170
        continue_page_start_y = 96
        bottom_margin = 74
        line_height = max(body_font_size * 1.85, 24)

        entry_index = 0
        toc_page_count = 0
        display_title = title.strip() or "目录"

        while entry_index < len(entries):
            toc_page = toc_doc.new_page(width=page_width, height=page_height)
            toc_page_count += 1

            if title_font_file:
                toc_page.insert_font(fontname=title_font_name, fontfile=title_font_file)
            if body_font_file:
                toc_page.insert_font(fontname=body_font_name, fontfile=body_font_file)

            if toc_page_count == 1:
                title_width = measure_text_width(display_title, title_font_size, title_font_name)
                title_x = max(left_margin, (page_width - title_width) / 2)
                toc_page.insert_text(
                    pymupdf.Point(title_x, title_y),
                    display_title,
                    fontsize=title_font_size,
                    fontname=title_font_name,
                    color=(0, 0, 0),
                )
                y = first_line_y
            else:
                y = continue_page_start_y

            while entry_index < len(entries):
                item_text, page_text = entries[entry_index]
                draw_toc_entry_line(
                    page=toc_page,
                    y=y,
                    item_text=item_text,
                    page_text=page_text,
                    left_x=left_margin,
                    right_x=right_margin,
                    font_size=body_font_size,
                    font_name=body_font_name,
                )

                y += line_height
                entry_index += 1
                if y > page_height - bottom_margin:
                    break

        if content_doc.page_count > 0:
            toc_doc.insert_pdf(content_doc, from_page=0, to_page=content_doc.page_count - 1)

        bookmarks = build_toc_bookmarks(
            entries=entries,
            toc_page_count=toc_page_count,
            content_page_count=content_doc.page_count,
            toc_title=display_title,
        )
        if len(bookmarks) > 1:
            toc_doc.set_toc(bookmarks)
        else:
            toc_doc.set_toc([[1, display_title, 1]])

        if output_pdf.exists():
            output_pdf.unlink()
        toc_doc.save(output_pdf)
        return output_pdf
    finally:
        toc_doc.close()
        content_doc.close()


def save_optimized_pdf(input_pdf: Path, output_pdf: Path) -> None:
    doc = pymupdf.open(input_pdf)
    try:
        if output_pdf.exists():
            output_pdf.unlink()
        try:
            # linear 参数在新版本 PyMuPDF 中已废弃，避免触发
            # "code=4: Linearisation is no longer supported"。
            doc.save(output_pdf, garbage=4, deflate=True, clean=True, use_objstms=1)
        except TypeError:
            # 兼容较旧版本：可能不支持 use_objstms。
            doc.save(output_pdf, garbage=4, deflate=True, clean=True)
        except Exception:
            # 兜底：只保留基础无损压缩，保证步骤4不因参数差异失败。
            doc.save(output_pdf, garbage=3, deflate=True)
    finally:
        doc.close()


def rasterize_pdf(
    input_pdf: Path,
    output_pdf: Path,
    scale: float,
    jpg_quality: int,
) -> None:
    src = pymupdf.open(input_pdf)
    dst = pymupdf.open()
    try:
        matrix = pymupdf.Matrix(scale, scale)
        for page_index in range(src.page_count):
            src_page = src.load_page(page_index)
            rect = src_page.rect
            pix = src_page.get_pixmap(matrix=matrix, alpha=False)
            try:
                image_stream = pix.tobytes("jpg", jpg_quality=jpg_quality)
            except Exception:
                image_stream = pix.tobytes("jpeg", jpg_quality=jpg_quality)

            dst_page = dst.new_page(width=rect.width, height=rect.height)
            dst_page.insert_image(rect, stream=image_stream)

        copy_bookmarks_between_docs(src, dst)

        if output_pdf.exists():
            output_pdf.unlink()
        dst.save(output_pdf, garbage=4, deflate=True)
    finally:
        dst.close()
        src.close()


def compress_pdf_to_target(
    input_pdf: Path,
    output_pdf: Path,
    target_mb: float,
    work_dir: Path,
    tolerance_percent: float = 10.0,
) -> Dict[str, Any]:
    source = validate_pdf_file(input_pdf)
    if target_mb <= 0:
        raise ValueError("目标压缩大小必须大于 0 MB。")
    if source.resolve() == output_pdf.resolve():
        raise ValueError("压缩输出路径不能与未压缩文件相同。")
    if tolerance_percent < 0:
        raise ValueError("允许超出比例不能为负数。")

    target_bytes = int(target_mb * 1024 * 1024)
    tolerance_ratio = tolerance_percent / 100.0
    allowed_upper_bytes = int(target_bytes * (1.0 + tolerance_ratio))
    source_size = source.stat().st_size

    candidates: List[Tuple[Path, int, str]] = []
    candidate_index = 0

    def add_candidate(path: Path, method: str) -> None:
        if path.exists():
            candidates.append((path, path.stat().st_size, method))

    def add_raster_candidate(scale: float, quality: int, tag: str) -> Tuple[Path, int, str]:
        nonlocal candidate_index
        candidate_path = work_dir / f"compress_raster_{candidate_index:03d}.pdf"
        candidate_index += 1
        rasterize_pdf(source, candidate_path, scale=scale, jpg_quality=quality)
        method = f"栅格压缩 scale={scale:.2f} quality={quality} {tag}".strip()
        add_candidate(candidate_path, method)
        return candidates[-1]

    optimized_path = work_dir / "compress_optimize.pdf"
    save_optimized_pdf(source, optimized_path)
    add_candidate(optimized_path, "优化保存")

    # 先用较高清晰度做“缩放探测”，找到最接近目标大小的缩放比例。
    probe_scales = [4.0, 3.5, 3.0, 2.5, 2.0, 1.75, 1.5, 1.25, 1.0, 0.85, 0.7]
    probe_results: List[Tuple[Path, int, str, float]] = []
    for scale in probe_scales:
        result = add_raster_candidate(scale, 96, tag="probe")
        probe_results.append((result[0], result[1], result[2], scale))

    if probe_results:
        closest_probe = min(probe_results, key=lambda item: abs(item[1] - target_bytes))
        refine_scale = float(closest_probe[3])
    else:
        refine_scale = 1.0

    # 在候选缩放比例上用二分搜索质量，尽量贴近目标容量。
    def refine_quality_for_scale(scale: float) -> None:
        low_q, high_q = 30, 98
        for _ in range(8):
            if low_q > high_q:
                break
            mid_q = (low_q + high_q) // 2
            _, mid_size, _ = add_raster_candidate(scale, mid_q, tag="refine")
            if mid_size > target_bytes:
                high_q = mid_q - 1
            else:
                low_q = mid_q + 1

    refine_quality_for_scale(refine_scale)

    # 再补一个相邻缩放比例，减少“目标 5MB 但结果过小”的情况。
    if refine_scale in probe_scales:
        idx = probe_scales.index(refine_scale)
        neighbor_scale = None
        if idx > 0:
            neighbor_scale = probe_scales[idx - 1]
        elif idx + 1 < len(probe_scales):
            neighbor_scale = probe_scales[idx + 1]

        if neighbor_scale is not None:
            refine_quality_for_scale(neighbor_scale)

    if not any(size <= allowed_upper_bytes for _, size, _ in candidates):
        aggressive_candidates = [
            max(0.35, refine_scale * 0.85),
            max(0.30, refine_scale * 0.70),
            0.60,
            0.50,
            0.40,
        ]
        seen_scales = set()
        for scale in aggressive_candidates:
            scale_key = round(scale, 3)
            if scale_key in seen_scales:
                continue
            seen_scales.add(scale_key)
            for quality in (28, 24, 20, 16, 12):
                add_raster_candidate(scale, quality, tag="aggressive")
                if any(size <= allowed_upper_bytes for _, size, _ in candidates):
                    break
            if any(size <= allowed_upper_bytes for _, size, _ in candidates):
                break

    if not candidates:
        raise RuntimeError("压缩失败：未生成任何候选文件。")

    within_tolerance = [item for item in candidates if item[1] <= allowed_upper_bytes]
    if within_tolerance:
        selected = min(within_tolerance, key=lambda item: (abs(item[1] - target_bytes), -item[1]))
    else:
        selected = min(candidates, key=lambda item: item[1])

    target_met = selected[1] <= allowed_upper_bytes
    strict_target_met = selected[1] <= target_bytes

    selected_path, selected_size, selected_method = selected
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    if output_pdf.exists():
        output_pdf.unlink()
    shutil.copy2(selected_path, output_pdf)

    for candidate_path, _, _ in candidates:
        if candidate_path.exists():
            candidate_path.unlink()

    return {
        "target_mb": target_mb,
        "tolerance_percent": tolerance_percent,
        "allowed_upper_mb": allowed_upper_bytes / (1024 * 1024),
        "source_mb": source_size / (1024 * 1024),
        "actual_mb": selected_size / (1024 * 1024),
        "target_met": target_met,
        "strict_target_met": strict_target_met,
        "method": selected_method,
    }


def render_first_page_preview(
    input_pdf: Path,
    output_image: Path,
    max_width: int = 260,
    max_height: int = 360,
) -> Path:
    doc = pymupdf.open(input_pdf)
    try:
        if doc.page_count <= 0:
            raise ValueError("该 PDF 没有可预览页面。")
        page = doc.load_page(0)
        rect = page.rect
        zoom = min(max_width / rect.width, max_height / rect.height)
        zoom = max(0.08, zoom)
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        output_image.parent.mkdir(parents=True, exist_ok=True)
        pix.save(output_image)
        return output_image
    finally:
        doc.close()


class PDFWorkflowApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF 扫描件处理工作台")
        self.geometry("1080x860")
        self.minsize(1020, 800)
        self.configure(bg="#f4f6fb")

        self.runtime_dir = Path(tempfile.mkdtemp(prefix="pdf_workflow_"))
        self.protocol("WM_DELETE_WINDOW", self.handle_close)

        self.merge_item_path_map: Dict[str, Path] = {}
        self.merge_item_pages_map: Dict[str, int] = {}
        self.dragging_item_id: Optional[str] = None
        self.merge_preview_photo: Optional[tk.PhotoImage] = None

        self.reference_pdf: Optional[Path] = None
        self.merged_pdf_path: Optional[Path] = None
        self.numbered_pdf_path: Optional[Path] = None
        self.final_uncompressed_path: Optional[Path] = None
        self.final_compressed_path: Optional[Path] = None
        self.auto_toc_entries: List[Tuple[str, str]] = []

        self.number_source_var = tk.StringVar()
        self.page_font_size_var = tk.IntVar(value=18)
        self.page_color_var = tk.StringVar(value="#000000")
        self.page_position_var = tk.StringVar(value="右下角")
        self.page_font_var = tk.StringVar(value=DEFAULT_PAGE_NUMBER_FONT)

        self.toc_source_var = tk.StringVar()
        self.toc_output_var = tk.StringVar()
        self.toc_title_var = tk.StringVar(value="目录")
        self.toc_title_font_var = tk.StringVar(value=DEFAULT_TOC_TITLE_FONT)
        self.toc_title_size_var = tk.IntVar(value=24)
        self.toc_body_font_var = tk.StringVar(value=DEFAULT_TOC_BODY_FONT)
        self.toc_body_size_var = tk.IntVar(value=14)
        self.toc_item_var = tk.StringVar()
        self.toc_page_var = tk.StringVar()

        self.compress_source_var = tk.StringVar()
        self.compress_output_var = tk.StringVar()
        self.target_size_mb_var = tk.StringVar(value="20")
        self.compress_tolerance_percent_var = tk.StringVar(value="10")

        self.status_var = tk.StringVar(
            value="步骤 1：先选择多个 PDF 并拖拽排序合并；之后依次添加页码、目录、压缩。"
        )

        self.setup_styles()
        self.build_ui()

    def setup_styles(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("App.TFrame", background="#f4f6fb")
        style.configure("Card.TFrame", background="#ffffff", relief="flat")
        style.configure(
            "Title.TLabel",
            background="#f4f6fb",
            foreground="#1f2a44",
            font=("Microsoft YaHei UI", 20, "bold"),
        )
        style.configure(
            "SubTitle.TLabel",
            background="#f4f6fb",
            foreground="#617089",
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "Label.TLabel",
            background="#ffffff",
            foreground="#2c3551",
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "Status.TLabel",
            background="#f4f6fb",
            foreground="#4d5a73",
            font=("Microsoft YaHei UI", 9),
        )

    def build_ui(self) -> None:
        wrapper = ttk.Frame(self, style="App.TFrame", padding=20)
        wrapper.pack(fill="both", expand=True)

        ttk.Label(wrapper, text="PDF 扫描件加页码、目录与压缩", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            wrapper,
            text=(
                "流程：1) 合并（支持拖拽排序+首页预览） 2) 添加页码 3) 生成目录 4) 按目标大小压缩。"
                "仅输出未压缩完整版与压缩版。"
            ),
            style="SubTitle.TLabel",
        ).pack(anchor="w", pady=(4, 12))

        self.notebook = ttk.Notebook(wrapper)
        self.notebook.pack(fill="both", expand=True)

        self.merge_tab = ttk.Frame(self.notebook, style="Card.TFrame", padding=16)
        self.number_tab = ttk.Frame(self.notebook, style="Card.TFrame", padding=16)
        self.toc_tab = ttk.Frame(self.notebook, style="Card.TFrame", padding=16)
        self.compress_tab = ttk.Frame(self.notebook, style="Card.TFrame", padding=16)

        self.notebook.add(self.merge_tab, text="步骤 1：合并 PDF")
        self.notebook.add(self.number_tab, text="步骤 2：添加页码")
        self.notebook.add(self.toc_tab, text="步骤 3：编写目录")
        self.notebook.add(self.compress_tab, text="步骤 4：压缩输出")

        self.notebook.tab(1, state="disabled")
        self.notebook.tab(2, state="disabled")
        self.notebook.tab(3, state="disabled")

        self.build_merge_tab()
        self.build_number_tab()
        self.build_toc_tab()
        self.build_compress_tab()

        ttk.Label(wrapper, textvariable=self.status_var, style="Status.TLabel").pack(anchor="w", pady=(10, 0))

    def create_action_button(
        self,
        parent,
        text: str,
        command,
        primary: bool = False,
        width: Optional[int] = None,
    ) -> tk.Button:
        fg = "#1f2a44"
        bg = "#eef2ff"
        active_bg = "#dfe7ff"
        border = "#bfcaef"

        if primary:
            fg = "#ffffff"
            bg = "#2f5cff"
            active_bg = "#244ddf"
            border = "#244ddf"

        button = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Microsoft YaHei UI", 10),
            fg=fg,
            bg=bg,
            activeforeground=fg,
            activebackground=active_bg,
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=10,
            pady=5,
            cursor="hand2",
        )
        button.configure(highlightbackground=border)
        if width is not None:
            button.configure(width=width)
        return button

    def build_path_row(
        self,
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        browse_command,
        readonly: bool = False,
    ) -> ttk.Entry:
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x", pady=(2, 10))

        ttk.Label(row, text=label, style="Label.TLabel").pack(anchor="w")
        input_group = ttk.Frame(row, style="Card.TFrame")
        input_group.pack(fill="x", pady=(6, 0))

        state = "readonly" if readonly else "normal"
        entry = ttk.Entry(input_group, textvariable=variable, state=state)
        entry.pack(side="left", fill="x", expand=True)

        if browse_command is not None:
            ttk.Button(input_group, text="浏览", command=browse_command).pack(side="left", padx=(8, 0))
        return entry

    # -----------------------------
    # 步骤1：合并
    # -----------------------------
    def build_merge_tab(self) -> None:
        tips = ttk.Label(
            self.merge_tab,
            text="选择多个 PDF 后，可在列表中拖拽调整顺序；右侧显示选中项第一页预览。",
            style="Label.TLabel",
        )
        tips.pack(anchor="w")

        toolbar = ttk.Frame(self.merge_tab, style="Card.TFrame")
        toolbar.pack(fill="x", pady=(10, 8))

        self.select_merge_files_button = self.create_action_button(
            toolbar,
            text="选择多个 PDF",
            command=self.select_merge_files,
            width=12,
        )
        self.select_merge_files_button.pack(side="left")

        self.remove_merge_file_button = self.create_action_button(
            toolbar,
            text="移除选中",
            command=self.remove_selected_merge_files,
            width=10,
        )
        self.remove_merge_file_button.pack(side="left", padx=(8, 0))

        self.clear_merge_files_button = self.create_action_button(
            toolbar,
            text="清空列表",
            command=self.clear_merge_files,
            width=10,
        )
        self.clear_merge_files_button.pack(side="left", padx=(8, 0))

        content = ttk.Frame(self.merge_tab, style="Card.TFrame")
        content.pack(fill="both", expand=True)

        left = ttk.Frame(content, style="Card.TFrame")
        left.pack(side="left", fill="both", expand=True)

        right = ttk.Frame(content, style="Card.TFrame", padding=(16, 0, 0, 0))
        right.pack(side="left", fill="y")

        self.merge_tree = ttk.Treeview(
            left,
            columns=("order", "name", "pages"),
            show="headings",
            height=14,
        )
        self.merge_tree.heading("order", text="顺序")
        self.merge_tree.heading("name", text="PDF 文件")
        self.merge_tree.heading("pages", text="页数")
        self.merge_tree.column("order", width=60, anchor="center")
        self.merge_tree.column("name", width=560, anchor="w")
        self.merge_tree.column("pages", width=90, anchor="center")

        merge_scroll = ttk.Scrollbar(left, orient="vertical", command=self.merge_tree.yview)
        self.merge_tree.configure(yscrollcommand=merge_scroll.set)

        self.merge_tree.pack(side="left", fill="both", expand=True)
        merge_scroll.pack(side="left", fill="y")

        self.merge_tree.bind("<<TreeviewSelect>>", self.on_merge_tree_select)
        self.merge_tree.bind("<ButtonPress-1>", self.on_merge_tree_press)
        self.merge_tree.bind("<B1-Motion>", self.on_merge_tree_motion)
        self.merge_tree.bind("<ButtonRelease-1>", self.on_merge_tree_release)

        ttk.Label(right, text="第一页预览", style="Label.TLabel").pack(anchor="w")
        self.preview_hint_label = ttk.Label(
            right,
            text="选中左侧文件可查看预览",
            style="SubTitle.TLabel",
        )
        self.preview_hint_label.pack(anchor="w", pady=(4, 8))

        self.preview_image_label = tk.Label(
            right,
            width=280,
            height=380,
            text="暂无预览",
            fg="#7a849b",
            bg="#f6f7fb",
            relief="solid",
            bd=1,
            anchor="center",
            justify="center",
            wraplength=240,
        )
        self.preview_image_label.pack(anchor="w")

        action_row = ttk.Frame(self.merge_tab, style="Card.TFrame", height=56)
        action_row.pack(fill="x", pady=(12, 0))
        action_row.pack_propagate(False)

        self.merge_execute_button = self.create_action_button(
            action_row,
            text="执行合并并进入下一步",
            command=self.handle_merge_pdfs,
            primary=True,
            width=24,
        )
        self.merge_execute_button.pack(side="right", pady=8)

    def get_merge_order_paths(self) -> List[Path]:
        return [self.merge_item_path_map[item_id] for item_id in self.merge_tree.get_children()]

    def update_merge_order_numbers(self) -> None:
        for index, item_id in enumerate(self.merge_tree.get_children(), start=1):
            values = list(self.merge_tree.item(item_id, "values"))
            if len(values) >= 3:
                values[0] = str(index)
                self.merge_tree.item(item_id, values=values)

    def reset_pipeline_after_merge_change(self) -> None:
        self.merged_pdf_path = None
        self.numbered_pdf_path = None
        self.final_uncompressed_path = None
        self.final_compressed_path = None
        self.auto_toc_entries = []

        self.number_source_var.set("")
        self.toc_source_var.set("")
        self.compress_source_var.set("")
        self.toc_output_var.set("")
        self.compress_output_var.set("")

        self.clear_toc_entries()

        self.notebook.tab(1, state="disabled")
        self.notebook.tab(2, state="disabled")
        self.notebook.tab(3, state="disabled")

    def select_merge_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="选择待合并的 PDF 文件（可多选）",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if not selected:
            return

        existing = {path.resolve() for path in self.merge_item_path_map.values()}
        added_count = 0

        for raw_path in selected:
            try:
                pdf_path = validate_pdf_file(raw_path)
                resolved = pdf_path.resolve()
                if resolved in existing:
                    continue

                page_count = get_pdf_page_count(pdf_path)
                item_id = self.merge_tree.insert(
                    "",
                    "end",
                    values=("0", pdf_path.name, str(page_count)),
                )
                self.merge_item_path_map[item_id] = pdf_path
                self.merge_item_pages_map[item_id] = page_count
                existing.add(resolved)
                added_count += 1
            except Exception as exc:
                messagebox.showwarning("文件跳过", str(exc))

        if added_count > 0:
            self.update_merge_order_numbers()
            if not self.reference_pdf:
                first_path = self.get_merge_order_paths()[0]
                self.reference_pdf = first_path

            first_item = self.merge_tree.get_children()[0]
            self.merge_tree.selection_set(first_item)
            self.show_preview_for_item(first_item)
            self.reset_pipeline_after_merge_change()
            self.status_var.set(f"已新增 {added_count} 个 PDF。可拖拽列表调整合并顺序。")

    def remove_selected_merge_files(self) -> None:
        selected_ids = self.merge_tree.selection()
        if not selected_ids:
            return

        for item_id in selected_ids:
            self.merge_item_path_map.pop(item_id, None)
            self.merge_item_pages_map.pop(item_id, None)
            self.merge_tree.delete(item_id)

        self.update_merge_order_numbers()
        self.reset_pipeline_after_merge_change()
        self.preview_image_label.configure(image="", text="暂无预览")
        self.merge_preview_photo = None

    def clear_merge_files(self) -> None:
        for item_id in self.merge_tree.get_children():
            self.merge_tree.delete(item_id)

        self.merge_item_path_map.clear()
        self.merge_item_pages_map.clear()
        self.reference_pdf = None
        self.reset_pipeline_after_merge_change()
        self.preview_image_label.configure(image="", text="暂无预览")
        self.merge_preview_photo = None

    def on_merge_tree_select(self, _event) -> None:
        selected = self.merge_tree.selection()
        if selected:
            self.show_preview_for_item(selected[0])

    def on_merge_tree_press(self, event) -> None:
        row_id = self.merge_tree.identify_row(event.y)
        self.dragging_item_id = row_id if row_id else None

    def on_merge_tree_motion(self, event) -> None:
        if not self.dragging_item_id:
            return
        target_id = self.merge_tree.identify_row(event.y)
        if target_id and target_id != self.dragging_item_id:
            target_index = self.merge_tree.index(target_id)
            self.merge_tree.move(self.dragging_item_id, "", target_index)
            self.update_merge_order_numbers()

    def on_merge_tree_release(self, _event) -> None:
        if self.dragging_item_id:
            self.update_merge_order_numbers()
        self.dragging_item_id = None

    def show_preview_for_item(self, item_id: str) -> None:
        pdf_path = self.merge_item_path_map.get(item_id)
        if not pdf_path:
            return

        preview_name = hashlib.md5(str(pdf_path).encode("utf-8")).hexdigest()[:12]
        preview_path = self.runtime_dir / f"preview_{preview_name}.png"
        try:
            if not preview_path.exists():
                render_first_page_preview(pdf_path, preview_path)

            photo = tk.PhotoImage(file=str(preview_path))
            self.merge_preview_photo = photo
            self.preview_image_label.configure(image=photo, text="")
        except Exception as exc:
            self.merge_preview_photo = None
            self.preview_image_label.configure(
                image="",
                text=f"预览失败\n{exc}",
            )

    def build_auto_toc_entries(self, merge_sources: Sequence[MergeSource]) -> List[Tuple[str, str]]:
        entries: List[Tuple[str, str]] = []
        current_page = 1
        for source in merge_sources:
            entries.append((source.path.stem, str(current_page)))
            current_page += source.page_count
        return entries

    def handle_merge_pdfs(self) -> None:
        try:
            ordered_paths = self.get_merge_order_paths()
            if len(ordered_paths) < 1:
                raise ValueError("请先选择至少一个 PDF 文件。")

            merge_sources = [
                MergeSource(path=path, page_count=get_pdf_page_count(path))
                for path in ordered_paths
            ]
            merged_path = self.runtime_dir / "01_merged.pdf"
            self.status_var.set("步骤 1 处理中：正在合并 PDF...")
            self.update_idletasks()

            merge_pdfs(ordered_paths, merged_path)
            self.reference_pdf = ordered_paths[0]
            self.merged_pdf_path = merged_path
            self.numbered_pdf_path = None
            self.final_uncompressed_path = None
            self.final_compressed_path = None

            self.auto_toc_entries = self.build_auto_toc_entries(merge_sources)
            self.clear_toc_entries()
            for item_text, page_text in self.auto_toc_entries:
                self.toc_tree.insert("", "end", values=(item_text, page_text))

            self.number_source_var.set(str(merged_path))
            default_uncompressed = build_default_uncompressed_path(self.reference_pdf)
            self.toc_output_var.set(str(default_uncompressed))
            self.compress_output_var.set(str(build_default_compressed_path(default_uncompressed)))

            self.notebook.tab(1, state="normal")
            self.notebook.tab(2, state="disabled")
            self.notebook.tab(3, state="disabled")
            self.notebook.select(1)

            self.status_var.set(
                "步骤 1 完成：已合并 PDF，并自动生成目录草稿（基于文件名+起始页码）。请进行步骤 2。"
            )
            messagebox.showinfo(
                "步骤 1 完成",
                "PDF 合并成功。\n"
                "已根据合并顺序自动生成目录初稿（文件名 + 起始页码），后续可在目录步骤中细调。",
            )
        except Exception as exc:
            self.status_var.set("步骤 1 失败，请检查输入后重试。")
            messagebox.showerror("步骤 1 失败", str(exc))

    # -----------------------------
    # 步骤2：页码
    # -----------------------------
    def build_number_tab(self) -> None:
        self.build_path_row(
            self.number_tab,
            "合并结果（步骤 1 自动生成）",
            self.number_source_var,
            None,
            readonly=True,
        )

        options = ttk.Frame(self.number_tab, style="Card.TFrame")
        options.pack(fill="x", pady=(10, 0))

        ttk.Label(options, text="页码字体", style="Label.TLabel").grid(row=0, column=0, sticky="w")
        self.page_font_box = ttk.Combobox(
            options,
            textvariable=self.page_font_var,
            values=list(FONT_OPTIONS.keys()),
            state="readonly",
            width=28,
        )
        self.page_font_box.grid(row=1, column=0, sticky="w", pady=(6, 0))

        ttk.Label(options, text="字体大小", style="Label.TLabel").grid(row=0, column=1, sticky="w", padx=(24, 0))
        self.page_size_box = ttk.Spinbox(
            options,
            from_=8,
            to=72,
            textvariable=self.page_font_size_var,
            width=8,
            justify="center",
        )
        self.page_size_box.grid(row=1, column=1, sticky="w", padx=(24, 0), pady=(6, 0))

        ttk.Label(options, text="页码颜色", style="Label.TLabel").grid(row=0, column=2, sticky="w", padx=(24, 0))
        color_group = ttk.Frame(options, style="Card.TFrame")
        color_group.grid(row=1, column=2, sticky="w", padx=(24, 0), pady=(6, 0))

        self.page_color_preview = tk.Label(
            color_group,
            width=3,
            height=1,
            bg=self.page_color_var.get(),
            relief="solid",
            bd=1,
        )
        self.page_color_preview.pack(side="left")

        self.select_color_button = self.create_action_button(
            color_group,
            text="选择颜色",
            command=self.select_page_color,
            width=10,
        )
        self.select_color_button.pack(side="left", padx=(8, 0))

        ttk.Label(options, text="页码位置", style="Label.TLabel").grid(row=0, column=3, sticky="w", padx=(24, 0))
        self.page_position_box = ttk.Combobox(
            options,
            textvariable=self.page_position_var,
            values=list(POSITION_OPTIONS.keys()),
            state="readonly",
            width=12,
        )
        self.page_position_box.grid(row=1, column=3, sticky="w", padx=(24, 0), pady=(6, 0))

        action_row = ttk.Frame(self.number_tab, style="Card.TFrame", height=56)
        action_row.pack(fill="x", pady=(14, 0))
        action_row.pack_propagate(False)

        self.number_execute_button = self.create_action_button(
            action_row,
            text="添加页码并进入下一步",
            command=self.handle_add_page_numbers,
            primary=True,
            width=22,
        )
        self.number_execute_button.pack(side="right", pady=8)

    def select_page_color(self) -> None:
        color_result = colorchooser.askcolor(title="选择页码颜色", color=self.page_color_var.get())
        if color_result[1]:
            self.page_color_var.set(color_result[1])
            self.page_color_preview.configure(bg=color_result[1])

    def handle_add_page_numbers(self) -> None:
        try:
            if not self.merged_pdf_path or not self.merged_pdf_path.exists():
                raise ValueError("请先完成步骤 1 的 PDF 合并。")

            font_size = int(self.page_font_size_var.get())
            if not 8 <= font_size <= 72:
                raise ValueError("字体大小需在 8 到 72 之间。")

            position_key = self.page_position_var.get()
            position = POSITION_OPTIONS.get(position_key)
            if not position:
                raise ValueError("请选择有效的页码位置。")

            numbered_path = self.runtime_dir / "02_numbered.pdf"
            self.status_var.set("步骤 2 处理中：正在添加页码...")
            self.update_idletasks()

            add_page_numbers(
                input_pdf=self.merged_pdf_path,
                output_pdf=numbered_path,
                font_size=font_size,
                color_hex=self.page_color_var.get(),
                position=position,
                font_choice=self.page_font_var.get(),
            )

            self.numbered_pdf_path = numbered_path
            self.final_uncompressed_path = None
            self.final_compressed_path = None
            self.compress_source_var.set("")

            self.toc_source_var.set(str(numbered_path))
            self.notebook.tab(2, state="normal")
            self.notebook.tab(3, state="disabled")
            self.notebook.select(2)

            self.status_var.set("步骤 2 完成：已添加页码。请在步骤 3 细化目录并生成未压缩完整版。")
            messagebox.showinfo("步骤 2 完成", "页码添加成功。\n请进入步骤 3 编写目录。")
        except Exception as exc:
            self.status_var.set("步骤 2 失败，请检查输入后重试。")
            messagebox.showerror("步骤 2 失败", str(exc))

    # -----------------------------
    # 步骤3：目录
    # -----------------------------
    def build_toc_tab(self) -> None:
        self.build_path_row(
            self.toc_tab,
            "已编号 PDF（步骤 2 自动生成）",
            self.toc_source_var,
            None,
            readonly=True,
        )
        self.build_path_row(self.toc_tab, "未压缩完整版输出路径", self.toc_output_var, self.select_toc_output_file)

        title_group = ttk.Frame(self.toc_tab, style="Card.TFrame")
        title_group.pack(fill="x", pady=(6, 0))

        ttk.Label(title_group, text="目录标题", style="Label.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(title_group, textvariable=self.toc_title_var, width=24).grid(row=1, column=0, sticky="w", pady=(6, 0))

        ttk.Label(title_group, text="标题字体", style="Label.TLabel").grid(row=0, column=1, sticky="w", padx=(20, 0))
        self.toc_title_font_box = ttk.Combobox(
            title_group,
            textvariable=self.toc_title_font_var,
            values=list(FONT_OPTIONS.keys()),
            state="readonly",
            width=26,
        )
        self.toc_title_font_box.grid(row=1, column=1, sticky="w", padx=(20, 0), pady=(6, 0))

        ttk.Label(title_group, text="标题字号", style="Label.TLabel").grid(row=0, column=2, sticky="w", padx=(20, 0))
        self.toc_title_size_box = ttk.Spinbox(
            title_group,
            from_=10,
            to=72,
            textvariable=self.toc_title_size_var,
            width=8,
            justify="center",
        )
        self.toc_title_size_box.grid(row=1, column=2, sticky="w", padx=(20, 0), pady=(6, 0))

        ttk.Label(title_group, text="目录字体", style="Label.TLabel").grid(row=0, column=3, sticky="w", padx=(20, 0))
        self.toc_body_font_box = ttk.Combobox(
            title_group,
            textvariable=self.toc_body_font_var,
            values=list(FONT_OPTIONS.keys()),
            state="readonly",
            width=26,
        )
        self.toc_body_font_box.grid(row=1, column=3, sticky="w", padx=(20, 0), pady=(6, 0))

        ttk.Label(title_group, text="目录字号", style="Label.TLabel").grid(row=0, column=4, sticky="w", padx=(20, 0))
        self.toc_body_size_box = ttk.Spinbox(
            title_group,
            from_=8,
            to=48,
            textvariable=self.toc_body_size_var,
            width=8,
            justify="center",
        )
        self.toc_body_size_box.grid(row=1, column=4, sticky="w", padx=(20, 0), pady=(6, 0))

        entry_card = ttk.Frame(self.toc_tab, style="Card.TFrame")
        entry_card.pack(fill="x", pady=(14, 0))

        ttk.Label(entry_card, text="目录项", style="Label.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(entry_card, text="页码", style="Label.TLabel").grid(row=0, column=1, sticky="w", padx=(12, 0))

        ttk.Entry(entry_card, textvariable=self.toc_item_var, width=56).grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(entry_card, textvariable=self.toc_page_var, width=12).grid(
            row=1,
            column=1,
            sticky="w",
            padx=(12, 0),
            pady=(6, 0),
        )

        self.add_toc_entry_button = self.create_action_button(
            entry_card,
            text="添加到目录列表",
            command=self.add_toc_entry,
            width=12,
        )
        self.add_toc_entry_button.grid(row=1, column=2, sticky="w", padx=(12, 0), pady=(6, 0))

        list_card = ttk.Frame(self.toc_tab, style="Card.TFrame")
        list_card.pack(fill="both", expand=True, pady=(10, 0))

        self.toc_tree = ttk.Treeview(list_card, columns=("item", "page"), show="headings", height=10)
        self.toc_tree.heading("item", text="目录项（第一列）")
        self.toc_tree.heading("page", text="对应页码（第二列）")
        self.toc_tree.column("item", width=690, anchor="w")
        self.toc_tree.column("page", width=180, anchor="center")

        toc_scroll = ttk.Scrollbar(list_card, orient="vertical", command=self.toc_tree.yview)
        self.toc_tree.configure(yscrollcommand=toc_scroll.set)
        self.toc_tree.pack(side="left", fill="both", expand=True)
        toc_scroll.pack(side="left", fill="y")

        action_row = ttk.Frame(self.toc_tab, style="Card.TFrame", height=56)
        action_row.pack(fill="x", pady=(12, 0))
        action_row.pack_propagate(False)

        self.remove_toc_button = self.create_action_button(
            action_row,
            text="删除选中项",
            command=self.remove_selected_toc_entry,
            width=12,
        )
        self.remove_toc_button.pack(side="left", padx=(0, 8), pady=8)

        self.clear_toc_button = self.create_action_button(
            action_row,
            text="恢复自动目录",
            command=self.restore_auto_toc_entries,
            width=12,
        )
        self.clear_toc_button.pack(side="left", pady=8)

        self.generate_toc_button = self.create_action_button(
            action_row,
            text="生成未压缩完整版并进入压缩",
            command=self.handle_generate_toc,
            primary=True,
            width=26,
        )
        self.generate_toc_button.pack(side="right", pady=8)

    def select_toc_output_file(self) -> None:
        suggested_name = "完整文档_未压缩.pdf"
        if self.reference_pdf:
            suggested_name = build_default_uncompressed_path(self.reference_pdf).name

        selected = filedialog.asksaveasfilename(
            title="选择未压缩完整版输出路径",
            defaultextension=".pdf",
            initialfile=suggested_name,
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if selected:
            self.toc_output_var.set(selected)
            self.compress_output_var.set(str(build_default_compressed_path(Path(selected))))

    def add_toc_entry(self) -> None:
        item = self.toc_item_var.get().strip()
        page = self.toc_page_var.get().strip()
        if not item:
            messagebox.showwarning("输入不完整", "请输入目录项内容。")
            return
        if not page:
            messagebox.showwarning("输入不完整", "请输入对应页码。")
            return

        self.toc_tree.insert("", "end", values=(item, page))
        self.toc_item_var.set("")
        self.toc_page_var.set("")

    def remove_selected_toc_entry(self) -> None:
        for item_id in self.toc_tree.selection():
            self.toc_tree.delete(item_id)

    def clear_toc_entries(self) -> None:
        for item_id in self.toc_tree.get_children():
            self.toc_tree.delete(item_id)

    def restore_auto_toc_entries(self) -> None:
        self.clear_toc_entries()
        for item_text, page_text in self.auto_toc_entries:
            self.toc_tree.insert("", "end", values=(item_text, page_text))

    def collect_toc_entries(self) -> List[Tuple[str, str]]:
        entries: List[Tuple[str, str]] = []
        for item_id in self.toc_tree.get_children():
            values = self.toc_tree.item(item_id, "values")
            if len(values) < 2:
                continue
            item_text = str(values[0]).strip()
            page_text = str(values[1]).strip()
            if not item_text or not page_text:
                raise ValueError("目录列表存在空值，请删除或补全后再生成。")
            entries.append((item_text, page_text))

        if not entries:
            raise ValueError("请先在目录列表中添加至少一条目录项。")
        return entries

    def handle_generate_toc(self) -> None:
        try:
            if not self.numbered_pdf_path or not self.numbered_pdf_path.exists():
                raise ValueError("请先完成步骤 2 添加页码。")

            title_size = int(self.toc_title_size_var.get())
            body_size = int(self.toc_body_size_var.get())
            if not 10 <= title_size <= 72:
                raise ValueError("标题字号需在 10 到 72 之间。")
            if not 8 <= body_size <= 48:
                raise ValueError("目录字号需在 8 到 48 之间。")

            entries = self.collect_toc_entries()
            output_text = self.toc_output_var.get().strip()
            if not output_text:
                if not self.reference_pdf:
                    raise ValueError("请设置未压缩完整版输出路径。")
                output_text = str(build_default_uncompressed_path(self.reference_pdf))
                self.toc_output_var.set(output_text)

            output_pdf = normalize_path(output_text)
            output_pdf.parent.mkdir(parents=True, exist_ok=True)

            self.status_var.set("步骤 3 处理中：正在生成目录并输出未压缩完整版...")
            self.update_idletasks()

            final_uncompressed = prepend_toc_pages(
                numbered_pdf=self.numbered_pdf_path,
                entries=entries,
                title=self.toc_title_var.get().strip() or "目录",
                output_pdf=output_pdf,
                title_font_choice=self.toc_title_font_var.get(),
                title_font_size=title_size,
                body_font_choice=self.toc_body_font_var.get(),
                body_font_size=body_size,
            )

            self.final_uncompressed_path = final_uncompressed
            self.compress_source_var.set(str(final_uncompressed))
            self.compress_output_var.set(str(build_default_compressed_path(final_uncompressed)))
            self.notebook.tab(3, state="normal")
            self.notebook.select(3)

            self.status_var.set("步骤 3 完成：未压缩完整版已输出，并已写入目录书签。请进行步骤 4 压缩。")
            messagebox.showinfo(
                "步骤 3 完成",
                f"未压缩完整版已生成（含目录书签）：\n{final_uncompressed}",
            )
        except Exception as exc:
            self.status_var.set("步骤 3 失败，请检查输入后重试。")
            messagebox.showerror("步骤 3 失败", str(exc))

    # -----------------------------
    # 步骤4：压缩
    # -----------------------------
    def build_compress_tab(self) -> None:
        self.build_path_row(
            self.compress_tab,
            "未压缩完整版（步骤 3 生成）",
            self.compress_source_var,
            None,
            readonly=True,
        )
        self.build_path_row(
            self.compress_tab,
            "压缩版输出路径",
            self.compress_output_var,
            self.select_compress_output_file,
        )

        option_row = ttk.Frame(self.compress_tab, style="Card.TFrame")
        option_row.pack(fill="x", pady=(8, 0))

        ttk.Label(option_row, text="目标压缩大小（MB）", style="Label.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(option_row, textvariable=self.target_size_mb_var, width=12).grid(
            row=1,
            column=0,
            sticky="w",
            pady=(6, 0),
        )

        ttk.Label(option_row, text="允许超出目标（%）", style="Label.TLabel").grid(row=0, column=1, sticky="w", padx=(20, 0))
        ttk.Entry(option_row, textvariable=self.compress_tolerance_percent_var, width=12).grid(
            row=1,
            column=1,
            sticky="w",
            padx=(20, 0),
            pady=(6, 0),
        )

        hint_text = (
            "说明：系统会优先输出最接近目标且尽量保留清晰度的结果；"
            "可设置允许超出比例，减少明显过压缩。"
        )
        ttk.Label(option_row, text=hint_text, style="SubTitle.TLabel").grid(
            row=0,
            column=2,
            rowspan=2,
            sticky="w",
            padx=(20, 0),
        )

        action_row = ttk.Frame(self.compress_tab, style="Card.TFrame", height=56)
        action_row.pack(fill="x", pady=(14, 0))
        action_row.pack_propagate(False)

        self.compress_execute_button = self.create_action_button(
            action_row,
            text="压缩并输出最终文件",
            command=self.handle_compress_pdf,
            primary=True,
            width=22,
        )
        self.compress_execute_button.pack(side="right", pady=8)

    def select_compress_output_file(self) -> None:
        if self.final_uncompressed_path:
            suggested = build_default_compressed_path(self.final_uncompressed_path).name
        else:
            suggested = "完整文档_压缩版.pdf"

        selected = filedialog.asksaveasfilename(
            title="选择压缩版输出路径",
            defaultextension=".pdf",
            initialfile=suggested,
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if selected:
            self.compress_output_var.set(selected)

    def handle_compress_pdf(self) -> None:
        try:
            if not self.final_uncompressed_path or not self.final_uncompressed_path.exists():
                raise ValueError("请先完成步骤 3，生成未压缩完整版。")

            target_mb = float(self.target_size_mb_var.get().strip())
            if target_mb <= 0:
                raise ValueError("目标压缩大小必须大于 0 MB。")
            tolerance_percent = float(self.compress_tolerance_percent_var.get().strip())
            if tolerance_percent < 0 or tolerance_percent > 100:
                raise ValueError("允许超出比例需在 0 到 100 之间。")

            output_text = self.compress_output_var.get().strip()
            if not output_text:
                output_text = str(build_default_compressed_path(self.final_uncompressed_path))
                self.compress_output_var.set(output_text)

            output_pdf = normalize_path(output_text)
            output_pdf.parent.mkdir(parents=True, exist_ok=True)

            self.status_var.set("步骤 4 处理中：正在压缩 PDF...")
            self.update_idletasks()

            result = compress_pdf_to_target(
                input_pdf=self.final_uncompressed_path,
                output_pdf=output_pdf,
                target_mb=target_mb,
                work_dir=self.runtime_dir,
                tolerance_percent=tolerance_percent,
            )

            self.final_compressed_path = output_pdf
            self.status_var.set(
                "步骤 4 完成：已输出未压缩完整版与压缩版最终文件。"
            )

            message = (
                f"未压缩完整版：\n{self.final_uncompressed_path}\n\n"
                f"压缩版：\n{output_pdf}\n\n"
                f"压缩方式：{result['method']}\n"
                f"原大小：{result['source_mb']:.2f} MB\n"
                f"目标：{result['target_mb']:.2f} MB\n"
                f"允许上限：{result['allowed_upper_mb']:.2f} MB（+{result['tolerance_percent']:.1f}%）\n"
                f"实际：{result['actual_mb']:.2f} MB"
            )

            if not result["target_met"]:
                message += "\n\n提示：目标较苛刻，当前结果仍高于允许上限，已输出可达最小结果。"
            elif not result["strict_target_met"]:
                message += "\n\n提示：结果略高于目标值，但在允许范围内，可保留更多清晰度。"

            messagebox.showinfo("步骤 4 完成", message)
        except Exception as exc:
            self.status_var.set("步骤 4 失败，请检查输入后重试。")
            messagebox.showerror("步骤 4 失败", str(exc))

    def handle_close(self) -> None:
        try:
            shutil.rmtree(self.runtime_dir, ignore_errors=True)
        finally:
            self.destroy()


if __name__ == "__main__":
    app = PDFWorkflowApp()
    app.mainloop()
