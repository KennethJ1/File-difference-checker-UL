import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, letter
import tempfile
import shutil
import os
from PIL import Image, ImageDraw
import numpy as np
from skimage.metrics import structural_similarity as ssim
from typing import Any, List, Dict, Tuple, Optional


def render_pdf_to_images(pdf_path: str, out_dir: str) -> list:
    """Render each page of pdf_path to a PNG in out_dir and return list of file paths."""
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        pix = doc.load_page(page_num).get_pixmap()
        img_path = os.path.join(out_dir, f"page_{page_num}_{os.path.basename(pdf_path)}.png")
        pix.save(img_path)
        images.append(img_path)
    return images


def extract_page_text(pdf_path: str) -> list:
    doc = fitz.open(pdf_path)
    return [page.get_text() for page in doc]


def highlight_text_differences(img_path: str, pdf_path: str, page_num: int, other_text: str, output_path: str, color=(255,0,0,30)) -> str:
    """Highlight words on the raster image (img_path) that differ between the page text and other_text.
    Uses PDF word coordinates from pdf_path at page_num to draw translucent rectangles on the image.
    """
    img = Image.open(img_path).convert("RGBA")
    overlay = Image.new("RGBA", img.size, (0,0,0,0))
    draw = ImageDraw.Draw(overlay)

    # load words and their coordinates from the PDF page
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num)
    words = page.get_text("words")  # list of tuples: (x0,y0,x1,y1, word, ...)

    page_words = [w[4] for w in words]
    set_page_words = set(page_words)
    set_other = set(str(other_text).split())
    diff_words = set_page_words.symmetric_difference(set_other)

    # fitz coordinates are in points relative to page; the pixmap we rendered earlier should match these dimensions
    for w in words:
        word_text = w[4]
        if word_text in diff_words:
            x0, y0, x1, y1 = w[0], w[1], w[2], w[3]
            # convert to int pixel coords
            try:
                rect = [int(x0), int(y0), int(x1), int(y1)]
                draw.rectangle(rect, fill=color)
            except Exception:
                continue

    img_highlighted = Image.alpha_composite(img, overlay)
    img_highlighted.save(output_path)
    return output_path


def highlight_image_differences(img1_path: str, img2_path: str, output_path: str,
                                color1=(255,0,0,40), color2=(0,255,0,40), threshold: float = 0.6, box_size: int = 16) -> str:
    """Produce a translucent overlay on img1 showing regions that differ from img2 using SSIM."""
    img1 = Image.open(img1_path).convert("RGBA")
    img2 = Image.open(img2_path).convert("RGBA")

    def remove_watermark(img):
        img_gray = img.convert("L")
        mask = img_gray.point(lambda x: 255 if x > 220 else 0)
        img_no_wm = img.copy()
        img_no_wm.paste((255,255,255,0), mask=mask)
        return img_no_wm

    img1_clean = remove_watermark(img1)
    img2_clean = remove_watermark(img2)

    arr1 = np.array(img1_clean.convert("L"))
    arr2 = np.array(img2_clean.convert("L"))
    _, diff_map = ssim(arr1, arr2, full=True)
    diff_map = (1 - diff_map)

    overlay = Image.new("RGBA", img1.size, (0,0,0,0))
    draw = ImageDraw.Draw(overlay)
    for y in range(0, img1.size[1], box_size):
        for x in range(0, img1.size[0], box_size):
            region = diff_map[y:y+box_size, x:x+box_size]
            if np.mean(region) > threshold:
                draw.rectangle([x, y, x+box_size, y+box_size], fill=color1)
    img_highlighted = Image.alpha_composite(img1, overlay)
    img_highlighted.save(output_path)
    return output_path


def save_side_by_side_pdf(images1, images2, texts1, texts2, output_path,
                          tmp_dir: str, vis_color=(255,255,0,20), text_color1=(255,0,0,30), text_color2=(0,255,0,30),
                          threshold: float = 0.6, box_size: int = 16, progress_cb=None):
    c = canvas.Canvas(output_path, pagesize=landscape(letter))
    width, height = landscape(letter)
    num_pages = min(len(images1), len(images2), len(texts1), len(texts2))

    def _progress(p: int):
        try:
            if callable(progress_cb):
                progress_cb(int(p))
        except Exception:
            pass

    for i in range(num_pages):
        _progress(30 + int(20 * (i / max(1, num_pages))))
        # Visual comparison highlight
        vis_img1 = os.path.join(tmp_dir, f"vis_file1_page_{i}.png")
        vis_img2 = os.path.join(tmp_dir, f"vis_file2_page_{i}.png")
        highlight_image_differences(images1[i], images2[i], vis_img1, color1=vis_color, threshold=threshold, box_size=box_size)
        highlight_image_differences(images2[i], images1[i], vis_img2, color1=vis_color, threshold=threshold, box_size=box_size)

        _progress(50 + int(20 * (i / max(1, num_pages))))
        # Text-based highlight overlays (use PDF page coordinates)
        highlighted_img1 = os.path.join(tmp_dir, f"highlighted_file1_page_{i}.png")
        highlighted_img2 = os.path.join(tmp_dir, f"highlighted_file2_page_{i}.png")
        # images1/2 correspond to rendered pages from original PDFs; we need the original pdf paths and page nums
        # We assume caller tracks page order consistently and will pass pdf paths as texts1/texts2 sources.
        # texts1/texts2 are full page text strings; we still need pdf paths to find word positions. Caller will set attributes in tmp metadata.
        # For compatibility with existing helpers, infer pdf paths from a small metadata file if present; otherwise skip text overlay.
        # To keep simple and robust, we will draw text diffs by checking words positions from the original PDFs via a naming convention stored in texts1/_paths.
        # For this implementation, texts1 and texts2 are tuples: (pdf_path, page_text)
        try:
            pdf1_path, page_text1 = texts1[i]
            pdf2_path, page_text2 = texts2[i]
            highlight_text_differences(vis_img1, pdf1_path, i, page_text2, highlighted_img1, color=text_color1)
            highlight_text_differences(vis_img2, pdf2_path, i, page_text1, highlighted_img2, color=text_color2)
            out_img1 = highlighted_img1
            out_img2 = highlighted_img2
        except Exception:
            # fallback to visual-only images
            out_img1 = vis_img1
            out_img2 = vis_img2

        c.drawImage(out_img1, 40, 40, width=(width/2)-60, height=height-80)
        c.drawImage(out_img2, (width/2)+20, 40, width=(width/2)-60, height=height-80)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(40, height-30, f"File 1 - Page {i+1} (Visual/Text Diff)")
        c.drawString((width/2)+20, height-30, f"File 2 - Page {i+1} (Visual/Text Diff)")
        c.setFont("Helvetica", 10)
        c.drawString(40, 20, "Legend: Yellow = Visual/Layout Diff, Red = Text removed/changed from File 1, Green = Text added/changed in File 2")
        c.showPage()

    c.save()

    # cleanup temp visual images (but keep rendered originals removed by caller)
    for i in range(num_pages):
        for fname in [os.path.join(tmp_dir, f"vis_file1_page_{i}.png"), os.path.join(tmp_dir, f"vis_file2_page_{i}.png"),
                      os.path.join(tmp_dir, f"highlighted_file1_page_{i}.png"), os.path.join(tmp_dir, f"highlighted_file2_page_{i}.png")]:
            try:
                if os.path.exists(fname):
                    os.remove(fname)
            except Exception:
                pass


def compare_pdfs(file_list: list, options: dict = None, **kwargs) -> Any:
    """Generalized PDF comparison entry point.
    - file_list: [file1, file2]
    - options: dict supporting keys: vis_color, text_color1, text_color2, threshold, box_size, output_path, progress_cb, return_meta
    Returns output_path (string) or (output_path, meta) when return_meta True.
    """
    opts = {}
    if options and isinstance(options, dict):
        opts.update(options)
    opts.update(kwargs)

    vis_color = tuple(opts.get("vis_color", (255,255,0,20)))
    text_color1 = tuple(opts.get("text_color1", (255,0,0,30)))
    text_color2 = tuple(opts.get("text_color2", (0,255,0,30)))
    threshold = float(opts.get("threshold", 0.6))
    box_size = int(opts.get("box_size", 16))
    output_path_opt = opts.get("output_path")
    progress_cb = opts.get("progress_cb")
    return_meta = bool(opts.get("return_meta", False))

    def _progress(p: int):
        try:
            if callable(progress_cb):
                progress_cb(int(p))
        except Exception:
            pass

    if len(file_list) < 2:
        raise ValueError("At least two PDF files required for comparison")

    file1, file2 = file_list[0], file_list[1]

    _progress(1)
    tmp_dir = tempfile.mkdtemp(prefix="pdfdiff_")
    try:
        # render both PDFs to images in tmp_dir
        images1 = render_pdf_to_images(file1, tmp_dir)
        images2 = render_pdf_to_images(file2, tmp_dir)
        _progress(20)

        # extract texts and pair with pdf paths so save_side_by_side_pdf can access coords
        raw_texts1 = extract_page_text(file1)
        raw_texts2 = extract_page_text(file2)
        texts1 = [(file1, t) for t in raw_texts1]
        texts2 = [(file2, t) for t in raw_texts2]

        _progress(30)
        # create side-by-side PDF with highlights
        if not output_path_opt:
            output_path = os.path.join(tmp_dir, f"pdf_comparison_{os.path.basename(file1)}_vs_{os.path.basename(file2)}.pdf")
        else:
            output_path = output_path_opt

        save_side_by_side_pdf(images1, images2, texts1, texts2, output_path,
                              tmp_dir, vis_color=vis_color, text_color1=text_color1, text_color2=text_color2,
                              threshold=threshold, box_size=box_size, progress_cb=progress_cb)
        _progress(95)

        result = output_path
        meta = {"pages_compared": min(len(images1), len(images2)), "files": (file1, file2)}

    finally:
        # remove rendered page images
        try:
            for p in os.listdir(tmp_dir):
                fp = os.path.join(tmp_dir, p)
                if os.path.isfile(fp) and fp.endswith(".png"):
                    try:
                        os.remove(fp)
                    except Exception:
                        pass
        except Exception:
            pass
        # do not remove tmp_dir if user asked to keep output in it
        if output_path_opt is None:
            # keep directory until function exits then remove whole dir
            try:
                shutil.rmtree(tmp_dir)
            except Exception:
                pass

    _progress(100)
    if return_meta:
        return result, meta
    return result


if __name__ == "__main__":
    file1 = "test_samples/DOC-001179_v1_0 WEB_r0 PF2150-FMD_PRODUCT_MANUAL_dr2.pdf"
    file2 = "test_samples/DOC-001179 PF2150-F Product Manual_DRAFT_13FEB2025.pdf"
    output_file = "test_samples/pdfDiff_SideBySide_VisualText.pdf"

    compare_pdfs([file1, file2], output_path=output_file)