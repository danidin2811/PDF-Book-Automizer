import cv2
import numpy as np
import fitz  # PyMuPDF
from pathlib import Path


def resize_and_crop(img, target_w_px, target_h_px):
    """
    Resizes and crops the image to match exact pixel dimensions while maintaining quality.
    """
    h, w = img.shape[:2]
    target_aspect = target_w_px / target_h_px
    input_aspect = w / h

    if input_aspect > target_aspect:
        # Image is too wide: scale by height and crop sides
        new_h = target_h_px
        new_w = int(w * (target_h_px / h))
        resized = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_LANCZOS4)
        start_x = (new_w - target_w_px) // 2
        return resized[:, start_x:start_x + target_w_px]
    else:
        # Image is too tall: scale by width and crop top/bottom
        new_w = target_w_px
        new_h = int(h * (target_w_px / w))
        resized = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_LANCZOS4)
        start_y = (new_h - target_h_px) // 2
        return resized[start_y:start_y + target_h_px, :]


def normalize_document_image(image_path: Path, output_path: Path, target_h_in: float, target_w_in: float = None):
    dpi = 300
    img_array = np.fromfile(str(image_path), np.uint8)
    img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)

    if img is None:
        return False

    # 1. Deskew logic remains the same
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    lines = cv2.HoughLinesP(edges, 1, np.pi / 180, 100, minLineLength=100, maxLineGap=10)

    angle = 0
    if lines is not None:
        angles = [np.arctan2(l[0][3] - l[0][1], l[0][2] - l[0][0]) * 180 / np.pi for l in lines]
        angle = np.median(angles)

    (h, w) = img.shape[:2]
    matrix = cv2.getRotationMatrix2D((w // 2, h // 2), angle, 1.0)
    rotated = cv2.warpAffine(img, matrix, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

    # 2. Crop to Content
    new_gray = cv2.cvtColor(rotated, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(new_gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    if not contours:
        return False

    c = max(contours, key=cv2.contourArea)
    x, y, w_c, h_c = cv2.boundingRect(c)
    cropped = rotated[y:y + h_c, x:x + w_c]

    # 3. Handle Scaling and DPI logic
    target_h_px = int(target_h_in * dpi)

    if target_w_in:
        # Exact Width and Height requested
        target_w_px = int(target_w_in * dpi)
        final_img = resize_and_crop(cropped, target_w_px, target_h_px)
    else:
        # Proportional scaling based on Height only
        aspect_ratio = w_c / h_c
        target_w_px = int(target_h_px * aspect_ratio)
        final_img = cv2.resize(cropped, (target_w_px, target_h_px), interpolation=cv2.INTER_LANCZOS4)

    # Unicode-safe save
    ext = output_path.suffix if output_path.suffix else '.png'
    is_success, buffer = cv2.imencode(ext, final_img)
    if is_success:
        with open(output_path, "wb") as f:
            f.write(buffer)
        return True
    return False


def pdf_page_to_normalized_image(pdf_path: Path):
    # Prompting for dimensions
    try:
        h_in = float(input("Enter desired height in inches (e.g. 6.107): "))
        w_raw = input("Enter desired width in inches (leave empty for proportional): ").strip()
        w_in = float(w_raw) if w_raw else None
    except ValueError:
        print("Invalid number entered.")
        return

    doc = fitz.open(pdf_path)
    pix = doc.load_page(0).get_pixmap(dpi=300)
    temp_img = pdf_path.with_suffix(".png")
    pix.save(str(temp_img))

    norm_path = pdf_path.with_name(f"{pdf_path.stem}_normalized.png")
    if normalize_document_image(temp_img, norm_path, h_in, w_in):
        print(f"Successfully saved {norm_path.name} at 300 DPI")
    temp_img.unlink()


if __name__ == "__main__":
    path = Path(input("Enter cover PDF path: ").replace('"', '').strip())
    pdf_page_to_normalized_image(path)