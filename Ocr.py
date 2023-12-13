import fitz
from PIL import Image
import io

def pdf_to_images(pdf_path):
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))
        images.append(img)
    return images
