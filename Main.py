from Ocr import pdf_to_images
from Ner import ocr_to_json
from Object_detecion import object_detection
from Docx import create_docx

pdf_path = '/home/leeladhar/Downloads/Red_Positive/1506.02640.pdf'
model_path = '/home/leeladhar/Downloads/Red_Positive/yolov8n.pt'
docx_path = '/home/leeladhar/Downloads/Red_Positive/docx.docx'

images = pdf_to_images(pdf_path)
ocr_data = ocr_to_json(images)
detection_data = object_detection(images, model_path)
create_docx(ocr_data, detection_data, docx_path)
