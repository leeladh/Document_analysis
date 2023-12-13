from fastapi import FastAPI, UploadFile
from fastapi.responses import FileResponse
import os
from Ocr import pdf_to_images
from Ner import ocr_to_json
from Object_detecion import object_detection
from Docx import create_docx


# Initialize FastAPI app
app = FastAPI()

model_path = 'yolov8n.pt'
# Add your existing functions here (pdf_to_images, ocr_to_json, object_detection, create_docx)

def process_pdf(pdf_path: str):
    images = pdf_to_images(pdf_path)
    ocr_data = ocr_to_json(images)
    detection_data = object_detection(images, model_path)
    docx_path = pdf_path.rsplit('.', 1)[0] + '.docx'
    create_docx(ocr_data, detection_data, docx_path)
    return docx_path

@app.post("/upload/")
async def upload_pdf(file: UploadFile):
    file_location = os.path.join(file.filename)
    try:
        with open(file_location, "wb+") as file_object:
            file_object.write(file.file.read())
        # Process the PDF
        docx_path = process_pdf(file_location)

        # Return the processed DOCX file
        return FileResponse(docx_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document', filename=os.path.basename(docx_path))

    except Exception as e:
        return {"Response": "Failed."}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="localhost", port=8000)
