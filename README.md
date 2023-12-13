Document Processing and Analysis
Project Overview
"Document Processing and Analysis" is a Python-based tool designed to process PDF documents and extract valuable information. It utilizes several advanced libraries to handle tasks such as converting PDF pages to images, performing Optical Character Recognition (OCR), and object detection using YOLO v8. The final output is a detailed DOCX file containing extracted text, recognized entities, and object detection results.

Key Libraries and Technologies
PyMuPDF: A Python library to access and interact with PDF files.
Pillow (PIL): A Python Imaging Library used for opening, manipulating, and saving many different image file formats.
pytesseract: An OCR tool for Python that recognizes text from images.
spaCy: An open-source software library for advanced Natural Language Processing (NLP).
python-docx: A Python library for creating and updating Microsoft Word (.docx) files.
ultralytics YOLO v8: An object detection model that identifies objects in images.
FastAPI: A modern, fast web framework for building APIs with Python.
Uvicorn: An ASGI server for Python.
OCR with Tesseract
Tesseract OCR is an optical character recognition engine for various operating systems. It is used here to extract text from images converted from PDF pages.

Object Detection with YOLO v8
YOLO (You Only Look Once) v8 is a state-of-the-art, real-time object detection system, implemented in this project to identify objects within the images extracted from the PDF documents.

Setup and Installation:
```
Python 3.x
```

Clone the Repository
Clone the project repository to your local machine using the following command:

```
git clone https://github.com/leeladh/Mypdfsystem/tree/pdf
```
Conda environment (recommended for managing dependencies):

```
conda activate pdfsystem
```
Install Required Libraries
Install all required libraries using the requirements.txt file:

```
pip install -r requirements.txt
```

Running the Application
To run the application, execute the following command in the root directory of the project:

```
Python3 app.py
```

