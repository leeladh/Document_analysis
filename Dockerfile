# Use an official Python runtime as a parent image
FROM python:3.8-slim

RUN apt-get update && apt-get install -y tesseract-ocr libtesseract-dev libgl1-mesa-glx

# Set the working directory in the container
WORKDIR /code/redpositive

# Copy the current directory contents into the container at /usr/src/app
COPY . .

# Install any needed packages specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

RUN python -m spacy download en_core_web_sm

EXPOSE 8000
# # Run app.py when the container launches
CMD ["uvicorn", "app:app", "--reload"]

#sudo docker run -it -p 8000:8000 mypdfapp:v2