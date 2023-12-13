import pytesseract
import spacy

nlp = spacy.load("en_core_web_sm")

def ocr_to_json(images):    
    json_data = {}
    for page_number, img in enumerate(images, start=1):
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        words = data['text']
        text = ' '.join(words)
        doc = nlp(text)
        entities = [{"word": ent.text, "label": ent.label_} for ent in doc.ents]
        json_data[str(page_number)] = {"words": words, "entities": entities}
    return json_data
