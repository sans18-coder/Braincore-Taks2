
from ultralytics import YOLO
from ultralytics.utils.plotting import Annotator, colors
import pandas as pd
from transformers import VisionEncoderDecoderModel,TrOCRProcessor
from PIL import Image
import glob as glob
import os
import cv2
od_model = YOLO('./OD-Model/best.pt')
names = od_model.names

processor = TrOCRProcessor.from_pretrained('microsoft/trocr-small-handwritten')
ocr_model = VisionEncoderDecoderModel.from_pretrained('./OCR-Model')

crop_dir_name = './temp'
data= './Test'
def crop(image_path):
  original_filename = os.path.splitext(os.path.basename(image_path))[0]
  image_path = os.path.abspath(image_path)
  im0 = cv2.imread(image_path)
  im0 = cv2.resize(im0, (775, 925))
  if im0 is None:
      raise ValueError(f"Error reading image file {image_path}")

  results = od_model.predict(im0, show=False)
  boxes = results[0].boxes.xyxy.cpu().tolist()
  clss = results[0].boxes.cls.cpu().tolist()
  annotator = Annotator(im0, line_width=2, example=names)

  if boxes is not None:
      for box, cls in zip(boxes, clss):
          class_name = names[int(cls)]
          annotator.box_label(box, color=colors(int(cls), True), label=class_name)
          crop_obj = im0[int(box[1]): int(box[3]), int(box[0]): int(box[2])]
          crop_filename = os.path.join(crop_dir_name, f"{original_filename}-{class_name}.png")
          cv2.imwrite(crop_filename, crop_obj) 
          
def read_image_to_RGB(image_path):
    image = Image.open(image_path).convert('RGB')
    return image

def ocr(image, processor, model):
    pixel_values = processor(image, return_tensors='pt').pixel_values
    generated_ids = model.generate(pixel_values)
    generated_text = processor.batch_decode(generated_ids, skip_special_tokens=True, clean_up_tokenization_spaces=True)[0]
    return generated_text


def input_text_to_excel(data_path):
    image_paths = glob.glob(data_path)
    
    data = {}
    
    for image_path in image_paths:
        image = read_image_to_RGB(image_path)
        text = ocr(image, processor, ocr_model)
        
        print(f"Text from {image_path}: {text}")
        
        filename = os.path.basename(image_path)
        image_name = filename.split('-')[0]
        print(f'Filename: {filename}, image_name: {image_name}') 
        
        if image_name not in data:
            data[image_name] = {'suara_paslon1': '', 'suara_paslon2': '', 'suara_paslon3': ''}
        
        if 'voting_paslon1' in filename:
            data[image_name]['suara_paslon1'] = text
        elif 'voting_paslon2' in filename:
            data[image_name]['suara_paslon2'] = text
        elif 'voting-paslon3' in filename:
            data[image_name]['suara_paslon3'] = text
    
    df_new = pd.DataFrame.from_dict(data, orient='index').reset_index()
    df_new.rename(columns={'index': 'image_name'}, inplace=True)
    
    print('New DataFrame:')
    print(df_new)  
    
    file_path =  './output_file/hasil_ocr.xlsx'
    if os.path.exists(file_path):
        df_existing = pd.read_excel(file_path)
        df_combined = pd.concat([df_existing, df_new]).drop_duplicates(subset='image_name', keep='last')
    else:
        df_combined = df_new
    
    print('Combined DataFrame:')
    print(df_combined)  # Debugging
    
    df_combined.to_excel(file_path, index=False)
    
    for file_path in image_paths:
        os.remove(file_path)
    
def execute(data_path):
    image_paths = glob.glob(data_path)
    for image_path in image_paths:
        crop(image_path)
        input_text_to_excel(os.path.join(crop_dir_name, '*'))