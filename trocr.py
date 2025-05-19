from transformers import TrOCRProcessor, VisionEncoderDecoderModel
from PIL import Image
import torch

image = Image.open("extracted_images/page_1.png").convert("RGB")

processor = TrOCRProcessor.from_pretrained("microsoft/trocr-base-handwritten")
model = VisionEncoderDecoderModel.from_pretrained("microsoft/trocr-base-handwritten")

pixel_values = processor(images=image, return_tensors="pt").pixel_values
print("here1")

generated_ids = model.generate(pixel_values)
print("here2")

generated_text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]
print("here3")

print("OCR Output:", generated_text)
print("here4")