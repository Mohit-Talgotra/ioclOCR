from transformers import TrOCRProcessor, VisionEncoderDecoderModel
from PIL import Image

image = Image.open("/home/talgotram/Repos/ioclOCR/extracted_images/page_1.png").convert("RGB")

processor = TrOCRProcessor.from_pretrained("microsoft/trocr-base-handwritten")
model = VisionEncoderDecoderModel.from_pretrained("microsoft/trocr-base-handwritten")

pixel_values = processor(images=image, return_tensors="pt").pixel_values
print(pixel_values)

generated_ids = model.generate(pixel_values)
print(generated_ids)

generated_text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]

print("OCR Output:", generated_text)