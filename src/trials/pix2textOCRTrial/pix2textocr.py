from pix2text import Pix2Text

p2t = Pix2Text(analyzer_config={"device": "cuda"})

image_path = "/home/talgotram/Repos/ioclOCR/extracted_images/page_1.png"

results = p2t.recognize(image_path, return_text=False)

print("Detected Text:")
print(results["text"])

print("\nDetected Segments:")
for seg in results["segments"]:
    print(f"Type: {seg['type']}, Text: {seg['text']}")