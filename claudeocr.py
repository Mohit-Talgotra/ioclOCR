import requests
import base64

API_KEY = "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"
API_URL = "https://api.anthropic.com/v1/messages"

def encode_image_to_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

image_b64 = encode_image_to_base64("extracted_images/page_1.png")

headers = {
    "x-api-key": API_KEY,
    "anthropic-version": "2023-06-01",
    "Content-Type": "application/json"
}

data = {
    "model": "claude-3-7-sonnet-20250219",
    "max_tokens": 1024,
    "messages": [
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": "Give me a digital form of the table in this image"
                },
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": image_b64
                    }
                }
            ]
        }
    ]
}

response = requests.post(API_URL, headers=headers, json=data)

print(response.json()["content"][0]["text"])