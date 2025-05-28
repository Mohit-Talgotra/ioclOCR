import requests
import base64

API_KEY = "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"  # Replace with your actual API key
API_URL = "https://api.anthropic.com/v1/messages"

def encode_image_to_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def send_image_to_claude(image_path, prompt_text):
    try:
        # Encode the image
        image_b64 = encode_image_to_base64(image_path)
        
        headers = {
            "x-api-key": API_KEY,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": "claude-3-7-sonnet-20250219",
            "max_tokens": 4096,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": prompt_text
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
        
        print("Sending request to Claude API...")
        response = requests.post(API_URL, headers=headers, json=data)
        
        # Check if the request was successful
        if response.status_code == 200:
            result = response.json()
            return result["content"][0]["text"]
        else:
            print(f"Error: API request failed with status code {response.status_code}")
            print(f"Response: {response.text}")
            return None
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    image_path = "/home/talgotram/Repos/ioclOCR/extracted_images/page_1.png"  # Update with your actual image path
    prompt = "Give me a digital form of the table in this image"
    
    result = send_image_to_claude(image_path, prompt)
    
    if result:
        print("\nClaude's Response:")
        print(result)