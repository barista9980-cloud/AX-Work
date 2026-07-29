import sys
from google import genai

sys.stdout.reconfigure(encoding='utf-8')

user_api_key = os.environ.get("GEMINI_API_KEY", "")

client = genai.Client(api_key=user_api_key)

print("Listing supported models for this API key...")
try:
    for m in client.models.list():
        print(f"Model: {m.name} | Supported Actions: {m.supported_actions}")
except Exception as e:
    print("Error listing models:", e)
