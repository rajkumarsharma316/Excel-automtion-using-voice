import os
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")
MODEL_NAME = os.getenv("MODEL_NAME", "models/gemini-pro-latest")

if not API_KEY:
    raise RuntimeError("❌ Missing GEMINI_API_KEY in .env")

genai.configure(api_key=API_KEY)
model = genai.GenerativeModel(MODEL_NAME)


def clean_json(text):
    """Remove ```json … ``` wrappers."""
    text = text.strip()
    if text.startswith("```"):
        text = text.replace("```json", "").replace("```", "").strip()
    return text


def interpret_command(text):
    prompt = f"""
You are an Excel voice assistant. Convert the spoken text into a JSON command.

User said: "{text}"

You MUST return ONLY valid JSON.

### Allowed actions:

Basic:
- "write" → {{"action":"write","cell":"A1","value":"hello"}}
- "delete_cell" → {{"action":"delete_cell","cell":"B5"}}
- "insert_row" → {{"action":"insert_row","row":3}}
- "insert_column" → {{"action":"insert_column","column":"C"}}

General Excel:
- "sum_column" → {{"action":"sum_column","column":"A"}}
- "format_bold" → {{"action":"format_bold","column":"B"}}
- "create_chart" → {{"action":"create_chart","x_column":"A","y_column":"B"}}
- "run_regression" → {{"action":"run_regression","x_column":"A","y_column":"B"}}
- "sort_column" → {{"action":"sort_column","column":"A","order":"asc"}}
- "filter_values" → {{"action":"filter_values","column":"A","condition":">10"}}

If the command is unclear ALWAYS return:
{{"action":"unknown"}}

Return JSON only. No explanation.
"""

    try:
        response = model.generate_content(prompt)
        raw = response.text.strip()
        print("\nAI Response:", raw)

        cleaned = clean_json(raw)
        print("Cleaned:", cleaned)

        return eval(cleaned)

    except Exception as e:
        print("❌ Error interpreting command:", e)
        return {"action": "unknown"}
