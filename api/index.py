from flask import Flask, request, render_template
import win32com.client
import pythoncom
from datetime import datetime
import json
import os
import requests

# -----------------------------------
# CREATE FLASK APP
# -----------------------------------

app = Flask(__name__)
# Prefer newer public models first, then fallback.
GEMINI_MODELS = [
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash",
    "gemini-1.5-flash-latest",
]


def execute_autocad_command(cad_command: str):
    pythoncom.CoInitialize()
    acad = win32com.client.Dispatch("AutoCAD.Application.26")
    acad.Visible = True
    doc = acad.ActiveDocument

    # Keep slight wait for command stability
    import time
    time.sleep(2)
    doc.SendCommand(cad_command + "\n")

    log = f"{datetime.now()} -> COMMAND executed | {cad_command}\n"
    with open("logs.txt", "a") as f:
        f.write(log)

    return log


def gemini_text_to_command(user_text: str) -> str:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY not set")

    prompt = f"""
You convert natural language into a single AutoCAD command string.
Return ONLY valid JSON in this exact shape:
{{"command":"<autocad_command_here>"}}
No markdown. No extra words.
If user asks for a circle and center is not provided, use center 0,0.
Example:
Input: create a circle of 5 radius
Output: {{"command":"_CIRCLE 0,0 5"}}

User input: {user_text}
"""

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1}
    }

    # Discover models available to this API key, then merge with local fallbacks.
    model_candidates = []
    try:
        list_url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
        list_response = requests.get(list_url, timeout=15)
        list_response.raise_for_status()
        models_payload = list_response.json()
        for model in models_payload.get("models", []):
            name = model.get("name", "")
            # name format is usually "models/<model-id>"
            if name.startswith("models/"):
                model_id = name.split("/", 1)[1]
            else:
                model_id = name
            methods = model.get("supportedGenerationMethods", [])
            if "generateContent" in methods and "gemini" in model_id:
                model_candidates.append(model_id)
    except requests.RequestException:
        # If listing fails, fallback to static list below.
        pass

    for static_model in GEMINI_MODELS:
        if static_model not in model_candidates:
            model_candidates.append(static_model)

    last_error = None
    result = None
    used_model = None
    for model_name in model_candidates:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        try:
            response = requests.post(url, json=payload, timeout=30)
            response.raise_for_status()
            result = response.json()
            used_model = model_name
            break
        except requests.RequestException as exc:
            last_error = exc

    if not result:
        raise RuntimeError(
            f"Gemini request failed for all discovered models. Last error: {last_error}"
        )

    ai_text = result["candidates"][0]["content"]["parts"][0]["text"].strip()
    parsed = json.loads(ai_text)
    cad_command = parsed.get("command", "").strip()
    if not cad_command:
        raise ValueError("Gemini did not return a valid command")
    if len(cad_command) > 200:
        raise ValueError("Generated command is too long")
    print(f"Gemini model used: {used_model}")
    return cad_command

# -----------------------------------
# HOME ROUTE
# -----------------------------------

@app.route("/")
def home():

    return {
        "status": "running",
        "message": "AutoCAD OFFSET MCP ACTIVE"
    }


@app.route("/chat")
def chat():
    return render_template("chat.html")

# -----------------------------------
# OFFSET ROUTE
# -----------------------------------

@app.route("/command", methods=["POST"])
def command():

    try:

        data = request.get_json(force=True)

        cad_command = data.get("command")

        if not cad_command:

            return {
                "status": "error",
                "message": "No command provided"
            }, 400

        print(f"Executing Command: {cad_command}")
        log = execute_autocad_command(cad_command)

        print(log)

        return {
            "status": "success",
            "command": cad_command
        }

    except Exception as e:

        print("ERROR:", e)

        return {
            "status": "error",
            "message": str(e)
        }, 500


@app.route("/ai-command", methods=["POST"])
def ai_command():

    try:
        data = request.get_json(force=True)
        user_text = data.get("text")

        if not user_text:
            return {
                "status": "error",
                "message": "No text provided"
            }, 400

        cad_command = gemini_text_to_command(user_text)
        print(f"AI Input: {user_text}")
        print(f"Generated Command: {cad_command}")

        log = execute_autocad_command(cad_command)
        print(log)

        return {
            "status": "success",
            "input": user_text,
            "command": cad_command
        }

    except Exception as e:
        print("ERROR:", e)
        return {
            "status": "error",
            "message": str(e)
        }, 500
# -----------------------------------
# START MCP SERVER
# -----------------------------------

if __name__ == "__main__":

    print("Starting AutoCAD OFFSET MCP...")

    app.run(
        host="127.0.0.1",
        port=5000,
        threaded=False
    )
