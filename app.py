from dotenv import load_dotenv
import os
from flask import Flask, request
from flask_cors import CORS
from dpformatter import dpformatter

load_dotenv()

app = Flask(__name__)
CORS(app)


@app.route("/", methods=["GET"])
def welcome():
    return "i am awake"


@app.route("/api/format", methods=["POST"])
def format_xlsx_base64():
    data = request.get_json()
    if "apikey" not in data or "xlsxBase64" not in data:
        return {"error": "bad request"}, 400
    if "apikey" in data and data["apikey"] != os.environ.get("APIKEY"):
        return {"error": "unauthorized request"}, 401
    return dpformatter(data["xlsxBase64"])
