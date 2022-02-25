import os
import requests
import json


def getoutpdf(xlsx_output_base64):
    url = "https://getoutpdf.com/api/convert/document-to-pdf"
    api_key = os.environ.get("GETOUTPDF_API_KEY")
    request_data = {"api_key": api_key, "document": xlsx_output_base64}

    res = requests.post(url, data=request_data)
    res_json = json.loads(res.text)

    return res_json["pdf_base64"]
