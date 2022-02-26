from __future__ import print_function
import os
import base64
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException


def cloudmersive(input_file):
    # Configure API key authorization: Apikey
    configuration = cloudmersive_convert_api_client.Configuration()
    configuration.api_key["Apikey"] = os.environ.get("CLOUDMERSIVE_API_KEY")
    # Uncomment below to setup prefix (e.g. Bearer) for API key, if needed
    # configuration.api_key_prefix['Apikey'] = 'Bearer'

    # create an instance of the API class
    api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(
        cloudmersive_convert_api_client.ApiClient(configuration)
    )

    try:
        # Convert Excel XLSX Spreadsheet to PDF
        api_response = api_instance.convert_document_xlsx_to_pdf(input_file)
    except ApiException as e:
        print(
            "Exception when calling ConvertDocumentApi->convert_document_xlsx_to_pdf: %s\n"
            % e
        )

    response_bytes = eval(api_response)
    pdf_base64 = base64.encodebytes(response_bytes).decode("utf-8")

    return pdf_base64
