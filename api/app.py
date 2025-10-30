import os
import shutil
from USAB_serverless import run_model
from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs
import tempfile

def handler(request):
    """
    Vercel Python serverless handler
    """
    try:
        # Uploaded file is in request['files']['file']
        uploaded_file = request.files.get('file')
        if not uploaded_file:
            return {"statusCode": 400, "body": "No file uploaded"}

        # Save uploaded input to temp
        input_path = os.path.join(tempfile.gettempdir(), uploaded_file.filename)
        with open(input_path, "wb") as f:
            f.write(uploaded_file.file.read())

        # Copy model template to temp
        model_template = "Bespoke Model - US - v2.xlsm"
        model_path = os.path.join(tempfile.gettempdir(), model_template)
        shutil.copyfile(model_template, model_path)

        # Output path
        output_path = os.path.join(tempfile.gettempdir(), "Processed_Model.xlsm")

        # Run model
        run_model(input_path, model_path, output_path)

        # Return the file
        with open(output_path, "rb") as f:
            content = f.read()

        return {
            "statusCode": 200,
            "headers": {
                "Content-Type": "application/vnd.ms-excel",
                "Content-Disposition": 'attachment; filename="Processed_Model.xlsm"'
            },
            "body": content,
            "isBase64Encoded": True
        }

    except Exception as e:
        return {"statusCode": 500, "body": str(e)}
