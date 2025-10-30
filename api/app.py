from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import shutil, os
from USAB_serverless import run_model

app = FastAPI()

@app.post("/")
async def run_model_api(file: UploadFile = File(...)):
    try:
        # Save uploaded input sheet to /tmp/
        input_path = f"/tmp/{file.filename}"
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Copy model template to /tmp/
        model_template = "Bespoke Model - US - v2.xlsm"
        model_path = f"/tmp/{model_template}"
        shutil.copyfile(model_template, model_path)

        # Define output path in /tmp/
        output_path = "/tmp/Processed_Model.xlsm"

        # Run the Excel processing
        run_model(input_path, model_path, output_path)

        # Return the processed Excel for download
        return FileResponse(
            path=output_path,
            filename="Processed_Model.xlsm",
            media_type='application/vnd.ms-excel'
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
