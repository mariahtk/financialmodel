from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import shutil
import os
from USAB_serverless import run_model

app = FastAPI()  # <-- Vercel looks for this

@app.post("/run-model/")
async def run_model_api(file: UploadFile = File(...)):
    input_path = f"/tmp/{file.filename}"
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    model_path = os.path.join(os.getcwd(), "Bespoke Model - US - v2.xlsm")
    output_path = "/tmp/output_model.xlsm"

    run_model(input_path, model_path, output_path)

    return FileResponse(
        path=output_path,
        filename="Processed_Model.xlsm",
        media_type='application/vnd.ms-excel'
    )
