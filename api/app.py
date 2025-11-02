# api/process_excel.py
import os
import shutil
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from openpyxl import load_workbook

app = FastAPI()

# Base model path (already in your Vercel project)
BASE_MODEL = "Bespoke Model - US - v2.xlsm"

# Map of input cells → model cells
CELL_MAP = {
    "F7": "E6",   # Address
    "F13": "E12", # Latitude
    "F15": "E14", # Longitude
    "F29": "E34", # Square footage
    "F37": "K10", # Market rent
    "F54": "K34", # Yes/No value
    "F56": "K36"  # Number of floors
}

@app.post("/process_excel/")
async def process_excel(file: UploadFile = File(...)):
    try:
        # Save uploaded input sheet temporarily
        input_path = f"/tmp/{file.filename}"
        with open(input_path, "wb") as f:
            f.write(await file.read())

        # Load input workbook
        wb_input = load_workbook(input_path, keep_vba=True)
        ws_input = wb_input.active  # or ws_input = wb_input["Sales Team Input Sheet"]

        # Copy base model to temp output path
        output_path = f"/tmp/Processed_Model.xlsm"
        shutil.copy2(BASE_MODEL, output_path)

        # Load the model workbook
        wb_model = load_workbook(output_path, keep_vba=True)
        ws_model = wb_model.active  # or ws_model = wb_model["Sales Team Input Sheet"]

        # Copy values from input sheet to model sheet
        for input_cell, model_cell in CELL_MAP.items():
            value = ws_input[input_cell].value
            if value is not None:
                ws_model[model_cell].value = value

        # Optional: process market rent dropdown logic
        try:
            market_rent = ws_input["F37"].value
            if market_rent is not None:
                if float(market_rent) == 15:
                    ws_model["K10"].value = "15 - 20"
                elif float(market_rent) == 20:
                    ws_model["K10"].value = "20 - 25"
        except Exception:
            pass

        # Save the modified model
        wb_model.save(output_path)

        return FileResponse(
            path=output_path,
            filename="Processed_Model.xlsm",
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12"
        )

    except Exception as e:
        return {"error": str(e)}
