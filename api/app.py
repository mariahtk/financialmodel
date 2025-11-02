import os
import shutil
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from openpyxl import load_workbook

app = FastAPI()

# Base model path
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
        # Save uploaded input sheet
        input_path = f"/tmp/{file.filename}"
        with open(input_path, "wb") as f:
            f.write(await file.read())

        # Load workbooks
        wb_input = load_workbook(input_path, keep_vba=True)
        wb_model = load_workbook(BASE_MODEL, keep_vba=True)

        ws_input = wb_input["Sales Team Input Sheet"]
        ws_model = wb_model["Sales Team Input Sheet"]

        # Copy mapped cells
        for input_cell, model_cell in CELL_MAP.items():
            value = ws_input[input_cell].value
            if value is not None:
                ws_model[model_cell].value = value

        # Market rent dropdown logic
        try:
            market_rent = ws_input["F37"].value
            if market_rent is not None:
                if float(market_rent) == 15:
                    ws_model["K10"].value = "15 - 20"
                elif float(market_rent) == 20:
                    ws_model["K10"].value = "20 - 25"
        except Exception:
            pass

        # Dynamic output filename using F7 + F9
        address = ws_input["F7"].value or "Processed"
        additional_info = ws_input["F9"].value or ""
        filename_base = f"{address} {additional_info}".strip()
        for ch in '<>:"/\\|?*':
            filename_base = filename_base.replace(ch, "_")
        output_filename = f"{filename_base}.xlsm"
        output_path = f"/tmp/{output_filename}"

        wb_model.save(output_path)

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12"
        )

    except Exception as e:
        return {"error": f"Unexpected error: {e}"}
