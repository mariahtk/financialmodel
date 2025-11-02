import os
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from openpyxl import load_workbook

app = FastAPI()

BASE_MODEL = "Bespoke Model - US - v2.xlsm"

CELL_MAP = {
    "F7": "E6", "F13": "E12", "F15": "E14",
    "F29": "E34", "F37": "K10", "F54": "K34", "F56": "K36"
}

@app.post("/process_excel/")
async def process_excel(file: UploadFile = File(...)):
    try:
        input_path = f"/tmp/{file.filename}"
        with open(input_path, "wb") as f:
            f.write(await file.read())

        wb_input = load_workbook(input_path, keep_vba=True)
        wb_model = load_workbook(BASE_MODEL, keep_vba=True)

        ws_input = wb_input["Sales Team Input Sheet"]
        ws_model = wb_model["Sales Team Input Sheet"]

        for input_cell, model_cell in CELL_MAP.items():
            value = ws_input[input_cell].value
            if value is not None:
                ws_model[model_cell].value = value

        try:
            market_rent = ws_input["F37"].value
            if market_rent is not None:
                if float(market_rent) == 15:
                    ws_model["K10"].value = "15 - 20"
                elif float(market_rent) == 20:
                    ws_model["K10"].value = "20 - 25"
        except Exception:
            pass

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


# New endpoint: download the base model directly
@app.get("/download_base_model/")
async def download_base_model():
    if not os.path.exists(BASE_MODEL):
        return {"error": "Base model file not found."}
    return FileResponse(
        path=BASE_MODEL,
        filename="Bespoke Model - US - v2.xlsm",
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12"
    )
