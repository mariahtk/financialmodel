import os
import re
import shutil
import win32com.client as win32

# --- Base directory ---
base_dir = r"C:\Users\Mariah.Krawchuk\Downloads"

# --- Find the correct Bespoke Input Sheet file (handles copies) ---
pattern = re.compile(r"Bespoke Input Sheet(?: \(\d+\))?\.xlsm$")
input_candidates = [f for f in os.listdir(base_dir) if pattern.match(f)]

if not input_candidates:
    raise FileNotFoundError("❌ No 'Bespoke Input Sheet' file found in Downloads.")

# Sort by modified time (newest first)
input_candidates.sort(key=lambda f: os.path.getmtime(os.path.join(base_dir, f)), reverse=True)

# Use the newest file
input_filename = input_candidates[0]
input_file = os.path.join(base_dir, input_filename)

# --- Model file path (US version) ---
model_file = os.path.join(base_dir, "Scripts and Models", "Bespoke Model - US - v2.xlsm")

# --- Start Excel ---
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    # --- Open the input workbook ---
    wb_input = excel.Workbooks.Open(input_file)
    ws_input = wb_input.Worksheets("Sales Team Input Sheet")

    # --- Read values from the input sheet ---
    address_raw = str(ws_input.Range("F7").Value).strip()
    additional_info = str(ws_input.Range("F9").Value).strip()
    latitude = ws_input.Range("F13").Value
    longitude = ws_input.Range("F15").Value
    sqft = ws_input.Range("F29").Value
    brand = ws_input.Range("F23").Value  # read if needed later
    market_rent = ws_input.Range("F37").Value
    yes_no_value = ws_input.Range("F54").Value
    num_floors = ws_input.Range("F56").Value

    if not address_raw:
        raise ValueError("Address in F7 is empty. Please fill it in.")

    # --- Replace "Dr" → "Drive" and "Blvd" → "Boulevard" for model workbook ---
    address_clean = address_raw.replace("Dr", "Drive").replace("Blvd", "Boulevard")

    # --- Create folder name = F7 + F9 ---
    folder_name = f"{address_clean} {additional_info}"
    invalid_chars = '<>:"/\\|?*'
    folder_name_clean = folder_name
    for ch in invalid_chars:
        folder_name_clean = folder_name_clean.replace(ch, "_")

    new_folder = os.path.join(base_dir, folder_name_clean)
    os.makedirs(new_folder, exist_ok=True)

    # --- Copy model workbook into new folder with filename = F7 only ---
    model_filename_clean = address_clean
    for ch in invalid_chars:
        model_filename_clean = model_filename_clean.replace(ch, "_")

    new_file = os.path.join(new_folder, f"{model_filename_clean}.xlsm")
    shutil.copy2(model_file, new_file)

    # --- Open the new model workbook ---
    wb_model = excel.Workbooks.Open(new_file)
    ws_model = wb_model.Worksheets("Sales Team Input Sheet")

    # --- Update building address ---
    ws_model.Range("E6").Value = address_clean

    # --- Update coordinates ---
    ws_model.Range("E12").Value = latitude
    ws_model.Range("E14").Value = longitude

    # --- Update square footage (F29 → E34) ---
    if sqft is not None:
        ws_model.Range("E34").Value = sqft

    # --- Handle F37 (market rent) → dropdown selection in K10 ---
    if market_rent is not None:
        try:
            market_rent_float = float(market_rent)
            if market_rent_float == 15:
                ws_model.Range("K10").Value = "15 - 20"
            elif market_rent_float == 20:
                ws_model.Range("K10").Value = "20 - 25"
            else:
                ws_model.Range("K10").Value = market_rent
        except ValueError:
            ws_model.Range("K10").Value = market_rent

    # --- Simply copy F54 text into K34 (no dropdown logic) ---
    if yes_no_value is not None:
        ws_model.Range("K34").Value = str(yes_no_value).strip()

    # --- Copy number of floors ---
    if num_floors is not None:
        ws_model.Range("K36").Value = num_floors

    # --- Save and close ---
    wb_model.Save()
    wb_model.Close(SaveChanges=True)
    wb_input.Close(SaveChanges=False)

    print(f"✅ Created new model from '{input_filename}' at:\n{new_file}")

except Exception as e:
    print(f"❌ Error: {e}")

finally:
    excel.Quit()


