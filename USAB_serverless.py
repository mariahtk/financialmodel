import openpyxl

def run_model(input_path, model_path, output_path):
    """
    Processes the input Excel and writes to the model Excel.
    """
    # Load input sheet
    wb_input = openpyxl.load_workbook(input_path, data_only=True)
    ws_input = wb_input.active  # adjust if needed

    # Load model template
    wb_model = openpyxl.load_workbook(model_path, keep_vba=True)
    ws_model = wb_model.active  # adjust if needed

    # Example mapping logic: copy column A from input to model
    for row in range(2, ws_input.max_row + 1):
        ws_model[f"A{row}"].value = ws_input[f"A{row}"].value

    # Save output
    wb_model.save(output_path)
