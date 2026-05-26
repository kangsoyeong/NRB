import openpyxl

FILE2 = "/root/.claude/uploads/24041d56-69af-4927-919e-4cbd89e59ed0/875360de-__________.xlsx"

wb2 = openpyxl.load_workbook(FILE2)
ws2 = wb2["Sheet1"]

# Check rows 5183-5187 (all values)
print("Rows 5183–5187 in File2 Sheet1:")
for i, row_cells in enumerate(ws2.iter_rows(min_row=5183, max_row=5187), start=5183):
    vals = [c.value for c in row_cells]
    non_null = [(j, v) for j, v in enumerate(vals) if v is not None and str(v).strip() != ""]
    print(f"  row {i}: non-null cols = {non_null}")
