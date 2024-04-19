import sheet_excavator, glob, json

files = glob.glob(r"D:\temp\*")
print(files)
extraction_details = [
    {"sheets":["Sheet1"], "extractions": [{"function":"single_cells", "instructions":{"1": "a1", "2": "b2", "3":"c3", "dato":"d4", "datotid":"e5"}}]},
    {"sheets":["Innledning"], "extractions": [{"function":"single_cells", "instructions":{"navn": "b34", "dato": "d34", "person":"d29", "telefon":"f29"}}]},
    {"sheets":["Generell info og kommentarer"], "extractions": [{"function":"single_cells", "instructions":{"field": "d7", "od-id": "m8"}}]},
    {"sheets":["Profil_1", "Profil_2"], "extractions": [{"function":"single_cells", "instructions":{"project_name": "h7"}}]}
]

results = sheet_excavator.excel_extract(files, extraction_details, 10)
for r in results:
    # Use json.dumps to print each result with indentation for better readability
    print(json.dumps(json.loads(r), indent=4))