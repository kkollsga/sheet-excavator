import sheet_excavator, glob, json

files = glob.glob(r"D:\temp\*")
print(files)
#files = ['D:\\temp\\File3.xlsm']
extraction_details = [
    {"sheets":["Sheet1"], "extractions": [
        {"function":"single_cells", "label": "single","instructions":{"1": "a1", "2": "b2", "3":"c3", "dato":"d4", "datotid":"e5"}}
    ]},
    {"sheets":["Innledning"], "extractions": [{"function":"single_cells", "instructions":{"navn": "b34", "dato": "d34", "person":"d29", "telefon":"f29"}}]},
    {"sheets":["Generell info og kommentarer"], "extractions": [
        {"function":"single_cells", "label": "single", "instructions":{"field": "d7", "od-id": "m8"}},
        {"function":"multirow_patterns", "label": "deposits", "instructions":{"start_row":28, "end_row":44, "unique_id":"B","columns":{"B":"Deposit", "C":"Discovery_well", "D":"Description", "E":"Oil_low", "F":"Oil_base", "G":"Oil_high", "H":"Cond_low", "I":"Cond_base", "J":"Cond_high", "K":"AssGass_low", "L":"AssGass_base", "M":"AssGass_high", "N":"FriGass_low", "O":"FriGass_base","P":"FriGass_high"}}},
    ]},
    {"sheets":["Profil_1", "Profil_2"], "extractions": [{"function":"single_cells", "label": "single","instructions":{"project_name": "h7"}}]}
]

results = sheet_excavator.excel_extract(files, extraction_details, 10)
print(json.dumps(json.loads(results), indent=4))