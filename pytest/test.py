import sheet_excavator, glob, json

files = glob.glob(r"D:\temp\*")
#files = ['D:\\temp\\25_2-10 S (Frigg-Gammadelta)_RNB2022.xlsm', 'D:\\temp\\35_12_2 (Grosbeak)_RNB2022.xlsm', 'D:\\temp\\Ormen Lange_RNB2022.xlsm', 'D:\\temp\\Sleipner Vest_RNB2022.xlsm', 'D:\\temp\\Statfjord_RNB2022.xlsm']
extraction_details = [
    {"sheets":["Generell info og kommentarer"], "extractions": [{"function":"single_cells", "instructions":{"field": "d7", "od-id": "m8"}}]},
    {"sheets":["Profil_1", "Profil_2"], "extractions": [{"function":"single_cells", "instructions":{"project_name": "h7"}}]}
]

results = sheet_excavator.excel_extract(files, extraction_details, 10)
print([json.loads(r) for r in results])