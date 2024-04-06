import sheet_excavator, glob

#files = glob.glob(r"D:\temp\*")
files = ['D:\\temp\\25_2-10 S (Frigg-Gammadelta)_RNB2022.xlsm', 'D:\\temp\\35_12_2 (Grosbeak)_RNB2022.xlsm', 'D:\\temp\\Ormen Lange_RNB2022.xlsm', 'D:\\temp\\Sleipner Vest_RNB2022.xlsm', 'D:\\temp\\Statfjord_RNB2022.xlsm']
extraction_details = [
    {"sheets":["Generell info og kommentarer"], "cells":["d7", "m8"], "value_name": ["field", "od-id"]},
    {"sheets":["Profil_1", "Profil_2"], "cells":["h7"], "value_name": ["project_name"]}
]
results = sheet_excavator.excel_extract(files,extraction_details)

print("RESULTS::::", results)