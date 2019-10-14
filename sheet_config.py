header = [["Klass [INFOMENTOR]", "Namn [INFOMENTOR]", "Personnummer [INFOMENTOR]", "Klass [EXTENS]", "Namn [EXTENS]", "Personnummer [EXTENS]"]]
reported_header = [["Klass [INFOMENTOR]", "Namn [INFOMENTOR]", "Personnummer [INFOMENTOR]", "Klass [EXTENS]", "Namn [EXTENS]", "Personnummer [EXTENS]", "KOMMENTAR"]]
# sheet_names = ["Samlat", "1A", "1B", "1C", "1D", "2A", "2B", "2C", "2D", "3A", "3B", "3C", "3D", "4A", "4B", "4C", "4D", "5A", "5B", "5C", "5D", "6A", "6B", "6C", "6D", "7A", "7B", "7C", "8A", "8B", "8C", "9A", "9B", "9C"]
sheet_names = ["Samlat", "Avvikelser"]
klasser = ["1A", "1B", "1C", "1D", "2A", "2B", "2C", "2D", "3A", "3B", "3C", "3D", "4A", "4B", "4C", "4D", "5A", "5B", "5C", "5D", "6A", "6B", "6C", "6D", "7A", "7B", "7C", "8A", "8B", "8C", "9A", "9B", "9C"]

##### BUILDING SHEETS #####
sheet_objects = {}
sheet_objects[0] = {
    "updateSheetProperties": {
            "properties": {
                "sheetId": 0,
                "title": sheet_names[0],
            },
            "fields": "title"
        }
}
for i, sheet in enumerate(sheet_names[1:]):
    # Add new sheet tabs
    sheet_objects[i+2] = {
        "addSheet": {
            "properties": {
                "title": sheet
            }
        }
    }

##########################


##### GENERATE OBJECT, COLUMN WIDTH TO THE DIFFERENT SHEETS #####
def generate_columns_update_object(sheet_dict):
    columns_object = {}
    list_index = 0
    for key, value in sheet_dict.items():
        sheet_name = key
        sheet_id = value
        columns = [150,150,200,150,150,150,200]

        for i, width in enumerate(columns):
            columns_object[list_index] = {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": i,
                        "endIndex": i+1
                    },
                    "properties": {
                        "pixelSize": width
                    },
                    "fields": "pixelSize"
                }
            }
            list_index += 1

    return columns_object
###################################################################