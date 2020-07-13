#! /usr/bin/env python3
import xlrd, json, os.path, requests as req
from progress.bar import Bar

#----------------- CONFIGURATION -----------------
API_KEY = ''
GEOCODING_ENABLED = True
GEOCODE_FIELD = 'address'

def get_col_names(sheet):
    row_size = sheet.row_len(0)
    col_values = sheet.row_values(0, 0, row_size)
    column_names = []

    for value in col_values:
        column_names.append(value)

    return column_names

def get_coordinates(address):
    url = 'https://maps.googleapis.com/maps/api/geocode/json?address={}'.format(address.replace(u"\u2018", "'").replace(u"\u2019", "'"))
    if API_KEY is not None:
        url = url + '&key={}'.format(API_KEY)
    res = req.get(url)
    res = res.json()
    
    if len(res['results']) == 0:
        print('No results returned for %s' %address)
        output = {
            "type": 'Point',
            "coordinates": []
        }
    else: 
        response = res['results'][0]
        output = {
            "type": 'Point',
            "coordinates": [
                response.get('geometry').get('location').get('lng'),
                response.get('geometry').get('location').get('lat')
            ]
        }
    
    return output

def get_row_data(row, column_names):
    row_data = {}
    counter = 0

    for cell in row:
        col = column_names[counter].lower().replace(' ', '_')
        if '.' in col:
            nested_cols = col.split('.')
            if nested_cols[0] not in row_data: 
                row_data[nested_cols[0]] = {}
            row_data[nested_cols[0]][nested_cols[1]] = cell.value
        else:
            if GEOCODING_ENABLED and GEOCODE_FIELD in col:
                row_data['location'] = get_coordinates(cell.value)
            row_data[col] = cell.value if '|' not in cell.value else cell.value.split('|')
        counter +=1

    return row_data

def get_sheet_data(sheet, column_names):
    nrows = sheet.nrows
    sheet_data = []
    
    bar = Bar('Processing', max=nrows)
    for idx in range(1, nrows):
        row = sheet.row(idx)
        rowdata = get_row_data(row, column_names)
        sheet_data.append(rowdata)
        bar.next()
    bar.finish()

    return sheet_data

def get_workbook_data(workbook):
    worksheet = workbook.sheet_by_index(0)
    column_names = get_col_names(worksheet)
    sheetdata = get_sheet_data(worksheet, column_names)

    return sheetdata

def main():
    filename = input("Enter the path to the filename")
    if os.path.isfile(filename):
        workbook = xlrd.open_workbook(filename)
        workbook_data = get_workbook_data(workbook)
        output = \
        open(filename.replace(".xlsx", ".json"), "w+")
        output.write(json.dumps(workbook_data, sort_keys=True, indent=2, separators=([',', ": "])))
        output.close()
        print ("%s was created" %output.name)
    else:
        print ("Sorry, that was not a valid filename")

main()