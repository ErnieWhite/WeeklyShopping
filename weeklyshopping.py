import csv
import datetime
import openpyxl
import pathlib


def main():
    p = pathlib.Path('./data')
    d = {}
    for file in [x for x in p.iterdir() if x.is_file()]:
        # print(file.stem, file.suffix)
        sheetname = file.stem[0:-9].upper()
        sheetdate = file.stem[len(file.stem)-8:]
        'YYYY/MM/DD'
        sheetdate = sheetdate[4:] + "/" + sheetdate[0:2] + "/" + sheetdate[2:4]
        print(sheetname)
        print(sheetdate)
        if sheetname not in d.keys():
            d[sheetname] = {}

        data = []
        if file.suffix in ['.xls', '.xlsx']:
            print(file)
            wb = openpyxl.load_workbook(file)
            ws = wb.worksheets[0]
            data = list(ws.values)
            data.pop(0)
            print(list(list(ws.values)[0]))
        elif file.suffix in ['.csv', '.tsv']:
            print(file)
            with open(file, 'r', newline='', encoding='utf-8-sig') as csv_file:
                reader = csv.reader(csv_file)
                row = next(reader)
                print(row)
                data = list(reader)

        item_id = -1
        unit_price = -1
        if sheetname == "FERGUSON":
            item_id = 3
            unit_price = 7
        elif sheetname == "HOMEDEPOT":
            item_id = 7
            unit_price = 13
        elif sheetname == "HOMEDEPOT-SECOND":
            item_id = 4
            unit_price = 10
        elif sheetname == "HOMEDEPOTMABIS":
            item_id = 4
            unit_price = 10
        elif sheetname == "LOWES":
            item_id = 4
            unit_price = 11
        else:
            print(f"I do not know about {sheetname.title()}.")
            continue
        for row in data:
            if row[item_id] not in d[sheetname]:
                d[sheetname][row[item_id]] = {
                    'row': row,
                    'pricedates': {}
                }
            d[sheetname][row[item_id]]['pricedates'][sheetdate] = row[unit_price]
    # print(d)
    for sheet, sheetdata in d.items():
        for item, itemdata in sheetdata.items():
            print(f'Sheetname : {sheet}\tItem : {item}\t', end='')
            for date in sorted(d[sheet][item]['pricedates']):
                print(f"Date : {date}\tUnit Price : {d[sheet][item]['pricedates'][date]}", end='\t')
            print()


if __name__ == '__main__':
    main()

