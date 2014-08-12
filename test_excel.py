import excel2json
wb = excel2json.WORKBOOK_TEMPLATES['site']
xlwb = wb.read('site.xlsx')
print wb.json(xlwb)
