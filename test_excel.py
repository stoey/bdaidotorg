import excel2json
import pprint

wb = excel2json.WORKBOOK_TEMPLATES['site']
xlwb = wb.read('site.xlsx')
pprint.pprint(wb.json(xlwb))
