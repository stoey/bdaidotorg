from collections import namedtuple
from datetime import datetime

import openpyxl

SUB_OBJECT = '__'


class WookbookTemplate(namedtuple('WookbookTemplate', 'sheets')):
    def excel(self):
        wb = openpyxl.Workbook()
        wb.remove_sheet(wb.active)
        for sheet in self.sheets:
            sheet.excel(wb)
        return wb

    def write(self, filename):
        wb = self.excel()
        wb.save(filename=filename)

    def json(self, excel_workbook):
        json = {}
        for sheet in self.sheets:
            json[sheet.name] = sheet.json(excel_workbook)
        return json


    def read(self, filename):
        return openpyxl.load_workbook(filename)

class Sheet(namedtuple('Sheet', 'name columns')):
    def excel(self, excel_workbook):
        ws = excel_workbook.create_sheet(title=self.name)
        for n, column in enumerate(self.columns):
            column.excel(ws, n + 1)

    def json(self, excel_workbook):
        ws = excel_workbook.get_sheet_by_name(self.name)
        json = []
        id_columns = [c for c in self.columns if SUB_OBJECT not in c.name]
        sub_columns = [c for c in self.columns if SUB_OBJECT in c.name]
        last_id = None
        obj = None
        for r, row in enumerate(ws.rows):
            r += 1 # Skip header
            values = tuple([col.json(ws, c, r) for c, col in enumerate(self.columns)])
            items = zip(self.columns, values)
            id = tuple([v for c,v in items if c in id_columns])
            if id != last_id:
                if obj is not None:
                    json.append(obj)
                obj = dict([i for i in items if i[0] not in sub_columns])
            else:
                sub_keys = set([c.name.partition(SUB_OBJECT)[0] for c in sub_columns])
                for sub_key in sub_keys:
                    obj[sub_key] = obj[sub_key] if sub_key in obj else []
                    sub_items = [(k.name.partition(SUB_OBJECT)[2], v) for k,v in items if k.name.startswith(sub_key)]
                    obj[sub_key].append(dict(sub_items))
            last_id = id
        if obj is not None:
            json.append(obj)
        return json


class Column(namedtuple('Column', 'name type')):
    def excel(self, excel_worksheet, index):
        excel_worksheet.cell(column=index, row=1).value = self.name

    def json(self, excel_worksheet, index, row):
        return excel_worksheet.cell(column=index + 1, row=row + 1).value


WORKBOOK_TEMPLATES = dict(
    site=WookbookTemplate((
        Sheet(
            name='nav',
            columns=(
                Column('name', unicode),
                Column('target', unicode),
            )
        ),
        Sheet(
            name='footer_columns',
            columns=(
                Column('section', unicode),
                Column('name', unicode),
                Column('target', unicode),
            ),
        ),
        Sheet(
            name='footer_links',
            columns=(
                Column('section', unicode),
                Column('top_link', unicode),
                Column('link%sname' % SUB_OBJECT, unicode),
                Column('link%starget' % SUB_OBJECT, unicode),
                Column('link%simage' % SUB_OBJECT, unicode),
            ),
        ),
        Sheet(
            name='jumbotron',
            columns=(
                Column('title', unicode),
                Column('expires', datetime),
                Column('background_image', unicode),
                Column('description', unicode),
                Column('action%stext' % SUB_OBJECT, unicode),
                Column('action%starget' % SUB_OBJECT, unicode),
            ),
        ),
    )),
    events=WookbookTemplate((
        Sheet(
            name='events',
            columns=(
                Column('name', unicode),
                Column('start', datetime),
                Column('end', datetime),
                Column('description', unicode),
                Column('image', unicode),
                Column('action%stext' % SUB_OBJECT, unicode),
                Column('action%starget' % SUB_OBJECT, unicode)
            ),
        ),
    )),
    pages=WookbookTemplate((
        Sheet(
            name='pages',
            columns=(
                Column('url', unicode),
                Column('title', unicode),
                Column('text', unicode),
                Column('image', unicode),
            ),
        ),
    )),
)



if __name__ == "__main__":
    for name, workbook in WORKBOOK_TEMPLATES.iteritems():
        filename = "static/%s.xlsx" % name
        workbook.write(filename)
        print "Wrote %s" % filename
