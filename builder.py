import json
import os

import pystache

BASE_PAGE_PATH = "_base"
TEMPLATE_DIR = "templates"
DATA_DIR = "data"
OUTPUT_DIR = "static/output"

class Page(object):
    def __init__(self, path):
        self.path = path

    @property
    def template(self):
        return os.path.join(TEMPLATE_DIR, self.path + ".mustache")

    @property
    def context(self):
        return os.path.join(DATA_DIR, self.path + ".json")

    @property
    def output(self):
        return os.path.join(OUTPUT_DIR, self.path)

def get_context(page, extra_context):
    context = {}
    if os.path.exists(page.context):
        if not os.path.isdir(page.context):
            with open(page.context) as context_file:
                context = json.load(context_file)
        else:
            for filename in os.listdir(page.context):
                with open(os.path.join(page.context, filename)) as context_file:
                    key = filename.partition('.')[0]
                    value = json.load(context_file)
                    context[key] = value
    context.update(extra_context)
    return context


def render_page(page, **extra_context):
    with open(page.template) as template_file:
       template = template_file.read()
    context = get_context(page, extra_context)
    if page.path != BASE_PAGE_PATH:
        body = pystache.render(template, context)
        return render_page(Page(BASE_PAGE_PATH), body=body)
    else:
        return pystache.render(template, context)

if __name__ == "__main__":
    for dirpath, dirnames, filenames in os.walk(TEMPLATE_DIR):
        filenames = [fn for fn in filenames if fn[0] not in '._']
        for filename in filenames:
            dirname = dirpath[len(TEMPLATE_DIR):].lstrip('/')
            path = os.path.join(dirname, filename.rpartition('.')[0])
            page = Page(path)
            out_dir = os.path.dirname(page.output)
            if not os.path.exists(out_dir):
                os.makedirs(out_dir)
            with open(page.output, 'w') as output_file:
                output_file.write(render_page(page))
            print "%s ---(%s)--> %s" % (page.context, page.template, page.output)
