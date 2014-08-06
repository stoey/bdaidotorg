import json
import os

import objectpath
import pystache

BASE_PAGE_PATH = "_base"
TEMPLATE_DIR = "templates"
DATA_DIR = "data"
OUTPUT_DIR = "static/output"

JSON_INCLUDE_PREFIX = '@'

FILE_HANDLERS = dict(
    mustache='MustacheHandler',
)

def get_handler(path):
    extension = path.rpartition('.')[2]
    handler_name = FILE_HANDLERS.get(extension, 'DefaultFileHandler')
    handler_class = globals()[handler_name]
    return handler_class(path)


class FileHandler(object):
    def __init__(self, path):
        self.path = path

    def process(self):
        # Ensure the parent directories exist
        out_dir = os.path.dirname(self.output_path)
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)
        # Perform OUTFILE < transform(INFILE)
        with open(os.path.join(TEMPLATE_DIR, self.path)) as infile:
            with open(self.output_path, 'wb') as outfile:
                outfile.write(self.transform(infile))

    def transform(self, f):
        raise NotImplementedError('Subclassed should define transform()')

    @property
    def output_name(self):
        return self.path

    @property
    def output_path(self):
        return os.path.join(OUTPUT_DIR, self.output_name)

    def __str__(self):
        return "%s ----> %s" % (self.path, self.output_name)


class DefaultFileHandler(FileHandler):
    def transform(self, f):
        return f.read()

class JsonIncluder(object):
    @staticmethod
    def is_include(key):
        return key.startswith(JSON_INCLUDE_PREFIX)

    @staticmethod
    def include_item(key, context):
        value = context[key]
        path, selector = value.partition('#')[::2]
        print key, path, selector, value
        with open(path) as f:
            data = json.load(f)
        if selector:
            tree = objectpath.Tree(data)
            # TODO - object path requires expression be a str
            # Probably a bug - if so fix it and sand a pull request upstream
            include = tree.execute(str(selector))
        else:
            include = data
        return key[len(JSON_INCLUDE_PREFIX):],include

    @staticmethod
    def update(json_obj):
        updates = {}
        for key in json_obj.iterkeys():
            print key
            if JsonIncluder.is_include(key):
                new_key, new_value = JsonIncluder.include_item(key, json_obj)
                updates[new_key] = new_value
        for key in updates.iterkeys():
            del json_obj[JSON_INCLUDE_PREFIX + key]
            json_obj[key] = updates[key]

class MustacheHandler(FileHandler):
    @property
    def extension(self):
        return self.path.rpartition('.')[2]

    @property
    def base_template(self):
        filename = "%s.%s" % (BASE_PAGE_PATH, self.extension)
        with open(os.path.join(TEMPLATE_DIR, filename)) as f:
            return f.read()

    @property
    def base_context(self):
        return self._context(BASE_PAGE_PATH + '.json')

    @property
    def context(self):
        return self._context(self.context_filename)

    @property
    def basename(self):
        return self.path.rpartition('.')[0]

    @property
    def context_filename(self):
        return self.basename + '.json'

    @staticmethod
    def _context(filename):
        context = {}
        path = os.path.join(DATA_DIR, filename)
        if os.path.exists(path):
            with open(path) as context_file:
                context = json.load(context_file)
        JsonIncluder.update(context)
        return context


    def _render(self, template):
        context = self.context
        body = pystache.render(template, context)
        base_context = self.base_context
        base_context.update(dict(body=body))
        return pystache.render(self.base_template, base_context)

    def transform(self, f):
        return self._render(f.read())

    @property
    def output_name(self):
        return self.basename

    def __str__(self):
        return "%s --[%s]--> %s" % (self.path, self.context_filename, self.output_name)

if __name__ == "__main__":
    for dirpath, dirnames, filenames in os.walk(TEMPLATE_DIR):
        filenames = [fn for fn in filenames if fn[0] not in '._']
        for filename in filenames:
            dirname = dirpath[len(TEMPLATE_DIR):].lstrip('/')
            extension = filename.rpartition('.')[2]
            path = os.path.join(dirname, filename)
            handler = get_handler(path)
            with open(handler.output_path, 'w') as output_file:
                handler.process()
            print handler
