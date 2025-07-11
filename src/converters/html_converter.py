from src.models.table_model import Sheet
from jinja2 import Environment, FileSystemLoader

class HTMLConverter:
    def __init__(self):
        self.env = Environment(loader=FileSystemLoader('src/templates'))

    def convert(self, sheet: Sheet) -> str:
        template = self.env.get_template('table_template.html')
        return template.render(sheet=sheet)
