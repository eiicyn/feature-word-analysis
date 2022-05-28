import time

from xlwings import App

from core.analysis import ANALYSIS_HANDLER_MAP
from core.template import Template, CATEGORY_TEMPLATE_MAP, DEFAULT_SHEET_NAME


def generate_template(path, category):
    template: Template = CATEGORY_TEMPLATE_MAP[category]

    with App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.add()
        sheet = wb.sheets[DEFAULT_SHEET_NAME]
        sheet.range("A1").value = template.CONTENT
        sheet.range("A1").expand()
        wb.save(rf'{path}')
        wb.close()


def generate_result(path, category, book):
    callable_method = ANALYSIS_HANDLER_MAP[category]
    callable_method(path, category, book)
