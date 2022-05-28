from typing import Any

from core.template import AdTermSearchTemplate, ABATermSearchTemplate, CompetitorTitleTemplate, \
    CATEGORY_TEMPLATE_MAP, DEFAULT_SHEET_NAME
from xlwings import App, Book
from pandas import DataFrame

_increase = 10000


def _analyze_competitor_title(path, category, book: Book):
    result_sheet = book.sheets[DEFAULT_SHEET_NAME]

    with App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(path)
        data_sheet = wb.sheets[DEFAULT_SHEET_NAME]

        df: DataFrame = data_sheet.range('A1').options(DataFrame,
                                                       header=1,
                                                       index=False,
                                                       expand="table").value
        data_list: Any = df.values.tolist()
        index_review_count_map = {}
        index_title_word_map = {}
        global_title_word = set()
        for i, val in enumerate(data_list):
            index_review_count_map[i + 1] = int(val[1])
            words = val[0].split(" ")
            index_title_word_map[i + 1] = set(words)
            global_title_word.update(words)

        result = [CATEGORY_TEMPLATE_MAP[category].CONTENT[0]]
        for w in global_title_word:
            review_count = 0
            for i in range(len(data_list)):
                index = i + 1
                if w in index_title_word_map[index]:
                    review_count += index_review_count_map[index]
            result.append([w, review_count])

        result_sheet.range("A1").value = result
        result_sheet.range("A1").expand()
        wb.close()


def _analyze_aba_term_search(path, category, book: Book):
    result_sheet = book.sheets[DEFAULT_SHEET_NAME]

    with App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(path)
        data_sheet = wb.sheets[DEFAULT_SHEET_NAME]

        df: DataFrame = data_sheet.range('A1').options(DataFrame,
                                                       header=1,
                                                       index=False,
                                                       expand="table").value
        data_list: Any = df.values.tolist()
        index_rank_map = {}
        index_search_word_map = {}
        global_search_word = set()
        for i, val in enumerate(data_list):
            # get reciprocal and times increase
            index_rank_map[i + 1] = float(1 / val[1]) * _increase
            words = val[0].split(" ")
            index_search_word_map[i + 1] = set(words)
            global_search_word.update(words)

        result = [CATEGORY_TEMPLATE_MAP[category].CONTENT[0]]
        for w in global_search_word:
            rank_count = 0
            for i in range(len(data_list)):
                index = i + 1
                if w in index_search_word_map[index]:
                    rank_count += index_rank_map[index]
            result.append([w, round(rank_count, 2)])

        result_sheet.range("A1").value = result
        result_sheet.range("A1").expand()
        wb.close()


def _analyze_ad_term_search(path, category, book: Book):
    result_sheet = book.sheets[DEFAULT_SHEET_NAME]

    with App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(path)
        data_sheet = wb.sheets[DEFAULT_SHEET_NAME]

        df: DataFrame = data_sheet.range('A1').options(DataFrame,
                                                       header=1,
                                                       index=False,
                                                       expand="table").value
        data_list: Any = df.values.tolist()
        index_data_map = {}
        index_search_word_map = {}
        global_search_word = set()
        for i, val in enumerate(data_list):
            data = (int(val[1]), int(val[2]), int(val[3]), float(val[4]), float(val[5]))
            index_data_map[i + 1] = data
            words = val[0].split(" ")
            index_search_word_map[i + 1] = set(words)
            global_search_word.update(words)

        result = [CATEGORY_TEMPLATE_MAP[category].CONTENT[0]]
        for w in global_search_word:
            impression_count, click_count, order_count, cost_count, sales_count = 0, 0, 0, 0, 0
            for i in range(len(data_list)):
                index = i + 1
                if w in index_search_word_map[index]:
                    impression_count += index_data_map[index][0]
                    click_count += index_data_map[index][1]
                    order_count += index_data_map[index][2]
                    cost_count += index_data_map[index][3]
                    sales_count += index_data_map[index][4]
            result.append([w, impression_count, click_count, order_count, cost_count, sales_count])

        result_sheet.range("A1").value = result
        result_sheet.range("A1").expand()
        wb.close()


ANALYSIS_HANDLER_MAP = {
    CompetitorTitleTemplate.NAME: _analyze_competitor_title,
    ABATermSearchTemplate.NAME: _analyze_aba_term_search,
    AdTermSearchTemplate.NAME: _analyze_ad_term_search,
}
