class Template:
    NAME = None
    COLUMNS = None
    CONTENT = None


class CompetitorTitleTemplate(Template):
    NAME = "竞品标题分析"
    COLUMNS = 2
    CONTENT = (("标题", "评论数"),
               ("PUMA Men's Suede Classic Sneaker", 215),
               ("Reebok Women's Princess Sneaker", 221),
               ("Tommy Hilfiger Women's Luster Sneaker", 99))


class ABATermSearchTemplate(Template):
    NAME = "ABA关键词搜索报告分析"
    COLUMNS = 2
    CONTENT = (("搜索词", "搜索频率排名"),
               ("air fryer", 43),
               ("ninja air fryer", 2169),
               ("air fryer oven", 7116))


class AdTermSearchTemplate(Template):
    NAME = "广告搜索词报告分析"
    COLUMNS = 6
    CONTENT = (("搜索词", "展现量", "点击量", "7天总订单数", "花费", "7天总销售额"),
               ("roti maker", 3542, 28, 2, 10.40, 85.98),
               ("india paratha pan", 2, 1, 1, 0.31, 43.99),
               ("roti tawa indian", 132, 5, 0, 0.29, 0.00))


DEFAULT_SHEET_NAME = "Sheet1"

CATEGORY_TEMPLATE_MAP = {
    CompetitorTitleTemplate.NAME: CompetitorTitleTemplate,
    ABATermSearchTemplate.NAME: ABATermSearchTemplate,
    AdTermSearchTemplate.NAME: AdTermSearchTemplate,
}
