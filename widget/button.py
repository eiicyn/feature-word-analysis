from PyQt5.QtWidgets import QPushButton


class CustomButtonStyle:
    TEXT = ""
    BASIC_STYLE = "QPushButton {\n" \
                  "    background-color: rgb(255, 181, 83);\n" \
                  "    border-radius:5px;\n" \
                  "    font:15px,微软雅黑;\n" \
                  "    color: rgb(0, 0, 0);\n" \
                  "}\n"
    HOVER_STYLE = "QPushButton:hover{\n" \
                  "    border: 3px solid #ff9900;\n" \
                  "    font:16px,微软雅黑;\n" \
                  "    font-weight: bold;\n" \
                  "}\n" \
                  ""

    @classmethod
    def get_style(cls, hover):
        return cls.BASIC_STYLE + cls.HOVER_STYLE if hover else cls.BASIC_STYLE


class ButtonBlock(CustomButtonStyle):
    BASIC_STYLE = "QPushButton {\n" \
                  "    background-color: rgb(255, 153, 0);\n" \
                  "    border-radius:5px;\n" \
                  "    font:16px,微软雅黑;\n" \
                  "    font-weight: bold;\n" \
                  "    color: rgb(0, 0, 0);\n" \
                  "}\n"


class ButtonSuccess(CustomButtonStyle):
    BASIC_STYLE = "QPushButton {\n" \
                  "    background-color: rgb(27, 213, 109);\n" \
                  "    border-radius:5px;\n" \
                  "    font:16px,微软雅黑;\n" \
                  "    font-weight: bold;\n" \
                  "    color: rgb(0, 0, 0);\n" \
                  "}\n"


class ButtonFailure(CustomButtonStyle):
    BASIC_STYLE = "QPushButton {\n" \
                  "    background-color: rgb(229, 64, 51);\n" \
                  "    border-radius:5px;\n" \
                  "    font:16px,微软雅黑;\n" \
                  "    font-weight: bold;\n" \
                  "    color: rgb(0, 0, 0);\n" \
                  "}\n"


class AnalysisButtonInit(CustomButtonStyle):
    TEXT = "开始分析"


class AnalysisButtonProcessing(ButtonBlock):
    TEXT = "取消分析"


class AnalysisButtonSuccess(ButtonSuccess):
    TEXT = "分析成功"


class AnalysisButtonFailure(ButtonFailure):
    TEXT = "分析失败"


class TemplateDownloadButtonInit(CustomButtonStyle):
    TEXT = "下载模板"


class ResultDownloadButtonInit(CustomButtonStyle):
    TEXT = "下载结果"


class DownloadButtonProcessing(ButtonBlock):
    TEXT = "下载中"


def set_button_meta(button: QPushButton, text, style):
    button.setText(text)
    button.setStyleSheet(style)
    button.repaint()
