import time
from typing import Any
from webbrowser import open

from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QWidget, QLineEdit, QProgressBar, QComboBox, QPushButton, QFileDialog
from xlwings import App, Book

from core import *
from widget import *

# do not remove [from ui import amazon_rc, arrow_rc]
from ui import amazon_rc, arrow_rc


class UiMainWindow:
    def __init__(self, main_window):
        self._data_file = ""
        self._analysis_result = Any
        self._runtime_excel_app: App = Any
        self._runtime_excel_book: Book = Any

        # widgets init
        self.mainWidget: QWidget = Any
        self.centerWidget: QWidget = Any
        self.uploadLineEdit: QLineEdit = Any
        self.analysisProgressBar: QProgressBar = Any
        self.templateComboBox: QComboBox = Any
        self.uploadFileButton: QPushButton = Any
        self.downloadResultButton: QPushButton = Any
        self.downloadTemplateButton: QPushButton = Any
        self.startAnalysisButton: QPushButton = Any
        self.logoLabel: CustomLabel = Any

        # setup all the ui
        self.setup_ui(main_window)

        # set up excluding pyqt
        self.uploadLineEdit.setReadOnly(True)
        self.analysisProgressBar.setVisible(False)
        self.downloadTemplateButton.clicked.connect(self._save_template_file)
        self.uploadFileButton.clicked.connect(self._upload_file)
        self.startAnalysisButton.clicked.connect(self._analysis)
        self.downloadResultButton.clicked.connect(self._save_result_file)
        self.logoLabel.connect_customized_slot(self._open_official_website)

    def setup_ui(self, main_window):
        main_window.setObjectName("mainWindow")
        main_window.resize(571, 453)
        main_window.setStyleSheet("background-color: rgb(2, 2, 2);")

        self.mainWidget = QWidget(main_window)
        self.mainWidget.setObjectName("mainWidget")
        self.centerWidget = QWidget(self.mainWidget)
        self.centerWidget.setGeometry(QtCore.QRect(20, 140, 531, 281))
        self.centerWidget.setStyleSheet("QWidget {\n"
                                        "    background-color: rgb(243, 243, 243);\n"
                                        "    border-radius:30px;\n"
                                        "}\n"
                                        "")
        self.centerWidget.setObjectName("centerWidget")
        self.uploadLineEdit = QLineEdit(self.centerWidget)
        self.uploadLineEdit.setGeometry(QtCore.QRect(30, 100, 341, 31))
        self.uploadLineEdit.setStyleSheet("QLineEdit {\n"
                                          "    border: 2px solid #ffb553;\n"
                                          "    border-radius:5px;\n"
                                          "    background-color: rgb(255, 255, 255);\n"
                                          "}")
        self.uploadLineEdit.setObjectName("uploadLineEdit")
        self.analysisProgressBar = QProgressBar(self.centerWidget)
        self.analysisProgressBar.setGeometry(QtCore.QRect(30, 150, 341, 31))
        self.analysisProgressBar.setStyleSheet("QProgressBar {\n"
                                               "    border-radius:5px;\n"
                                               "    border-style: solid;\n"
                                               "    background-color: rgb(255, 255, 255);\n"
                                               "    text-align: center;\n"
                                               "}\n"
                                               "\n"
                                               "QProgressBar::chunk {\n"
                                               "    border-radius:5px;\n"
                                               "    background-color: rgb(255, 181, 83);\n"
                                               "}\n"
                                               "")
        self.analysisProgressBar.setProperty("value", 0)
        self.analysisProgressBar.setObjectName("analysisProgressBar")
        self.templateComboBox = QComboBox(self.centerWidget)
        self.templateComboBox.setGeometry(QtCore.QRect(30, 50, 341, 31))
        self.templateComboBox.setStyleSheet("QComboBox {\n"
                                            "    background-color: rgb(255, 255, 255);\n"
                                            "    border: 2px solid #ffb553;\n"
                                            "    border-radius: 5px;\n"
                                            "    padding-left: 10px;\n"
                                            "    font:16px,微软雅黑;\n"
                                            "    color: rgb(0, 0, 0);\n"
                                            "}\n"
                                            "\n"
                                            "QComboBox:drop-down {\n"
                                            "    border: 0px;\n"
                                            "}\n"
                                            "\n"
                                            "QComboBox:on {\n"
                                            "    border: 4px solid #ff9900;\n"
                                            "}\n"
                                            "\n"
                                            "QComboBox:down-arrow {\n"
                                            "    image: url(:/arrow/dropdown.png);\n"
                                            "    width: 12px;\n"
                                            "    height: 12px;\n"
                                            "}\n"
                                            "\n"
                                            "\n"
                                            "QListView {\n"
                                            "    background-color: rgb(255, 255, 255);\n"
                                            "    padding: 5px;\n"
                                            "    outline: 0px;\n"
                                            "    color: rgb(0, 0, 0);\n"
                                            "}\n"
                                            "\n"
                                            "\n"
                                            "\n"
                                            "")
        self.templateComboBox.setPlaceholderText("")
        self.templateComboBox.setObjectName("templateComboBox")
        self.templateComboBox.addItem("")
        self.templateComboBox.addItem("")
        self.templateComboBox.addItem("")
        self.templateComboBox.addItem("")
        self.templateComboBox.addItem("")
        self.uploadFileButton = QPushButton(self.centerWidget)
        self.uploadFileButton.setGeometry(QtCore.QRect(400, 100, 101, 31))
        self.uploadFileButton.setStyleSheet(CustomButtonStyle.get_style(hover=True))
        self.uploadFileButton.setObjectName("uploadFileButton")
        self.downloadResultButton = QPushButton(self.centerWidget)
        self.downloadResultButton.setGeometry(QtCore.QRect(400, 200, 101, 31))
        self.downloadResultButton.setStyleSheet(CustomButtonStyle.get_style(hover=True))
        self.downloadResultButton.setObjectName("downloadResultButton")
        self.downloadTemplateButton = QPushButton(self.centerWidget)
        self.downloadTemplateButton.setGeometry(QtCore.QRect(400, 50, 101, 31))
        self.downloadTemplateButton.setStyleSheet(CustomButtonStyle.get_style(hover=True))
        self.downloadTemplateButton.setObjectName("downloadTemplateButton")
        self.startAnalysisButton = QPushButton(self.centerWidget)
        self.startAnalysisButton.setGeometry(QtCore.QRect(400, 150, 101, 31))
        self.startAnalysisButton.setStyleSheet(CustomButtonStyle.get_style(hover=True))
        self.startAnalysisButton.setObjectName("startAnalysisButton")
        self.uploadLineEdit.raise_()
        self.analysisProgressBar.raise_()
        self.uploadFileButton.raise_()
        self.downloadResultButton.raise_()
        self.downloadTemplateButton.raise_()
        self.startAnalysisButton.raise_()
        self.templateComboBox.raise_()
        self.logoLabel = CustomLabel(self.mainWidget)
        self.logoLabel.setGeometry(QtCore.QRect(-80, -70, 441, 291))
        self.logoLabel.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.logoLabel.setStyleSheet("image: url(:/company/global.png);\n"
                                     "background-color: rgb(2, 2, 2);\n"
                                     "")
        self.logoLabel.setText("")
        self.logoLabel.setObjectName("logoLabel")
        self.logoLabel.raise_()
        self.centerWidget.raise_()
        main_window.setCentralWidget(self.mainWidget)

        self.translate_ui(main_window)
        self.templateComboBox.setCurrentIndex(0)

        QtCore.QMetaObject.connectSlotsByName(main_window)

    def translate_ui(self, main_window):
        _translate = QtCore.QCoreApplication.translate
        main_window.setWindowTitle(_translate("mainWindow", APP_NAME))
        self.templateComboBox.setItemText(0, _translate("mainWindow", "请选择分析类型"))
        self.templateComboBox.setItemText(1, _translate("mainWindow", "竞品标题分析"))
        self.templateComboBox.setItemText(2, _translate("mainWindow", "评论内容分析"))
        self.templateComboBox.setItemText(3, _translate("mainWindow", "ABA关键词搜索报告分析"))
        self.templateComboBox.setItemText(4, _translate("mainWindow", "广告搜索词报告分析"))
        self.uploadFileButton.setText(_translate("mainWindow", "选择文件"))
        self.downloadResultButton.setText(_translate("mainWindow", "下载结果"))
        self.downloadTemplateButton.setText(_translate("mainWindow", "下载模板"))
        self.startAnalysisButton.setText(_translate("mainWindow", "开始分析"))

    def _upload_file(self):
        event_result = QFileDialog.getOpenFileName(self.centerWidget, caption="选择要分析的Excel文件",
                                                   filter="Excel文件 (*.xls, *.xlsx)")
        print(event_result)
        if event_result[0]:
            self._data_file = event_result[0]
            self.uploadLineEdit.setText(self._data_file)

    def _save_template_file(self):
        category = self.templateComboBox.currentText()
        if category not in CATEGORY_TEMPLATE_MAP:
            return
        print(category)
        event_result = QFileDialog.getSaveFileName(self.centerWidget, caption="Save File",
                                                   filter="Excel 2007+ (*.xlsx)")
        print(event_result)
        if event_result[0]:
            set_button_meta(self.downloadTemplateButton, DownloadButtonProcessing.TEXT,
                            DownloadButtonProcessing.get_style(hover=False))
            generate_template(event_result[0], category)
            set_button_meta(self.downloadTemplateButton, TemplateDownloadButtonInit.TEXT,
                            TemplateDownloadButtonInit.get_style(hover=True))

    def _save_result_file(self):
        # TODO 如无分析结果则不能下载

        event_result = QFileDialog.getSaveFileName(self.centerWidget, caption="Save File",
                                                   filter="Excel 2007+ (*.xlsx)")
        print(event_result)
        if event_result[0]:
            self._runtime_excel_book.save(rf'{event_result[0]}')
            self._gc_runtime_excel()
            self.startAnalysisButton.setText("开始分析")

    def _analysis(self):
        category = self.templateComboBox.currentText()
        if category not in CATEGORY_TEMPLATE_MAP:
            return
        # TODO 如果没上传文件不能分析dialog
        print(category)

        set_button_meta(self.startAnalysisButton, AnalysisButtonProcessing.TEXT,
                        AnalysisButtonProcessing.get_style(hover=False))
        self._init_runtime_excel()
        self._runtime_excel_book = self._runtime_excel_app.books.add()

        generate_result(self._data_file, category, self._runtime_excel_book)
        set_button_meta(self.startAnalysisButton, AnalysisButtonSuccess.TEXT,
                        AnalysisButtonSuccess.get_style(hover=False))

    def _init_runtime_excel(self):
        self._gc_runtime_excel()
        self._runtime_excel_app = App(visible=False, add_book=False)
        self._runtime_excel_app.display_alerts = False
        self._runtime_excel_app.screen_updating = False

    def _gc_runtime_excel(self):
        if self._runtime_excel_book and self._runtime_excel_book != Any:
            self._runtime_excel_book.close()
        if self._runtime_excel_app and self._runtime_excel_app != Any:
            self._runtime_excel_app.kill()
        self._runtime_excel_book = None
        self._runtime_excel_app = None

    @staticmethod
    def _open_official_website():
        try:
            open("https://gs.amazon.cn/", new=2, autoraise=True)
        except:
            pass
