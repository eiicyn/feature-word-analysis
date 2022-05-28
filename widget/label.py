from PyQt5 import QtWidgets, QtCore


class CustomLabel(QtWidgets.QLabel):
    _button_clicked_signal = QtCore.pyqtSignal()

    def __init__(self, parent=None):
        super(CustomLabel, self).__init__(parent)

    def mouseReleaseEvent(self, QMouseEvent):
        self._button_clicked_signal.emit()

    def connect_customized_slot(self, func):
        self._button_clicked_signal.connect(func)
