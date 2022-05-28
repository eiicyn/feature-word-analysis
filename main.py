if __name__ == "__main__":
    from sys import exit, argv
    from PyQt5 import QtWidgets

    from ui import *

    app = QtWidgets.QApplication(argv)
    main_window = QtWidgets.QMainWindow()
    ui = UiMainWindow(main_window)
    main_window.show()

    exit(app.exec_())
