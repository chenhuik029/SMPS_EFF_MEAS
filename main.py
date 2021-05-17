import function_UI
from PyQt5.QtWidgets import QApplication, QMainWindow
import sys


def main():
    app = QApplication(sys.argv)
    ui_display = function_UI.FixedVIN_VarVOUT_UI()  # If future there is a MAIN UI, then change it!
    ui_display.show()
    app.exec_()


if __name__ == "__main__":
    main()
