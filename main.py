from PyQt5.QtWidgets import QApplication
from gui import FileManagerApp

if __name__ == "__main__":
    app = QApplication([])
    window = FileManagerApp()
    window.show()
    app.exec_()