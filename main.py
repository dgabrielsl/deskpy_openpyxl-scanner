import os
import sys

from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.QtCore import *

os.system('cls')

class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init()
        self.page()
        self.show()
    
    def init(self):
        self.setWindowIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirHomeIcon)))
        self.setWindowTitle('DeskPy')
        self.setMinimumWidth(900)
        self.setMinimumHeight(500)

    def page(self):
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20,20,20,20)
        widget.setLayout(layout)
        self.setCentralWidget(widget)
        # All widgets down here.
        layout.addStretch()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    # app.setStyleSheet("""
    #     """)
    win = Main()
    sys.exit(app.exec())