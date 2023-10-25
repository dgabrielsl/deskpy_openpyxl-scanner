import os
import sys
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.QtCore import *
from deskpy_openpyxl import Excel

class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init()
        self.page()
        self.show()
    
    def init(self):
        self.setWindowIcon(QIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogNewFolder)))
        self.setWindowTitle('DeskPy')
        self.setMinimumWidth(1080)

    def page(self):
        widget = QWidget()
        self.wlayout = QVBoxLayout()
        self.wlayout.setContentsMargins(20,20,20,20)
        widget.setLayout(self.wlayout)
        self.setCentralWidget(widget)
        h1 = QLabel('DeskPy - openpyxl')
        h1.setStyleSheet('padding: 1em; color: #AFA; background: #142; border: 4px groove #CCC; border-radius: 12px; font-size: 15px; font-weight: 600;')
        h1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.wlayout.addWidget(h1)
        path_description_1 = QLabel('Directorio de expedientes a escanear:')
        path_description_1.setMinimumWidth(220)
        self.path_1 = QLineEdit()
        self.path_1.setPlaceholderText('*')
        self.path_1.setReadOnly(True)
        change_path_1 = QPushButton('Search')
        change_path_1.setMinimumWidth(100)
        change_path_1.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        change_path_1.clicked.connect(lambda:self.filedialog(self.path_1))
        g1 = QHBoxLayout()
        g1.setContentsMargins(0,0,0,20)
        g1.addWidget(self.path_1)
        g1.addWidget(change_path_1)
        self.wlayout.addWidget(path_description_1)
        self.wlayout.addLayout(g1)
        title_description = QLabel('Guardar el archivo como:')
        title_description.setMinimumWidth(220)
        self.excel_output_name = QLineEdit()
        self.excel_output_name.setPlaceholderText('Título del documento')
        g2a = QHBoxLayout()
        g2a.setContentsMargins(0,0,0,20)
        g2a.addWidget(title_description)
        g2a.addWidget(self.excel_output_name)
        path_description_2 = QLabel('Guardar el reporte de Excel en la carpeta:')
        path_description_2.setMinimumWidth(220)
        self.path_2 = QLineEdit()
        self.path_2.setPlaceholderText('*')
        self.path_2.setReadOnly(True)
        change_path_2 = QPushButton('Search')
        change_path_2.setMinimumWidth(100)
        change_path_2.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        change_path_2.clicked.connect(lambda:self.filedialog(self.path_2))
        g2b = QHBoxLayout()
        g2b.setContentsMargins(0,0,0,20)
        g2b.addWidget(path_description_2)
        g2b.addWidget(self.path_2)
        g2b.addWidget(change_path_2)
        self.launch = QPushButton('Start Scan')
        self.launch.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.launch.setStyleSheet('padding: 12px; margin-bottom: 30px; font-size: 14px; font-weight: 600;')
        self.launch.clicked.connect(self.deploy_app)
        self.wlayout.addLayout(g2a)
        self.wlayout.addLayout(g2b)
        self.wlayout.addWidget(self.launch)



        self.path_1.setText('C:/Users/gabriel.solano/Documents/Drive/Lab/c) Samples/openpyxl - Reporte de archivos/')
        self.path_1.setText('C:/Users/dgabr/OneDrive/Documentos/Multimoney (cloud)/Lab/c) Samples/openpyxl - Reporte de archivos/')
        self.path_2.setText('C:/Users/dgabr/Downloads/')
        self.excel_output_name.setText('New report from DeskPy')





    def filedialog(self, record_in):
        get_dirname = QFileDialog.getExistingDirectory()
        get_dirname += '/'
        if get_dirname == '/': get_dirname = '-'
        else: record_in.setText(get_dirname)

    def deploy_app(self):
        try: os.remove(f'{self.path_2.text()}{self.excel_output_name.text()}.xlsx')
        except: pass
        self.check = os.path.exists(f'{self.path_2.text()}{self.excel_output_name.text()}.xlsx')
        if self.check == False:
            Excel.new_book(self, self.path_1, self.excel_output_name, self.path_2)
            Excel.get_tree(self)
        Excel.sck_folder(self)
        try: self.next_folder = next(self.iterator_tree)
        except StopIteration:
            self.launch.setDisabled(True)
            self.launch.setStyleSheet('padding: 12px; margin-bottom: 30px; font-size: 14px; font-weight: 600; color: #888; background: #333; border: 4px ridge #DDD;')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet("""
        QLineEdit{padding: 6px; border: 4px groove #BBB; border-radius: 5px;}
        QPushButton{padding: 8px; color: #0C5; background: #010; border-radius: 12px; border: 2px solid #ACA;}
        QPushButton:hover{background: #021;}
        QPlainTextEdit{padding: 5px; color: #0F8; background: #111; border: 4px ridge #CCC; font-size: 11px;}
        """)
    win = Main()
    sys.exit(app.exec())