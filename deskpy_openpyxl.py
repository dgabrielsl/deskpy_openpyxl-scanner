import os
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Alignment

class Excel():
    def new_book(self, p1, fw, p2):
        self.to_scan = p1.text()
        self.final_workbook = fw.text()
        self.to_save = p2.text()
        workbook = Workbook()
        actsheet = workbook.active
        actsheet.append(['#', 'ID', 'CLIENTE', 'PAGARÉ', 'OTROS DOCUMENTOS', 'INFORMACIÓN GENERAL', 'APROBACIONES CREDITICIAS', 'INF. ANÁLISIS CAPACIDAD', 'RESULTADOS DE ANÁLISIS', 'DOCUMENTOS FALTANTES'])
        workbook.save(f'{self.to_save}/{self.final_workbook}.xlsx')

    def get_tree(self):
        self.tree = os.listdir(self.to_scan)
        self.iterator_tree = iter(self.tree)
        self.next_folder = next(self.iterator_tree)

    def sck_folder(self):
        self.main_folder = []
        for leaf in self.tree:
            self.main_folder.append(f'{self.to_scan}{self.next_folder}/')
            try: self.next_folder = next(self.iterator_tree)
            except: pass

        self.empties = []

        for sub_folder in self.main_folder:
            print(f'{self.main_folder.index(sub_folder)}: {sub_folder}')
            ftitle = sub_folder.split('/')
            ftitle = ftitle[-2]

            print('• 0. OTROS DOCUMENTOS')
            try: 
                self.folder_zr = os.listdir(f'{sub_folder}0. OTROS DOCUMENTOS')
                for file in self.folder_zr:
                    print(f'\t{file.replace(ftitle,"")}')
            except: pass


            print('\n• 1. INFORMACIÓN GENERAL')
            try: 
                self.folder_st = os.listdir(f'{sub_folder}1. INFORMACIÓN GENERAL')
                for file in self.folder_st:
                    print(f'\t{file.replace(ftitle,"")}')
            except: pass


            print('\n• 2. APROBACIONES CREDITICIAS')
            try: 
                self.folder_nd = os.listdir(f'{sub_folder}2. APROBACIONES CREDITICIAS')
                for file in self.folder_nd:
                    print(f'\t{file.replace(ftitle,"")}')
            except: pass


            print('\n• 3. INFORMACIÓN PARA ANÁLISIS CAPACIDAD')
            try: 
                self.folder_rd = os.listdir(f'{sub_folder}3. INFORMACIÓN PARA ANÁLISIS CAPACIDAD')
                for file in self.folder_rd:
                    print(f'\t{file.replace(ftitle,"")}')
            except: pass


            print('\n• 4. RESULTADOS DE ANÁLISIS')
            try: 
                self.folder_th = os.listdir(f'{sub_folder}4. RESULTADOS DE ANÁLISIS')
                for file in self.folder_th:
                    print(f'\t{file.replace(ftitle,"")}')
            except: pass

            print('\n\n')


