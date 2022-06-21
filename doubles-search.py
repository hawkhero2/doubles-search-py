import os
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog
from openpyxl import Workbook
from openpyxl import load_workbook

class App(QMainWindow):
    
    def __init__(self):
        super().__init__()
        self.title = 'Find Doubles'
        self.left : int = 500
        self.top : int = 500
        self.height : int = 250
        self.width : int = 250
        self.initUI()
        


    def initUI(self):
        self.statusBar().showMessage('Select excel')
        browse_button = QPushButton('Browse', self)
        browse_button.clicked.connect(self.browse_file)
        browse_button.move(50,25)
        self.show()

    def browse_file(self):
        fileName, _ = QFileDialog.getOpenFileName(self,"Select Excel", "","Excel Files(*.xlsx)") # Grab File

        desktop = os.path.expanduser("~\Desktop\\") # current user desktop path
        base_xl = Workbook()
        result_xl = Workbook()
        
        base_xl = load_workbook(fileName,data_only=True)
        base_xl_sheet = base_xl.active
        # make a new list to store the IDs
        column_to_check = []
        
        result_xl.create_sheet("result")
        result_xl_sheet = result_xl["result"]
        
        for val in base_xl_sheet.iter_rows(min_row=1, max_row=base_xl_sheet.max_row, min_col=2, max_col=2, values_only=True):
            column_to_check.append(val)
        for id in base_xl_sheet.iter_rows(min_row=1, max_row=base_xl_sheet.max_row, min_col=1, max_col=1, values_only=True):
            if (id not in column_to_check):
                result_xl_sheet.append(id)
        base_xl.close()
        result_xl.save(desktop + 'result.xlsx')
        self.statusBar().showMessage('Process Finished')
        
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = App()
    sys.exit(app.exec_())