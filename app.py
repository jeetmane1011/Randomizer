from PyQt5 import uic, QtCore, QtGui
from PyQt5.QtWidgets import QMainWindow, QWidget, QCheckBox, QApplication, QAction, QStyle, QFrame, QLabel, QPushButton, QTableWidget,QTabWidget, QLineEdit, QSpinBox
import sys
import os
import pandas as pd
import numpy as np
from datetime import date
from tendo import singleton
import DialogHandler
from ExcelHandler import Workbook
import TableHandler

try:
    single = singleton.SingleInstance()
except:
    sys.exit()

class UI(QMainWindow):
    def __init__(self):
        super().__init__()

        # load XML on the window 
        uic.loadUi("randomizer.ui", self)  

        self.wb = None
        self.all_names_df = None
        self.avail_names_df = None
        self.new_subset_df = None
        self.random_name = None

        self.mainWidget = self.findChild(QWidget, "centralwidget")
        self.tab_widget = self.findChild(QTabWidget, "tabWidget")
        self.select_file_btn = self.findChild(QPushButton, "select_file_btn")
        self.filepath_label = self.findChild(QLabel, "filepath_label")
        self.names_table = self.findChild(QTableWidget, "names_tableWidget")
        self.all_names_table = self.findChild(QTableWidget, "all_names_table")
        self.all_names_row_label = self.findChild(QLabel, "all_names_row_label")
        self.reset_btn = self.findChild(QPushButton, "reset_all_btn")
        self.rows_info_label = self.findChild(QLabel, "rows_label")
        self.max_rows_label = self.findChild(QLabel, "max_rows_label")
        self.new_subset_rows_label = self.findChild(QLabel, "new_subset_rows_label")
        self.row_spinBox = self.findChild(QSpinBox, "rows_spinBox")
        self.sheet_name_label = self.findChild(QLabel, "event_name_label")
        self.sheet_name_lineEdit = self.findChild(QLineEdit, "event_name_lineEdit")
        self.rand_name_lineEdit = self.findChild(QLineEdit, "rand_name_Linedit")
        self.save_xl_btn = self.findChild(QPushButton, "save_xl_btn")
        self.rand_name_btn = self.findChild(QPushButton, "random_name_btn")
        self.save_frame = self.findChild(QFrame, "save_frame")
        self.gen_name_frame = self.findChild(QFrame, "generate_name_frame")
        self.sheets_table = self.findChild(QTableWidget, "sheets_tableWidget")

        self.setWindowIcon(QtGui.QIcon('randomIcon.ico'))

        
        self.save_xl_btn.setIcon(self.style().standardIcon(getattr(QStyle, "SP_DialogSaveButton")))

        self.select_file_btn.clicked.connect(self.get_data_from_file)
        self.save_xl_btn.clicked.connect(self.save_to_new_sheet)
        self.rand_name_btn.clicked.connect(self.generate_random_name)
        self.reset_btn.clicked.connect(self.reset_subset)


        # function for closing window for default close button   
        quit = QAction("Quit", self)
        quit.triggered.connect(self.closeEvent)

        self.config_before_render_table()
        self.config_before_random_name()  

        # show window
        self.show()


    def get_data_from_file(self):
        filepath = DialogHandler.selectExcelFile(self.mainWidget)

        if filepath == "":
            return

        self.wb = Workbook(filepath)
        
        self.filepath_label.setText(os.path.basename(self.wb.filepath))
        self.all_names_df = self.wb.get_all_names()

        if len(self.wb.sheets) == 1:
            self.avail_names_df = self.all_names_df
        else:
            self.avail_names_df = self.wb.get_avail_names(self.all_names_df)    

        QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)

        TableHandler.draw_table(self.all_names_table, self.all_names_df)
        TableHandler.draw_table(self.names_table, self.avail_names_df)
        TableHandler.draw_table(self.sheets_table, pd.DataFrame(self.wb.sheets_info))

        self.config_after_render_table()   
        self.config_before_random_name()  
         
        QApplication.restoreOverrideCursor()



    def generate_random_name(self):    
        if self.row_spinBox.value() == 0:
            return
        
        # create subset df
        self.create_random_list()
        # get random from the subset
        ind = np.random.randint(0,self.new_subset_df.shape[0])
        self.random_name = self.new_subset_df.iloc[ind]
        # shift rand_name to 1st row
        self.new_subset_df = pd.concat([self.new_subset_df.iloc[ind:], self.new_subset_df.iloc[:ind]], ignore_index=True)
        # set config
        self.config_after_random_name()
    

    def create_random_list(self):
        QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)

        random_ind_list = np.random.choice(
                range(self.avail_names_df.shape[0]),
                self.row_spinBox.value(),
                replace=False
            ).tolist()
        
        self.new_subset_df = self.avail_names_df.iloc[random_ind_list]
        self.new_subset_df.reset_index(drop=True, inplace=True)
        TableHandler.draw_table(self.names_table, self.new_subset_df)
        
        QApplication.restoreOverrideCursor()
    

    def config_before_render_table(self):
        self.tab_widget.setCurrentIndex(0)
        self.tab_widget.setEnabled(False)
        self.all_names_row_label.setText("0")
        self.max_rows_label.setText("")
        self.set_table_dim_label(0,0)
        self.filepath_label.setText("")


    def config_after_render_table(self):
        total_rows = self.all_names_df.shape[0]
        avail_rows = self.avail_names_df.shape[0]

        self.tab_widget.setEnabled(True)
        self.tab_widget.setCurrentIndex(0)
        self.row_spinBox.setMinimum(0)
        self.row_spinBox.setMaximum(avail_rows)
        self.sheets_table.setHorizontalHeaderLabels(["Sheet name", "Rows", "Candidate"])
        self.all_names_row_label.setText(f"{total_rows}")
        self.set_table_dim_label(avail_rows, total_rows)
        self.max_rows_label.setText(f"(max {avail_rows})")


    def config_before_random_name(self):
        self.reset_btn.setDisabled(True)
        self.row_spinBox.setValue(0)
        self.gen_name_frame.show()
        self.save_frame.hide()
        self.select_file_btn.setDisabled(False)


    def config_after_random_name(self):
        self.reset_btn.setDisabled(False)
        self.sheet_name_lineEdit.setText("")
        self.gen_name_frame.hide()
        self.save_frame.show()
        self.select_file_btn.setDisabled(True)
        self.new_subset_rows_label.setText(f"{self.new_subset_df.shape[0]}")
        self.rand_name_lineEdit.setText(self.wb.get_candidate_name(self.random_name.to_frame().T))   
    

    def save_to_new_sheet(self):
        sheet_name = self.sheet_name_lineEdit.text()

        if sheet_name == "":
            sheet_name = f"Sheet{len(self.wb.sheets)+1}_" + date.today().strftime("%d%b%Y")

        if sheet_name in self.wb.sheets:
            DialogHandler.show_error_messagebox(self.mainWidget,"Name already exists. Please Try Again !!!")
            self.sheet_name_lineEdit.clear()
            return
        
        try:
            self.wb.write_to_new_sheet(sheet_name, self.new_subset_df)
        except PermissionError as pe:
            DialogHandler.show_error_messagebox(self.mainWidget, "The file is already open in background. Close it to save !!!")
        except Exception as e:
            DialogHandler.show_error_messagebox(self.mainWidget, str(e))
        else:
            DialogHandler.save_success_messagebox(self.mainWidget, sheet_name)
            self.reset_app()


    def reset_subset(self):
        
        QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)

        self.new_subset_df = None
        self.random_name = None
        TableHandler.draw_table(self.names_table, self.avail_names_df)
        self.config_before_random_name()

        QApplication.restoreOverrideCursor()


    def reset_app(self):
        self.wb = None
        self.all_names_df = None
        self.avail_names_df = None
        self.new_subset_df = None
        self.random_name = None
        TableHandler.clear_table(self.names_table)
        TableHandler.clear_table(self.all_names_table)
        TableHandler.clear_table(self.sheets_table)
        self.config_before_render_table()
        self.config_before_random_name()

    
    def closeEvent(self, event):
        if self.random_name is not None:
            value = DialogHandler.save_and_quit_messagebox(self.mainWidget)
            if value == 1:
                self.save_to_new_sheet()
                return
            elif value == 2:
                event.ignore()
                return
    
        event.accept()


    def set_table_dim_label(self, avail_rows, all_rows):
        self.rows_info_label.setText(f'Available: {avail_rows} of {all_rows}')


if __name__=="__main__":
    app = QApplication(sys.argv)
    UIWindow = UI()
    app.exec_()