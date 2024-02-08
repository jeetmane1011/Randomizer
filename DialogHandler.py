from PyQt5.QtWidgets import QFileDialog, QMessageBox

def selectExcelFile(parent):
    fileInfo = QFileDialog.getOpenFileName(parent, \
                                            "Select Excel File", \
                                            "C:/", \
                                            "Excel file (*.xlsx);;Excel file (*.xlr);;Excel file (*.xls);;Excel file (*xlsm);;Excel file (*xlsb)") 
    # return filepath
    return fileInfo[0] 

# show error box
def show_error_messagebox(parent, message):
    msg = QMessageBox(parent=parent)
    msg.setIcon(QMessageBox.Warning)
    msg.setWindowTitle("Error")
    msg.setText(message)
    msg.setStandardButtons(QMessageBox.Retry)
    msg.exec_()

# when saved successfully
def save_success_messagebox(parent, sheetname):
    msg = QMessageBox(parent=parent)
    msg.setIcon(QMessageBox.Information)
    msg.setText(f"Sheet '{sheetname}' saved successfully !!!")
    msg.setWindowTitle("Saved")
    msg.setStandardButtons(QMessageBox.Ok)
    msg.exec_()

def save_and_quit_messagebox(parent):
    msg = QMessageBox(parent=parent)
    msg.setIcon(QMessageBox.Question)
    msg.setText("There are some unsaved changes. Save and Quit ?")
    msg.setWindowTitle("Save Changes")
    msg.setStandardButtons(QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
    ret = msg.exec_()
    if ret == QMessageBox.Save:
        return 1
    elif ret == QMessageBox.Discard:
        return 0
    else:
        return 2
