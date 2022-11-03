import sys
import os
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QInputDialog,
    QLineEdit,
    QFileDialog,
    QFormLayout,
    QLabel,
    QLineEdit,
    QDialogButtonBox,
    QVBoxLayout,
    QPushButton,
    QGridLayout,
    QErrorMessage,
    QMessageBox
)
from PyQt5.QtCore import QRegExp, Qt
from PyQt5.QtGui import QRegExpValidator
import pandas as pd


VALIDATION_ERROR = 'Validation Error'

class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = '🇰🇭 CSV file converter'
        self.left = 10
        self.top = 10
        self.width = 640
        self.height = 480
        self.fileNameInput = QLineEdit(self)
        self.filePathInput = QLineEdit(self)
        self.errorDialog = QErrorMessage(self)
        self.popUpDialog = QMessageBox(self)
        self.btnBox = QDialogButtonBox()
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.resize(270*2, 120)

        dlgLayout = QVBoxLayout()
        # Create a form layout and add widgets
        gridLayout = QGridLayout()
        gridLayout.addWidget(QLabel("File Name:"), 0,0)
        gridLayout.addWidget(self.fileNameInput, 0,1)

        # Row: 1
        btnBrowse = QPushButton('Browse')
        btnBrowse.clicked.connect(self.openFileNameDialog)
        gridLayout.addWidget(btnBrowse, 0,2)

        # Row: 2
        gridLayout.addWidget(QLabel("File Path:"), 1,0)
        gridLayout.addWidget(self.filePathInput, 1,1)

        # Row: 4
        gridLayout.addWidget(QLabel("All right reserve Ⓒ Sagittech Inc."), 3,1)

        # Add a button box
        self.btnBox.setStandardButtons(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )

        # On OK clicked
        self.btnBox.accepted.connect(self.doConvertingFile)

        #On Cancel clicked
        self.btnBox.rejected.connect(self.closeApp)

        # Set the layout on the dialog
        dlgLayout.addLayout(gridLayout)
        dlgLayout.addWidget(self.btnBox)

        #
        self.errorDialog.setWindowModality(Qt.WindowModal)

        self.setLayout(dlgLayout)
        self.center()

        self.show()

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"File explorer", "","All Files (*);;CSV Files (*.csv)", options=options)
        if fileName:
            print(fileName)
            head, tail = os.path.split(fileName)
            self.fileNameInput.setText(tail)
            self.filePathInput.setText(head)

    def openFileNamesDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(self,"QFileDialog.getOpenFileNames()", "","All Files (*);;Python Files (*.py)", options=options)
        if files:
            print(files)

    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","All Files (*);;Text Files (*.txt)", options=options)
        if fileName:
            print(fileName)

    def center(self):
        frameGm = self.frameGeometry()
        screen = QApplication.desktop().screenNumber(QApplication.desktop().cursor().pos())
        centerPoint = QApplication.desktop().screenGeometry(screen).center()
        frameGm.moveCenter(centerPoint)
        self.move(frameGm.topLeft())

    def showPopUp(self, title='Something happen', text='Everything gonna be alright', icon=QMessageBox.Information):
        self.popUpDialog.setWindowTitle(title)
        self.popUpDialog.setText(text)
        self.popUpDialog.setIcon(icon)
        self.popUpDialog.show()

    def doConvertingFile(self):
        print('Starting ...')
        button = self.btnBox.button(QDialogButtonBox.Ok)
        button.setEnabled(False)

        if self.fileNameInput.text() == '':
            self.showPopUp(VALIDATION_ERROR, 'No file is inputted or chosen', QMessageBox.Warning)
            button.setEnabled(True)
            return

        if not self.fileNameInput.text().endswith('.xlsx'):
            self.showPopUp(VALIDATION_ERROR, 'Only XLSX file is allowed', QMessageBox.Warning)
            button.setEnabled(True)
            return

        try:
            input_file = os.path.join(self.filePathInput.text(), self.fileNameInput.text())
            new_directory = os.path.join(os.path.dirname(input_file), 'converted_csv')
            os.makedirs(new_directory, exist_ok=True)

            read_file = pd.read_excel(input_file)
            new_file = os.path.join(new_directory, os.path.splitext(self.fileNameInput.text())[0] + '.csv')

            # Write the dataframe object
            # into csv file
            read_file.to_csv(new_file,
                            index = None,
                            header=True)

            self.showPopUp('For your information', 'Conversion completed, cheers !!!')

        except Exception as err:
            self.showPopUp('Error', str(err), QMessageBox.Critical)

        button.setEnabled(True)

    def closeApp(self):
        sys.exit(0)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
