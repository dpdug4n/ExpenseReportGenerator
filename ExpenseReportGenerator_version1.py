#Exepense Report Generator by Patrick Dugan
#Ain't fancy but it works.

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
import os, glob
from PIL import Image
import PyPDF2 


class Ui_ExpenseReportGenerator( QObject):
    def setupUi(self, ExpenseReportGenerator):
        ExpenseReportGenerator.setObjectName("ExpenseReportGenerator")
        ExpenseReportGenerator.setEnabled(True)
        ExpenseReportGenerator.resize(400, 250)
        ExpenseReportGenerator.setMinimumSize(QtCore.QSize(400, 250))
        ExpenseReportGenerator.setMaximumSize(QtCore.QSize(400, 250))
        self.tabWidget = QtWidgets.QTabWidget(ExpenseReportGenerator)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 361, 211))
        self.tabWidget.setObjectName("tabWidget")
        self.tabInfo = QtWidgets.QWidget()
        self.tabInfo.setObjectName("tabInfo")
        self.txtInfo = QtWidgets.QTextBrowser(self.tabInfo)
        self.txtInfo.setGeometry(QtCore.QRect(0, 0, 391, 271))
        self.txtInfo.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.txtInfo.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.txtInfo.setOpenExternalLinks(True)
        self.txtInfo.setObjectName("txtInfo")
        self.tabWidget.addTab(self.tabInfo, "")
        self.tabGenerator = QtWidgets.QWidget()
        self.tabGenerator.setObjectName("tabGenerator")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.tabGenerator)
        self.verticalLayout.setObjectName("verticalLayout")
        self.lblPDF = QtWidgets.QLabel(self.tabGenerator)
        self.lblPDF.setObjectName("lblPDF")
        self.verticalLayout.addWidget(self.lblPDF)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.lineEditPDF = QtWidgets.QLineEdit(self.tabGenerator)
        self.lineEditPDF.setObjectName("lineEditPDF")
        self.horizontalLayout_2.addWidget(self.lineEditPDF)
        self.btnPDF = QtWidgets.QPushButton(self.tabGenerator)
        self.btnPDF.setObjectName("btnPDF")
        self.horizontalLayout_2.addWidget(self.btnPDF)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.lblReceiptFolder = QtWidgets.QLabel(self.tabGenerator)
        self.lblReceiptFolder.setObjectName("lblReceiptFolder")
        self.verticalLayout.addWidget(self.lblReceiptFolder)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.lineEditReceiptFolder = QtWidgets.QLineEdit(self.tabGenerator)
        self.lineEditReceiptFolder.setObjectName("lineEditReceiptFolder")
        self.horizontalLayout.addWidget(self.lineEditReceiptFolder)
        self.btnReceiptFolder = QtWidgets.QPushButton(self.tabGenerator)
        self.btnReceiptFolder.setObjectName("btnReceiptFolder")
        self.horizontalLayout.addWidget(self.btnReceiptFolder)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.btnGenerate = QtWidgets.QPushButton(self.tabGenerator)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.btnGenerate.setFont(font)
        self.btnGenerate.setObjectName("btnGenerate")
        self.verticalLayout.addWidget(self.btnGenerate)
        self.tabWidget.addTab(self.tabGenerator, "")

        self.retranslateUi(ExpenseReportGenerator)
        self.tabWidget.setCurrentIndex(1)
        self.btnPDF.clicked.connect(self.browsePDF)
        self.btnReceiptFolder.clicked.connect(self.browseReceipts)
        self.lineEditPDF.returnPressed.connect(self.returnBrowsePDF)
        self.lineEditReceiptFolder.returnPressed.connect(self.returnBrowseReceipts)
        self.btnGenerate.clicked.connect(self.writeExpenseReport)
        QtCore.QMetaObject.connectSlotsByName(ExpenseReportGenerator)

    def retranslateUi(self, ExpenseReportGenerator):
        _translate = QtCore.QCoreApplication.translate
        ExpenseReportGenerator.setWindowTitle(_translate("ExpenseReportGenerator", "Expense Report Generator"))
        self.txtInfo.setHtml(_translate("ExpenseReportGenerator", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:14pt; font-weight:600; text-decoration: underline;\">Expense Report Generator</span></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:10pt;\">by Patrick Dugan</span></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:10pt;\">Source Code available on: </span></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><a href=\"https://github.com/dpdug4n/ExpenseReportGenerator\"><span style=\" font-size:10pt; text-decoration: underline; color:#0000ff;\">GitHub</span></a></p>\n"
"<p align=\"center\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:10pt;\"><br /></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:10pt; font-weight:600; text-decoration: underline;\">HOW TO</span></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:10pt;\">Select the PDF file of the expense report spreadsheet.</span></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:10pt;\"> Select the folder containing images of reciepts.</span></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:10pt;\">Click \'Generate\'.</span></p>\n"
"<p align=\"center\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:10pt;\"><br /></p>\n"
"<p align=\"center\" style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:10pt;\"><br /></p></body></html>"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tabInfo), _translate("ExpenseReportGenerator", "Information"))
        self.lblPDF.setText(_translate("ExpenseReportGenerator", "PDF file:"))
        self.btnPDF.setText(_translate("ExpenseReportGenerator", "Browse"))
        self.lblReceiptFolder.setText(_translate("ExpenseReportGenerator", "Receipt Image Folder:"))
        self.btnReceiptFolder.setText(_translate("ExpenseReportGenerator", "Browse"))
        self.btnGenerate.setText(_translate("ExpenseReportGenerator", "Generate"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tabGenerator), _translate("ExpenseReportGenerator", "Generator"))

    #do stuff here
    @pyqtSlot()
    def returnBrowsePDF( self ):
        pass

    @pyqtSlot()
    def browsePDF( self ):
        options = QtWidgets.QFileDialog.Options()

        fileNamePDF, _ = QtWidgets.QFileDialog.getOpenFileName(
                        None,
                        "Select Expense Report PDF",
                        "",
                        "PDF Files (*.pdf)",
                        options=options)
        self.lineEditPDF.setText(fileNamePDF)

    @pyqtSlot()
    def returnBrowseReceipts( self ):
        pass

    @pyqtSlot()
    def browseReceipts( self ):
        fileNameReceiptFolder = QtWidgets.QFileDialog.getExistingDirectory(
                        None,
                        "Select Folder with Receipt Images",
                        "",
                        QtWidgets.QFileDialog.ShowDirsOnly,
                        )
        self.lineEditReceiptFolder.setText(fileNameReceiptFolder)
    
    @pyqtSlot()
    def writeExpenseReport( self ):
        expenseReportFilename = self.lineEditPDF.text()
        msgBox = QtWidgets.QMessageBox()
        msgBox.setIcon(QtWidgets.QMessageBox.Information)
        msgBox.setText('Please wait for a few seconds after clicking okay \n Your Expense Report will be in the receipt folder. \n '+
            'Make sure you select the correct folder and only receipts images & the expense report is in the folder.\n'+
            'Everthing in the folder but the Expense Report will be deleted.')
        def makePdf(imageDir, SaveToDir):
            os.chdir(imageDir)
            try:
                for j in os.listdir(os.getcwd()):
                    os.chdir(imageDir)
                    fname, fext = os.path.splitext(j)
                    newfilename = fname + ".pdf"
                    im = Image.open(fname + fext)
                    if im.mode == "RGBA":
                        im = im.convert("RGB")
                    os.chdir(SaveToDir)
                    if not os.path.exists(newfilename):
                        im.save(newfilename, "PDF", resolution=100.0)
                    im.close()
            except Exception as e:
                print(e)
        def merger():
            pdfs2merge= [expenseReportFilename, ]
            for _ in os.listdir(os.getcwd()):
                if _.endswith(".pdf"):
                    pdfs2merge.append(_)
            
            pdfWriter = PyPDF2.PdfFileWriter()
            for filename in pdfs2merge:
                pdfFileObj = open(filename, 'rb')
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                for pageNum in range(pdfReader.numPages):
                    pageObj = pdfReader.getPage(pageNum)
                    pdfWriter.addPage(pageObj)
                pdfFileObj.close()
            pdfOutput = open('Expense Report.pdf', 'wb')
            pdfWriter.write(pdfOutput)
            pdfOutput.close()
        imageDir = self.lineEditReceiptFolder.text() 
        SaveToDir = self.lineEditReceiptFolder.text() 
        msgBox.exec()
        makePdf(imageDir, SaveToDir)
        merger()
        for _ in os.listdir(os.getcwd()):
            try:
                if not _.startswith('Expense Report'):
                    os.remove(_)
            except Exception as e: print(e)

        
          


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ExpenseReportGenerator = QtWidgets.QDialog()
    ui = Ui_ExpenseReportGenerator()
    ui.setupUi(ExpenseReportGenerator)
    ExpenseReportGenerator.show()
    sys.exit(app.exec_())
