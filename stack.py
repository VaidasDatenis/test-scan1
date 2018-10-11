import sys
import os
from PyQt5.QtGui     import QIcon
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog,
                             QPushButton, QTextEdit, QLineEdit, QToolTip,
                             QMessageBox, QProgressBar, QWidget,
                             QHBoxLayout, QVBoxLayout)
from PyQt5.QtCore    import QCoreApplication, QBasicTimer, QThread, pyqtSignal


class AThread(QThread):

    # Use the signals to send the information.
    updateSignal = pyqtSignal(list, int)

    def __init__(self, folder, allFiles):
        super().__init__()
        self.folder  = folder
        self.cons    = 1 if allFiles<100 else allFiles//100+1
        self.indPbar = 1
        self.ind     = 0
        self.files   = []

    def run(self):
        for root, dirs, files in os.walk(self.folder):
            self.updateSignal.emit(["\n{}:".format(root),], self.indPbar)
            for file in files:
                if self.cons == 1:
                    self.indPbar += 1
                    self.updateSignal.emit([file,], self.indPbar)
                else:
                    self.files.append("{}".format(file))
                    self.ind += 1

                if  self.ind >= 100:
                    self.indPbar += 1
                    self.updateSignal.emit(self.files, self.indPbar)
                    self.ind = 0
                    self.cons -= 1
                    self.files = []
                QThread.msleep(1)
            self.updateSignal.emit(self.files, self.indPbar)
            self.files = []
        self.updateSignal.emit(self.files, 100)


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon("C:/Users/dava8001/AppData/Local/Programs/Python/Python36-32/test-scan1/barcode.png"))
        self.setWindowTitle("Barcode Scanner")
        self.resize(500, 400)
        self.allFiles = 0

        self.InitWindow()

    def InitWindow(self):
        self.button = QPushButton("Choose Folder Path")
        self.button.clicked.connect(self.getFolder)
        self.lineEdit = QLineEdit("C:/") # "C:/Users/dava8001/Desktop"
        self.lineEdit.setReadOnly(True)

        self.btnScan  = QPushButton("Scan")
        self.btnScan.clicked.connect(
            lambda folder=self.lineEdit.text(): self.scan(self.lineEdit.text()))
        self.btnScan.setEnabled(False)

        self.btnClose = QPushButton("Close")
        self.btnClose.clicked.connect(self.CloseApp)
        self.textEdit = QTextEdit(self)
        self.pbar     = QProgressBar(self)

        centralWidget = QWidget()
        self.setCentralWidget(centralWidget)

        layoutH = QHBoxLayout()
        layoutH.addWidget(self.button)
        layoutH.addStretch(1)
        layoutH.addWidget(self.btnScan)

        layoutV = QVBoxLayout(centralWidget)
        layoutV.addLayout(layoutH)
        layoutV.addWidget(self.lineEdit)
        layoutV.addWidget(self.textEdit)
        layoutV.addWidget(self.pbar)
        layoutV.addWidget(self.btnClose)

    def scan(self, folder):

        self.btnScan.setEnabled(False)
        self.thread = AThread(folder, self.allFiles)
        self.thread.finished.connect(self.closeW)
        self.thread.start()
        self.thread.updateSignal.connect(self.update)

    def update(self, texts, val):
        text = "\n".join(texts)
        self.textEdit.append(text)
        self.pbar.setValue(val)

    def closeW(self):
        self.textEdit.append("--- Directory scan complete! ---")
        self.btnScan.setEnabled(True)

    def getFolder(self):
        options = QFileDialog.DontResolveSymlinks | QFileDialog.ShowDirsOnly
        folder  = QFileDialog.getExistingDirectory(self,
                            "Open Folder",
                            self.lineEdit.text(),
                            options=options)
        if folder:
            self.pbar.setValue(0)
            self.lineEdit.setText(folder)

            self.textEdit.append("\n{}:".format(folder))
            self.allFiles = 0
            for root, dirs, files in os.walk(folder):
                self.allFiles += len(files)
            self.textEdit.append("    total files: {}".format(self.allFiles))
            self.btnScan.setEnabled(True)
        else:
            self.textEdit.append("Sorry, choose a directory to scan!")
            self.btnScan.setEnabled(False)

    def CloseApp(self):
        reply = QMessageBox.question(self, "Close Message",
                                     "Are you sure you want to close?",
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = Window()
    window.show()
    sys.exit(app.exec())