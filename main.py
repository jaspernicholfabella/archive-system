import sys
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets
import subprocess

ui, _ = loadUiType('archive_system.ui')


class MainApp(QMainWindow, ui):
    PDFJS = 'file:///C:/__Jasper__/archive-system/pdfjs/web/viewer.html'
    PDF = ''
    PDF = 'file:///C:/__Jasper__/archive-system/pdfjs/web/sample.pdf'


    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Handle_UI_Changes()
        self.Handle_Buttons()

    def Handle_UI_Changes(self):
        self.tabWidget.tabBar().setVisible(False)
        self.test_engine.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, self.PDF)))

    def Handle_Buttons(self):
        QtWebEngineWidgets.QWebEngineProfile.defaultProfile().downloadRequested.connect(self.on_download_request)


    @QtCore.pyqtSlot(QtWebEngineWidgets.QWebEngineDownloadItem)
    def on_download_request(self,download):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save File", "sample.pdf", "*.pdf"
        )
        if path:
            download.setPath(path)
            download.accept()


def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()


if __name__ == '__main__':
    main()