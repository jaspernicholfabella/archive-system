import sys
from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets



class Window(QtWebEngineWidgets.QWebEngineView):
    PDFJS = 'file:///C:/__Jasper__/archive-system/pdfjs/web/viewer.html'
    PDF = 'file:///C:/__Jasper__/archive-system/pdfjs/web/sample.pdf'
    def __init__(self):
        super(Window, self).__init__()
        self.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, self.PDF)))

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = Window()
    window.setGeometry(600, 50, 800, 600)
    window.show()
    sys.exit(app.exec_())