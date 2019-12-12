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

    def defaults(self):
        self.tabWidget.setCurrentIndex(0)
        self.design_1.setEnabled(False)
        self.design_2.setEnabled(False)
        self.design_3.setEnabled(False)
        self.design_4.setEnabled(False)
        self.design_5.setEnabled(False)
        self.design_6.setEnabled(False)
        self.design_7.setEnabled(False)
        self.design_8.setEnabled(False)
        self.design_9.setEnabled(False)
        self.design_10.setEnabled(False)
        self.dock_home.setVisible(False)
        self.dock_archive.setVisible(False)
        self.dock_messages.setVisible(False)
        self.dock_upload.setVisible(False)
        self.dock_settings.setVisible(False)
        self.dock_exit.setVisible(False)
        self.dock_home_2.setVisible(False)
        self.dock_archive_2.setVisible(False)
        self.dock_messages_2.setVisible(False)
        self.dock_upload_2.setVisible(False)
        self.dock_exit_2.setVisible(False)
        self.login_error_message.setText('')
        self.login_username.setText('')
        self.login_password.setText('')
        self.mail_messages_number.setVisible(False)
        self.mail_chat_number.setVisible(False)
        self.user_mail_messages_number.setVisible(False)
        self.user_mail_chat_number.setVisible(False)


    def Handle_UI_Changes(self):
        self.defaults()
        self.tabWidget.tabBar().setVisible(False)
        self.mail_tab_widget.tabBar().setVisible(False)
        self.user_mail_tab_widget.tabBar().setVisible(False)
        self.left_pane.tabBar().setVisible(False)
        #pdf_engine_startups
        self.archive_web_engine.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, '')))

        #settings
        self.settings_table_widget_accounts.setEditTriggers(QTableWidget.NoEditTriggers)
        self.settings_table_widget_accounts.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
        self.settings_table_widget_accounts.setColumnHidden(0,True)
        #mail
        self.mail_tab_widget.setCurrentIndex(0)
        self.mail_inbox_tab_widget.setCurrentIndex(0)
        self.mail_approved_webwidget.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, '')))
        self.mail_rejected_webwidget.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, '')))

        #usermail
        self.user_mail_tab_widget.setCurrentIndex(0)
        self.user_mail_inbox_tab_widget.setCurrentIndex(0)
        self.user_mail_approved_webwidget.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, '')))
        self.user_mail_rejected_webwidget.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, '')))


    def Handle_Buttons(self):
        #dock icons
        self.dock_home.clicked.connect(lambda: self.tabWidget.setCurrentIndex(1))
        self.dock_archive.clicked.connect(lambda: self.tabWidget.setCurrentIndex(2))
        self.dock_messages.clicked.connect(self.dock_messages_action)
        self.dock_upload.clicked.connect(lambda: self.tabWidget.setCurrentIndex(4))
        self.dock_settings.clicked.connect(lambda : self.tabWidget.setCurrentIndex(5))
        self.dock_exit.clicked.connect(self.dock_exit_action)
        #dock2 icons
        self.dock_home_2.clicked.connect(lambda: self.tabWidget.setCurrentIndex(1))
        self.dock_archive_2.clicked.connect(lambda: self.tabWidget.setCurrentIndex(2))
        self.dock_messages_2.clicked.connect(self.dock_messages_2_action)
        self.dock_upload_2.clicked.connect(lambda: self.tabWidget.setCurrentIndex(4))
        self.dock_exit_2.clicked.connect(self.dock_exit_action)
        #login
        self.login_button.clicked.connect(self.login_button_action)
        self.login_username.textChanged.connect(lambda: self.login_error_message.setText(''))
        self.login_password.textChanged.connect(lambda: self.login_error_message.setText(''))
        #mail
        self.mail_inbox_button.clicked.connect(lambda: self.mail_tab_widget.setCurrentIndex(0))
        self.mail_compose_button.clicked.connect(lambda: self.mail_tab_widget.setCurrentIndex(1))
        self.mail_chat_button.clicked.connect(lambda: self.mail_tab_widget.setCurrentIndex(2))
        #user_mail
        self.user_mail_inbox_button.clicked.connect(lambda: self.user_mail_tab_widget.setCurrentIndex(0))
        self.user_mail_compose_button.clicked.connect(lambda: self.user_mail_tab_widget.setCurrentIndex(1))
        self.user_mail_chat_button.clicked.connect(lambda: self.user_mail_tab_widget.setCurrentIndex(2))
        self.user_mail_request_button.clicked.connect(lambda: self.user_mail_tab_widget.setCurrentIndex(3))

        QtWebEngineWidgets.QWebEngineProfile.defaultProfile().downloadRequested.connect(self.on_download_request)


    def dock_messages_action(self):
        self.tabWidget.setCurrentIndex(3)
        self.mail_tab_widget.setCurrentIndex(0)
        self.mail_inbox_tab_widget.setCurrentIndex(0)

    def dock_messages_2_action(self):
        self.tabWidget.setCurrentIndex(6)
        self.user_mail_tab_widget.setCurrentIndex(0)
        self.user_mail_inbox_tab_widget.setCurrentIndex(0)


    def dock_exit_action(self):
        buttonReply = QMessageBox.question(self, '', "Are you sure you want to Exit?",
                                           QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if buttonReply == QMessageBox.Yes:
            self.show()
            self.defaults()
        else:
            print('No clicked.')

    def set_dock_visible_admin(self):
        self.left_pane.setCurrentIndex(0)
        self.dock_home.setVisible(True)
        self.dock_archive.setVisible(True)
        self.dock_messages.setVisible(True)
        self.dock_upload.setVisible(True)
        self.dock_settings.setVisible(True)
        self.dock_exit.setVisible(True)

    def set_dock_visible_user(self):
        self.left_pane.setCurrentIndex(1)
        self.dock_home_2.setVisible(True)
        self.dock_archive_2.setVisible(True)
        self.dock_messages_2.setVisible(True)
        self.dock_upload_2.setVisible(True)
        self.dock_exit_2.setVisible(True)

    def login_button_action(self):
        if self.login_username.text() == 'admin' and self.login_password.text() == 'admin':
            self.set_dock_visible_admin()
            self.tabWidget.setCurrentIndex(1)
            self.upload_mark_confidential_widget.setEnabled(True)
            self.archive_delete_button.setEnabled(True)
        elif self.login_username.text() == 'user' and self.login_password.text() == 'user':
            self.set_dock_visible_user()
            self.tabWidget.setCurrentIndex(1)
            self.upload_mark_confidential_widget.setEnabled(False)
            self.archive_delete_button.setEnabled(False)
        else:
            self.login_error_message.setText('Wrong username or password!!')


    @QtCore.pyqtSlot(QtWebEngineWidgets.QWebEngineDownloadItem)
    def on_download_request(self,download):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save File", "sample.pdf", "*.pdf")
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