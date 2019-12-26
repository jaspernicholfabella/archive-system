import sys
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets
from sqlalchemy import asc
from fpdf import FPDF
from PIL import Image
import sqlconn as sqc
import win32com.client
import shutil
from cryptography.fernet import Fernet
from collections import OrderedDict
import datetime
import os
ui, _ = loadUiType('archive_system.ui')
accounts_ui, _ = loadUiType('add_account.ui')
doctype_ui, _ = loadUiType('add_document_type.ui')
docinfo_ui, _ = loadUiType('document_information.ui')
routing_ui, _ = loadUiType('routing.ui')
upload_ui, _ = loadUiType('save_to_archive.ui')
#globals
global useraccount
global useraccount_name
global settings_account_table
global settings_doctype_table
global sharedrive
global docinfo_selected_item
global docinfo_dict
global mail_tab_widget
global mail_inbox_tab_widget
global mail_inbox_waiting_approve
global mail_inbox_waiting_reject
global mail_inbox_waiting_list
global mail_inbox_approved_list
global mail_inbox_rejected_list
global mail_inbox_text_browser
global mail_approved_text_browser
global mail_rejected_text_browser
global mail_inbox_waiting_view
global mail_inbox_waiting_attached_name
global mail_inbox_approved_button_container
global mail_inbox_approved_comment


class Routing_Dialogue(QDialog,routing_ui):
    dir = ''
    from_who = ''
    sender = ''
    edit_id = 0
    routing_type = ''
    def __init__(self,parent=None):
        super(Routing_Dialogue,self).__init__(parent)
        self.setupUi(self)

    def ShowDialogue(self,id,sender,dir,from_who,routing_type):
        self.dir = dir
        self.sender = sender
        self.edit_id = id
        self.from_who = from_who
        self.routing_type = routing_type
        if routing_type == 'accepted':
            self.routing_label.setText('Accept Request')
            self.routing_label_comment.setText('Comments : ')
        elif routing_type == 'rejected':
            self.routing_label.setText('Reject Request')
            self.routing_label_comment.setText('Reasons : ')

        self.buttonBox.accepted.connect(self.mail_urequest_send_action)
        self.mail_urequest_upload_images.clicked.connect(self.mail_urequest_upload_images_action)
        self.mail_urequest_upload_documents.clicked.connect(self.mail_urequest_upload_documents_action)

    def mail_urequest_upload_images_action(self):
        self.mail_urequest_upload_filetype = 'img'
        self.mail_urequest_upload_list_dictionary = {}
        self.mail_urequest_upload_list.clear()
        self.mail_urequest_upload_list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        image_files, _ = QFileDialog.getOpenFileNames(self, "Open Images",
                                                      '',
                                                      "Image File (*.jpg *.png)", options=options)
        for image_file in image_files:
            temp = image_file.split('/')
            self.mail_urequest_upload_list_dictionary.update({temp[len(temp) - 1]: image_file})
            self.mail_urequest_upload_list.addItem(temp[len(temp) - 1])

    def mail_urequest_upload_documents_action(self):
        self.mail_urequest_upload_filetype = 'doc'
        self.mail_urequest_upload_list_dictionary = {}
        self.mail_urequest_upload_list.clear()

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        doc_name, _ = QFileDialog.getOpenFileName(self, "Open Document",
                                                  '',
                                                  "Document File (*.docx *.doc *.xls *.xlsx *.pdf)",
                                                  options=options)
        temp = doc_name.split('/')
        self.mail_urequest_upload_list_dictionary.update({temp[len(temp) - 1]: doc_name})
        self.mail_urequest_upload_list.addItem(temp[len(temp) - 1])

    def mail_urequest_send_action(self):
        global useraccount
        have_attached = False
        filerror = False
        alias = ''
        filetype = ''
        iseditable = False

        if self.mail_urequest_upload_list.count() > 0:
            if self.mail_urequest_upload_filetype == 'img':
                have_attached = True
                filetype = 'pdf'
                iseditable = False
                dir = self.dir
                alias = 'reply_' + str(self.from_who + '-routing-' + datetime.datetime.now().strftime(
                    "%m_%d_%Y_%H_%M_%S")).replace('-', '_').replace(' ', '_').lower()
                try:
                    os.makedirs(dir)
                except:
                    print('directory already exists.')
                pdf = FPDF()
                for i in range(self.mail_urequest_upload_list.count()):
                    imageFile = self.mail_urequest_upload_list_dictionary[
                        self.mail_urequest_upload_list.item(i).text()]
                    cover = Image.open(imageFile)
                    width, height = cover.size
                    width, height = float(width * 0.264583), float(height * 0.264583)
                    pdf_size = {'P': {'w': 210, 'h': 297}, 'L': {'w': 297, 'h': 210}}
                    orientation = 'P' if width < height else 'L'
                    width = width if width < pdf_size[orientation]['w'] else pdf_size[orientation]['w']
                    height = height if height < pdf_size[orientation]['h'] else pdf_size[orientation]['h']
                    pdf.add_page(orientation=orientation)
                    pdf.image(imageFile, 0, 0, width, height)
                pdf.output(dir + '/' + alias + '.pdf', "F")

            elif self.mail_urequest_upload_filetype == 'doc':
                have_attached = True
                filerror = False
                dir = self.settings_sharedrive_loc.text()
                alias = str(self.from_who + '-routing-' + datetime.datetime.now().strftime(
                    "%m_%d_%Y_%H_%M_%S")).replace('-', '_').replace(' ', '_').lower()
                try:
                    os.makedirs(dir)
                except:
                    print('directory already exists.')

                doc = self.mail_urequest_upload_list_dictionary[self.mail_urequest_upload_list.item(0).text()]
                if '.doc' in doc or '.docx' in doc:
                    filetype = 'docx'
                    iseditable = True
                    try:
                        word = win32com.client.gencache.EnsureDispatch("Word.Application")
                        worddoc = word.Documents.Open(os.path.abspath(doc))
                        worddoc.SaveAs(os.path.abspath(dir + '\\' + alias + ".pdf"), FileFormat=17)
                        worddoc.Close()
                        word.Quit()
                        shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.docx'))
                    except:
                        filerror = True
                        msg = QMessageBox()
                        msg.setIcon(QMessageBox.Critical)
                        msg.setText("File Conversion Failed, try Restart Applicaton!")
                        msg.setInformativeText('File Conversion Error')
                        msg.setWindowTitle("Error")
                        msg.exec_()
                elif '.xls' in doc or '.xlsx' in doc:
                    have_attached = True
                    filetype = 'xlsx'
                    iseditable = True
                    try:
                        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
                        wb = excel.Workbooks.Open(os.path.abspath(doc))
                        count = wb.Sheets.Count
                        ws_index_list = []
                        for i in range(1, count + 1):
                            ws_index_list.append(i)
                        wb.WorkSheets(ws_index_list).Select()
                        # Save
                        wb.ActiveSheet.ExportAsFixedFormat(0, os.path.abspath(dir + '\\' + alias + '.pdf'))
                        wb.Close()
                        excel.Quit()
                        shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.xlsx'))
                    except:
                        filerror = True
                        msg = QMessageBox()
                        msg.setIcon(QMessageBox.Critical)
                        msg.setText("File Conversion Failed, try Restart Applicaton!")
                        msg.setInformativeText('File Conversion Error')
                        msg.setWindowTitle("Error")
                        msg.exec_()
                elif '.pdf' in doc:
                    filetype = 'pdf'
                    iseditable = False
                    shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.pdf'))
        else:
            have_attached = False

        if filerror == False:
            engine = sqc.Database().engine
            conn = engine.connect()
            archive_mail = sqc.Database().archive_mail
            ins = archive_mail.update().where(archive_mail.c.mailid == self.edit_id).\
                values(
                sender=useraccount_name,
                reciever=self.sender,
                isseen=False,
                status=self.routing_type,
                status_message=self.routing_text.toPlainText(),
                reply_have_attached =have_attached,
                reply_attached_alias = alias,
                reply_is_editable = iseditable,
                reply_filetype = filetype
            )
            conn.execute(ins)
            self.mail_inbox_refresh()
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Reply was not Sent!")
            msg.setInformativeText('Something went Wrong')
            msg.setWindowTitle("Error")
            msg.exec_()

    mail_inbox_dictionary = {}
    def mail_inbox_refresh(self):
        global mail_tab_widget
        global mail_inbox_tab_widget
        global mail_inbox_waiting_approve
        global mail_inbox_waiting_reject
        global mail_inbox_waiting_list
        global mail_inbox_approved_list
        global mail_inbox_rejected_list
        global mail_inbox_text_browser
        global mail_approved_text_browser
        global mail_rejected_text_browser
        global mail_inbox_waiting_view
        global mail_inbox_waiting_attached_name
        global mail_inbox_approved_button_container
        global mail_inbox_approved_comment
        mail_inbox_approved_comment.setText('')
        mail_inbox_approved_button_container.setVisible(False)
        mail_tab_widget.setCurrentIndex(0)
        mail_inbox_tab_widget.setCurrentIndex(0)
        mail_inbox_waiting_approve.setVisible(False)
        mail_inbox_waiting_reject.setVisible(False)
        mail_inbox_waiting_list.clear()
        mail_inbox_approved_list.clear()
        mail_inbox_rejected_list.clear()
        mail_inbox_text_browser.setText('')
        mail_approved_text_browser.setText('')
        mail_rejected_text_browser.setText('')
        mail_inbox_waiting_view.setVisible(False)
        mail_inbox_waiting_attached_name.setText('')
        engine = sqc.Database().engine
        conn = engine.connect()
        archive_mail = sqc.Database().archive_mail
        s = archive_mail.select()
        s_value = conn.execute(s)

        for val in s_value:
            self.mail_inbox_dictionary.update({
                str(val[0]):{
                    'sender':val[1],
                    'reciever':val[2],
                    'date_sent':val[3],
                    'from_who':val[4],
                    'subject':val[5],
                    'action':val[6],
                    'have_attached':val[7],
                    'attached_alias':val[8],
                    'isseen':val[9],
                    'iseditable':val[10],
                    'filetype':val[11],
                    'status':val[12],
                    'status_message':val[13]}
            })
        self.mail_inbox_dictionary = OrderedDict(sorted(self.mail_inbox_dictionary.items(), reverse=True))

        for key in self.mail_inbox_dictionary.keys():
            if self.mail_inbox_dictionary[key]['status'] == 'wait':
                mail_inbox_waiting_list.addItem('{}-(MSGID{})-[{}]'.format(self.mail_inbox_dictionary[key]['sender'],key,self.mail_inbox_dictionary[key]['date_sent']))
            elif self.mail_inbox_dictionary[key]['status'] == 'accepted':
                mail_inbox_approved_list.addItem('{}-(MSGID{})-[{}]'.format(self.mail_inbox_dictionary[key]['sender'], key,self.mail_inbox_dictionary[key]['date_sent']))
            elif self.mail_inbox_dictionary[key]['status'] == 'rejected':
                mail_inbox_rejected_list.addItem('{}-(MSGID{})-[{}]'.format(self.mail_inbox_dictionary[key]['sender'], key,self.mail_inbox_dictionary[key]['date_sent']))

class Docinfo_Dialogue(QDialog,docinfo_ui):

    def __init__(self,parent=None):
        super(Docinfo_Dialogue,self).__init__(parent)
        self.setupUi(self)
        self.ShowDialogue()

    def ShowDialogue(self):
        self.docinfo_name.setText(docinfo_selected_item)
        self.docinfo_type.setText(docinfo_dict[docinfo_selected_item]['doctype'])
        self.docinfo_description.setPlainText(docinfo_dict[docinfo_selected_item]['description'])
        self.docinfo_date.setText(str(docinfo_dict[docinfo_selected_item]['date_uploaded']))

class Accounts_Dialogue(QDialog,accounts_ui):
    edit_id = 0
    operationType = ''

    def __init__(self,parent=None):
        super(Accounts_Dialogue,self).__init__(parent)
        self.setupUi(self)

    def ShowDialogue(self,id,username,password,operationType=''):
        self.username.setText(username)
        self.password.setText(password)
        self.edit_id = id
        self.operationType = operationType
        self.buttonBox.accepted.connect(self.ok_button)

    def ok_button(self):
        engine = sqc.Database().engine
        archive_admin = sqc.Database().archive_admin
        conn = engine.connect()

        if self.operationType == 'edit':
            self.account_label.setText('Edit Account')
            s = archive_admin.update().where(archive_admin.c.userid == self.edit_id).\
                values(username = self.username.text(),
                       password = self.password.text(),
                       previlage = self.previlage_combo.currentText())
            conn.execute(s)
            self.show_settings()

        elif self.operationType == 'add':
            self.account_label.setText('Add Account')
            s = archive_admin.insert().values(
                username=self.username.text(),
                password=self.password.text(),
                previlage=self.previlage_combo.currentText())
            conn.execute(s)
            self.show_settings()

        conn.close()

    def show_settings(self):
        global settings_account_table
        settings_account_table.setRowCount(0)
        engine = sqc.Database().engine
        archive_admin = sqc.Database().archive_admin
        conn= engine.connect()
        #admin_table
        s = archive_admin.select().order_by(asc(archive_admin.c.username))
        s_value = conn.execute(s)
        table = settings_account_table
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[2])))
            table.setItem(row_position, 3, QTableWidgetItem(str(val[3])))
        conn.close()

class Doctype_Dialogue(QDialog,doctype_ui):
    edit_id = 0
    operationType = ''

    def __init__(self,parent=None):
        super(Doctype_Dialogue,self).__init__(parent)
        self.setupUi(self)

    def ShowDialogue(self,id,doctype,operationType=''):
        self.doctype.setText(doctype)
        self.edit_id = id
        self.operationType = operationType
        self.buttonBox.accepted.connect(self.ok_button)

    def ok_button(self):
        engine = sqc.Database().engine
        archive_doctype = sqc.Database().archive_doctype
        conn = engine.connect()

        if self.operationType == 'edit':
            self.doctype_label.setText('Edit Document Type')
            s = archive_doctype.update().where(archive_doctype.c.doctype_id == self.edit_id).\
                values(document_type = self.doctype.text())
            conn.execute(s)
            self.show_settings()

        elif self.operationType == 'add':
            self.doctype_label.setText('Add Document Type')
            s = archive_doctype.insert().values(
                document_type=self.doctype.text())
            conn.execute(s)
            self.show_settings()

        conn.close()

    def show_settings(self):
        global settings_doctype_table
        global upload_doctype
        global archive_options
        settings_doctype_table.setRowCount(0)
        engine = sqc.Database().engine
        archive_doctype = sqc.Database().archive_doctype
        conn= engine.connect()
        #doctype_table
        s = archive_doctype.select().order_by(asc(archive_doctype.c.document_type))
        s_value = conn.execute(s)
        table = settings_doctype_table
        upload_doctype.clear()
        archive_options.clear()
        archive_options.addItem('All')
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            upload_doctype.addItem(val[1])
            archive_options.addItem(val[1])
        conn.close()


'''
------------------------------------------------------------------------------------------------------------------------
MAIN CLASS
------------------------------------------------------------------------------------------------------------------------
'''

class MainApp(QMainWindow, ui):
    PDFJS = 'file:///C:/__Jasper__/archive-system/pdfjs/web/viewer.html'
    PDF = ''


    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Handle_UI_Changes()
        self.Global_Variables()
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

        self.show_settings()
        self.upload_document_name.setText('')

    def Global_Variables(self):
        global settings_account_table
        settings_account_table = self.settings_account_table
        global settings_doctype_table
        settings_doctype_table = self.settings_doctype_table
        global sharedrive
        sharedrive = self.settings_sharedrive_loc.text()
        global archive_options
        archive_options = self.archive_options
        global upload_doctype
        upload_doctype = self.upload_doctype

        global mail_tab_widget
        mail_tab_widget = self.mail_tab_widget
        global mail_inbox_tab_widget
        mail_inbox_tab_widget = self.mail_inbox_tab_widget
        global mail_inbox_waiting_approve
        mail_inbox_waiting_approve = self.mail_inbox_waiting_approve
        global mail_inbox_waiting_reject
        mail_inbox_waiting_reject = self.mail_inbox_waiting_reject
        global mail_inbox_waiting_list
        mail_inbox_waiting_list = self.mail_inbox_waiting_list
        global mail_inbox_approved_list
        mail_inbox_approved_list = self.mail_inbox_approved_list
        global mail_inbox_rejected_list
        mail_inbox_rejected_list = self.mail_inbox_rejected_list
        global mail_inbox_text_browser
        mail_inbox_text_browser = self.mail_inbox_text_browser
        global mail_approved_text_browser
        mail_approved_text_browser = self.mail_approved_text_browser
        global mail_rejected_text_browser
        mail_rejected_text_browser = self.mail_rejected_text_browser
        global mail_inbox_waiting_view
        mail_inbox_waiting_view = self.mail_inbox_waiting_view
        global mail_inbox_waiting_attached_name
        mail_inbox_waiting_attached_name = self.mail_inbox_waiting_attached_name
        global mail_inbox_approved_button_container
        mail_inbox_approved_button_container = self.mail_inbox_approved_button_container
        global mail_inbox_approved_comment
        mail_inbox_approved_comment = self.mail_inbox_approved_comment

    def Handle_UI_Changes(self):

        self.defaults()
        self.tabWidget.tabBar().setVisible(False)
        self.mail_tab_widget.tabBar().setVisible(False)

        self.left_pane.tabBar().setVisible(False)
        #pdf_engine_startups
        self.archive_web_engine.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, self.PDF)))

        #settings
        self.settings_account_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.settings_account_table.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
        self.settings_account_table.setColumnHidden(0,True)

        self.settings_doctype_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.settings_doctype_table.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
        self.settings_doctype_table.setColumnHidden(0,True)

        #mail
        self.mail_tab_widget.setCurrentIndex(0)
        self.mail_inbox_tab_widget.setCurrentIndex(0)
        self.mail_compose_button.setVisible(False)
        self.mail_chat_button.setVisible(False)

    def Handle_Buttons(self):
        QtWebEngineWidgets.QWebEngineProfile.defaultProfile().downloadRequested.connect(self.on_download_request)
        #dock icons
        self.dock_home.clicked.connect(lambda: self.tabWidget.setCurrentIndex(1))
        self.dock_archive.clicked.connect(self.dock_archive_action)
        self.dock_messages.clicked.connect(self.dock_messages_action)
        self.dock_upload.clicked.connect(lambda: self.tabWidget.setCurrentIndex(4))
        self.dock_settings.clicked.connect(lambda : self.tabWidget.setCurrentIndex(5))
        self.dock_settings.clicked.connect(self.show_settings)
        self.dock_exit.clicked.connect(self.dock_exit_action)
        #dock2 icons
        self.dock_home_2.clicked.connect(lambda: self.tabWidget.setCurrentIndex(1))
        self.dock_archive_2.clicked.connect(self.dock_archive_action)
        self.dock_messages_2.clicked.connect(self.dock_messages_2_action)
        self.dock_upload_2.clicked.connect(lambda: self.tabWidget.setCurrentIndex(4))
        self.dock_exit_2.clicked.connect(self.dock_exit_action)
        #login
        self.login_button.clicked.connect(self.login_button_action)
        self.login_username.textChanged.connect(lambda: self.login_error_message.setText(''))
        self.login_password.textChanged.connect(lambda: self.login_error_message.setText(''))
        #settings
        self.settings_add_account.clicked.connect(self.settings_add_account_action)
        self.settings_edit_account.clicked.connect(lambda: self.settings_edit_account_action(self.settings_account_table))
        self.settings_delete_account.clicked.connect(lambda: self.settings_delete_account_action(self.settings_account_table))
        self.settings_add_doctype.clicked.connect(self.settings_add_doctype_action)
        self.settings_edit_doctype.clicked.connect(lambda: self.settings_edit_doctype_action(self.settings_doctype_table))
        self.settings_delete_doctype.clicked.connect(lambda: self.settings_delete_doctype_action(self.settings_doctype_table))
        self.settings_edit_share_drive.clicked.connect(self.settings_edit_sharedrive_action)
        #upload
        self.upload_image_button.clicked.connect(self.upload_image_button_action)
        self.upload_save_button.clicked.connect(self.upload_save_button_action)
        self.upload_document_button.clicked.connect(self.upload_document_button_action)
        #archive
        self.archive_doclist.doubleClicked.connect(self.archive_doclist_action)
        self.archive_delete_button.clicked.connect(self.archive_delete_button_action)
        self.archive_edit_button.clicked.connect(self.archive_edit_button_action)
        self.archive_options.currentTextChanged.connect(self.archive_options_action)
        self.archive_search.textChanged.connect(self.archive_search_action)
        self.archive_search_alphabetically.clicked.connect(self.archive_search_alphabetically_action)
        self.archive_search_historically.clicked.connect(self.archive_search_historically_action)
        # mail
        self.mail_inbox_button.clicked.connect(self.mail_inbox_button_action)
        self.mail_compose_button.clicked.connect(lambda: self.mail_tab_widget.setCurrentIndex(1))
        self.mail_request_button.clicked.connect(self.mail_request_button_action)
        self.mail_urequest_upload_images.clicked.connect(self.mail_urequest_upload_images_action)
        self.mail_urequest_upload_documents.clicked.connect(self.mail_urequest_upload_documents_action)
        self.mail_urequest_send.clicked.connect(self.mail_urequest_send_action)
        self.mail_inbox_waiting_list.clicked.connect(self.mail_inbox_waiting_list_action)
        self.mail_inbox_waiting_view.clicked.connect(self.mail_inbox_waiting_view_action)
        self.mail_inbox_waiting_approve.clicked.connect(self.mail_inbox_waiting_approve_action)
        self.mail_inbox_waiting_reject.clicked.connect(self.mail_inbox_waiting_reject_action)
        self.mail_inbox_approved_list.clicked.connect(self.mail_inbox_approved_list_action)
        self.mail_inbox_approved_view.clicked.connect(self.mail_inbox_approved_view_action)
        # chat
        self.mail_compose_button.clicked.connect(self.mail_compose_button_action)
        self.mail_compose_upload_images.clicked.connect(self.mail_compose_upload_images_action)
        self.mail_compose_upload_documents.clicked.connect(self.mail_compose_upload_documents_action)

    ##
    # DOCK
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

    ##
    # LOGIN
    def login_button_action(self):
        global useraccount
        global useraccount_name
        username = self.login_username.text()
        password = self.login_password.text()

        engine = sqc.Database().engine
        archive_admin = sqc.Database().archive_admin

        conn = engine.connect()
        s = archive_admin.select()
        s_value = conn.execute(s)

        for val in s_value:
            if str(username).lower() == str(val[1]).lower() and str(password).lower() == str(val[2]).lower():
                if val[3] == 'admin':
                    self.set_dock_visible_admin()
                    self.tabWidget.setCurrentIndex(1)
                    self.upload_mark_confidential_widget.setEnabled(True)
                    self.archive_delete_button.setEnabled(True)
                    useraccount = 'admin'
                    useraccount_name = val[1]
                elif val[3] == 'user':
                    self.set_dock_visible_user()
                    self.tabWidget.setCurrentIndex(1)
                    self.upload_mark_confidential_widget.setEnabled(False)
                    self.archive_delete_button.setEnabled(False)
                    useraccount = 'user'
                    useraccount_name = val[1]
            else:
                self.login_username.setText('')
                self.login_password.setText('')
                self.login_error_message.setText('Wrong username or password!!')

        conn.close()

    ##
    # SETTINGS
    def show_settings(self):
        self.settings_account_table.setRowCount(0)
        self.settings_doctype_table.setRowCount(0)
        engine = sqc.Database().engine
        archive_admin = sqc.Database().archive_admin
        archive_doctype = sqc.Database().archive_doctype
        conn= engine.connect()
        #admin_table
        s = archive_admin.select().order_by(asc(archive_admin.c.username))
        s_value = conn.execute(s)
        table = self.settings_account_table
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[2])))
            table.setItem(row_position, 3, QTableWidgetItem(str(val[3])))

        #doctype_table
        self.upload_doctype.clear()
        self.archive_options.clear()
        s = archive_doctype.select().order_by(asc(archive_doctype.c.document_type))
        s_value = conn.execute(s)
        table = self.settings_doctype_table
        self.archive_options.addItem('All')
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            self.upload_doctype.addItem(val[1])
            self.archive_options.addItem(val[1])
        #sharedrive
        archive_sharedrive = sqc.Database().archive_sharedrive
        s = archive_sharedrive.select()
        s_value = conn.execute(s)
        for val in s_value:
            self.settings_sharedrive_loc.setText(str(val[1]))

        conn.close()

    def settings_edit_account_action(self,table):
        try:
            r = table.currentRow()
            id = table.item(r,0).text()
            username = table.item(r,1).text()
            password = table.item(r,2).text()
            ad = Accounts_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id,username,password,operationType='edit')
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

    def settings_add_account_action(self):
        try:
            ad = Accounts_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id,'','',operationType='add')
        except:
            pass

    def settings_delete_account_action(self,table):
        try:
            r = table.currentRow()
            id = table.item(r, 0).text()
            engine = sqc.Database().engine
            conn = engine.connect()
            archive_admin = sqc.Database().archive_admin
            s = archive_admin.delete().where(archive_admin.c.userid == id)
            conn.execute(s)
            conn.close()
            self.show_settings()
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

    def settings_edit_doctype_action(self,table):
        try:
            r = table.currentRow()
            id = table.item(r,0).text()
            doctype = table.item(r,1).text()
            ad = Doctype_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id,doctype,operationType='edit')
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

    def settings_add_doctype_action(self):
        try:
            ad = Doctype_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id,'',operationType='add')
        except:
            pass

    def settings_delete_doctype_action(self,table):
        try:
            r = table.currentRow()
            id = table.item(r, 0).text()
            engine = sqc.Database().engine
            conn = engine.connect()
            archive_doctype = sqc.Database().archive_doctype
            s = archive_doctype.delete().where(archive_doctype.c.doctype_id == id)
            conn.execute(s)
            conn.close()
            self.show_settings()
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

    def settings_edit_sharedrive_action(self):
        file = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        engine = sqc.Database().engine
        conn = engine.connect()
        archive_sharedrive = sqc.Database().archive_sharedrive

        s = archive_sharedrive.update().where(archive_sharedrive.c.sdid == 1).\
            values(sharedrive = '{}/archive_data'.format(file))

        conn.execute(s)
        self.show_settings()

    ##
    # UPLOAD
    upload_filetype = ''
    upload_list_dictionary = {}
    upload_progress_value = 0
    def upload_clear(self):
        self.upload_document_name.setText('')
        self.upload_mark_confidential.setCheckState(QtCore.Qt.Unchecked)
        self.upload_doctype.setCurrentIndex(0)
        self.upload_description.setText('')
        self.upload_progress_value = 0
        self.upload_progress_bar.setValue(0)

    def upload_image_button_action(self):
        self.upload_clear()
        self.upload_filetype = 'img'
        self.upload_list_dictionary = {}
        self.upload_list.clear()
        self.upload_list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        image_files, _ = QFileDialog.getOpenFileNames(self,"Open Images",
                                                      '',
                                                      "Image File (*.jpg *.png)",options=options)

        for image_file in image_files:
            temp = image_file.split('/')
            self.upload_list_dictionary.update({temp[len(temp) - 1]:image_file})
            self.upload_list.addItem(temp[len(temp) - 1])

    def upload_document_button_action(self):
        self.upload_clear()
        self.upload_filetype = 'doc'
        self.upload_list.clear()
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        doc_name, _ = QFileDialog.getOpenFileName(self,"Open Document",
                                                      '',
                                                      "Document File (*.docx *.doc *.xls *.xlsx *.pdf)",options=options)
        temp = doc_name.split('/')
        self.upload_list_dictionary.update({temp[len(temp) - 1]:doc_name})
        self.upload_list.addItem(temp[len(temp) - 1])

    def upload_save_button_action(self):
        for i in range(self.upload_list.count()):
            print(self.upload_list_dictionary[self.upload_list.item(i).text()])

        if self.upload_list.count() > 0:
            if len(self.upload_document_name.text()) > 1:
                if self.upload_filetype == 'img':

                    dir = self.settings_sharedrive_loc.text()
                    alias = str(self.upload_document_name.text()).replace('-', '_').replace(' ', '_').lower()
                    try:
                        os.makedirs(dir)
                    except:
                        print('directory already exists.')
                    pdf = FPDF()
                    for i in range(self.upload_list.count()):
                        imageFile = self.upload_list_dictionary[self.upload_list.item(i).text()]
                        cover = Image.open(imageFile)
                        width, height = cover.size
                        width, height = float(width * 0.264583), float(height * 0.264583)
                        pdf_size = {'P': {'w': 210, 'h': 297}, 'L': {'w': 297, 'h': 210}}
                        orientation = 'P' if width < height else 'L'
                        width = width if width < pdf_size[orientation]['w'] else pdf_size[orientation]['w']
                        height = height if height < pdf_size[orientation]['h'] else pdf_size[orientation]['h']
                        pdf.add_page(orientation=orientation)
                        pdf.image(imageFile, 0, 0, width, height)
                        self.upload_progress_value += 90/self.upload_list.count()
                        self.upload_progress_bar.setValue(self.upload_progress_value)


                    pdf.output(dir+'/'+alias+'.pdf',"F")

                    ##for encryption
                    # key = Fernet.generate_key()
                    # cipher_suite = Fernet(key)
                    # cipher_text = cipher_suite.encrypt('{}.pdf'.format(alias).encode())
                    # cipher_string = str(cip)
                    # pdf.output(dir +'/'+ str(cipher_text) +'.ams', "F")
                    # #plain_text = cipher_suite.decrypt(cipher_text)

                    engine = sqc.Database().engine
                    conn = engine.connect()
                    archive_document = sqc.Database().archive_document
                    s = archive_document.select().where(archive_document.c.alias == alias)
                    s_value = conn.execute(s)
                    x = 0
                    for val in s_value:
                        x+=1
                    if x < 1:
                        if self.upload_mark_confidential.checkState() == QtCore.Qt.Checked:
                            isconfidential=True
                        else:
                            isconfidential=False
                        ins = archive_document.insert().values(
                            docname = self.upload_document_name.text(),
                            isconfidential = isconfidential,
                            doctype = self.upload_doctype.currentText(),
                            description = self.upload_description.toPlainText(),
                            alias = alias,
                            iseditable = False,
                            filetype = 'pdf',
                            date_uploaded = datetime.datetime.utcnow()
                        )
                        conn.execute(ins)
                        self.upload_progress_bar.setValue(100)
                    else:
                        msg = QMessageBox()
                        msg.setIcon(QMessageBox.Critical)
                        msg.setText("Document Name Error!")
                        msg.setInformativeText('Document Name Exists')
                        msg.setWindowTitle("Error")
                        msg.exec_()


                elif self.upload_filetype == 'doc':
                    filerror = False
                    dir = self.settings_sharedrive_loc.text()
                    alias = str(self.upload_document_name.text()).replace('-', '_').replace(' ', '_').lower()
                    try:
                        os.makedirs(dir)
                    except:
                        print('directory already exists.')

                    doc = self.upload_list_dictionary[self.upload_list.item(i).text()]
                    if '.doc' in doc or '.docx' in doc:
                        filetype = 'docx'
                        iseditable = True
                        try:
                            word = win32com.client.DispatchEx("Word.Application")
                            worddoc = word.Documents.Open(os.path.abspath(doc))
                            worddoc.SaveAs(os.path.abspath(dir+'\\'+alias+".pdf"), FileFormat=17)
                            worddoc.Close()
                            word.Quit()
                            self.upload_progress_value += 90
                            self.upload_progress_bar.setValue(self.upload_progress_value)
                            shutil.copyfile(os.path.abspath(doc),os.path.abspath(dir+'\\'+alias+'.docx'))
                        except:
                            filerror = True
                            msg = QMessageBox()
                            msg.setIcon(QMessageBox.Critical)
                            msg.setText("File Conversion Failed, try Restart Applicaton!")
                            msg.setInformativeText('File Conversion Error')
                            msg.setWindowTitle("Error")
                            msg.exec_()
                    elif '.xls' in doc or '.xlsx' in doc:
                        filetype = 'xlsx'
                        iseditable = True
                        try:
                            excel = win32com.client.DispatchEx("Excel.Application")
                            wb = excel.Workbooks.Open(os.path.abspath(doc))
                            count = wb.Sheets.Count
                            ws_index_list = []
                            for i in range(1,count + 1):
                                ws_index_list.append(i)
                            wb.WorkSheets(ws_index_list).Select()
                            #Save
                            wb.ActiveSheet.ExportAsFixedFormat(0,os.path.abspath(dir + '\\' + alias + '.pdf'))
                            wb.Close()
                            excel.Quit()
                            self.upload_progress_value += 90
                            self.upload_progress_bar.setValue(self.upload_progress_value)
                            shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.xlsx'))
                        except:
                            filerror = True
                            msg = QMessageBox()
                            msg.setIcon(QMessageBox.Critical)
                            msg.setText("File Conversion Failed, try Restart Applicaton!")
                            msg.setInformativeText('File Conversion Error')
                            msg.setWindowTitle("Error")
                            msg.exec_()
                    elif '.pdf' in doc:
                        filetype = 'pdf'
                        iseditable = False
                        shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.pdf'))
                        self.upload_progress_value += 90
                        self.upload_progress_bar.setValue(self.upload_progress_value)

                    if filerror == False:
                        engine = sqc.Database().engine
                        conn = engine.connect()
                        archive_document = sqc.Database().archive_document
                        s = archive_document.select().where(archive_document.c.alias == alias)
                        s_value = conn.execute(s)
                        x = 0
                        for val in s_value:
                            x+=1
                        if x < 1:
                            if self.upload_mark_confidential.checkState() == QtCore.Qt.Checked:
                                isconfidential=True
                            else:
                                isconfidential=False
                            ins = archive_document.insert().values(
                                docname = self.upload_document_name.text(),
                                isconfidential = isconfidential,
                                doctype = self.upload_doctype.currentText(),
                                description = self.upload_description.toPlainText(),
                                alias = alias,
                                iseditable = iseditable,
                                filetype = filetype,
                                date_uploaded = datetime.datetime.utcnow()
                            )
                            conn.execute(ins)
                            self.upload_progress_bar.setValue(100)
                        else:
                            msg = QMessageBox()
                            msg.setIcon(QMessageBox.Critical)
                            msg.setText("Document Name Error!")
                            msg.setInformativeText('Document Name Exists')
                            msg.setWindowTitle("Error")
                            msg.exec_()

            else:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Error")
                msg.setInformativeText('Document Name Empty')
                msg.setWindowTitle("Error")
                msg.exec_()
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Uploads Yet')
            msg.setWindowTitle("Error")
            msg.exec_()

    ##
    # ARCHIVE
    archive_dictionary = {}
    def dock_archive_action(self):
        global useraccount
        self.archive_dictionary = {}
        self.archive_doclist.clear()
        self.archive_edit_button.setVisible(False)
        self.tabWidget.setCurrentIndex(6)
        progress_value = 0
        engine = sqc.Database().engine
        conn = engine.connect()
        archive_document = sqc.Database().archive_document
        if useraccount == 'admin':
            s = archive_document.select()
        elif useraccount == 'user':
            s = archive_document.select().where(archive_document.c.isconfidential == False)
        s_value = conn.execute(s)
        x_value = conn.execute(s)
        d_value = conn.execute(s)
        dir = self.settings_sharedrive_loc.text()
        x = 0
        for val in x_value:
            x+= 1

        for val in s_value:
            progress_value += 100 / x
            self.load_archive_progressbar.setValue(progress_value)
            alias = val[5]
            index = val[0]
            #check if alias.pdf exists in the sharedrive. if doc does not exists. delete the data from the database
            if os.path.exists(os.path.abspath(dir+'\\'+alias+'.pdf')):
                print(alias, '- exists')
                #do some cryptography here#
                #-------------------------#
            else:
                print(alias, '-does not exists')
                d = archive_document.delete().where(archive_document.c.docid == val[0])
                conn.execute(d)

        self.load_archive_progressbar.setValue(100)
        for val in d_value:
            self.archive_dictionary.update({val[1] :
                {'id' : val[0],
                'isconfidential' : val[2],
                'doctype' : val[3],
                'description' : val[4],
                'alias' : val[5],
                'iseditable':val[6],
                'filetype':val[7],
                'date_uploaded':val[8]}})

        for key in self.archive_dictionary.keys():
            self.archive_doclist.addItem(key)

        self.archive_doclist.installEventFilter(self)

        self.tabWidget.setCurrentIndex(2)
        self.archive_search_alphabetically_action()


        # Sort nested dictionary by key
        #res = sorted(test_dict.items(), key=lambda x: x[1]['marks'])

    def archive_doclist_action(self):
        dir = self.settings_sharedrive_loc.text()
        alias = self.archive_dictionary[self.archive_doclist.currentItem().text()]['alias']
        iseditable = self.archive_dictionary[self.archive_doclist.currentItem().text()]['iseditable']
        if iseditable == True:
            self.archive_edit_button.setVisible(True)
        else:
            self.archive_edit_button.setVisible(False)
        path_to_pdf = os.path.abspath(dir+'\\'+alias+'.pdf')
        self.archive_web_engine.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, path_to_pdf)))

    def eventFilter(self, source, event):
        global docinfo_selected_item
        global docinfo_dict
        if (event.type() == QtCore.QEvent.ContextMenu and
            source is self.archive_doclist):
            menu = QtWidgets.QMenu()
            menu.addAction('Information')
            if menu.exec_(event.globalPos()):
                item = source.itemAt(event.pos())
                docinfo_selected_item = item.text()
                docinfo_dict = self.archive_dictionary
                ##############
                d = Docinfo_Dialogue(self)
                d.show()

            return True
        return super(MainApp, self).eventFilter(source, event)

    def archive_delete_button_action(self):
        dir = self.settings_sharedrive_loc.text()
        id = self.archive_dictionary[self.archive_doclist.currentItem().text()]['id']
        alias = self.archive_dictionary[self.archive_doclist.currentItem().text()]['alias']
        iseditable = self.archive_dictionary[self.archive_doclist.currentItem().text()]['iseditable']
        filetype = self.archive_dictionary[self.archive_doclist.currentItem().text()]['filetype']
        if os.path.exists(os.path.abspath(dir + '\\' + alias + '.pdf')):
            try:
                if iseditable:
                    os.remove(dir + '\\' + alias + '.pdf')
                    os.remove((dir + '\\' + alias + '.'+filetype))
                else:
                    os.remove(dir + '\\' + alias + '.pdf')
            except:
                pass
        self.archive_web_engine.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, '')))
        self.dock_archive_action()

    def archive_edit_button_action(self):
        dir = self.settings_sharedrive_loc.text()
        alias = self.archive_dictionary[self.archive_doclist.currentItem().text()]['alias']
        editable_filetype = self.archive_dictionary[self.archive_doclist.currentItem().text()]['filetype']
        os.startfile(os.path.abspath(dir+'\\'+alias+'.'+editable_filetype))

    def archive_options_action(self):
        self.archive_doclist.clear()
        self.archive_edit_button.setVisible(False)

        if archive_options.currentText() == 'All':
            for key in self.archive_dictionary.keys():
                self.archive_doclist.addItem(key)
        else:
            for key in self.archive_dictionary.keys():
                if self.archive_options.currentText().lower() in self.archive_dictionary[key]['doctype'].lower():
                    self.archive_doclist.addItem(key)
        self.archive_search_alphabetically_action()

    def archive_search_action(self):
        self.archive_doclist.clear()
        self.archive_edit_button.setVisible(False)

        if self.archive_search.text() == '':
            for key in self.archive_dictionary.keys():
                self.archive_doclist.addItem(key)
        else:
            for key in self.archive_dictionary.keys():
                if self.archive_search.text().lower() in key.lower():
                    self.archive_doclist.addItem(key)
        self.archive_search_alphabetically_action()

    def archive_search_alphabetically_action(self):
        items = []
        for index in range(self.archive_doclist.count()):
            items.append(self.archive_doclist.item(index).text())
        self.archive_doclist.clear()
        items = sorted(items,key=lambda v: v.upper())
        for item in items:
            self.archive_doclist.addItem(item)
        self.archive_web_engine.load(QtCore.QUrl.fromUserInput('%s?file=%s' % (self.PDFJS, '')))

    def archive_search_historically_action(self):
        pass

    ##MAIL
    def dock_messages_action(self):
        self.tabWidget.setCurrentIndex(3)
        self.mail_tab_widget.setCurrentIndex(0)
        self.mail_inbox_tab_widget.setCurrentIndex(0)
        self.mail_request_button.setVisible(False)
        self.mail_inbox_waiting_reject.setVisible(True)
        self.mail_inbox_waiting_approve.setVisible(True)
        self.mail_inbox_button_action()

    def dock_messages_2_action(self):
        self.tabWidget.setCurrentIndex(3)
        self.mail_tab_widget.setCurrentIndex(0)
        self.mail_inbox_tab_widget.setCurrentIndex(0)
        self.mail_request_button.setVisible(True)
        self.mail_inbox_waiting_reject.setVisible(False)
        self.mail_inbox_waiting_approve.setVisible(False)
        self.mail_inbox_button_action()


    def mail_request_button_action(self):
        self.mail_tab_widget.setCurrentIndex(3)
        self.mail_urequest_from.setText('')
        self.mail_urequest_subject.setPlainText('')
        self.mail_urequest_action_to_be_taken.setPlainText('')
        self.mail_urequest_upload_list.clear()
        self.mail_urequest_combo.clear()
        engine = sqc.Database().engine
        conn = engine.connect()
        archive_admin = sqc.Database().archive_admin
        s = archive_admin.select().where(archive_admin.c.previlage == 'admin')
        s_value = conn.execute(s)
        for val in s_value:
            self.mail_urequest_combo.addItem(val[1])

    mail_urequest_upload_list_dictionary = {}
    mail_urequest_upload_filetype = ''
    def mail_urequest_upload_images_action(self):
        self.mail_urequest_upload_filetype = 'img'
        self.mail_urequest_upload_list_dictionary = {}
        self.mail_urequest_upload_list.clear()
        self.mail_urequest_upload_list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        image_files, _ = QFileDialog.getOpenFileNames(self,"Open Images",
                                                      '',
                                                      "Image File (*.jpg *.png)",options=options)
        for image_file in image_files:
            temp = image_file.split('/')
            self.mail_urequest_upload_list_dictionary.update({temp[len(temp) - 1]:image_file})
            self.mail_urequest_upload_list.addItem(temp[len(temp) - 1])

    def mail_urequest_upload_documents_action(self):
        self.mail_urequest_upload_filetype = 'doc'
        self.mail_urequest_upload_list_dictionary = {}
        self.mail_urequest_upload_list.clear()

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        doc_name, _ = QFileDialog.getOpenFileName(self,"Open Document",
                                                      '',
                                                      "Document File (*.docx *.doc *.xls *.xlsx *.pdf)",options=options)
        temp = doc_name.split('/')
        self.mail_urequest_upload_list_dictionary.update({temp[len(temp) - 1]:doc_name})
        self.mail_urequest_upload_list.addItem(temp[len(temp) - 1])

    def mail_urequest_send_action(self):
        global useraccount_name
        have_attached = False
        filerror = False
        alias = ''
        filetype = ''
        iseditable = False

        if self.mail_urequest_upload_list.count() > 0:
            if self.mail_urequest_upload_filetype == 'img':
                have_attached = True
                filetype = 'pdf'
                iseditable = False
                dir = self.settings_sharedrive_loc.text()
                alias = str(self.mail_urequest_from.text() + '-routing-'+datetime.datetime.now().strftime("%m_%d_%Y_%H_%M_%S")).replace('-', '_').replace(' ', '_').lower()
                try:
                    os.makedirs(dir)
                except:
                    print('directory already exists.')
                pdf = FPDF()
                for i in range(self.mail_urequest_upload_list.count()):
                    imageFile = self.mail_urequest_upload_list_dictionary[self.mail_urequest_upload_list.item(i).text()]
                    cover = Image.open(imageFile)
                    width, height = cover.size
                    width, height = float(width * 0.264583), float(height * 0.264583)
                    pdf_size = {'P': {'w': 210, 'h': 297}, 'L': {'w': 297, 'h': 210}}
                    orientation = 'P' if width < height else 'L'
                    width = width if width < pdf_size[orientation]['w'] else pdf_size[orientation]['w']
                    height = height if height < pdf_size[orientation]['h'] else pdf_size[orientation]['h']
                    pdf.add_page(orientation=orientation)
                    pdf.image(imageFile, 0, 0, width, height)
                pdf.output(dir+'/'+alias+'.pdf',"F")

            elif self.mail_urequest_upload_filetype == 'doc':
                have_attached = True
                filerror = False
                dir = self.settings_sharedrive_loc.text()
                alias = str(self.mail_urequest_from.text() + '-routing-'+datetime.datetime.now().strftime("%m_%d_%Y_%H_%M_%S")).replace('-', '_').replace(' ', '_').lower()
                try:
                    os.makedirs(dir)
                except:
                    print('directory already exists.')

                doc = self.mail_urequest_upload_list_dictionary[self.mail_urequest_upload_list.item(0).text()]
                if '.doc' in doc or '.docx' in doc:
                    filetype = 'docx'
                    iseditable = True
                    try:
                        word = win32com.client.gencache.EnsureDispatch("Word.Application")
                        worddoc = word.Documents.Open(os.path.abspath(doc))
                        worddoc.SaveAs(os.path.abspath(dir+'\\'+alias+".pdf"), FileFormat=17)
                        worddoc.Close()
                        word.Quit()
                        shutil.copyfile(os.path.abspath(doc),os.path.abspath(dir+'\\'+alias+'.docx'))
                    except:
                        filerror = True
                        msg = QMessageBox()
                        msg.setIcon(QMessageBox.Critical)
                        msg.setText("File Conversion Failed, try Restart Applicaton!")
                        msg.setInformativeText('File Conversion Error')
                        msg.setWindowTitle("Error")
                        msg.exec_()
                elif '.xls' in doc or '.xlsx' in doc:
                    have_attached = True
                    filetype = 'xlsx'
                    iseditable = True
                    try:
                        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
                        wb = excel.Workbooks.Open(os.path.abspath(doc))
                        count = wb.Sheets.Count
                        ws_index_list = []
                        for i in range(1,count + 1):
                            ws_index_list.append(i)
                        wb.WorkSheets(ws_index_list).Select()
                        #Save
                        wb.ActiveSheet.ExportAsFixedFormat(0,os.path.abspath(dir + '\\' + alias + '.pdf'))
                        wb.Close()
                        excel.Quit()
                        shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.xlsx'))
                    except:
                        filerror = True
                        msg = QMessageBox()
                        msg.setIcon(QMessageBox.Critical)
                        msg.setText("File Conversion Failed, try Restart Applicaton!")
                        msg.setInformativeText('File Conversion Error')
                        msg.setWindowTitle("Error")
                        msg.exec_()
                elif '.pdf' in doc:
                    filetype = 'pdf'
                    iseditable = False
                    shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.pdf'))
        else:
            have_attached = False

        if filerror == False:
            engine = sqc.Database().engine
            conn = engine.connect()
            archive_mail = sqc.Database().archive_mail
            ins = archive_mail.insert().values(
                sender = useraccount_name,
                reciever = self.mail_urequest_combo.currentText(),
                date_sent = datetime.datetime.utcnow(),
                from_who = self.mail_urequest_from.text(),
                subject = self.mail_urequest_subject.toPlainText(),
                action = self.mail_urequest_action_to_be_taken.toPlainText(),
                have_attached = have_attached,
                attached_alias = alias,
                isseen = False,
                iseditable = iseditable,
                filetype = filetype,
                status = 'wait',
                status_message='',
                reply_have_attached=False,
                reply_attached_alias=False
            )
            conn.execute(ins)
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Message Sent ot {}!".format(self.mail_urequest_combo.currentText()))
            msg.setWindowTitle("Info")
            msg.exec_()

            self.mail_request_button_action()
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Message Not Sent!")
            msg.setInformativeText('Something went Wrong')
            msg.setWindowTitle("Error")
            msg.exec_()

    def mail_inbox_text_browser_source(self, date_recieved_value, from_value, subject_value, action_to_be_taken_value):
        tempstr = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">' \
                  '<html><head><meta name="qrichtext" content="1" /><style type="text/css">' \
                  'p, li { white-space: pre-wrap; }' \
                  '</style></head><body style=" font-family:\'Segoe UI\'; font-size:8.25pt; font-weight:400; font-style:normal;">' \
                  '<p style=" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" font-size:10pt; font-weight:600;">' \
                  'Date Recieved</span><span style=" font-size:10pt;"> : ' \
                  + str(date_recieved_value) + '' \
                                          '<br /></span><span style=" font-size:10pt; font-weight:600;">From</span><span style=" font-size:10pt;"> : ' \
                  + from_value + '<br /><br /></span><span style=" font-size:10pt; font-weight:600;">' \
                                 'Subject / Specific Concern</span><span style=" font-size:10pt;"> :</span></p>' \
                                 '<p style=" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" font-size:10pt;">' \
                  + subject_value + '</span></p>' \
                                    '<p style=" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" font-size:10pt; font-weight:600;">' \
                                    'Action to be taken</span><span style=" font-size:10pt;"> :</span></p><p style=" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" font-size:10pt;">' \
                  + action_to_be_taken_value + '</span></p><p style="-qt-paragraph-type:empty; margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:10pt;"><br /></p></body></html>'

        return tempstr

    mail_inbox_dictionary = {}
    def mail_inbox_button_action(self):
        self.mail_inbox_approved_comment.setPlainText('')
        self.mail_inbox_approved_button_container.setVisible(False)
        self.mail_inbox_waiting_view.setVisible(False)
        self.mail_inbox_waiting_attached_name.setText('')
        self.mail_tab_widget.setCurrentIndex(0)
        self.mail_inbox_tab_widget.setCurrentIndex(0)
        self.mail_inbox_waiting_approve.setVisible(False)
        self.mail_inbox_waiting_reject.setVisible(False)
        self.mail_inbox_text_browser.setText('')
        self.mail_approved_text_browser.setText('')
        self.mail_rejected_text_browser.setText('')
        self.mail_inbox_waiting_list.clear()
        self.mail_inbox_approved_list.clear()
        self.mail_inbox_rejected_list.clear()
        engine = sqc.Database().engine
        conn = engine.connect()
        archive_mail = sqc.Database().archive_mail
        s = archive_mail.select()
        s_value = conn.execute(s)

        for val in s_value:
            self.mail_inbox_dictionary.update({
                str(val[0]):{
                    'sender':val[1],
                    'reciever':val[2],
                    'date_sent':val[3],
                    'from_who':val[4],
                    'subject':val[5],
                    'action':val[6],
                    'have_attached':val[7],
                    'attached_alias':val[8],
                    'isseen':val[9],
                    'iseditable':val[10],
                    'filetype':val[11],
                    'status':val[12],
                    'status_message':val[13],
                    'reply_have_attached':val[14],
                    'reply_attached_alias':val[15],
                    'reply_is_editable':val[16],
                    'reply_filetype':val[17]

                }
            })
        self.mail_inbox_dictionary = OrderedDict(sorted(self.mail_inbox_dictionary.items(), reverse=True))

        for key in self.mail_inbox_dictionary.keys():
            if self.mail_inbox_dictionary[key]['status'] == 'wait':
                self.mail_inbox_waiting_list.addItem('{}-(MSGID{})-[{}]'.format(self.mail_inbox_dictionary[key]['sender'],key,self.mail_inbox_dictionary[key]['date_sent']))
            elif self.mail_inbox_dictionary[key]['status'] == 'accepted':
                self.mail_inbox_approved_list.addItem('{}-(MSGID{})-[{}]'.format(self.mail_inbox_dictionary[key]['sender'], key,self.mail_inbox_dictionary[key]['date_sent']))
            elif self.mail_inbox_dictionary[key]['status'] == 'rejected':
                self.mail_inbox_rejected_list.addItem('{}-(MSGID{})-[{}]'.format(self.mail_inbox_dictionary[key]['sender'], key,self.mail_inbox_dictionary[key]['date_sent']))

    def mail_inbox_waiting_list_action(self):
        global useraccount
        self.mail_inbox_waiting_view.setVisible(True)
        temp = self.mail_inbox_waiting_list.currentItem().text()
        temp = temp.split('(')[1]
        temp = temp.split(')')[0]
        id = temp.replace('MSGID','')
        self.mail_inbox_text_browser.setText(self.mail_inbox_text_browser_source(self.mail_inbox_dictionary[id]['date_sent'],self.mail_inbox_dictionary[id]['from_who'],self.mail_inbox_dictionary[id]['subject'],self.mail_inbox_dictionary[id]['action']))
        self.mail_inbox_waiting_attached_name.setText(str(self.mail_inbox_dictionary[id]['attached_alias']))

        if self.mail_inbox_dictionary[id]['have_attached'] == False:
            self.mail_inbox_waiting_view.setVisible(False)
        else:
            self.mail_inbox_waiting_view.setVisible(True)

        if useraccount == 'admin':
            self.mail_inbox_waiting_approve.setVisible(True)
            self.mail_inbox_waiting_reject.setVisible(True)

    def mail_inbox_waiting_view_action(self):
        temp = self.mail_inbox_waiting_list.currentItem().text()
        temp = temp.split('(')[1]
        temp = temp.split(')')[0]
        id = temp.replace('MSGID', '')
        dir = self.settings_sharedrive_loc.text()
        alias = self.mail_inbox_dictionary[id]['attached_alias']
        filetype =self.mail_inbox_dictionary[id]['filetype']
        os.startfile(os.path.abspath(dir+'\\'+alias+'.'+filetype))

    def mail_inbox_waiting_approve_action(self):

        name = ''
        temp = self.mail_inbox_waiting_list.currentItem().text()
        temp = temp.split('(')[1]
        temp = temp.split(')')[0]
        id = temp.replace('MSGID', '')
        engine = sqc.Database().engine
        conn = engine.connect()
        archive_mail = sqc.Database().archive_mail
        s = archive_mail.select().where(archive_mail.c.mailid == id)
        s_value = conn.execute(s)
        for val in s_value:
            name = val[1]
            from_who = val[4]
        try:
            ad = Routing_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id,name,self.settings_sharedrive_loc.text(),from_who,'accepted')
        except:
            pass

    def mail_inbox_waiting_reject_action(self):

        name = ''
        temp = self.mail_inbox_waiting_list.currentItem().text()
        temp = temp.split('(')[1]
        temp = temp.split(')')[0]
        id = temp.replace('MSGID', '')
        engine = sqc.Database().engine
        conn = engine.connect()
        archive_mail = sqc.Database().archive_mail
        s = archive_mail.select().where(archive_mail.c.mailid == id)
        s_value = conn.execute(s)
        for val in s_value:
            name = val[1]
            from_who = val[4]
        try:
            ad = Routing_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id,name,self.settings_sharedrive_loc.text(),from_who,'rejected')
        except:
            pass

    def mail_inbox_approved_list_action(self):
        global useraccount
        self.mail_inbox_approved_button_container.setVisible(True)
        temp = self.mail_inbox_approved_list.currentItem().text()
        temp = temp.split('(')[1]
        temp = temp.split(')')[0]
        id = temp.replace('MSGID','')
        self.mail_approved_text_browser.setText(self.mail_inbox_text_browser_source(self.mail_inbox_dictionary[id]['date_sent'],self.mail_inbox_dictionary[id]['from_who'],self.mail_inbox_dictionary[id]['subject'],self.mail_inbox_dictionary[id]['action']))
        self.mail_inbox_approved_comment.setPlainText(self.mail_inbox_dictionary[id]['status_message'])

        if self.mail_inbox_dictionary[id]['reply_have_attached'] == False:
            self.mail_inbox_approved_view.setVisible(False)
            #self.mail_inbox_approved_save.setVisible(False)
            if useraccount != 'admin':
                self.mail_inbox_approved_delete.setVisible(False)
            else:
                self.mail_inbox_approved_delete.setVisible(True)
        else:
            self.mail_inbox_approved_view.setVisible(True)
            #self.mail_inbox_approved_save.setVisible(True)
            if useraccount != 'admin':
                self.mail_inbox_approved_delete.setVisible(False)
            else:
                self.mail_inbox_approved_delete.setVisible(True)

    def mail_inbox_approved_view_action(self):
        temp = self.mail_inbox_approved_list.currentItem().text()
        temp = temp.split('(')[1]
        temp = temp.split(')')[0]
        id = temp.replace('MSGID', '')
        dir = self.settings_sharedrive_loc.text()
        alias = self.mail_inbox_dictionary[id]['reply_attached_alias']
        filetype =self.mail_inbox_dictionary[id]['reply_filetype']
        os.startfile(os.path.abspath(dir+'\\'+alias+'.'+filetype))

    ##COMPOSE
    def mail_compose_refresh(self):
        self.mail_compose_subject.setText('')
        self.mail_compose_content.setPlainText('')
        self.mail_compose_upload_list.clear()

    def mail_compose_button_action(self):
        global useraccount
        self.mail_tab_widget.setCurrentIndex(1)
        engine = sqc.Database().engine
        conn = engine.connect()
        archive_admin = sqc.Database().archive_admin
        if useraccount == 'user':
            s = archive_admin.select().where(archive_admin.c.previlage == 'admin')
        elif useraccount == 'admin':
            s = archive_admin.select().where(archive_admin.c.previlage == 'user')
        s_value = conn.execute(s)
        for val in s_value:
            self.mail_compose_combo.addItem(val[1])
        self.mail_compose_refresh()

    mail_compose_upload_list_dictionary = {}
    mail_compose_upload_filetype = ''

    def mail_compose_upload_images_action(self):
        self.mail_compose_upload_filetype = 'img'
        self.mail_compose_upload_list_dictionary = {}
        self.mail_compose_upload_list.clear()
        self.mail_compose_upload_list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        image_files, _ = QFileDialog.getOpenFileNames(self, "Open Images",
                                                      '',
                                                      "Image File (*.jpg *.png)", options=options)
        for image_file in image_files:
            temp = image_file.split('/')
            self.mail_compose_upload_list_dictionary.update({temp[len(temp) - 1]: image_file})
            self.mail_compose_upload_list.addItem(temp[len(temp) - 1])

    def mail_compose_upload_documents_action(self):
        self.mail_compose_upload_filetype = 'doc'
        self.mail_compose_upload_list_dictionary = {}
        self.mail_compose_upload_list.clear()

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        doc_name, _ = QFileDialog.getOpenFileName(self, "Open Document",
                                                  '',
                                                  "Document File (*.docx *.doc *.xls *.xlsx *.pdf)", options=options)
        temp = doc_name.split('/')
        self.mail_compose_upload_list_dictionary.update({temp[len(temp) - 1]: doc_name})
        self.mail_compose_upload_list.addItem(temp[len(temp) - 1])

    def mail_compose_send_action(self):
        global useraccount_name
        have_attached = False
        filerror = False
        alias = ''
        filetype = ''
        iseditable = False

        if self.mail_compose_upload_list.count() > 0:
            if self.mail_compose_upload_filetype == 'img':
                have_attached = True
                filetype = 'pdf'
                iseditable = False
                dir = self.settings_sharedrive_loc.text()
                alias = str('message-' + datetime.datetime.now().strftime(
                    "%m_%d_%Y_%H_%M_%S")).replace('-', '_').replace(' ', '_').lower()
                try:
                    os.makedirs(dir)
                except:
                    print('directory already exists.')
                pdf = FPDF()
                for i in range(self.mail_urequest_upload_list.count()):
                    imageFile = self.mail_urequest_upload_list_dictionary[self.mail_urequest_upload_list.item(i).text()]
                    cover = Image.open(imageFile)
                    width, height = cover.size
                    width, height = float(width * 0.264583), float(height * 0.264583)
                    pdf_size = {'P': {'w': 210, 'h': 297}, 'L': {'w': 297, 'h': 210}}
                    orientation = 'P' if width < height else 'L'
                    width = width if width < pdf_size[orientation]['w'] else pdf_size[orientation]['w']
                    height = height if height < pdf_size[orientation]['h'] else pdf_size[orientation]['h']
                    pdf.add_page(orientation=orientation)
                    pdf.image(imageFile, 0, 0, width, height)
                pdf.output(dir + '/' + alias + '.pdf', "F")

            elif self.mail_urequest_upload_filetype == 'doc':
                have_attached = True
                filerror = False
                dir = self.settings_sharedrive_loc.text()
                alias = str(self.mail_urequest_from.text() + '-routing-' + datetime.datetime.now().strftime(
                    "%m_%d_%Y_%H_%M_%S")).replace('-', '_').replace(' ', '_').lower()
                try:
                    os.makedirs(dir)
                except:
                    print('directory already exists.')

                doc = self.mail_urequest_upload_list_dictionary[self.mail_urequest_upload_list.item(0).text()]
                if '.doc' in doc or '.docx' in doc:
                    filetype = 'docx'
                    iseditable = True
                    try:
                        word = win32com.client.gencache.EnsureDispatch("Word.Application")
                        worddoc = word.Documents.Open(os.path.abspath(doc))
                        worddoc.SaveAs(os.path.abspath(dir + '\\' + alias + ".pdf"), FileFormat=17)
                        worddoc.Close()
                        word.Quit()
                        shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.docx'))
                    except:
                        filerror = True
                        msg = QMessageBox()
                        msg.setIcon(QMessageBox.Critical)
                        msg.setText("File Conversion Failed, try Restart Applicaton!")
                        msg.setInformativeText('File Conversion Error')
                        msg.setWindowTitle("Error")
                        msg.exec_()
                elif '.xls' in doc or '.xlsx' in doc:
                    have_attached = True
                    filetype = 'xlsx'
                    iseditable = True
                    try:
                        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
                        wb = excel.Workbooks.Open(os.path.abspath(doc))
                        count = wb.Sheets.Count
                        ws_index_list = []
                        for i in range(1, count + 1):
                            ws_index_list.append(i)
                        wb.WorkSheets(ws_index_list).Select()
                        # Save
                        wb.ActiveSheet.ExportAsFixedFormat(0, os.path.abspath(dir + '\\' + alias + '.pdf'))
                        wb.Close()
                        excel.Quit()
                        shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.xlsx'))
                    except:
                        filerror = True
                        msg = QMessageBox()
                        msg.setIcon(QMessageBox.Critical)
                        msg.setText("File Conversion Failed, try Restart Applicaton!")
                        msg.setInformativeText('File Conversion Error')
                        msg.setWindowTitle("Error")
                        msg.exec_()
                elif '.pdf' in doc:
                    filetype = 'pdf'
                    iseditable = False
                    shutil.copyfile(os.path.abspath(doc), os.path.abspath(dir + '\\' + alias + '.pdf'))
        else:
            have_attached = False

        if filerror == False:
            engine = sqc.Database().engine
            conn = engine.connect()
            archive_mail = sqc.Database().archive_mail
            ins = archive_mail.insert().values(
                sender=useraccount_name,
                reciever=self.mail_urequest_combo.currentText(),
                date_sent=datetime.datetime.utcnow(),
                from_who=self.mail_urequest_from.text(),
                subject=self.mail_urequest_subject.toPlainText(),
                action=self.mail_urequest_action_to_be_taken.toPlainText(),
                have_attached=have_attached,
                attached_alias=alias,
                isseen=False,
                iseditable=iseditable,
                filetype=filetype,
                status='wait',
                status_message='',
                reply_have_attached=False,
                reply_attached_alias=False
            )
            conn.execute(ins)
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Message Sent ot {}!".format(self.mail_urequest_combo.currentText()))
            msg.setWindowTitle("Info")
            msg.exec_()

            self.mail_request_button_action()
        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Message Not Sent!")
            msg.setInformativeText('Something went Wrong')
            msg.setWindowTitle("Error")
            msg.exec_()







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