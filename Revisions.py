import wx  # wxWidgets used as the GUI
from wx import xrc  # allows the loading and access of xrc file (xml) that describes GUI
from wx.html import HtmlEasyPrinting
import wx.grid as gridlib
from wxPython.calendar import *

ctrl = xrc.XRCCTRL  # define a shortined function name (just for convienience)

# for sending emails
import smtplib
from email import Encoders
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.MIMEMultipart import MIMEMultipart
from email.Utils import formatdate

import os
import glob
import sys
import time
import datetime as dt
from threading import Thread

import Database
import General
import OrderFileOpeners
import Ecrs


def on_click_add_new_revision(event):
    # check that entered related ecr id is valid
    related_ecr = ctrl(General.app.main_frame, 'text:related_ecr').GetValue()
    if related_ecr != '':
        cursor = Database.connection.cursor()
        if not cursor.execute('SELECT TOP 1 id FROM ecrs WHERE id = \'{}\''.format(related_ecr)).fetchone():
            wx.MessageBox(
                'The related ECR ID: {} was not found in the database.\n\nCorrect the ECR ID or clear that field to revise\nwithout relating the revision to an ECR.'.format(
                    related_ecr), 'Error', wx.OK | wx.ICON_WARNING)
            return

    notebook = ctrl(General.app.main_frame, 'notebook:revisions')
    item_entries = []
    for index in range(notebook.GetPageCount()):
        item_entries.append(notebook.GetPageText(index).strip().split(' ')[0])

    General.app.init_submit_revision_dialog(item_entries)


def on_click_document_tree(event):
    # if a valid document is selected, as in not a catagory heading, close
    # the dialog and 'return' what was selected (in a format the matches the database)
    pt = event.GetPosition()
    tree = event.GetEventObject()
    item, flags = tree.HitTest(pt)
    '''
    if item:
        if tree.GetItemText(item)[-1:] != ' ':
            tree.Expand(item)
        else:
            if tree.GetItemText(item)[-2:] != '  ':
                General.app.selected_revision_document = '{}>{}'.format(tree.GetItemText(tree.GetItemParent(item)).strip(), tree.GetItemText(item).strip())
            else:
                General.app.selected_revision_document = '{}'.format(tree.GetItemText(item).strip())

            General.app.revision_document_selection_dialog.Destroy()
    '''

    if item:
        if tree.GetItemText(item)[-1:] == ' ':
            document = ''

            # climb back up the tree and build the document's "full name"
            try:
                parent_item = item
                while 1:
                    document = '{}>{}'.format(tree.GetItemText(parent_item).strip(), document)
                    parent_item = tree.GetItemParent(parent_item)
            except:
                pass

            document = document[:-1]
            General.app.selected_revision_document = document.replace('>', ' > ')
            General.app.revision_document_selection_dialog.Destroy()

        else:
            tree.Expand(item)

    event.Skip()


'''
def on_click_select_revision_document(event):
    button = event.GetEventObject()
    new_label = General.app.init_revision_document_selection_dialog()
    if new_label == None: new_label = 'Select'
    
    if new_label == 'BOM':
        ctrl(General.app.new_revision_dialog, 'label:rec_dollars').Enable()
        ctrl(General.app.new_revision_dialog, 'text:rec_dollars').Enable()
    else:
        ctrl(General.app.new_revision_dialog, 'label:rec_dollars').Disable()
        ctrl(General.app.new_revision_dialog, 'text:rec_dollars').Disable()
        ctrl(General.app.new_revision_dialog, 'text:rec_dollars').SetValue('')
    
    if new_label == 'Select':
        ctrl(General.app.new_revision_dialog, 'button:submit_revision').Disable()
    else:
        ctrl(General.app.new_revision_dialog, 'button:submit_revision').Enable()
    
    button.SetLabel(str(new_label))
    button.GetParent().GetSizer().Layout()
'''


def on_click_print_revisions(event):
    notebook = ctrl(General.app.main_frame, 'notebook:revisions')
    item_entries = []
    for index in range(notebook.GetPageCount()):
        item_entries.append(notebook.GetPageText(index).strip().split(' ')[0])

    cursor = Database.connection.cursor()

    for item in item_entries:
        column_names = Database.get_table_column_names('revisions', presentable=True)

        revisions = cursor.execute('SELECT * FROM revisions WHERE item = \'{}\''.format(item)).fetchall()
        if revisions == None:
            continue

        # Revisions for Item Number: {}
        # <hr>

        html_to_print = '''<style type=\"text/css\">td{{font-family:Arial; color:black; font-size:8pt;}}</style>
            <table border="1" cellspacing="0"><tr>
            '''.format(item)

        blacklisted_columns = ['dollars_reconciled', 'related_ecr']

        for column_name in column_names:
            if column_name.replace(' ', '_').lower() not in blacklisted_columns:
                html_to_print += '<th align=\"right\" valign=\"top\">{}</th>'.format(column_name.replace(' ', '&nbsp;'))

        html_to_print += '</tr>'

        for revision in revisions:
            html_to_print += '<tr>'

            for index, column_value in enumerate(revision):
                if column_names[index].replace(' ', '_').lower() not in blacklisted_columns:
                    if column_names[index] == 'When Revised':
                        column_value = General.format_date_nicely(column_value)

                    if column_names[index] != 'Description':
                        column_value = str(column_value).replace(' ', '&nbsp;')

                    html_to_print += '<td align=\"left\" valign=\"top\">{}</td>'.format(column_value)

            html_to_print += '</tr>'

        html_to_print += '</table>'

        printer = HtmlEasyPrinting()
        printer.GetPrintData().SetPaperId(wx.PAPER_LETTER)
        printer.GetPrintData().SetOrientation(wx.LANDSCAPE)
        printer.SetStandardFonts(9)
        printer.GetPageSetupData().SetMarginTopLeft((0, 0))
        printer.GetPageSetupData().SetMarginBottomRight((0, 0))
        printer.PrintText(html_to_print)


def on_click_submit_revision(event):
    # temp = dt.datetime.now()

    notebook = ctrl(General.app.main_frame, 'notebook:revisions')

    item_entries = []
    for index in range(notebook.GetPageCount()):
        item_entries.append(notebook.GetPageText(index).strip().split(' ')[0])

    cursor = Database.connection.cursor()

    table = ctrl(General.app.new_revision_dialog, 'panel:table').GetChildren()[0]

    new_revision_records = []

    for row in range(table.GetNumberRows()):
        document = table.GetCellValue(row, 0).strip().replace(' > ', '>')
        description = table.GetCellValue(row, 1).strip().replace("'", "''").replace('\"', "''''")
        if document == '(click to select document)':
            continue

        if document != '' and description == '':
            wx.MessageBox('You must enter a change description for document: {}'.format(document), 'Hint',
                          wx.OK | wx.ICON_WARNING)
            return

        # print 'row: {}, value: {}'.format(row, table.GetCellValue(row, 0))
        if document != '' and description != '':
            new_revision_records.append((document, description))

    if ctrl(General.app.new_revision_dialog, 'choice:revision_reasons').GetStringSelection() == '':
        wx.MessageBox('Select a reason for these revisions first.', 'Hint', wx.OK | wx.ICON_WARNING)
        return

    when_revised = str(dt.datetime.today())[:19]
    dollars_reconciled = None

    # document = ctrl(General.app.new_revision_dialog, 'button:select_document').GetLabel().replace(' > ', '>')
    show_revision_message = False
    revision_message = 'Notifications are about to be sent out for your revised documents.\n\n'
    for new_revision_record in new_revision_records:
        document_reminder = cursor.execute(
            'SELECT TOP 1 submission_reminder FROM revision_document_choices WHERE document = \'{}\''.format(
                new_revision_record[0])).fetchone()
        if document_reminder == None:
            continue
        else:
            document_reminder = document_reminder[0]

        if document_reminder != None:
            revision_message += '* ' + document_reminder + ' now if you haven\'t already.\n'
            show_revision_message = True

    if show_revision_message == True:
        wx.MessageBox(revision_message, 'Reminder', wx.OK | wx.ICON_WARNING)

    revision_email_list = cursor.execute('SELECT email FROM employees WHERE gets_revision_notice = 1').fetchall()
    popsheet_email_list = cursor.execute('SELECT email FROM employees WHERE gets_pop_sheet_notice = 1').fetchall()
    sender = \
    cursor.execute('SELECT TOP 1 email FROM employees WHERE name = \'{}\''.format(General.app.current_user)).fetchone()[
        0]

    related_ecr = ctrl(General.app.main_frame, 'text:related_ecr').GetValue()
    if related_ecr == '': related_ecr = 0

    for item_entry in item_entries:
        for new_revision_record in new_revision_records:
            level = cursor.execute(
                'SELECT TOP 1 level FROM revisions WHERE item = \'{}\' AND document = \'{}\' ORDER BY when_revised DESC'.format(
                    item_entry, new_revision_record[0])).fetchone()
            if level == None:
                level = 1
            else:
                level = level[0] + 1

            dollars_reconciled = 0
            if new_revision_record[0] == 'BOM':
                if ctrl(General.app.new_revision_dialog, 'text:rec_dollars').GetValue() != '':
                    dollars_reconciled = ctrl(General.app.new_revision_dialog, 'text:rec_dollars').GetValue()

            sql = 'INSERT INTO revisions (item, document, level, description, reason, who_revised, when_revised, dollars_reconciled, Production_Plant, related_ecr) VALUES ('
            sql += '\'{}\', '.format(item_entry)
            sql += '\'{}\', '.format(new_revision_record[0])
            sql += '{}, '.format(level)
            sql += '\'{}\', '.format(new_revision_record[1])
            sql += '\'{}\', '.format(
                ctrl(General.app.new_revision_dialog, 'choice:revision_reasons').GetStringSelection())
            sql += '\'{}\', '.format(General.app.current_user)
            sql += '\'{}\', '.format(when_revised)
            sql += '{}, '.format(dollars_reconciled)
            sql += '\'{}\', '.format(Ecrs.Prod_Plant)
            sql += '\'{}\')'.format(related_ecr)
            

            # print sql
            cursor.execute(sql)
            Database.connection.commit()

            # set rec dollars to null if it's None... cause i like perfection
            latest_id = cursor.execute('SELECT MAX(id) FROM revisions').fetchone()[0]
            cursor.execute(
                'UPDATE revisions SET dollars_reconciled = NULL WHERE id = \'{}\' AND dollars_reconciled = 0'.format(
                    latest_id))
            cursor.execute(
                'UPDATE revisions SET related_ecr = NULL WHERE id = \'{}\' AND related_ecr = 0'.format(latest_id))

            revision = cursor.execute('SELECT TOP 1 * FROM revisions WHERE id = \'{}\''.format(latest_id)).fetchone()
            order = cursor.execute(
                'SELECT TOP 1 * FROM {} WHERE item = \'{}\''.format(Ecrs.table_used, revision[1])).fetchone()

            Database.connection.commit()

            # takes a little longer to send an email so put it in a serperate thread so user
            # doesn't have to wait around :)
            Thread(target=send_revision_email, args=(revision, order, revision_email_list, sender)).start()

            # send out popsheet notices especially!
            if revision[2] == 'Pop Sheet':
                Thread(target=send_popsheet_email, args=(revision, order, popsheet_email_list, sender)).start()

            # check if BOM was uploaded if revision done on it
            # if revision[2] == 'BOM':
            #	Thread(target=check_for_BOM_net_change, args=(item_entry, )).start()

    field = ctrl(General.app.main_frame, 'text:revision_items').GetValue().replace('0', '*')
    ctrl(General.app.main_frame, 'text:revision_items').SetValue(field)
    field = ctrl(General.app.main_frame, 'text:revision_items').GetValue().replace('*', '0')
    ctrl(General.app.main_frame, 'text:revision_items').SetValue(field)

    General.app.new_revision_dialog.Destroy()


# print 'time took to submit rev: {}'.format((dt.datetime.now() - temp))



def on_click_table_cell(event):
    table = event.GetEventObject()

    if event.GetCol() == 0:
        document = General.app.init_revision_document_selection_dialog()

        # check if this doc has already been entered into the table... don't want that
        for row in range(table.GetNumberRows()):
            if document == table.GetCellValue(row, 0).strip():
                wx.MessageBox('You already entered a revision for document: {}'.format(document), 'Hint',
                              wx.OK | wx.ICON_WARNING)
                return

        ###if document == 'BOM':
        #	ctrl(General.app.new_revision_dialog, 'label:rec_dollars').Enable()
        #	ctrl(General.app.new_revision_dialog, 'text:rec_dollars').Enable()

        if document != None:
            table.SetCellValue(event.GetRow(), event.GetCol(), document)
        else:
            table.SetCellValue(event.GetRow(), event.GetCol(), '')
            table.SetCellValue(event.GetRow(), event.GetCol() + 1, '')

        ctrl(General.app.new_revision_dialog, 'button:submit_revision').Enable()

    # event.GetRow()
    event.Skip()


def on_entry_revision_items(event):
    field = ctrl(General.app.main_frame, 'text:revision_items').GetValue() + ','
    # make it so entries can be space or comma delimited
    field = field.replace(' ', ', ').upper()
    entries = field.split(',')
    valid_entries = []

    cursor = Database.connection.cursor()

    for entry in entries:
        entry = entry.strip()
        if entry != '':
            # see if entry is in the order database
            order = cursor.execute(
                "SELECT TOP 1 sales_order FROM {} WHERE item LIKE '%{}%'".format(Ecrs.table_used, entry)).fetchone()
            if order != None:
                if len(entry) == 9 or \
                        (len(entry) == 7 and entry[:2] == '01') or \
                        (len(entry) == 8 and entry[:2] != 'KW') or \
                                len(entry) == 18:
                    entry = cursor.execute(
                        "SELECT TOP 1 item FROM {} WHERE item LIKE '%{}%'".format(Ecrs.table_used, entry)).fetchone()[0]
                    valid_entries.append('{} ({})'.format(entry, order[0]))

    if len(valid_entries) == 0:
        ctrl(General.app.main_frame, 'button:add_new_revision').Disable()
    else:
        ctrl(General.app.main_frame, 'button:add_new_revision').Enable()

    notebook = ctrl(General.app.main_frame, 'notebook:revisions')
    page_indexes_to_remove = []

    for index in range(notebook.GetPageCount()):
        keep_page = False
        for entry in valid_entries:
            if entry.split(' ')[0] == notebook.GetPageText(index).strip().split(' ')[0]:
                keep_page = True

        if keep_page == False:
            page_indexes_to_remove.append(index)

    # remove any tabs that the text field doesn't ask for
    for index in reversed(page_indexes_to_remove):
        # print notebook.GetPageText(index)
        wx.FindWindowByName(notebook.GetPageText(index).strip()).Destroy()

        notebook.RemovePage(index)
        ctrl(General.app.main_frame, 'text:related_ecr').SetValue('')
    # print "might be a memory leak here cause im not deleteing the children of the page im removing"

    pages_to_add = []

    for entry in valid_entries:
        page_already_present = False
        for index in range(notebook.GetPageCount()):
            if entry.split(' ')[0] == notebook.GetPageText(index).strip().split(' ')[0]:
                page_already_present = True

        if page_already_present == False:
            pages_to_add.append(entry)

    for page_label in pages_to_add:
        panel = wx.Panel(notebook, wx.ID_ANY, name=page_label)
        notebook.AddPage(panel, '  {}  '.format(page_label), select=True)

        list_control = wx.ListCtrl(panel, -1,
                                   style=wx.LC_REPORT | wx.BORDER_NONE | wx.LC_VRULES | wx.LC_HRULES | wx.LC_SINGLE_SEL)

        list_control.Bind(wx.EVT_CHAR, onCharEvent)

        column_names = Database.get_table_column_names('revisions', presentable=True)
        for index, column_name in enumerate(column_names):
            list_control.InsertColumn(index, column_name)

        # in order to color the first doc (the latest revision for that doc type), we must
        # keep track if we have seen it yet as we loop through the revision records
        documents_seen_above = []

        revisions = cursor.execute("SELECT * FROM revisions WHERE item LIKE '%{}%' ORDER BY when_revised DESC".format(
            page_label.split(' ')[0])).fetchall()
        # revisions = cursor.execute("SELECT * FROM revisions WHERE item = \'{}\' ORDER BY document ASC, level DESC".format(page_label.split(' ')[0])).fetchall()
        for revision_index, revision in enumerate(revisions):
            list_control.InsertStringItem(sys.maxint, '#')

            if revision[2] not in documents_seen_above:
                list_control.SetItemBackgroundColour(revision_index, '#EFFACF')
                documents_seen_above.append(revision[2])

            for column_index, column_value in enumerate(revision):
                if column_names[column_index] == 'Document':
                    column_value = column_value.replace('>', ' > ')

                if column_names[column_index] == 'When Revised':
                    column_value = General.format_date_nicely(column_value)

                if column_value != None:
                    list_control.SetStringItem(revision_index, column_index, str(column_value).replace('\n', ' \\ '))

        # print documents_seen_above

        for column_index in range(len(column_names)):
            list_control.SetColumnWidth(column_index, wx.LIST_AUTOSIZE_USEHEADER)

        # hide the Id and Item columns
        list_control.SetColumnWidth(0, 0)
        list_control.SetColumnWidth(1, 0)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(list_control, 1, wx.EXPAND | wx.ALL)

        panel.SetSizer(sizer)
        panel.SetAutoLayout(True)
        panel.Layout()

    ctrl(General.app.main_frame, 'text:revision_items').SetFocus()


def onCharEvent(event):
    keycode = event.GetKeyCode()
    controlDown = event.CmdDown()
    altDown = event.AltDown()
    shiftDown = event.ShiftDown()

    list = event.GetEventObject()

    if keycode == 5:
        cursor = Database.connection.cursor()
        revision_id = list.GetItem(list.GetFirstSelected(), 0).GetText()
        revision = cursor.execute("SELECT * FROM revisions WHERE id = {}".format(revision_id)).fetchall()[0]
        # print 'revision', revision
        order = cursor.execute("SELECT * FROM {} WHERE item = '{}'".format(Ecrs.table_used, revision[1])).fetchall()[0]
        # print 'order', order
        users_email = \
        cursor.execute("SELECT email FROM employees WHERE name = '{}'".format(General.app.current_user)).fetchall()[0][
            0]
        # print 'users_email', users_email
        send_revision_email(revision, order, [(users_email,), ], users_email)

        wx.MessageBox("A copy of this revision notice has been sent to {}".format(users_email), 'Info',
                      wx.OK | wx.ICON_INFORMATION)

    event.Skip()


def on_select_revision_reason(event):
    selection = event.GetEventObject().GetStringSelection()

    if selection == 'BOM Reconciliation':
        ctrl(General.app.new_revision_dialog, 'label:rec_dollars').Enable()
        ctrl(General.app.new_revision_dialog, 'text:rec_dollars').Enable()
    else:
        ctrl(General.app.new_revision_dialog, 'label:rec_dollars').Disable()
        ctrl(General.app.new_revision_dialog, 'text:rec_dollars').Disable()
        ctrl(General.app.new_revision_dialog, 'text:rec_dollars').SetValue('')


def on_selection_change_document_tree(event):
    # the tree be actin a little funky...
    # the desired behavior is to have nothing selected if a tree catagory is expanded.
    # so try to see if something in the tree is selected. if it is, unselect everything.
    # the 'if nothing selected' part prevents the event from being called a bagillion times.
    tree = event.GetEventObject()
    try:
        tree.GetItemText(tree.GetSelection())
        tree.UnselectAll()
    except:
        pass


def on_size_document_table(event):
    table = event.GetEventObject()
    table.SetColSize(1, table.GetSize()[0] - table.GetColSize(0) - wx.SystemSettings.GetMetric(wx.SYS_VSCROLL_X))
    # table.SetColSize(1, table.GetSize()[0] - table.GetColSize(0))

    # print wx.SystemSettings.GetMetric(wx.SYS_VSCROLL_X)
    # print wx.SystemSettings.GetMetric(wx.SYS_VSCROLL_ARROW_X)
    event.Skip()


def send_revision_email(revision, order, email_list, sender):
    '''
    threaded_connection = Database.connect_to_database()
    cursor = threaded_connection.cursor()

    #cursor = Database.connection.cursor()
    revision = cursor.execute('SELECT * FROM revisions WHERE id = \'{}\' LAMIT 1'.format(revision_id)).fetchone()
    order = cursor.execute('SELECT * FROM {} WHERE item = \'{}\' LAMIT 1'.format(Ecrs.table_used, revision[1])).fetchone()
    email_list = cursor.execute('SELECT email FROM employees WHERE gets_revision_notice = \'yes\''.format(revision[1])).fetchall()
    sender = cursor.execute('SELECT email FROM employees WHERE name = \'{}\' LAMIT 1'.format(General.app.current_user)).fetchone()[0]
    '''

    server = smtplib.SMTP('mailrelay.lennoxintl.com')
    email_string = ''
    for email in email_list:
        email_string += '; {}'.format(email[0])
    email_string = email_string[2:]

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = email_string
    msg["Subject"] = 'Revision Notice: {} ({}-{}) {}'.format(order[0], order[1], order[2],
                                                             revision[2].replace('>', ' > '))
    # msg["Subject"] = 'Revision Notice: {} ({}-{}) {}'.format(order[0], order[1], order[2], revision[2].replace('>', ' > '))
    msg['Date'] = formatdate(localtime=True)

    shortcuts = ''
    try:
        order_directory = OrderFileOpeners.get_order_directory(order[1])
    except Exception as e:
        print e
        order_directory = None

    if order_directory:
        shortcuts = '<a href=\"file:///{}\">Open Order Folder</a>'.format(order_directory)
    # add a 'open revised document' also
    ###<a href=\"file:///R:\\Design_Eng\\Orders\\Orders_2012\\W-X-Y-Z\\Walmart\\212031\\212031-BOM_CA.xls">Open Revised Document</a>

    revision_id = revision[0]
    related_ecr = revision[9]
    # if related_ecr == None: related_ecr = ''
    document = revision[2].replace('>', ' > ')
    level = revision[3]
    reason = revision[5]
    description = revision[4]

    item_number = order[0]
    sales_order = '{}-{}'.format(order[1], order[2])
    customer = order[5]
    location = '{}, {}'.format(order[8], order[9])
    model = order[11]

    target_date = order[23]
    if target_date:
        target_date = target_date.strftime('%m/%d/%Y')
    else:
        target_date = ''

    produced_date = order[19]
    if produced_date:
        produced_date = produced_date.strftime('%m/%d/%Y')
    else:
        produced_date = ''

    # size="3"
    body_html = '''<style type=\"text/css\">td{{font-family:Arial; color:black; font-size:12pt;}}</style>
        <font face=\"arial\">
        {}
        <hr>
        <table border="0">
        <tr><td align=\"right\">Item&nbsp;Number:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Sales&nbsp;Order:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Customer:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Location:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Model:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Target&nbsp;Date:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Produced&nbsp;Date:&nbsp;</td><td>{}</td></tr>
        </table>
        <hr>
        <table border="0">
        <tr><td align=\"right\">Revision&nbsp;ID:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Related&nbsp;ECR:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Reason:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Document:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Revision&nbsp;Level:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\" valign=\"top\">Description:&nbsp;</td><td>{}</td></tr>
        </table>
        '''.format(shortcuts, item_number, sales_order, customer, location, model, target_date, produced_date,
                   revision_id, related_ecr, reason, document, level, description)

    body = MIMEMultipart('alternative')
    body.attach(MIMEText(body_html, 'html'))
    msg.attach(body)

    # print email_string

    try:
        server.sendmail(sender, email_list, msg.as_string())
    except Exception, e:
        wx.MessageBox('Unable to send email. Error: {}'.format(e), 'An error occurred!', wx.OK | wx.ICON_ERROR)

    server.close()


'''
def check_for_BOM_net_change(item):
    print 'checking for BOM net change report'
    path_to_check = 'S:\Engineering\BOM_Uploads'
    
    new_net_change_found = False
    for file_found in glob.iglob(os.path.join(path_to_check, '*{}*'.format(item))):
        when_modified = dt.datetime.fromtimestamp(os.path.getmtime(file_found))
        date_modified = when_modified.date()
        print '{}	{}'.format(when_modified, file_found)
        
        if date_modified == dt.date.today():
            new_net_change_found = True
            
    if new_net_change_found == False:
        wx.MessageBox('A BOM net change report was NOT found in {}\nfor item number {} today.\n\nMake sure that the BOM was uploaded correctly.'.format(path_to_check, item), 'Warning!', wx.OK | wx.ICON_WARNING)
'''


def send_popsheet_email(revision, order, email_list, sender):
    '''
    threaded_connection = Database.connect_to_database()
    cursor = threaded_connection.cursor()

    #cursor = Database.connection.cursor()
    revision = cursor.execute('SELECT * FROM revisions WHERE id = \'{}\' LAMIT 1'.format(revision_id)).fetchone()
    order = cursor.execute('SELECT * FROM {} WHERE item = \'{}\' LAMIT 1'.format(Ecrs.table_used, revision[1])).fetchone()
    email_list = cursor.execute('SELECT email FROM employees WHERE gets_revision_notice = \'yes\''.format(revision[1])).fetchall()
    sender = cursor.execute('SELECT email FROM employees WHERE name = \'{}\' LAMIT 1'.format(General.app.current_user)).fetchone()[0]
    '''

    server = smtplib.SMTP('mailrelay.lennoxintl.com')
    email_string = ''
    for email in email_list:
        email_string += '; {}'.format(email[0])
    email_string = email_string[2:]

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = email_string
    msg["Subject"] = 'Pop Sheet Notice: {} ({}-{})'.format(order[0], order[1], order[2])
    # msg["Subject"] = 'Revision Notice: {} ({}-{}) {}'.format(order[0], order[1], order[2], revision[2].replace('>', ' > '))
    msg['Date'] = formatdate(localtime=True)

    shortcuts = ''
    try:
        order_directory = OrderFileOpeners.get_order_directory(order[1])
    except Exception as e:
        print e
        order_directory = None

    if order_directory:
        shortcuts = '<a href=\"file:///{}\">Open Order Folder</a>'.format(order_directory)
    # add a 'open revised document' also
    ###<a href=\"file:///R:\\Design_Eng\\Orders\\Orders_2012\\W-X-Y-Z\\Walmart\\212031\\212031-BOM_CA.xls">Open Revised Document</a>

    revision_id = revision[0]
    related_ecr = revision[9]
    # if related_ecr == None: related_ecr = ''
    document = revision[2].replace('>', ' > ')
    level = revision[3]
    reason = revision[5]
    description = revision[4]

    item_number = order[0]
    sales_order = '{}-{}'.format(order[1], order[2])
    customer = order[5]
    location = '{}, {}'.format(order[8], order[9])
    model = order[11]

    target_date = order[23]
    if target_date:
        target_date = target_date.strftime('%m/%d/%Y')
    else:
        target_date = ''

    produced_date = order[19]
    if produced_date:
        produced_date = produced_date.strftime('%m/%d/%Y')
    else:
        produced_date = ''

    # size="3"
    body_html = '''<style type=\"text/css\">td{{font-family:Arial; color:black; font-size:12pt;}}</style>
        <font face=\"arial\">
        {}
        <hr>
        <table border="0">
        <tr><td align=\"right\">Item&nbsp;Number:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Sales&nbsp;Order:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Customer:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Location:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Model:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Target&nbsp;Date:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Produced&nbsp;Date:&nbsp;</td><td>{}</td></tr>
        </table>
        <hr>
        <table border="0">
        <tr><td align=\"right\">Revision&nbsp;ID:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Related&nbsp;ECR:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Reason:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Document:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\">Revision&nbsp;Level:&nbsp;</td><td>{}</td></tr>
        <tr><td align=\"right\" valign=\"top\">Description:&nbsp;</td><td>{}</td></tr>
        </table>
        '''.format(shortcuts, item_number, sales_order, customer, location, model, target_date, produced_date,
                   revision_id, related_ecr, reason, document, level, description)

    body = MIMEMultipart('alternative')
    body.attach(MIMEText(body_html, 'html'))
    msg.attach(body)

    # print email_string

    try:
        server.sendmail(sender, email_list, msg.as_string())
    except Exception, e:
        wx.MessageBox('Unable to send email. Error: {}'.format(e), 'An error occurred!', wx.OK | wx.ICON_ERROR)

    server.close()
