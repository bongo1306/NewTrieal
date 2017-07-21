#!/usr/bin/env python
# -*- coding: utf8 -*-
import wx  # wxWidgets used as the GUI
from wx.html import HtmlEasyPrinting
from wx import xrc  # allows the loading and access of xrc file (xml) that describes GUI
import wx.grid as gridlib
from wxPython.calendar import *

ctrl = xrc.XRCCTRL  # define a shortined function name (just for convienience)

import traceback
import xlwt  # for writing data to excel files

# for sending emails
import smtplib
from email import Encoders
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.MIMEMultipart import MIMEMultipart
 from email.Utils import formatdate

import shutil  # for coping files (like ecr attachments)

import pyodbc  # for connecting to dbworks database

from threading import Thread  # so a slow querry won't make the gui lag (ONLY FOR READS NOT WRITES!)
import sys
import os
import stat
import time
import datetime as dt

import Database
import General
import OrderFileOpeners

# import Printer


# reasons_needing_approval = ['Part Substitution', 'Customer Change', 'Agency Approval', 'Platform Change', 'Continuing Improvement', 'Spec Alignment']

reasons_needing_approval = [
    "Customer Change",
]


def export_for_approval(event):
    '''
    #prompt user to choose where to save
    default_file_name = "ECRs for committee {}".format(str(dt.date.today().strftime('%m-%d-%y')))
    save_dialog = wx.FileDialog(General.app.main_frame, message="Export file as ...",
                            defaultDir=os.path.join(os.path.expanduser("~"), "Desktop"),
                            defaultFile=default_file_name, wildcard="Excel Spreadsheet (*.xls)|*.xls", style=wx.SAVE|wx.OVERWRITE_PROMPT)

    #show the save dialog and get user's input... if not canceled
    if save_dialog.ShowModal() == wx.ID_OK:
        save_path = save_dialog.GetPath()
        save_dialog.Destroy()
    else:
        save_dialog.Destroy()
        return

    #save results data to excel
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('ECRs')
    '''

    ctrl(General.app.main_frame, 'button:export_for_committee').Disable()

    # make that excel read only yo
    os.chmod(General.resource_path('CommitteeECRs.xlsm'), stat.S_IREAD)

    cursor = Database.connection.cursor()

    approval_reasons_string = ''
    for reason in reasons_needing_approval:
        approval_reasons_string += "reason = '{}' OR ".format(reason)
    approval_reasons_string = approval_reasons_string[:-4]

    records = cursor.execute('''
		SELECT 
			ecrs.id, 
			ecrs.reference_number, 
			employee_owner.owner,
			ecrs.reason,
			ecrs.reason,
			ecrs.status,
			ecrs.priority,
			ecrs.request, 
			ecrs.resolution, 
			ecrs.department,
			ecrs.who_requested, 
			ecrs.when_requested,
			ecrs.when_needed
		FROM ecrs LEFT JOIN employee_owner ON employee_owner.employee = ecrs.who_requested WHERE
			ecrs.status='Open' AND
			ecrs.approval_stage='New Request, needs reviewing'
		ORDER BY employee_owner.owner ASC, ecrs.when_requested DESC		
		''').fetchall()

    new_records = []
    for ecr_index, ecr in enumerate(records):
        reason_code = ecr[3]
        if reason_code in reasons_needing_approval:
            new_records.append(ecr)

    ecrs_data = new_records

    '''
    ecrs_data = cursor.execute(''
        SELECT 
            ecrs.id, 
            ecrs.reference_number, 
            employee_owner.owner,
            ecrs.reason,
            ecrs.reason,
            ecrs.status,
            ecrs.priority,
            ecrs.request, 
            ecrs.resolution, 
            ecrs.department,
            ecrs.who_requested, 
            ecrs.when_requested,
            ecrs.when_needed
        FROM ecrs LEFT JOIN employee_owner ON employee_owner.employee = ecrs.who_requested WHERE
            ecrs.when_requested >= '2/18/2013' AND
            ecrs.approval_stage != 'Approved'
        ORDER BY employee_owner.owner ASC, ecrs.when_requested DESC		
        '').fetchall()
    #ORDER BY status DESC, priority, when_needed	
    #		({})
    #	ORDER BY when_requested DESC		
    #	''.format(approval_reasons_string)).fetchall()'''

    headers = ['ID', 'Ref#', 'Owner', 'Primary Code', 'Secondary Code', 'Status', 'Priority', 'Request',
               'Changes Required', 'Department', 'Who Requested', 'When Requested', 'When Needed']

    with open(General.resource_path("CommitteeECRs.txt"), "w") as text_file:
        # write out headers
        text_file.write('{}\n'.format('`````'.join(headers)))

        # write out data
        for ecr_index, ecr_data in enumerate(ecrs_data):
            formatted_ecr_data = []

            for field_index, field in enumerate(ecr_data):
                if field_index == 2:
                    # field = 'owna!'
                    pass

                if field_index == 3:
                    secondary_code = field

                    primary_code = cursor.execute(
                        "SELECT primary_code FROM secondary_ecr_reason_codes WHERE code = '{}'".format(
                            secondary_code)).fetchall()
                    # print primary_code

                    if primary_code:
                        primary_code = primary_code[0][0]

                    if not primary_code:
                        # old secondary code mappings to primary
                        if secondary_code == 'AE Revision': primary_code = 'Order Revision'
                        if secondary_code == 'Agency Approval': primary_code = 'Regulatory'
                        if secondary_code == 'BOM Reconciliation': primary_code = 'Order Revision'
                        if secondary_code == 'Continuing Improvement': primary_code = 'Cost Reduction / VAVE'
                        if secondary_code == 'Customer Change': primary_code = 'Customer Requested Change'
                        if secondary_code == 'Engineering Error': primary_code = 'Corrective Actions'
                        if secondary_code == 'Invalid Request': primary_code = 'Other'
                        if secondary_code == 'New Part Number': primary_code = 'Other'
                        if secondary_code == 'Obsolete Part Number': primary_code = 'Part Substitution'
                        if secondary_code == 'Other': primary_code = 'Other'
                        if secondary_code == 'Part Substitution': primary_code = 'Cost Reduction / VAVE'
                        if secondary_code == 'Platform Change': primary_code = 'Cost Reduction / VAVE'
                        if secondary_code == 'Product Development': primary_code = 'Cost Reduction / VAVE'
                        if secondary_code == 'Product Improvement': primary_code = 'Cost Reduction / VAVE'
                        if secondary_code == 'Proposal Drawing': primary_code = 'Quote / Information Request'
                        if secondary_code == 'Recall': primary_code = 'Other'
                        if secondary_code == 'Request for Information': primary_code = 'Quote / Information Request'
                        if secondary_code == 'Spec Alignment': primary_code = 'Cost Reduction / VAVE'
                        if secondary_code == 'Unit Reallocation': primary_code = 'Other'

                    field = primary_code

                if type(field) == dt.datetime:
                    formatted_ecr_data.append(field.strftime('%m/%d/%Y %I:%M %p').replace(' 11:59 PM', ''))
                else:
                    formatted_ecr_data.append(str(field).replace('None', '').replace('\n', '~~~~~'))

            text_file.write('{}\n'.format('`````'.join(formatted_ecr_data)))

    os.startfile(General.resource_path('CommitteeECRs.xlsm'))

    ctrl(General.app.main_frame, 'button:export_for_committee').Enable()

    '''
    headers = ['ID', 'Ref#', 'Item#', 'Department', 'Request', 'Changes Required', 'Priority', 'Who Requested', 'When Requested', 'When Needed', 'Who Approved 1st Stage', 'Who Approved 2nd Stage']

    #write out headers
    for header_index, header in enumerate(headers):
        worksheet.write(0, header_index, header)

    #write out data
    for ecr_index, ecr_data in enumerate(ecrs_data):
        for field_index, field in enumerate(ecr_data):
            if type(field) == dt.datetime:
                worksheet.write(ecr_index+1, field_index, field.strftime('%m/%d/%Y %I:%M %p').replace(' 11:59 PM', ''))
            else:
                worksheet.write(ecr_index+1, field_index, field)

    workbook.save(save_path)

    wx.MessageBox('Export completed.', 'Info', wx.OK | wx.ICON_INFORMATION)
    '''


def get_similar_items(item):
    # find items that could be similarly affected by the issue in this ECR
    try:
        cursor = Database.connection.cursor()

        customer, family = cursor.execute("SELECT customer, family FROM orders WHERE item='{}'".format(item)).fetchone()

        sql = '''
			SELECT
				orders.sales_order, 
				orders.item,
				item_responsibilities2.project_lead,
				item_responsibilities2.mechanical_engineer,
				item_responsibilities2.electrical_engineer,
				item_responsibilities2.structural_engineer
			FROM 
				orders LEFT JOIN item_responsibilities2 ON orders.item = item_responsibilities2.item
			WHERE
				orders.item <> '{}' AND 
				orders.customer = '{}' AND
				orders.family = '{}' AND
				orders.date_produced IS NULL AND
				orders.date_de_released IS NOT NULL AND
				orders.is_canceled = 0
			ORDER BY
				orders.item
		'''.format(item, customer.replace("'", "''"), family)
        similar_items = cursor.execute(sql).fetchall()

        sales_orders = list(set(zip(*similar_items)[0]))
        sales_orders.sort()

        similar_text = ""

        for so in sales_orders:
            similar_text += '({}) '.format(so)

            for similar_entry in similar_items:
                if similar_entry[0] == so:
                    similar_text += '{}, '.format(similar_entry[1])

            similar_text = similar_text[:-2]
            similar_text += '     '

        try:
            similar_text = similar_text[:-4]
        except:
            pass

        return (similar_text, similar_items)

    except Exception as e:
        print e

    return None


def on_click_approve_selected_ers(event):
    print 'APPROVE SELECTED ECRS'
    list_ctrl = ctrl(General.app.main_frame, 'list:committee_ecrs')

    cursor = Database.connection.cursor()

    for row in range(list_ctrl.GetItemCount()):
        if list_ctrl.IsSelected(row):
            ecr_id = list_ctrl.GetItem(row, 0).GetText()

            cursor.execute("UPDATE ecrs SET approval_stage = 'Approved' WHERE id = {}".format(ecr_id))

    Database.connection.commit()

    refresh_open_ecrs_list()
    refresh_committee_ecrs_list()


def on_click_approve_1_for_modify(event):
    button = event.GetEventObject()

    if button.GetLabel() == 'Approve':
        button.SetLabel('Unapprove')
        ctrl(General.app.modify_ecr_dialog, 'label:who_approved_first').SetLabel(General.app.current_user)
    else:
        button.SetLabel('Approve')
        ctrl(General.app.modify_ecr_dialog, 'label:who_approved_first').SetLabel('')

    ctrl(General.app.modify_ecr_dialog, 'panel:committee').Layout()


def on_click_approve_2_for_modify(event):
    button = event.GetEventObject()

    if button.GetLabel() == 'Approve':
        button.SetLabel('Unapprove')
        ctrl(General.app.modify_ecr_dialog, 'label:who_approved_second').SetLabel(General.app.current_user)
    else:
        button.SetLabel('Approve')
        ctrl(General.app.modify_ecr_dialog, 'label:who_approved_second').SetLabel('')

    ctrl(General.app.modify_ecr_dialog, 'panel:committee').Layout()


def on_click_approve_1_for_close(event):
    button = event.GetEventObject()

    if button.GetLabel() == 'Approve':
        button.SetLabel('Unapprove')
        ctrl(General.app.close_ecr_dialog, 'label:who_approved_first').SetLabel(General.app.current_user)
    else:
        button.SetLabel('Approve')
        ctrl(General.app.close_ecr_dialog, 'label:who_approved_first').SetLabel('')

    ctrl(General.app.close_ecr_dialog, 'panel:committee').Layout()


def on_click_approve_2_for_close(event):
    button = event.GetEventObject()

    if button.GetLabel() == 'Approve':
        button.SetLabel('Unapprove')
        ctrl(General.app.close_ecr_dialog, 'label:who_approved_second').SetLabel(General.app.current_user)
    else:
        button.SetLabel('Approve')
        ctrl(General.app.close_ecr_dialog, 'label:who_approved_second').SetLabel('')

    ctrl(General.app.close_ecr_dialog, 'panel:committee').Layout()


def get_cleaned_list_headers(list):
    headers = []
    for col in range(list.GetColumnCount()):
        headers.append(list.GetColumn(col).GetText().replace(u'↓', '').replace(u'↑', '').strip())

    return headers


def set_list_headers(list, headers):
    list.DeleteAllColumns()
    for header_index, header in enumerate(headers):
        list.InsertColumn(header_index, header)


'''
def check_description_for_parts(event):
	entry = ctrl(General.app.main_frame, 'text:text:mentioned_parts').GetValue()

    try:
        conn = pyodbc.connect('DSN=BPCS;')
        cursor = conn.cursor()

		#get part number descriptions from BPCS
		entry = xx.REQUEST_DESCRIPTION

		partNumbers = []

		for i in range(0, len(entry)):
			if entry[i:i+8].find(' ') == -1:
				if entry[i].isdigit() and entry[i+5:i+8].isdigit():
					if len(entry[i:i+8]) == 8:
						partNumbers.append(entry[i:i+8].upper())
					#print entry[i:i+8]

		if len(partNumbers) == 0:
			continue

		try:
			entry += "\n"

			for i in partNumbers:
				#print i
				SQL = 'SELECT IIM.IDSCE, IIM.IDESC FROM RDBDIRE.V60BPCSF.IIM IIM WHERE (IIM.IPROD=\'' + i + '\')'
				fetched = cursor.execute(SQL).fetchone()
				try:
					entry += '\n-----\n' + i + '\t' + fetched[0].strip(' ') + '\t' + fetched[1].strip(' ')
				except:
					entry += '\n-----\n' + 'BPCS ERROR'

			xx.REQUEST_DESCRIPTION = entry
		except:
			print 'HEY and error: ' + str(fetched)


        #close BPCS database
        cursor.close()
        conn.close()
    except:
        print 'Failed to connect to DBworks database'
'''


def check_reference_field(event):
    entry = ctrl(General.app.new_ecr_dialog, 'text:reference_number').GetValue()

    # ctrl(General.app.main_frame, 'statusbar:main').SetStatusText('')
    clear_useful_info_panel()

    # determine reference type by entry length and other defining characteristics
    if (len(entry) == 7) and (entry[0] == '0'):
        # it might be an item number, search for it in the DB
        sql = "SELECT * FROM orders WHERE item LIKE '%{}%'".format(entry)
        thread = Thread(target=Database.query_one, args=(sql, update_useful_info_panel))
        thread.start()

    elif '-' in entry:
        # it might be a sales order with specified line up, search for it in the DB
        sales_order = entry.split('-')[0]
        line_up = 1
        if entry.split('-')[1] != '':
            line_up = int(float(entry.split('-')[1]))
        sql = "SELECT * FROM orders WHERE sales_order LIKE '%{}%' AND line_up = '{}'".format(sales_order, line_up)
        thread = Thread(target=Database.query_one, args=(sql, update_useful_info_panel))
        thread.start()

    elif (len(entry) == 9) and (entry[0:3] == 'KW0'):
        # it might be an item number with KW prefix, search for it in the DB
        sql = "SELECT * FROM orders WHERE item LIKE '%{}%'".format(entry)
        thread = Thread(target=Database.query_one, args=(sql, update_useful_info_panel))
        thread.start()

    elif len(entry) == 8 and entry[0] == '2':
        # it might be a sales order, search for it in the DB
        sql = "SELECT * FROM orders WHERE item LIKE '%{}%'".format(entry)
        thread = Thread(target=Database.query_one, args=(sql, update_useful_info_panel))
        thread.start()


    elif len(entry) == 10:
        # it might be a serial number, search for it in the DB
        sql = "SELECT * FROM orders WHERE serial = \'{}\'".format(entry)
        thread = Thread(target=Database.query_one, args=(sql, update_useful_info_panel))
        thread.start()


    elif len(entry) == 6:
        # it might be a sales order, search for it in the DB
        sql = "SELECT * FROM orders WHERE sales_order LIKE '%{}%' OR quote = \'{}\'".format(entry, entry)
        thread = Thread(target=Database.query_one, args=(sql, update_useful_info_panel))
        thread.start()

    elif len(entry) == 8 and entry[0] == '5':
        # it might be a sales order, search for it in the DB
        sql = "SELECT * FROM orders WHERE sales_order LIKE '%{}%'".format(entry)
        thread = Thread(target=Database.query_one, args=(sql, update_useful_info_panel))
        thread.start()

    else:
        clear_useful_info_panel()


def clear_useful_info_panel():
    ctrl(General.app.new_ecr_dialog, 'label:item_number').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:sales_order').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:line_up').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:serial').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:quote').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:model').SetLabel('...')

    ctrl(General.app.new_ecr_dialog, 'label:customer').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:store_name').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:store_id').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:city').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:state').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:country').SetLabel('...')

    # ctrl(General.app.new_ecr_dialog, 'label:mechanical').SetLabel('...')
    # ctrl(General.app.new_ecr_dialog, 'label:electrical').SetLabel('...')
    # ctrl(General.app.new_ecr_dialog, 'label:structural').SetLabel('...')
    # ctrl(General.app.new_ecr_dialog, 'label:program').SetLabel('...')

    ctrl(General.app.new_ecr_dialog, 'label:order_status').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:date_released').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:date_produced').SetLabel('...')
    ctrl(General.app.new_ecr_dialog, 'label:date_shipped').SetLabel('...')


'''
def export_search_results(event):
	#prompt user to choose where to save
	save_dialog = wx.FileDialog(General.app.main_frame, message="Export file as ...", 
							defaultDir=os.path.join(os.path.expanduser("~"), "Desktop"), 
							defaultFile="ecr_search_results", wildcard="Excel Spreadsheet (*.xls)|*.xls", style=wx.SAVE|wx.OVERWRITE_PROMPT)

	#show the save dialog and get user's input... if not canceled
	if save_dialog.ShowModal() == wx.ID_OK:
		save_path = save_dialog.GetPath()
		save_dialog.Destroy()
	else:
		save_dialog.Destroy()
		return

	#save results data to excel
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet('ECR Search Results')

	results_list = ctrl(General.app.main_frame, 'list:results')
	#current_item = results_list.GetTopItem()

	for row in range(results_list.GetItemCount()):
		for col in range(results_list.GetColumnCount()):
			worksheet.write(row, col, results_list.GetItem(row, col).GetText())

	workbook.save(save_path)

	wx.MessageBox('Export completed.', 'Info', wx.OK | wx.ICON_INFORMATION)
'''


def hide_search_ecrs_dialog(event):
    event.GetEventObject().Hide()


def hide_things_based_on_user_department():
    cursor = Database.connection.cursor()

    department = cursor.execute(
        'SELECT TOP 1 department FROM employees WHERE name = \'{}\''.format(General.app.current_user)).fetchone()[0]

    if (department != 'Design Engineering') and (department != 'Applications Engineering'):
        ctrl(General.app.main_frame, 'static:mechanical').Hide()
        ctrl(General.app.main_frame, 'static:electrical').Hide()
        ctrl(General.app.main_frame, 'static:structural').Hide()
        ctrl(General.app.main_frame, 'label:order_panel_mechanical').Hide()
        ctrl(General.app.main_frame, 'label:order_panel_electrical').Hide()
        ctrl(General.app.main_frame, 'label:order_panel_structural').Hide()
        ctrl(General.app.main_frame, 'line:order').Hide()

        ctrl(General.app.main_frame, 'line:ecr').Hide()
        ctrl(General.app.main_frame, 'static:claimed_by').Hide()
        ctrl(General.app.main_frame, 'static:assigned_to').Hide()
        ctrl(General.app.main_frame, 'label:ecr_panel_claimed_by').Hide()
        ctrl(General.app.main_frame, 'label:ecr_panel_assigned_to').Hide()

        ctrl(General.app.main_frame, 'line:ecr1').Hide()
        ctrl(General.app.main_frame, 'button:claim').Hide()
        ctrl(General.app.main_frame, 'button:modify').Hide()
        ctrl(General.app.main_frame, 'button:assign').Hide()
        ctrl(General.app.main_frame, 'button:close').Hide()

        ctrl(General.app.main_frame, 'line:ecr2').Hide()
        ctrl(General.app.main_frame, 'button:add_revisions_with_ecr').Hide()


def on_click_attatch_document(event):
    dialog = wx.FileDialog(None, style=wx.OPEN | wx.MULTIPLE)

    if dialog.ShowModal() == wx.ID_OK:
        file_paths = dialog.GetPaths()

        # write the file names to the text box and write the full file
        # name with path to the hidden text box next to it so we can read it later
        files_string = ''
        files_with_path_string = ''

        for path in file_paths:
            files_string += '; {}'.format(path.split('\\')[-1])
            files_with_path_string += ';{}'.format(path)

        files_string = files_string[2:]
        files_with_path_string = files_with_path_string[1:]

        ctrl(General.app.new_ecr_dialog, 'text:attached_document').SetValue(files_string)
        ctrl(General.app.new_ecr_dialog, 'text:attached_document_paths').SetValue(files_with_path_string)
    else:
        ctrl(General.app.new_ecr_dialog, 'text:attached_document').SetValue('')
        ctrl(General.app.new_ecr_dialog, 'text:attached_document_paths').SetValue('')

    dialog.Destroy()


def on_click_claim_ecr(event):
    ecr_id = ctrl(General.app.main_frame, 'label:ecr_panel_id').GetLabel()
    if ecr_id == '':
        return

    cursor = Database.connection.cursor()
    cursor.execute('''
		UPDATE ecrs SET who_claimed=\'{}\', when_claimed=\'{}\' WHERE id=\'{}\'
		'''.format(General.app.current_user, str(dt.datetime.today())[:19], ecr_id))
    Database.connection.commit()

    refresh_my_ecrs_list()
    refresh_open_ecrs_list()
    refresh_closed_ecrs_list()
    refresh_my_assigned_ecrs_list()
    refresh_committee_ecrs_list()
    populate_ecr_panel(ecr_id)


def on_click_close_ecr(event):
    if ctrl(General.app.close_ecr_dialog, 'choice:ecr_reason').GetStringSelection() == 'Engineering Error':
        if ctrl(General.app.close_ecr_dialog, 'choice:who_errored').GetStringSelection() == '':
            wx.MessageBox('Since this is an Engineering Error, you must select who errored before closing the ECR.',
                          'Hint', wx.OK | wx.ICON_WARNING)
            return

        if ctrl(General.app.close_ecr_dialog, 'choice:ecr_component').GetStringSelection() == '':
            wx.MessageBox(
                'Since this is an Engineering Error, you must select the applicable component before closing the ECR.',
                'Hint', wx.OK | wx.ICON_WARNING)
            return

        if ctrl(General.app.close_ecr_dialog, 'choice:ecr_sub_system').GetStringSelection() == '':
            wx.MessageBox(
                'Since this is an Engineering Error, you must select the applicable sub system before closing the ECR.',
                'Hint', wx.OK | wx.ICON_WARNING)
            return

    if ctrl(General.app.close_ecr_dialog, 'text:resolution').GetValue().strip() == '':
        wx.MessageBox('Please enter a descriptive resolution before closing the ECR.', 'Hint', wx.OK | wx.ICON_WARNING)
        return

    need_by_date = ctrl(General.app.close_ecr_dialog, 'calendar:ecr_need_by').GetDate()
    need_by_date = dt.date(need_by_date.GetYear(), need_by_date.GetMonth() + 1, need_by_date.GetDay())

    cursor = Database.connection.cursor()

    # department = cursor.execute('SELECT department FROM employees WHERE name = \'{}\' LAMIT 1'.format(General.app.current_user)).fetchone()[0]

    ##attachment_string = ctrl(General.app.close_ecr_dialog, 'text:attached_document_paths').GetValue()
    '''
    try:
        if attachment_string != '':
            last_id = cursor.execute('SELECT MAX(id) FROM ecrs').fetchone()[0]

            attachment_files = attachment_string.split(';')

            attachment_string = ''
            for attachment in attachment_files:
                attachment_string += ';{}_{}'.format(last_id+1, attachment.split('\\')[-1])

                shutil.copyfile(attachment, '{}\\{}_{}'.format(General.attachment_directory, last_id+1, attachment.split('\\')[-1]))

            attachment_string = attachment_string[1:]
    except:
        wx.MessageBox('ECR could not be submitted with those attachments for some reason...\n\n{}'.format(traceback.format_exc()), 'An error occurred!', wx.OK | wx.ICON_ERROR)
        return
    '''

    type = General.app.ecr_type
    reference_number = ctrl(General.app.close_ecr_dialog, 'text:reference_number').GetValue().replace("'",
                                                                                                      "''").replace(
        '\"', "''''")
    item_number = Database.get_item_from_ref(reference_number)
    reason = ctrl(General.app.close_ecr_dialog, 'choice:ecr_reason').GetStringSelection()
    document = ctrl(General.app.close_ecr_dialog, 'choice:ecr_document').GetStringSelection()
    request = ctrl(General.app.close_ecr_dialog, 'text:description').GetValue().replace("'", "''").replace('\"', "''''")
    when_needed = need_by_date
    resolution = ctrl(General.app.close_ecr_dialog, 'text:resolution').GetValue().replace("'", "''").replace('\"',
                                                                                                             "''''")
    who_errored = ctrl(General.app.close_ecr_dialog, 'choice:who_errored').GetStringSelection().replace("'",
                                                                                                        "''").replace(
        '\"', "''''")
    who_closed = General.app.current_user
    when_closed = str(dt.datetime.today())[:19]
    id = General.app.close_ecr_dialog.GetTitle().split(' ')[-1]
    who_approved_first = ctrl(General.app.close_ecr_dialog, 'label:who_approved_first').GetLabel()
    who_approved_second = ctrl(General.app.close_ecr_dialog, 'label:who_approved_second').GetLabel()
    priority = ctrl(General.app.close_ecr_dialog, 'spin:priority').GetValue()

    # track if someone is changing the reason code
    previous_reason_code = cursor.execute("SELECT reason FROM ecrs WHERE id = '{}'".format(id)).fetchone()[0]
    new_reason_code = reason
    if previous_reason_code != new_reason_code:
        sql = "INSERT INTO ecr_reason_code_changes (ecr_id, who_changed, when_changed, previous_code, new_code) VALUES ("
        sql += "{}, ".format(id)
        sql += "'{}', ".format(General.app.current_user)
        sql += "'{}', ".format(str(dt.datetime.today())[:19])
        sql += "'{}', ".format(previous_reason_code)
        sql += "'{}')".format(new_reason_code)
        cursor.execute(sql)

    sql = 'UPDATE ecrs SET '
    sql += 'type=\'{}\', '.format(type)
    sql += 'reference_number=\'{}\', '.format(reference_number)

    if item_number == None:
        sql += 'item=NULL, '
    else:
        sql += 'item=\'{}\', '.format(item_number)

    sql += 'reason=\'{}\', '.format(reason)
    sql += 'document=\'{}\', '.format(document)
    sql += 'request=\'{}\', '.format(request)
    sql += 'when_needed=\'{} 23:59:00\', '.format(when_needed)
    sql += 'resolution=\'{}\', '.format(resolution)

    if who_errored != '':
        sql += 'who_errored=\'{}\', '.format(who_errored)
    else:
        sql += 'who_errored=NULL, '

    sql += 'status=\'Closed\', '
    sql += 'who_closed=\'{}\', '.format(who_closed)
    sql += 'when_closed=\'{}\', '.format(when_closed)

    # if who_approved_first != '':
    #	sql += 'who_approved_first=\'{}\', '.format(who_approved_first)
    # if who_approved_second != '':
    #	sql += 'who_approved_second=\'{}\', '.format(who_approved_second)

    if who_approved_first == '':
        sql += 'who_approved_first = NULL, '
    else:
        sql += 'who_approved_first=\'{}\', '.format(who_approved_first)

    if who_approved_second == '':
        sql += 'who_approved_second = NULL, '
    else:
        sql += 'who_approved_second=\'{}\', '.format(who_approved_second)

    sql += 'approval_stage=\'{}\', '.format(ctrl(General.app.close_ecr_dialog, 'choice:stage').GetStringSelection())

    component = ctrl(General.app.close_ecr_dialog, 'choice:ecr_component').GetStringSelection()
    if component != '':
        sql += "component='{}', ".format(component.replace("'", "''").replace('\"', "''''"))
    else:
        sql += "component=NULL, "

    sub_system = ctrl(General.app.close_ecr_dialog, 'choice:ecr_sub_system').GetStringSelection()
    if sub_system != '':
        sql += "sub_system='{}', ".format(sub_system.replace("'", "''").replace('\"', "''''"))
    else:
        sql += "sub_system=NULL, "

    severity = ctrl(General.app.close_ecr_dialog, 'choice:ecr_severity').GetStringSelection()
    if severity == 'High':
        sql += "severity=1.0, "
    elif severity == 'Medium':
        sql += "severity=0.5, "
    elif severity == 'Low':
        sql += "severity=0.1, "
    else:
        sql += "severity=1.0, "

    sql += 'priority=\'{}\' '.format(priority)

    sql += 'WHERE id=\'{}\''.format(id)

    cursor.execute(sql)
    Database.connection.commit()

    cursor.execute("UPDATE ecrs SET attachment = NULL WHERE attachment = ''")
    Database.connection.commit()

    ecr = cursor.execute(
        'SELECT TOP 1 id, reference_number, request, resolution, who_requested, who_errored, item, type, document FROM ecrs WHERE id = \'{}\''.format(
            id)).fetchone()
    # order = Database.get_order_data_from_ref(reference_number)
    order = cursor.execute("SELECT TOP 1 * FROM orders WHERE item = \'{}\'".format(ecr[6])).fetchone()

    # takes a little longer to send an email so put it in a serperate thread so user
    # doesn't have to wait around :)
    sender = \
    cursor.execute('SELECT TOP 1 email FROM employees WHERE name = \'{}\''.format(General.app.current_user)).fetchone()[
        0]

    # email person who originally entered the ECR
    reciever = cursor.execute('SELECT TOP 1 email FROM employees WHERE name = \'{}\''.format(
        ecr[4].replace("'", "''").replace('\"', "''''"))).fetchone()[0]
    Thread(target=send_ecr_closed_email, args=(ecr, order, reciever, sender)).start()

    # email the engineer who errored
    if who_errored != '':
        reciever = cursor.execute('SELECT TOP 1 email FROM employees WHERE name = \'{}\''.format(
            ecr[5].replace("'", "''").replace('\"', "''''"))).fetchone()[0]
        Thread(target=send_ecr_soe_email, args=(ecr, order, reciever, sender)).start()

    # email people who worked on items that may be similarly affected by this ECR
    if ctrl(General.app.close_ecr_dialog, 'checkbox:similar_ecrs').GetValue() == True:
        similar_items_data = get_similar_items(item_number)
        if similar_items_data:
            reciever_names = []

            ecr_type = ecr[7]

            for data in similar_items_data[1]:
                sales_order, item, project_lead, mechanical_engineer, electrical_engineer, structural_engineer = data

                if type == 'Mechanical':
                    reciever_names.append(mechanical_engineer)

                elif type == 'Electrical':
                    reciever_names.append(electrical_engineer)

                elif type == 'Structural':
                    reciever_names.append(structural_engineer)

                elif type == 'Other':
                    reciever_names.append(project_lead)

                # if no one assigned to that post... call out project lead
                if reciever_names[-1] == None:
                    reciever_names.append(project_lead)

            reciever_emails = []
            reciever_names = list(set(reciever_names))

            for reciever_name in reciever_names:
                if reciever_name:
                    result = cursor.execute("SELECT TOP 1 email FROM employees WHERE name = '{}'".format(
                        reciever_name.replace("'", "''"))).fetchone()
                    if result:
                        reciever_emails.append(result[0])

            reciever_emails = list(set(reciever_emails))

            Thread(target=send_similar_items_email,
                   args=(order, ecr, similar_items_data[1], reciever_emails, sender)).start()

    refresh_my_ecrs_list()
    refresh_open_ecrs_list()
    refresh_closed_ecrs_list()
    refresh_my_assigned_ecrs_list()
    refresh_committee_ecrs_list()
    populate_ecr_panel(id)

    General.app.close_ecr_dialog.Destroy()


def on_click_open_assign_ecr_form(event):
    ecr_id = ctrl(General.app.main_frame, 'label:ecr_panel_id').GetLabel()
    if ecr_id == '':
        return
    General.app.init_assign_ecr_dialog(ecr_id)


def on_click_open_attachments(event):
    ecr_id = ctrl(General.app.main_frame, 'label:ecr_panel_id').GetLabel()
    if ecr_id == '':
        return

    cursor = Database.connection.cursor()
    attachment_files = cursor.execute('SELECT TOP 1 attachment FROM ecrs WHERE id = \'{}\''.format(ecr_id)).fetchone()[
        0]

    for attachment_file in attachment_files.split(';'):
        try:
            os.startfile('{}\\{}'.format(General.attachment_directory, attachment_file))
        except:
            wx.MessageBox('Could not open file: {}\\{}'.format(General.attachment_directory, attachment_file), 'Error',
                          wx.OK | wx.ICON_ERROR)


def on_click_open_close_ecr_form(event):
    ecr_id = ctrl(General.app.main_frame, 'label:ecr_panel_id').GetLabel()
    if ecr_id == '':
        return
    General.app.init_close_ecr_dialog(ecr_id)


def on_click_open_email_ecr_form(event):
    General.app.init_email_ecr_dialog()


def on_click_open_modify_ecr_form(event):
    ecr_id = ctrl(General.app.main_frame, 'label:ecr_panel_id').GetLabel()
    if ecr_id == '':
        return
    General.app.init_modify_ecr_dialog(ecr_id)


def on_click_open_new_ecr_form(event):
    General.app.init_new_ecr_dialog()


def on_click_open_duplicate_ecr_form(event):
    ecr_id = ctrl(General.app.main_frame, 'label:ecr_panel_id').GetLabel()
    if ecr_id == '':
        return
    General.app.init_new_ecr_dialog(ecr_id)


def on_click_open_search_ecrs_form(event):
    General.app.init_search_ecrs_dialog()


def on_click_print_ecr(event):
    item_number = ctrl(General.app.main_frame, 'label:order_panel_item_number').GetLabel()
    ecr_id = ctrl(General.app.main_frame, 'label:ecr_panel_id').GetLabel()
    if ecr_id == '':
        return

    cursor = Database.connection.cursor()
    # order = cursor.execute('SELECT TOP 1 item, sales_order, line_up, customer, city, state, model, mechanical_by, electrical_by, structural_by FROM orders WHERE item = \'{}\''.format(item_number)).fetchone()
    ecr = cursor.execute(
        'SELECT TOP 1 id, reference_number, document, reason, who_requested, department, when_needed, request, resolution FROM ecrs WHERE id = \'{}\''.format(
            ecr_id)).fetchone()

    order = cursor.execute('''
		SELECT TOP 1
			orders.item,
			orders.sales_order,
			orders.line_up,
			orders.customer,
			orders.city,
			orders.state,
			orders.model,
			orders.mechanical_by,
			orders.electrical_by,
			orders.structural_by,
			item_responsibilities2.mechanical_cad_designer,
			item_responsibilities2.electrical_cad_designer,
			item_responsibilities2.structural_cad_designer
		FROM orders LEFT JOIN item_responsibilities2 ON orders.item = item_responsibilities2.item WHERE
			orders.item = '{}'	
		'''.format(item_number)).fetchone()

    html_to_print = '''<style type=\"text/css\">td{{font-family:Arial; color:black; font-size:12pt;}}</style>'''

    if order != None:
        item, sales_order, line_up, customer, city, state, model, \
        mechanical_by, electrical_by, structural_by, \
        mechanical_cad_designer, electrical_cad_designer, structural_cad_designer = order

        # use who done its from item_responsibilities if there... otherwise use from filemaker
        if mechanical_cad_designer: mechanical_by = mechanical_cad_designer
        if electrical_cad_designer: electrical_by = electrical_cad_designer
        if structural_cad_designer: structural_by = structural_cad_designer

        html_to_print += '''
		<table border="0">
		<tr><td align=\"right\">Item&nbsp;Number:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Sales&nbsp;Order:&nbsp;</td><td>{}-{}</td></tr>
		<tr><td align=\"right\">Customer:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Location:&nbsp;</td><td>{}, {}</td></tr>
		<tr><td align=\"right\">Model:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Mechanical&nbsp;by:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Electrical&nbsp;by:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Structural&nbsp;by:&nbsp;</td><td>{}</td></tr>
		</table>
		<hr>
		'''.format(item, sales_order, line_up, customer, city, state, model, mechanical_by, electrical_by,
                   structural_by)

    html_to_print += '''
		<table border="0">
		<tr><td align=\"right\">ECR&nbsp;ID:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">ReferenceNo:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Document:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Reason:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Who&nbsp;Requested:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Department:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">When&nbsp;Needed:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\" valign=\"top\">Request:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\" valign=\"top\">Resolution:&nbsp;</td><td>{}</td></tr>
		</table>
		'''.format(ecr[0], ecr[1], ecr[2], ecr[3], ecr[4], ecr[5], General.format_date_nicely(ecr[6]), ecr[7], ecr[8])

    printer = HtmlEasyPrinting()
    printer.GetPrintData().SetPaperId(wx.PAPER_LETTER)
    printer.PrintText(html_to_print)


def on_click_submit_ecr(event):
    # check that field entries are valid_boundary
    if ctrl(General.app.new_ecr_dialog, 'choice:ecr_reason').GetStringSelection() == '':
        wx.MessageBox('You must select a reason for request from the drop down menu\nbefore submitting a new ECR.',
                      'Hint', wx.OK | wx.ICON_WARNING)
        return

    if ctrl(General.app.new_ecr_dialog, 'choice:ecr_document').GetStringSelection() == '':
        wx.MessageBox('You must select a document from the drop down menu\nbefore submitting a new ECR.', 'Hint',
                      wx.OK | wx.ICON_WARNING)
        return

    if ctrl(General.app.new_ecr_dialog, 'text:description').GetValue().strip() == '':
        wx.MessageBox(
            'You must enter a request description before submitting a new ECR.\nThe more descriptive, the faster the answer.',
            'Hint', wx.OK | wx.ICON_WARNING)
        return

    need_by_date = ctrl(General.app.new_ecr_dialog, 'calendar:ecr_need_by').GetDate()
    need_by_date = dt.date(need_by_date.GetYear(), need_by_date.GetMonth() + 1, need_by_date.GetDay())
    if need_by_date < dt.date.today():
        wx.MessageBox('You cannot request a need by date from the past...', 'Hint', wx.OK | wx.ICON_WARNING)
        return

    cursor = Database.connection.cursor()

    department = cursor.execute(
        'SELECT TOP 1 department FROM employees WHERE name = \'{}\''.format(General.app.current_user)).fetchone()[0]

    attachment_string = ctrl(General.app.new_ecr_dialog, 'text:attached_document_paths').GetValue()

    try:
        if attachment_string != '':
            last_id = cursor.execute('SELECT MAX(id) FROM ecrs').fetchone()[0]

            attachment_files = attachment_string.split(';')

            attachment_string = ''
            for attachment in attachment_files:
                attachment_string += ';{}_{}'.format(last_id + 1, attachment.split('\\')[-1])

                shutil.copyfile(attachment, '{}\\{}_{}'.format(General.attachment_directory, last_id + 1,
                                                               attachment.split('\\')[-1]))

            attachment_string = attachment_string[1:]
    except:
        wx.MessageBox(
            'ECR could not be submitted with those attachments for some reason...\n\n{}'.format(traceback.format_exc()),
            'An error occurred!', wx.OK | wx.ICON_ERROR)
        return

    reference_number = ctrl(General.app.new_ecr_dialog, 'text:reference_number').GetValue().replace("'", "''").replace(
        '\"', "''''")
    item_number = Database.get_item_from_ref(reference_number)

    new_id = cursor.execute("SELECT MAX(id) FROM ecrs").fetchone()[0] + 1

    if item_number != None:
        sql = 'INSERT INTO ecrs (id, status, reference_number, item, document, reason, department, who_requested, type, request, attachment, when_requested, when_needed) VALUES ('
        sql += '{}, '.format(new_id)
        sql += '\'Open\', '
        sql += '\'{}\', '.format(reference_number)
        sql += '\'{}\', '.format(item_number)
        sql += '\'{}\', '.format(ctrl(General.app.new_ecr_dialog, 'choice:ecr_document').GetStringSelection())
        sql += '\'{}\', '.format(ctrl(General.app.new_ecr_dialog, 'choice:ecr_reason').GetStringSelection())
        sql += '\'{}\', '.format(department)
        sql += '\'{}\', '.format(General.app.current_user)
        sql += '\'{}\', '.format(General.app.ecr_type)
        sql += '\'{}\', '.format(
            ctrl(General.app.new_ecr_dialog, 'text:description').GetValue().replace("'", "''").replace('\"', "''''"))
        sql += '\'{}\', '.format(attachment_string)
        sql += '\'{}\', '.format(str(dt.datetime.today())[:19])
        sql += '\'{} 23:59:00\')'.format(need_by_date)
    else:
        sql = 'INSERT INTO ecrs (id, status, reference_number, document, reason, department, who_requested, type, request, attachment, when_requested, when_needed) VALUES ('
        sql += '{}, '.format(new_id)
        sql += '\'Open\', '
        sql += '\'{}\', '.format(reference_number)
        sql += '\'{}\', '.format(ctrl(General.app.new_ecr_dialog, 'choice:ecr_document').GetStringSelection())
        sql += '\'{}\', '.format(ctrl(General.app.new_ecr_dialog, 'choice:ecr_reason').GetStringSelection())
        sql += '\'{}\', '.format(department)
        sql += '\'{}\', '.format(General.app.current_user)
        sql += '\'{}\', '.format(General.app.ecr_type)
        sql += '\'{}\', '.format(
            ctrl(General.app.new_ecr_dialog, 'text:description').GetValue().replace("'", "''").replace('\"', "''''"))
        sql += '\'{}\', '.format(attachment_string)
        sql += '\'{}\', '.format(str(dt.datetime.today())[:19])
        sql += '\'{} 23:59:00\')'.format(need_by_date)

    # print sql
    cursor.execute(sql)
    cursor.execute("UPDATE ecrs SET attachment = NULL WHERE attachment = ''")

    Database.connection.commit()

    refresh_my_ecrs_list()
    refresh_open_ecrs_list()
    General.app.new_ecr_dialog.Destroy()


def on_click_modify_ecr(event):
    need_by_date = ctrl(General.app.modify_ecr_dialog, 'calendar:ecr_need_by').GetDate()
    need_by_date = dt.date(need_by_date.GetYear(), need_by_date.GetMonth() + 1, need_by_date.GetDay())

    cursor = Database.connection.cursor()

    # department = cursor.execute('SELECT department FROM employees WHERE name = \'{}\' LAMIT 1'.format(General.app.current_user)).fetchone()[0]

    ##attachment_string = ctrl(General.app.modify_ecr_dialog, 'text:attached_document_paths').GetValue()
    '''
    try:
        if attachment_string != '':
            last_id = cursor.execute('SELECT MAX(id) FROM ecrs').fetchone()[0]

            attachment_files = attachment_string.split(';')

            attachment_string = ''
            for attachment in attachment_files:
                attachment_string += ';{}_{}'.format(last_id+1, attachment.split('\\')[-1])

                shutil.copyfile(attachment, '{}\\{}_{}'.format(General.attachment_directory, last_id+1, attachment.split('\\')[-1]))

            attachment_string = attachment_string[1:]
    except:
        wx.MessageBox('ECR could not be submitted with those attachments for some reason...\n\n{}'.format(traceback.format_exc()), 'An error occurred!', wx.OK | wx.ICON_ERROR)
        return
    '''

    ecr_id = General.app.modify_ecr_dialog.GetTitle().split(' ')[-1]

    # track if someone is changing the reason code
    previous_reason_code = cursor.execute("SELECT reason FROM ecrs WHERE id = '{}'".format(ecr_id)).fetchone()[0]
    new_reason_code = ctrl(General.app.modify_ecr_dialog, 'choice:ecr_reason').GetStringSelection()
    if previous_reason_code != new_reason_code:
        sql = "INSERT INTO ecr_reason_code_changes (ecr_id, who_changed, when_changed, previous_code, new_code) VALUES ("
        sql += "{}, ".format(ecr_id)
        sql += "'{}', ".format(General.app.current_user)
        sql += "'{}', ".format(str(dt.datetime.today())[:19])
        sql += "'{}', ".format(previous_reason_code)
        sql += "'{}')".format(new_reason_code)
        cursor.execute(sql)

    reference_number = ctrl(General.app.modify_ecr_dialog, 'text:reference_number').GetValue().replace("'",
                                                                                                       "''").replace(
        '\"', "''''")
    item_number = Database.get_item_from_ref(reference_number)

    sql = 'UPDATE ecrs SET '
    sql += 'type=\'{}\', '.format(General.app.ecr_type)
    sql += 'reference_number=\'{}\', '.format(reference_number)

    if item_number == None:
        sql += 'item=NULL, '
    else:
        sql += 'item=\'{}\', '.format(item_number)

    sql += 'reason=\'{}\', '.format(ctrl(General.app.modify_ecr_dialog, 'choice:ecr_reason').GetStringSelection())
    sql += 'document=\'{}\', '.format(ctrl(General.app.modify_ecr_dialog, 'choice:ecr_document').GetStringSelection())
    sql += 'request=\'{}\', '.format(
        ctrl(General.app.modify_ecr_dialog, 'text:description').GetValue().replace("'", "''").replace('\"', "''''"))
    sql += 'when_needed=\'{} 23:59:00\', '.format(need_by_date)
    sql += 'resolution=\'{}\', '.format(
        ctrl(General.app.modify_ecr_dialog, 'text:resolution').GetValue().replace("'", "''").replace('\"', "''''"))

    if ctrl(General.app.modify_ecr_dialog, 'choice:who_errored').GetStringSelection() != '':
        sql += 'who_errored=\'{}\', '.format(
            ctrl(General.app.modify_ecr_dialog, 'choice:who_errored').GetStringSelection().replace("'", "''").replace(
                '\"', "''''"))
    else:
        sql += 'who_errored=NULL, '

    sql += 'who_modified=\'{}\', '.format(General.app.current_user)
    sql += 'when_modified=\'{}\', '.format(format(str(dt.datetime.today())[:19]))

    # if ctrl(General.app.modify_ecr_dialog, 'label:who_approved_first').GetLabel() != '':
    #	sql += 'who_approved_first=\'{}\', '.format(ctrl(General.app.modify_ecr_dialog, 'label:who_approved_first').GetLabel())

    # if ctrl(General.app.modify_ecr_dialog, 'label:who_approved_second').GetLabel() != '':
    #	sql += 'who_approved_second=\'{}\', '.format(ctrl(General.app.modify_ecr_dialog, 'label:who_approved_second').GetLabel())


    if ctrl(General.app.modify_ecr_dialog, 'label:who_approved_first').GetLabel() == '':
        sql += 'who_approved_first = NULL, '
    else:
        sql += 'who_approved_first=\'{}\', '.format(
            ctrl(General.app.modify_ecr_dialog, 'label:who_approved_first').GetLabel())

    if ctrl(General.app.modify_ecr_dialog, 'label:who_approved_second').GetLabel() == '':
        sql += 'who_approved_second = NULL, '
    else:
        sql += 'who_approved_second=\'{}\', '.format(
            ctrl(General.app.modify_ecr_dialog, 'label:who_approved_second').GetLabel())

    sql += 'approval_stage=\'{}\', '.format(ctrl(General.app.modify_ecr_dialog, 'choice:stage').GetStringSelection())

    component = ctrl(General.app.modify_ecr_dialog, 'choice:ecr_component').GetStringSelection()
    if component != '':
        sql += "component='{}', ".format(component.replace("'", "''").replace('\"', "''''"))
    else:
        sql += "component=NULL, "

    sub_system = ctrl(General.app.modify_ecr_dialog, 'choice:ecr_sub_system').GetStringSelection()
    if sub_system != '':
        sql += "sub_system='{}', ".format(sub_system.replace("'", "''").replace('\"', "''''"))
    else:
        sql += "sub_system=NULL, "

    severity = ctrl(General.app.modify_ecr_dialog, 'choice:ecr_severity').GetStringSelection()
    if severity == 'High':
        sql += "severity=1.0, "
    elif severity == 'Medium':
        sql += "severity=0.5, "
    elif severity == 'Low':
        sql += "severity=0.1, "
    else:
        sql += "severity=1.0, "

    sql += 'priority=\'{}\' '.format(ctrl(General.app.modify_ecr_dialog, 'spin:priority').GetValue())
    sql += 'WHERE id=\'{}\''.format(ecr_id)

    cursor.execute(sql)
    Database.connection.commit()

    cursor.execute("UPDATE ecrs SET attachment = NULL WHERE attachment = ''")
    Database.connection.commit()

    refresh_my_ecrs_list()
    refresh_open_ecrs_list()
    refresh_closed_ecrs_list()
    refresh_my_assigned_ecrs_list()
    refresh_committee_ecrs_list()

    populate_ecr_panel(ecr_id)

    General.app.modify_ecr_dialog.Destroy()


def on_click_add_revisions_with_ecr(event):
    item_number = ctrl(General.app.main_frame, 'label:order_panel_item_number').GetLabel()
    ecr_id = ctrl(General.app.main_frame, 'label:ecr_panel_id').GetLabel()

    if item_number != '':
        ctrl(General.app.main_frame, 'text:revision_items').SetValue(item_number)
        ctrl(General.app.main_frame, 'text:related_ecr').SetValue(ecr_id)
        ctrl(General.app.main_frame, 'notebook:main').SetSelection(1)


def on_text_ecr_description(event):
    if General.app.dbworks_connection != None:
        entry = ctrl(General.app.new_ecr_dialog, 'text:description').GetValue()

        words = entry.replace(' ', '|').replace('.', '|').replace(',', '|').replace('\n', '|').split('|')[:-1]

        for word in words:
            if word not in zip(*General.app.list_of_checked_words_entered)[0]:

                # do a few simple checks to make sure then entry is in item number format before querying the DB
                if len(word) >= 8:
                    if word[-8:][2].isdigit() == False:
                        if word[-8:][1].isdigit() == True:
                            # see if it's a part number in dbworks...
                            # print 'hitting database... for word {}'.format(word)
                            results = General.app.dbworks_cursor.execute(
                                'SELECT TOP 1 DESCRIPTION FROM DOCUMENT WHERE KW_PART_NUMBER=\'{}\''.format(
                                    word[-8:])).fetchone()
                            if results != None:
                                if results[0] != '':
                                    General.app.list_of_checked_words_entered.append((word, results[0]))
                            else:
                                General.app.list_of_checked_words_entered.append((word, None))

        # remove checked words if deleted from text box
        for index, checked_word in enumerate(zip(*General.app.list_of_checked_words_entered)[0]):
            if checked_word not in words:
                General.app.list_of_checked_words_entered[index] = (None, None)

        mentioned_string = ''
        for word_tuple in General.app.list_of_checked_words_entered:
            if word_tuple[1] != None:
                mentioned_string += '{}:   {}\n'.format(word_tuple[0].upper()[-8:], word_tuple[1])

        ctrl(General.app.new_ecr_dialog, 'text:mentioned_parts').SetValue(mentioned_string)


def on_select_assign_ecr(event):
    name = event.GetEventObject().GetStringSelection().replace("'", "''")
    ecr_id = General.app.assign_ecr_dialog.GetTitle().split(' ')[-1]

    cursor = Database.connection.cursor()
    if name != '':
        cursor.execute('''
			UPDATE ecrs SET who_assigned=\'{}\', when_assigned=\'{}\' WHERE id=\'{}\'
			'''.format(name, str(dt.datetime.today())[:19], ecr_id))
    else:
        cursor.execute('''
			UPDATE ecrs SET who_assigned=NULL, when_assigned=NULL WHERE id=\'{}\'
			'''.format(ecr_id))

    Database.connection.commit()

    ecr = cursor.execute(
        'SELECT TOP 1 id, reference_number, request, resolution, who_requested, who_errored, item FROM ecrs WHERE id = \'{}\''.format(
            ecr_id)).fetchone()
    # order = Database.get_order_data_from_ref(ecr[1])
    order = cursor.execute("SELECT TOP 1 * FROM orders WHERE item = \'{}\'".format(ecr[6])).fetchone()

    if name != '':
        # takes a little longer to send an email so put it in a serperate thread so user
        # doesn't have to wait around :)
        sender = cursor.execute(
            'SELECT TOP 1 email FROM employees WHERE name = \'{}\''.format(General.app.current_user)).fetchone()[0]

        # email person who was assigned the ECR
        reciever = cursor.execute('SELECT TOP 1 email FROM employees WHERE name = \'{}\''.format(name)).fetchone()[0]
        Thread(target=send_ecr_assigned_email, args=(ecr, order, reciever, sender)).start()

    refresh_my_ecrs_list()
    refresh_open_ecrs_list()
    refresh_closed_ecrs_list()
    refresh_my_assigned_ecrs_list()
    refresh_committee_ecrs_list()
    populate_ecr_panel(ecr_id)
    General.app.assign_ecr_dialog.Destroy()


def on_select_ecr_item(event):
    item = event.GetEventObject()

    if item.Name == 'list:open_ecrs':
        populate_ecr_order_panel(item_number=item.GetItem(item.GetFirstSelected(), 5).GetText())
    else:
        populate_ecr_order_panel(item_number=item.GetItem(item.GetFirstSelected(), 3).GetText())

    populate_ecr_panel(id=item.GetItem(item.GetFirstSelected(), 0).GetText())

    if item.Name == 'list:my_ecrs':
        if item.GetItem(item.GetFirstSelected(), 0).GetText() == "[more]":
            refresh_my_ecrs_list(limit=(item.GetItemCount() - 1) * 2)

    elif item.Name == 'list:closed_ecrs':
        if item.GetItem(item.GetFirstSelected(), 0).GetText() == "[more]":
            refresh_closed_ecrs_list(limit=(item.GetItemCount() - 1) * 2)

    elif item.Name == 'list:my_assigned_ecrs':
        if item.GetItem(item.GetFirstSelected(), 0).GetText() == "[more]":
            refresh_my_assigned_ecrs_list(limit=(item.GetItemCount() - 1) * 2)

    elif item.Name == 'list:committee_ecrs':
        if item.GetItem(item.GetFirstSelected(), 0).GetText() == "[more]":
            refresh_committee_ecrs_list(limit=(item.GetItemCount() - 1) * 2)


def on_select_ecr_reason(event):
    # change the lead time based on what reason was selected
    selection = ctrl(General.app.new_ecr_dialog, 'choice:ecr_reason').GetStringSelection()

    cursor = Database.connection.cursor()
    # cursor.execute("SELECT lead_time FROM ecr_reason_choices WHERE reason=\'{}\'".format(selection))
    cursor.execute("SELECT lead_time FROM secondary_ecr_reason_codes WHERE code=\'{}\'".format(selection))

    lead_time_date = dt.datetime.today() + dt.timedelta(cursor.fetchone()[0])
    ctrl(General.app.new_ecr_dialog, 'calendar:ecr_need_by').SetDate(
        wx.DateTimeFromDMY(lead_time_date.day, lead_time_date.month - 1, lead_time_date.year))


def on_choice_set_severity_default(event):
    dialog = wx.GetTopLevelParent(event.GetEventObject())

    discipline = General.app.ecr_type
    component = ctrl(dialog, 'choice:ecr_component').GetStringSelection()
    document = ctrl(dialog, 'choice:ecr_document').GetStringSelection()

    cursor = Database.connection.cursor()
    cursor.execute(
        "SELECT severity_default FROM ecr.severity_defaults WHERE discipline='{}' and component='{}' and document='{}'".format(
            discipline, component, document))

    try:
        severity_default = float(cursor.fetchone()[0])

        if severity_default == 1.0:
            ctrl(dialog, 'choice:ecr_severity').SetStringSelection('High')
        elif severity_default == 0.5:
            ctrl(dialog, 'choice:ecr_severity').SetStringSelection('Medium')
        elif severity_default == 0.1:
            ctrl(dialog, 'choice:ecr_severity').SetStringSelection('Low')
        else:
            ctrl(dialog, 'choice:ecr_severity').SetStringSelection('High')

    except:
        ctrl(dialog, 'choice:ecr_severity').SetStringSelection('High')


'''
def populate_ecr_fields_from_list(event, tab):
	item = event.GetEventObject()	
	ctrl(General.app.main_frame, 'text:'+tab+'_description').SetValue(item.GetItem(item.GetFirstSelected(), 4).GetText())
	ctrl(General.app.main_frame, 'text:'+tab+'_resolution').SetValue(item.GetItem(item.GetFirstSelected(), 5).GetText())

	if item.GetItem(item.GetFirstSelected(), 0).GetText() == "(More)":
		print item.GetItem(item.GetFirstSelected(), 1).GetText()
		refresh_my_ecrs_list(limit=int(item.GetItem(item.GetFirstSelected(), 1).GetText()))
'''


def populate_ecr_order_panel(item_number):
    cursor = Database.connection.cursor()

    # item_number =
    # cursor.execute("SELECT * FROM orders WHERE item=\'{}\'".format(item_number))
    # query_result = Database.get_order_data_from_ref(ref_number)
    # query_result = cursor.execute("SELECT TOP 1 * FROM orders WHERE item = \'{}\'".format(item_number)).fetchone()

    query_result = cursor.execute('''
		SELECT TOP 1
			orders.item,
			orders.sales_order,
			orders.line_up,
			orders.serial,
			orders.customer,
			orders.mechanical_by,
			orders.electrical_by,
			orders.structural_by,
			item_responsibilities2.mechanical_cad_designer,
			item_responsibilities2.electrical_cad_designer,
			item_responsibilities2.structural_cad_designer,
			orders.model
		FROM orders LEFT JOIN item_responsibilities2 ON orders.item = item_responsibilities2.item WHERE
			orders.item = '{}'	
		'''.format(item_number)).fetchone()

    # replace NULLs in results with blank string
    if query_result != None:
        query_result = ['' if x == None else x for x in query_result]

    if query_result != None:
        item, sales_order, line_up, serial, customer, \
        mechanical_by, electrical_by, structural_by, \
        mechanical_cad_designer, electrical_cad_designer, structural_cad_designer, model = query_result

        if mechanical_cad_designer == '':
            real_mechanical_by = mechanical_by
        else:
            real_mechanical_by = mechanical_cad_designer

        if electrical_cad_designer == '':
            real_electrical_by = electrical_by
        else:
            real_electrical_by = electrical_cad_designer

        if structural_cad_designer == '':
            real_structural_by = structural_by
        else:
            real_structural_by = structural_cad_designer

        ctrl(General.app.main_frame, 'label:order_panel_item_number').SetLabel(str(item))
        ctrl(General.app.main_frame, 'label:order_panel_sales_order').SetLabel(str(sales_order) + '-' + str(line_up))
        ctrl(General.app.main_frame, 'label:order_panel_serial').SetLabel(str(serial))
        ctrl(General.app.main_frame, 'label:order_panel_customer').SetLabel(str(customer))
        ctrl(General.app.main_frame, 'label:order_panel_mechanical').SetLabel(str(real_mechanical_by))
        ctrl(General.app.main_frame, 'label:order_panel_electrical').SetLabel(str(real_electrical_by))
        ctrl(General.app.main_frame, 'label:order_panel_structural').SetLabel(str(real_structural_by))
        ctrl(General.app.main_frame, 'label:order_panel_model').SetLabel(str(model))
    else:
        ctrl(General.app.main_frame, 'label:order_panel_item_number').SetLabel('')
        ctrl(General.app.main_frame, 'label:order_panel_sales_order').SetLabel('')
        ctrl(General.app.main_frame, 'label:order_panel_serial').SetLabel('')
        ctrl(General.app.main_frame, 'label:order_panel_customer').SetLabel('')
        ctrl(General.app.main_frame, 'label:order_panel_mechanical').SetLabel('')
        ctrl(General.app.main_frame, 'label:order_panel_electrical').SetLabel('')
        ctrl(General.app.main_frame, 'label:order_panel_structural').SetLabel('')
        ctrl(General.app.main_frame, 'label:order_panel_model').SetLabel('')


def populate_ecr_panel(id):
    if id != None and id != '[more]':
        cursor = Database.connection.cursor()
        cursor.execute(
            "SELECT id, reference_number, type, reason, document, who_claimed, who_assigned, request, resolution, attachment, status FROM ecrs WHERE id=\'{}\'".format(
                id))
        query_result = cursor.fetchone()
    else:
        query_result = None

    # replace NULLs in results with blank string
    if query_result != None:
        query_result = ['' if x == None else x for x in query_result]

    if query_result != None:
        ctrl(General.app.main_frame, 'label:ecr_panel_id').SetLabel(str(query_result[0]))
        ctrl(General.app.main_frame, 'label:ecr_panel_reference_number').SetLabel(str(query_result[1]))
        ctrl(General.app.main_frame, 'label:ecr_panel_change_is').SetLabel(str(query_result[2]))
        ctrl(General.app.main_frame, 'label:ecr_panel_reason').SetLabel(str(query_result[3]))
        ctrl(General.app.main_frame, 'label:ecr_panel_document').SetLabel(str(query_result[4]))
        ctrl(General.app.main_frame, 'label:ecr_panel_claimed_by').SetLabel(str(query_result[5]))
        ctrl(General.app.main_frame, 'label:ecr_panel_assigned_to').SetLabel(str(query_result[6]))
        ctrl(General.app.main_frame, 'text:ecr_panel_description').SetValue(str(query_result[7]))
        ctrl(General.app.main_frame, 'text:ecr_panel_resolution').SetValue(str(query_result[8]))

        # if attachments in ecr
        if query_result[9] != '':
            print query_result[9]
            ctrl(General.app.main_frame, 'button:open_attachment').Show()
            ctrl(General.app.main_frame, 'button:open_attachment').GetParent().Layout()
        else:
            ctrl(General.app.main_frame, 'button:open_attachment').Hide()

        # if already closed
        if query_result[10] == 'Closed':
            ctrl(General.app.main_frame, 'button:claim').Disable()
            ctrl(General.app.main_frame, 'button:assign').Disable()
            ctrl(General.app.main_frame, 'button:close').Disable()
        else:
            ctrl(General.app.main_frame, 'button:claim').Enable()
            ctrl(General.app.main_frame, 'button:assign').Enable()
            ctrl(General.app.main_frame, 'button:close').Enable()

        # if already claimed by this user, disable
        if query_result[5] == General.app.current_user:
            ctrl(General.app.main_frame, 'button:claim').Disable()


    else:
        ctrl(General.app.main_frame, 'label:ecr_panel_id').SetLabel('')
        ctrl(General.app.main_frame, 'label:ecr_panel_reference_number').SetLabel('')
        ctrl(General.app.main_frame, 'label:ecr_panel_change_is').SetLabel('')
        ctrl(General.app.main_frame, 'label:ecr_panel_reason').SetLabel('')
        ctrl(General.app.main_frame, 'label:ecr_panel_document').SetLabel('')
        ctrl(General.app.main_frame, 'label:ecr_panel_claimed_by').SetLabel('')
        ctrl(General.app.main_frame, 'label:ecr_panel_assigned_to').SetLabel('')
        ctrl(General.app.main_frame, 'text:ecr_panel_description').SetValue('')
        ctrl(General.app.main_frame, 'text:ecr_panel_resolution').SetValue('')

        ctrl(General.app.main_frame, 'button:open_attachment').Hide()


def radio_button_selected(event, type):
    if General.app.new_ecr_dialog:
        # add document options to choice box based on ecr type selected
        cursor = Database.connection.cursor()
        cursor.execute("SELECT document FROM ecr_document_choices WHERE type=\'{}\' OR type=\'*\'".format(type))
        ctrl(General.app.new_ecr_dialog, 'choice:ecr_document').Clear()
        ctrl(General.app.new_ecr_dialog, 'choice:ecr_document').AppendItems(zip(*cursor.fetchall())[0])
    General.app.ecr_type = type


# print 'good!'



def refresh_closed_ecrs_list(event=None, limit=15):
    closed_ecr_list = ctrl(General.app.main_frame, 'list:closed_ecrs')

    # clear out the list
    closed_ecr_list.DeleteAllItems()
    set_list_headers(closed_ecr_list, get_cleaned_list_headers(closed_ecr_list))

    column_names = Database.get_table_column_names('ecrs', presentable=True)

    if closed_ecr_list.GetColumn(0) == None:
        for index, column_name in enumerate(column_names):
            if column_name == 'Reference Number':
                closed_ecr_list.InsertColumn(index, 'ReferenceNo')
            elif column_name == 'Document':
                # kinda ghetto but replace the field where document would be with sales order
                closed_ecr_list.InsertColumn(index, 'Sales Order')
            else:
                closed_ecr_list.InsertColumn(index, column_name)

    # query the database
    cursor = Database.connection.cursor()

    user_department = cursor.execute(
        'SELECT TOP 1 department FROM employees WHERE name = \'{}\''.format(General.app.current_user)).fetchone()[0]

    cursor.execute("SELECT TOP {} * FROM ecrs WHERE status = \'Closed\' ORDER BY when_closed DESC".format(limit))
    records = cursor.fetchall()

    for ecr_index, ecr in enumerate(records):
        closed_ecr_list.InsertStringItem(sys.maxint, '#')

        # get order data from reference number
        sales_order = None
        reference_number = ecr[2]
        if reference_number != None:
            # order_data = Database.get_order_data_from_ref(reference_number)
            order_data = cursor.execute("SELECT TOP 1 * FROM orders WHERE item = \'{}\'".format(ecr[3])).fetchone()

            if order_data != None:
                sales_order = str(order_data[1]) + '-' + str(order_data[2])

        for column_index, column_value in enumerate(ecr):
            try:
                if column_names[column_index] == 'When Requested':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Needed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)[:8]

                if column_names[column_index] == 'When Closed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Modified':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Claimed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Assigned':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'Document':
                    column_value = sales_order

                if column_value != None:
                    closed_ecr_list.SetStringItem(ecr_index, column_index, str(column_value).replace('\n', ' \\ '))
            except:
                print "### Error adding record to list:"
                print column_value
                print sys.exc_info()

    # last row allows user to load more records
    if len(records) == limit:
        closed_ecr_list.InsertStringItem(sys.maxint, ' ')
        closed_ecr_list.SetStringItem(limit, 0, "[more]")
    # closed_ecr_list.SetStringItem(i+1, 1, str(limit*2))

    if user_department == 'Design Engineering':
        columns_to_hide = ['item_number', 'status', 'type']
    else:
        columns_to_hide = ['item_number', 'status', 'reason', 'department', 'type', 'when_modified', 'who_errored',
                           'who_claimed', 'when_claimed', 'who_assigned', 'when_assigned', 'who_modified']

    for column_index, column_name in enumerate(column_names):
        if column_name.lower().replace(' ', '_') in columns_to_hide:
            closed_ecr_list.SetColumnWidth(column_index, 0)
        else:
            if column_name.lower().replace(' ', '_') == 'request':
                closed_ecr_list.SetColumnWidth(column_index, 400)
            elif column_name.lower().replace(' ', '_') == 'resolution':
                closed_ecr_list.SetColumnWidth(column_index, 400)
            else:
                closed_ecr_list.SetColumnWidth(column_index, wx.LIST_AUTOSIZE_USEHEADER)

    '''
    #create columns for the list
    if closed_ecr_list.GetColumn(0) == None:
        closed_ecr_list.InsertColumn(0, 'ID')
        closed_ecr_list.InsertColumn(1, 'ReferenceNo')
        closed_ecr_list.InsertColumn(2, 'Sales Order')
        closed_ecr_list.InsertColumn(3, 'Request')
        closed_ecr_list.InsertColumn(4, 'Resolution')
        closed_ecr_list.InsertColumn(5, 'When Requested')
        closed_ecr_list.InsertColumn(6, 'When Closed')

    #query the database
    cursor = Database.connection.cursor()
    cursor.execute("SELECT * FROM ecrs WHERE status = \'Closed\' ORDER BY when_closed DESC LIMIT {}".format(limit))
    records = cursor.fetchall()

    for i in range(len(records)):
        try:
            row = records[i]

            #get order data from reference number
            sales_order = None
            reference_number = row[2]
            if reference_number != None:
                order_data = Database.get_order_data_from_ref(reference_number)

                if order_data != None:
                    sales_order = str(order_data[1]) +'-'+ str(order_data[2])

            #populate row fields
            closed_ecr_list.InsertStringItem(sys.maxint, '#')
            closed_ecr_list.SetStringItem(i, 0, str(row[0]))	#id
            if reference_number != None:
                closed_ecr_list.SetStringItem(i, 1, reference_number)
            if sales_order != None:
                closed_ecr_list.SetStringItem(i, 2, sales_order)
            closed_ecr_list.SetStringItem(i, 3, str(row[8]))
            if row[9] != None:
                closed_ecr_list.SetStringItem(i, 4, str(row[9]))

            #requested date
            dt_object = time.strptime(str(row[11]), "%Y-%m-%d %H:%M:%S") #to python time object
            closed_ecr_list.SetStringItem(i, 5, time.strftime("%m/%d/%y   %I:%M %p", dt_object))

            #closed date
            if row[13] != None:
                dt_object = time.strptime(str(row[13]), "%Y-%m-%d %H:%M:%S") #to python time object
                closed_ecr_list.SetStringItem(i, 6, time.strftime("%m/%d/%y   %I:%M %p", dt_object))

            #color the row based on it's status
            if str(row[1]) == 'Open':
                closed_ecr_list.SetItemBackgroundColour(i, '#FFF0B7')

        except:
            print "### Error adding record to list:"
            print sys.exc_info()

    #last row allows user to load more records
    if len(records) == limit:
        closed_ecr_list.InsertStringItem(sys.maxint, ' ')
        closed_ecr_list.SetStringItem(i+1, 0, "[more]")
        #closed_ecr_list.SetStringItem(i+1, 1, str(limit*2))

    #set some columns to a good width
    closed_ecr_list.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
    closed_ecr_list.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
    closed_ecr_list.SetColumnWidth(2, wx.LIST_AUTOSIZE_USEHEADER)
    closed_ecr_list.SetColumnWidth(3, wx.LIST_AUTOSIZE_USEHEADER) #or 400?
    closed_ecr_list.SetColumnWidth(4, wx.LIST_AUTOSIZE_USEHEADER)
    closed_ecr_list.SetColumnWidth(5, wx.LIST_AUTOSIZE_USEHEADER)
    closed_ecr_list.SetColumnWidth(6, wx.LIST_AUTOSIZE_USEHEADER)
    '''

    # change label saying when last refreshed
    now = time.localtime()
    ctrl(General.app.main_frame, 'label:closed_ecrs_last_updated').SetLabel(
        'List last updated {}'.format(time.strftime("%I:%M %p", now).strip('0')))


def refresh_my_ecrs_list(event=None, limit=15):
    my_ecr_list = ctrl(General.app.main_frame, 'list:my_ecrs')

    # clear out the list
    my_ecr_list.DeleteAllItems()
    set_list_headers(my_ecr_list, get_cleaned_list_headers(my_ecr_list))

    column_names = Database.get_table_column_names('ecrs', presentable=True)

    if my_ecr_list.GetColumn(0) == None:
        for index, column_name in enumerate(column_names):
            if column_name == 'Reference Number':
                my_ecr_list.InsertColumn(index, 'ReferenceNo')
            elif column_name == 'Document':
                # kinda ghetto but replace the field where document would be with sales order
                my_ecr_list.InsertColumn(index, 'Sales Order')
            else:
                my_ecr_list.InsertColumn(index, column_name)

    # query the database
    cursor = Database.connection.cursor()

    user_department = cursor.execute(
        "SELECT TOP 1 department FROM employees WHERE name = \'{}\'".format(General.app.current_user)).fetchone()[0]

    cursor.execute("SELECT TOP {} * FROM ecrs WHERE who_requested = \'{}\' ORDER BY when_requested DESC".format(limit,
                                                                                                                General.app.current_user))
    records = cursor.fetchall()

    for ecr_index, ecr in enumerate(records):
        my_ecr_list.InsertStringItem(sys.maxint, '#')

        if str(ecr[1]) == 'Open':
            my_ecr_list.SetItemBackgroundColour(ecr_index, '#FFF0B7')

        # get order data from reference number
        sales_order = None
        reference_number = ecr[2]
        if reference_number != None:
            # order_data = Database.get_order_data_from_ref(reference_number)
            order_data = cursor.execute("SELECT TOP 1 * FROM orders WHERE item = \'{}\'".format(ecr[3])).fetchone()

            if order_data != None:
                sales_order = str(order_data[1]) + '-' + str(order_data[2])

        for column_index, column_value in enumerate(ecr):
            try:
                if column_names[column_index] == 'When Requested':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Needed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)[:8]

                if column_names[column_index] == 'When Closed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Modified':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Claimed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Assigned':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'Document':
                    column_value = sales_order

                if column_value != None:
                    my_ecr_list.SetStringItem(ecr_index, column_index, str(column_value).replace('\n', ' \\ '))
            except:
                print "### Error adding record to list:"
                print '---', ecr
                print sys.exc_info()

    # last row allows user to load more records
    if len(records) == limit:
        my_ecr_list.InsertStringItem(sys.maxint, ' ')
        my_ecr_list.SetStringItem(limit, 0, "[more]")
    # my_ecr_list.SetStringItem(i+1, 1, str(limit*2))

    if user_department == 'Design Engineering':
        columns_to_hide = ['item_number', 'type']
    else:
        columns_to_hide = ['item_number', 'reason', 'department', 'who_requested', 'type', 'attachment',
                           'when_modified', 'who_errored', 'who_claimed', 'when_claimed', 'who_assigned',
                           'when_assigned', 'who_modified']

    '''
    #set the Id column header to a name as long as an ECR number because although the code is supposed
    # to autofit to the rows longest width, it apparently doesn't make it wide enough on some user's computers...
    # namely Bradshaws.
    header = my_ecr_list.GetColumn(0) 
    header.SetText('1022118') 
    my_ecr_list.SetColumn(0, header)
    '''

    for column_index, column_name in enumerate(column_names):
        if column_name.lower().replace(' ', '_') in columns_to_hide:
            my_ecr_list.SetColumnWidth(column_index, 0)
        else:
            if column_name.lower().replace(' ', '_') == 'request':
                my_ecr_list.SetColumnWidth(column_index, 400)
            elif column_name.lower().replace(' ', '_') == 'resolution':
                my_ecr_list.SetColumnWidth(column_index, 400)
            else:
                my_ecr_list.SetColumnWidth(column_index, wx.LIST_AUTOSIZE_USEHEADER)

    '''
    #return the Id column header value back to the name it should be.
    header = my_ecr_list.GetColumn(0) 
    header.SetText('Id')
    my_ecr_list.SetColumn(0, header)
    '''

    # change label saying when last refreshed
    now = time.localtime()
    ctrl(General.app.main_frame, 'label:my_ecrs_last_updated').SetLabel(
        'List last updated {}'.format(time.strftime("%I:%M %p", now).strip('0')))


def refresh_my_assigned_ecrs_list(event=None, limit=15):
    my_assigned_ecr_list = ctrl(General.app.main_frame, 'list:my_assigned_ecrs')

    # clear out the list
    my_assigned_ecr_list.DeleteAllItems()
    set_list_headers(my_assigned_ecr_list, get_cleaned_list_headers(my_assigned_ecr_list))

    column_names = Database.get_table_column_names('ecrs', presentable=True)

    if my_assigned_ecr_list.GetColumn(0) == None:
        for index, column_name in enumerate(column_names):
            if column_name == 'Reference Number':
                my_assigned_ecr_list.InsertColumn(index, 'ReferenceNo')
            elif column_name == 'Document':
                # kinda ghetto but replace the field where document would be with sales order
                my_assigned_ecr_list.InsertColumn(index, 'Sales Order')
            else:
                my_assigned_ecr_list.InsertColumn(index, column_name)

    # query the database
    cursor = Database.connection.cursor()

    user_department = cursor.execute(
        "SELECT TOP 1 department FROM employees WHERE name = \'{}\'".format(General.app.current_user)).fetchone()[0]

    cursor.execute(
        "SELECT TOP {} * FROM ecrs WHERE who_assigned = \'{}\' AND status='Open' ORDER BY when_requested DESC".format(
            limit, General.app.current_user))
    records = cursor.fetchall()

    for ecr_index, ecr in enumerate(records):
        my_assigned_ecr_list.InsertStringItem(sys.maxint, '#')

        if str(ecr[1]) == 'Open':
            my_assigned_ecr_list.SetItemBackgroundColour(ecr_index, '#FFF0B7')

        # get order data from reference number
        sales_order = None
        reference_number = ecr[2]
        if reference_number != None:
            # order_data = Database.get_order_data_from_ref(reference_number)
            order_data = cursor.execute("SELECT TOP 1 * FROM orders WHERE item = \'{}\'".format(ecr[3])).fetchone()

            if order_data != None:
                sales_order = str(order_data[1]) + '-' + str(order_data[2])

        for column_index, column_value in enumerate(ecr):
            try:
                if column_names[column_index] == 'When Requested':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Needed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)[:8]

                if column_names[column_index] == 'When Closed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Modified':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Claimed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Assigned':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'Document':
                    column_value = sales_order

                if column_value != None:
                    my_assigned_ecr_list.SetStringItem(ecr_index, column_index, str(column_value).replace('\n', ' \\ '))
            except:
                print "### Error adding record to list:"
                print '---', ecr
                print sys.exc_info()

    # last row allows user to load more records
    if len(records) == limit:
        my_assigned_ecr_list.InsertStringItem(sys.maxint, ' ')
        my_assigned_ecr_list.SetStringItem(limit, 0, "[more]")
    # my_assigned_ecr_list.SetStringItem(i+1, 1, str(limit*2))

    if user_department == 'Design Engineering':
        columns_to_hide = ['item_number', 'type']
    else:
        columns_to_hide = ['item_number', 'reason', 'department', 'who_requested', 'type', 'attachment',
                           'when_modified', 'who_errored', 'who_claimed', 'when_claimed', 'who_assigned',
                           'when_assigned', 'who_modified']

    for column_index, column_name in enumerate(column_names):
        if column_name.lower().replace(' ', '_') in columns_to_hide:
            my_assigned_ecr_list.SetColumnWidth(column_index, 0)
        else:
            if column_name.lower().replace(' ', '_') == 'request':
                my_assigned_ecr_list.SetColumnWidth(column_index, 400)
            elif column_name.lower().replace(' ', '_') == 'resolution':
                my_assigned_ecr_list.SetColumnWidth(column_index, 400)
            else:
                my_assigned_ecr_list.SetColumnWidth(column_index, wx.LIST_AUTOSIZE_USEHEADER)

    # change label saying when last refreshed
    now = time.localtime()
    ctrl(General.app.main_frame, 'label:my_assigned_ecrs_last_updated').SetLabel(
        'List last updated {}'.format(time.strftime("%I:%M %p", now).strip('0')))


def refresh_open_ecrs_list(event=None, limit=100):
    open_ecr_list = ctrl(General.app.main_frame, 'list:open_ecrs')

    # clear out the list
    open_ecr_list.DeleteAllItems()
    set_list_headers(open_ecr_list, get_cleaned_list_headers(open_ecr_list))

    column_names = ["Id", "Who Claimed", "Who Assigned", "Sales Order", "Item", "Production Order",
                    "Request", "Who Requested", "When Needed", "When Requested", "ReferenceNo",
                    "Reason", "Department", "Type", "Who Modified", "When Modified",
                    "When Claimed", "When Assigned", "Attachment", "Resolution", ]

    if open_ecr_list.GetColumn(0) == None:
        for index, column_name in enumerate(column_names):
            open_ecr_list.InsertColumn(index, column_name)

    # query the database
    cursor = Database.connection.cursor()

    user_department = cursor.execute(
        'SELECT TOP 1 department FROM employees WHERE name = \'{}\''.format(General.app.current_user)).fetchone()[0]

    cursor.execute('''
		SELECT
			ecrs.id,
			ecrs.who_claimed,
			ecrs.who_assigned,
			orders.sales_order,
			orders.line_up,
			ecrs.item,
			ecrs.request,
			ecrs.who_requested,
			ecrs.when_needed,
			ecrs.when_requested,
			ecrs.reference_number,
			ecrs.reason,
			ecrs.department,
			ecrs.type,
			ecrs.who_modified,
			ecrs.when_modified,
			ecrs.when_claimed,
			ecrs.when_assigned,
			ecrs.attachment,
			ecrs.resolution,
			ecrs.approval_stage
		FROM 
			dbo.ecrs
		LEFT JOIN      
			dbo.orders ON ecrs.item = orders.item
		WHERE 
			ecrs.status = 'Open'
		ORDER BY 
			orders.sales_order ASC, orders.line_up ASC''')

    records = cursor.fetchall()

    # purge out ECRs that need approval and don't yet have it...
    new_records = []
    for ecr in records:
        reason = ecr[11]
        approval_stage = ecr[20]

        if reason not in reasons_needing_approval:
            new_records.append(ecr)

        else:
            if approval_stage != 'New Request, needs reviewing':
                new_records.append(ecr)

    records = new_records

    ecrs_open = 0
    ecrs_late = 0
    ecrs_due_today = 0

    for ecr_index, ecr in enumerate(records):
        open_ecr_list.InsertStringItem(sys.maxint, '#')

        ecrs_open += 1

        when_needed_dt = dt.datetime.strptime(str(ecr[column_names.index('When Needed')]),
                                              "%Y-%m-%d %H:%M:%S")  # to python time object
        # color ECRs red if late
        if when_needed_dt < dt.datetime.today():
            open_ecr_list.SetItemBackgroundColour(ecr_index, '#FF9999')
            ecrs_late += 1

        # color them orange if due today
        if (when_needed_dt.month == dt.datetime.today().month) and (when_needed_dt.day == dt.datetime.today().day):
            open_ecr_list.SetItemBackgroundColour(ecr_index, '#FFF0B7')
            ecrs_due_today += 1

        formatted_record = []
        for field in ecr:
            if field == None:
                field = ''

            elif isinstance(field, dt.datetime):
                field = field.strftime('%m/%d/%y %I:%M %p')

            else:
                pass

            formatted_record.append(field)

        for column_index, column_value in enumerate(formatted_record):
            try:
                if column_names[column_index] == 'When Needed':
                    column_value = column_value[:8]

                if column_value != None:
                    open_ecr_list.SetStringItem(ecr_index, column_index, str(column_value).replace('\n', ' \\ '))
            except:
                print "### Error adding record to list:"
                print 'yoyoy: ', column_value
                print sys.exc_info()

    # last row allows user to load more records
    if len(records) == limit:
        open_ecr_list.InsertStringItem(sys.maxint, ' ')
        open_ecr_list.SetStringItem(limit, 0, "[more]")
    # open_ecr_list.SetStringItem(i+1, 1, str(limit*2))

    '''
    if user_department == 'Design Engineering':
        columns_to_hide = ['item_number', 'status', 'type']
    else:
        columns_to_hide = ['item_number', 'status', 'reason', 'department', 'type', 'attachment', 'when_modified', 'when_closed',  'who_errored', 'who_claimed', 'when_claimed', 'who_assigned', 'when_assigned', 'who_modified']
    '''

    for column_index, column_name in enumerate(column_names):
        if column_name == 'Request':
            open_ecr_list.SetColumnWidth(column_index, 400)
        elif column_name == 'Resolution':
            open_ecr_list.SetColumnWidth(column_index, 300)
        elif column_name == 'Attachment':
            open_ecr_list.SetColumnWidth(column_index, 200)
        else:
            open_ecr_list.SetColumnWidth(column_index, wx.LIST_AUTOSIZE_USEHEADER)

    # change label saying when last refreshed
    now = time.localtime()
    ctrl(General.app.main_frame, 'label:open_ecrs_last_updated').SetLabel(
        'List last updated {}'.format(time.strftime("%I:%M %p", now).strip('0')))

    # say how many we have open, late and due
    ctrl(General.app.main_frame, 'label:ecr_stats').SetLabel(
        '{} open, {} late, {} due today'.format(ecrs_open, ecrs_late, ecrs_due_today))


def refresh_committee_ecrs_list(event=None, limit=100):
    committee_ecr_list = ctrl(General.app.main_frame, 'list:committee_ecrs')

    # clear out the list
    committee_ecr_list.DeleteAllItems()
    set_list_headers(committee_ecr_list, get_cleaned_list_headers(committee_ecr_list))

    column_names = Database.get_table_column_names('ecrs', presentable=True)

    if committee_ecr_list.GetColumn(0) == None:
        for index, column_name in enumerate(column_names):
            if column_name == 'Reference Number':
                committee_ecr_list.InsertColumn(index, 'ReferenceNo')
            elif column_name == 'Document':
                # kinda ghetto but replace the field where document would be with sales order
                committee_ecr_list.InsertColumn(index, 'Sales Order')
            else:
                committee_ecr_list.InsertColumn(index, column_name)

    # query the database
    cursor = Database.connection.cursor()

    # user_department = cursor.execute("SELECT TOP 1 department FROM employees WHERE name = \'{}\'".format(General.app.current_user)).fetchone()[0]

    cursor.execute(
        "SELECT TOP {} * FROM ecrs WHERE status='Open' AND approval_stage='New Request, needs reviewing' ORDER BY when_requested DESC".format(
            limit, General.app.current_user))
    records = cursor.fetchall()

    new_records = []
    for ecr_index, ecr in enumerate(records):
        reason_code = ecr[5]
        if reason_code in reasons_needing_approval:
            new_records.append(ecr)

    records = new_records

    for ecr_index, ecr in enumerate(records):

        committee_ecr_list.InsertStringItem(sys.maxint, '#')

        # if str(ecr[1]) == 'Open':
        #	committee_ecr_list.SetItemBackgroundColour(ecr_index, '#FFF0B7')

        # get order data from reference number
        sales_order = None
        reference_number = ecr[2]
        if reference_number != None:
            # order_data = Database.get_order_data_from_ref(reference_number)
            order_data = cursor.execute("SELECT TOP 1 * FROM orders WHERE item = \'{}\'".format(ecr[3])).fetchone()

            if order_data != None:
                sales_order = str(order_data[1]) + '-' + str(order_data[2])

        for column_index, column_value in enumerate(ecr):
            try:
                if column_names[column_index] == 'When Requested':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Needed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)[:8]

                if column_names[column_index] == 'When Closed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Modified':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Claimed':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'When Assigned':
                    if column_value != None: column_value = General.format_date_nicely(column_value)

                if column_names[column_index] == 'Document':
                    column_value = sales_order

                if column_value != None:
                    committee_ecr_list.SetStringItem(ecr_index, column_index, str(column_value).replace('\n', ' \\ '))
            except:
                print "### Error adding record to list:"
                print '---', ecr
                print sys.exc_info()

    # last row allows user to load more records
    if len(records) == limit:
        committee_ecr_list.InsertStringItem(sys.maxint, ' ')
        committee_ecr_list.SetStringItem(limit, 0, "[more]")
    # committee_ecr_list.SetStringItem(i+1, 1, str(limit*2))

    '''
    if user_department == 'Design Engineering':
        columns_to_hide = ['item_number', 'type']
    else:
        columns_to_hide = ['item_number', 'reason', 'department', 'who_requested', 'type', 'attachment', 'when_modified', 'who_errored', 'who_claimed', 'when_claimed', 'who_assigned', 'when_assigned', 'who_modified']
        '''
    columns_to_hide = []

    for column_index, column_name in enumerate(column_names):
        if column_name.lower().replace(' ', '_') in columns_to_hide:
            committee_ecr_list.SetColumnWidth(column_index, 0)
        else:
            if column_name.lower().replace(' ', '_') == 'request':
                committee_ecr_list.SetColumnWidth(column_index, 400)
            elif column_name.lower().replace(' ', '_') == 'resolution':
                committee_ecr_list.SetColumnWidth(column_index, 400)
            else:
                committee_ecr_list.SetColumnWidth(column_index, wx.LIST_AUTOSIZE_USEHEADER)

    '''
    #return the Id column header value back to the name it should be.
    header = committee_ecr_list.GetColumn(0) 
    header.SetText('Id')
    committee_ecr_list.SetColumn(0, header)
    '''

    # change label saying when last refreshed
    now = time.localtime()
    ctrl(General.app.main_frame, 'label:committee_ecrs_last_updated').SetLabel(
        'List last updated {}'.format(time.strftime("%I:%M %p", now).strip('0')))


def reset_search_fields():
    for i in range(11):
        ctrl(General.app.search_ecrs_dialog, 'choice:search_condition' + str(i)).SetStringSelection(' ')
        ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(i)).SetValue('')

    # set 'from' dates to a low value...
    ctrl(General.app.search_ecrs_dialog, 'date_picker:requested_from').SetValue(wx.DateTimeFromDMY(1, 0, 1900))
    ctrl(General.app.search_ecrs_dialog, 'date_picker:needed_from').SetValue(wx.DateTimeFromDMY(1, 0, 1900))
    ctrl(General.app.search_ecrs_dialog, 'date_picker:closed_from').SetValue(wx.DateTimeFromDMY(1, 0, 1900))
    ctrl(General.app.search_ecrs_dialog, 'date_picker:modified_from').SetValue(wx.DateTimeFromDMY(1, 0, 1900))

    # set 'to' dates to a year ahead
    date_list = ctrl(General.app.search_ecrs_dialog, 'date_picker:requested_to').GetValue().Format("%d-%m-%Y").split(
        '-')
    date_wx_format = wx.DateTimeFromDMY(int(date_list[0]), int(date_list[1]) - 1, int(date_list[2]) + 1)
    ctrl(General.app.search_ecrs_dialog, 'date_picker:requested_to').SetValue(date_wx_format)
    ctrl(General.app.search_ecrs_dialog, 'date_picker:needed_to').SetValue(date_wx_format)
    ctrl(General.app.search_ecrs_dialog, 'date_picker:closed_to').SetValue(date_wx_format)
    ctrl(General.app.search_ecrs_dialog, 'date_picker:modified_to').SetValue(date_wx_format)

    ctrl(General.app.search_ecrs_dialog, 'choice:sort_by').SetStringSelection('When requested')
    ctrl(General.app.search_ecrs_dialog, 'choice:sort_in').SetSelection(0)
    ctrl(General.app.search_ecrs_dialog, 'choice:limit').SetStringSelection('25 records')


def search_condition_selected(event, index):
    # clear out the search field if no condition selected
    if ctrl(General.app.search_ecrs_dialog, 'choice:search_condition' + str(index)).GetSelection() == 0:
        ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(index)).SetValue('')


def search_ecrs(event):
    event.GetEventObject().SetLabel('Searching...')
    results_list = ctrl(General.app.main_frame, 'list:results')

    # clear out the list
    results_list.DeleteAllItems()

    column_names = Database.get_table_column_names('ecrs', presentable=True)

    # create columns for the list
    if results_list.GetColumn(0) == None:
        for index, column_name in enumerate(column_names):
            results_list.InsertColumn(index, column_name)

        '''
        results_list.InsertColumn(0, 'ID')
        results_list.InsertColumn(1, 'Status')
        results_list.InsertColumn(2, 'ReferenceNo')
        results_list.InsertColumn(3, 'Document')
        results_list.InsertColumn(4, 'Reason')
        results_list.InsertColumn(5, 'Department')
        results_list.InsertColumn(6, 'who_requested')
        results_list.InsertColumn(7, 'Type')
        results_list.InsertColumn(8, 'Description')
        results_list.InsertColumn(9, 'Resolution')
        results_list.InsertColumn(10, 'When Requested')
        results_list.InsertColumn(11, 'When Needed')
        results_list.InsertColumn(12, 'When Closed')
        results_list.InsertColumn(13, 'When Modified')
        results_list.InsertColumn(14, 'Who Errored')
        results_list.InsertColumn(15, 'Who Claimed')
        results_list.InsertColumn(16, 'When Claimed')
        results_list.InsertColumn(17, 'Who Assigned To')
        '''

    # generate SQL query from search fields
    sql = "SELECT "

    # limit the records pulled if desired
    if ctrl(General.app.search_ecrs_dialog, 'choice:limit').GetStringSelection() != '(no limit)':
        sql += "TOP {} ".format(
            int(ctrl(General.app.search_ecrs_dialog, 'choice:limit').GetStringSelection().split(' ')[0]))
    sql += "* FROM ecrs WHERE "

    table_columns = ['id', 'status', 'reference_number', 'document', 'reason', 'department', 'who_requested', 'type',
                     'request', 'resolution', 'who_errored', 'who_claimed', 'who_assigned', 'when_requested',
                     'when_needed', 'when_closed', 'when_modified']

    ##a better way to do it would be like...
    # table_columns = Database.get_table_column_names('ecrs')
    # table_columns.remove('attachment')

    for i in range(13):
        condition = ctrl(General.app.search_ecrs_dialog, 'choice:search_condition' + str(i)).GetStringSelection()
        if condition == '=':
            sql += table_columns[i] + '=\'{}\' AND '.format(
                ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(i)).GetValue())
        if condition == '>':
            sql += table_columns[i] + '>\'{}\' AND '.format(
                ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(i)).GetValue())
        if condition == '<':
            sql += table_columns[i] + '<\'{}\' AND '.format(
                ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(i)).GetValue())
        if condition == '>=':
            sql += table_columns[i] + '>=\'{}\' AND '.format(
                ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(i)).GetValue())
        if condition == '<=':
            sql += table_columns[i] + '<=\'{}\' AND '.format(
                ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(i)).GetValue())
        if condition == 'not':
            sql += table_columns[i] + '<>\'{}\' AND '.format(
                ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(i)).GetValue())
        if condition == 'contains':
            sql += table_columns[i] + ' LIKE \'%{}%\' AND '.format(
                ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(i)).GetValue())

    # set minimum date?
    date = ctrl(General.app.search_ecrs_dialog, 'date_picker:requested_from').GetValue().Format("%Y-%m-%d")
    if date != '1900-01-01':
        sql += 'when_requested>\'{} 00:00:00\' AND '.format(date)
    date = ctrl(General.app.search_ecrs_dialog, 'date_picker:needed_from').GetValue().Format("%Y-%m-%d")
    if date != '1900-01-01':
        sql += 'when_needed>\'{} 00:00:00\' AND '.format(date)
    date = ctrl(General.app.search_ecrs_dialog, 'date_picker:closed_from').GetValue().Format("%Y-%m-%d")
    if date != '1900-01-01':
        sql += 'when_closed>\'{} 00:00:00\' AND '.format(date)
    date = ctrl(General.app.search_ecrs_dialog, 'date_picker:modified_from').GetValue().Format("%Y-%m-%d")
    if date != '1900-01-01':
        sql += 'when_modified>\'{} 00:00:00\' AND '.format(date)

    # set maximum date
    now = dt.datetime.now()
    if ctrl(General.app.search_ecrs_dialog, 'date_picker:requested_to').GetValue().Format(
            "%d-%m-%Y") != "%02d-%02d-%04d" % (now.day, now.month, now.year + 1):
        sql += 'when_requested<\'{} 23:59:59\' AND '.format(
            ctrl(General.app.search_ecrs_dialog, 'date_picker:requested_to').GetValue().Format("%Y-%m-%d"))
    if ctrl(General.app.search_ecrs_dialog, 'date_picker:needed_to').GetValue().Format(
            "%d-%m-%Y") != "%02d-%02d-%04d" % (now.day, now.month, now.year + 1):
        sql += 'when_needed<\'{} 23:59:59\' AND '.format(
            ctrl(General.app.search_ecrs_dialog, 'date_picker:needed_to').GetValue().Format("%Y-%m-%d"))
    if ctrl(General.app.search_ecrs_dialog, 'date_picker:closed_to').GetValue().Format(
            "%d-%m-%Y") != "%02d-%02d-%04d" % (now.day, now.month, now.year + 1):
        sql += 'when_closed<\'{} 23:59:59\' AND '.format(
            ctrl(General.app.search_ecrs_dialog, 'date_picker:closed_to').GetValue().Format("%Y-%m-%d"))
    if ctrl(General.app.search_ecrs_dialog, 'date_picker:modified_to').GetValue().Format(
            "%d-%m-%Y") != "%02d-%02d-%04d" % (now.day, now.month, now.year + 1):
        sql += 'when_modified<\'{} 23:59:59\' AND '.format(
            ctrl(General.app.search_ecrs_dialog, 'date_picker:modified_to').GetValue().Format("%Y-%m-%d"))

    # if there is a trailing "AND " on the string then get rid of it
    if sql[-4:] == "AND ":
        sql = sql[:-4]

    # if there is a trailing "WHERE" on the string then get rid of it
    if sql[-6:] == "WHERE ":
        sql = sql[:-6]

    # specify sort order
    value = ctrl(General.app.search_ecrs_dialog, 'choice:sort_by').GetSelection()
    column = Database.get_table_column_names('ecrs')[value]
    sql += "ORDER BY {} ".format(column)

    if ctrl(General.app.search_ecrs_dialog, 'choice:sort_in').GetSelection() == 0:
        sql += 'DESC '
    else:
        sql += 'ASC '

    ##limit the records pulled if desired
    # if ctrl(General.app.search_ecrs_dialog, 'choice:limit').GetStringSelection() != '(no limit)':
    #	sql += "LIMIT {}".format(int(ctrl(General.app.search_ecrs_dialog, 'choice:limit').GetStringSelection().split(' ')[0]))

    # query the database
    cursor = Database.connection.cursor()
    cursor.execute(sql)
    records = cursor.fetchall()

    for ecr_index, ecr in enumerate(records):
        results_list.InsertStringItem(sys.maxint, '#')

        for column_index, column_value in enumerate(ecr):
            if column_names[column_index] == 'When Requested':
                column_value = General.format_date_nicely(column_value)
            if column_names[column_index] == 'When Needed':
                column_value = General.format_date_nicely(column_value)
            if column_names[column_index] == 'When Closed':
                column_value = General.format_date_nicely(column_value)
            if column_names[column_index] == 'When Modified':
                column_value = General.format_date_nicely(column_value)
            if column_names[column_index] == 'When Claimed':
                column_value = General.format_date_nicely(column_value)

            if column_value != None:
                results_list.SetStringItem(ecr_index, column_index, str(column_value).replace('\n', ' \\ '))

    # print documents_seen_above

    for column_index in range(len(column_names)):
        results_list.SetColumnWidth(column_index, wx.LIST_AUTOSIZE_USEHEADER)

    '''
    for i in range(len(records)):
        row = records[i]

        #populate row fields
        results_list.InsertStringItem(sys.maxint, '#')
        for j in range(10):
            if row[j] != None:
                results_list.SetStringItem(i, j, str(row[j]))

        #requested date
        dt_object = time.strptime(str(row[11]), "%Y-%m-%d %H:%M:%S") #to python time object
        results_list.SetStringItem(i, 10, time.strftime("%m/%d/%y   %I:%M %p", dt_object))

        #needed date
        dt_object = time.strptime(str(row[12]), "%Y-%m-%d %H:%M:%S") #to python time object
        results_list.SetStringItem(i, 11, time.strftime("%m/%d/%y   %I:%M %p", dt_object))

        #closed date
        if row[13] != None:
            dt_object = time.strptime(str(row[13]), "%Y-%m-%d %H:%M:%S") #to python time object
            results_list.SetStringItem(i, 12, time.strftime("%m/%d/%y   %I:%M %p", dt_object))

        #closed date
        if row[14] != None:
            dt_object = time.strptime(str(row[14]), "%Y-%m-%d %H:%M:%S") #to python time object
            results_list.SetStringItem(i, 13, time.strftime("%m/%d/%y   %I:%M %p", dt_object))

        #the rest of the entries
        for j in range(14, 18):
            if row[j+1] != None:
                results_list.SetStringItem(i, j, str(row[j+1]))

    #set some columns to a good width
    results_list.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(2, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(3, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(4, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(5, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(6, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(7, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(8, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(9, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(10, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(11, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(12, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(13, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(14, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(15, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(16, wx.LIST_AUTOSIZE_USEHEADER)
    results_list.SetColumnWidth(17, wx.LIST_AUTOSIZE_USEHEADER)
    '''

    event.GetEventObject().SetLabel('Begin Search')
    notebook = ctrl(General.app.main_frame, 'notebook:ecrs')
    notebook.SetSelection(notebook.GetPageCount() - 1)

    General.app.search_ecrs_dialog.Hide()


def search_value_entered(event, index):
    # if condition not specified when they type a value, force it to '='
    if ctrl(General.app.search_ecrs_dialog, 'choice:search_condition' + str(index)).GetSelection() == 0:
        if ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(index)).GetValue() != '':
            ctrl(General.app.search_ecrs_dialog, 'choice:search_condition' + str(index)).SetSelection(1)

    # although, if what they typed ends up being blank, clear the condition
    if ctrl(General.app.search_ecrs_dialog, 'combo:search_value' + str(index)).GetValue() == '':
        ctrl(General.app.search_ecrs_dialog, 'choice:search_condition' + str(index)).SetSelection(0)


def send_ecr_assigned_email(ecr, order, reciever, sender):
    server = smtplib.SMTP('mailrelay.lennoxintl.com')

    shortcuts = ''

    item_number = ''
    sales_order = ''
    customer = ''
    location = ''
    model = ''

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = reciever
    if order != None:
        msg["Subject"] = 'Assigned ECR: {}, SO: {}-{}, RefNo: {}'.format(ecr[0], order[1], order[2], ecr[1])

        order_directory = OrderFileOpeners.get_order_directory(order[1])

        if order_directory:
            shortcuts = '<a href=\"file:///{}\">Open Order Folder</a>'.format(order_directory)

        item_number = order[0]
        sales_order = '{}-{}'.format(order[1], order[2])
        customer = order[5]
        location = '{}, {}'.format(order[8], order[9])
        model = order[11]
    else:
        msg["Subject"] = 'Assigned ECR: {}, RefNo: {}'.format(ecr[0], ecr[1])

    msg['Date'] = formatdate(localtime=True)

    # size="3"
    body_html = '''<style type=\"text/css\">td{{font-family:Arial; color:black; font-size:12pt;}}</style>
		<font face=\"arial\">
		You have been assigned this ECR<br><br>
		{}
		<hr>
		<table border="0">
		<tr><td align=\"right\">Item&nbsp;Number:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Sales&nbsp;Order:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Customer:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Location:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Model:&nbsp;</td><td>{}</td></tr>
		</table>
		<hr>
		<table border="0">
		<tr><td align=\"right\">ECR&nbsp;ID:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">RefNo:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\" valign=\"top\">Request:&nbsp;</td><td>{}</td></tr>
		</table>
		'''.format(shortcuts, item_number, sales_order, customer, location, model, ecr[0], ecr[1], ecr[2])

    body = MIMEMultipart('alternative')
    body.attach(MIMEText(body_html, 'html'))
    msg.attach(body)

    # print email_string

    try:
        server.sendmail(sender, reciever, msg.as_string())
    except Exception, e:
        wx.MessageBox('Unable to send email. Error: {}'.format(e), 'An error occurred!', wx.OK | wx.ICON_ERROR)

    server.close()


def send_similar_items_email(order_data, ecr_data, similar_items_data, recievers, sender):
    server = smtplib.SMTP('mailrelay.lennoxintl.com')

    reciever_email_string = ''
    for email in recievers:
        reciever_email_string += '; {}'.format(email)
    reciever_email_string = reciever_email_string[2:]

    # print "reciever_email_string:", reciever_email_string

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = reciever_email_string
    # msg["Subject"] = 'Possibly Pertinent Items for Closed ECR'
    msg["Subject"] = 'Closed ECR might apply to other items'

    # order_directory = OrderFileOpeners.get_order_directory(order[1])

    # if order_directory:
    #	shortcuts = '<a href=\"file:///{}\">Open Order Folder</a>'.format(order_directory)


    msg['Date'] = formatdate(localtime=True)

    # size="3" #<font size="10" face="Calibri">
    body_html = '''<style type="text/css">
		td{font-family:Calibri; color:black; font-size:11pt;}
		BODY{font-family:Calibri; color:black; font-size:11pt;}
		</style>\n'''

    ecr_id = ecr_data[0]
    ecr_type = ecr_data[7]
    reference_number = ecr_data[1]
    document = ecr_data[8]
    request = ecr_data[2]
    resolution = ecr_data[3]

    body_html += '''
		<table border="0">
		<tr><td align="right">ECR&nbsp;ID:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right">Type:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right">RefNo:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right">Document:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right" valign="top">Request:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right" valign="top">Resolution:&nbsp;</td><td>{}</td></tr>
		</table>
		<hr>\n
		'''.format(ecr_id, ecr_type, reference_number, document, request, resolution)

    item_number = order_data[0]
    sales_order = '{}-{}'.format(order_data[1], order_data[2])
    customer = order_data[5]
    location = '{}, {}'.format(order_data[8], order_data[9])
    model = order_data[11]
    family = order_data[21]

    body_html += '''
		<table border="0">
		<tr><td align="right">Item&nbsp;Number:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right">Sales&nbsp;Order:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right">Customer:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right">Location:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right">Model:&nbsp;</td><td>{}</td></tr>
		<tr><td align="right">Family:&nbsp;</td><td>{}</td></tr>
		</table>
		<br>\n
		'''.format(item_number, sales_order, customer, location, model, family)

    body_html += "These are items of the same customer and family for this ECR that are released but not yet produced:<br>\n"

    body_html += '''<table border="0">'''

    for similar_item_data in similar_items_data:
        sales_order, item, project_lead, mechanical_engineer, electrical_engineer, structural_engineer = similar_item_data

        order_directory = OrderFileOpeners.get_order_directory(sales_order)
        if order_directory:
            sales_order = '''<a href=\"file:///{}\">{}</a>'''.format(order_directory, sales_order)

        if ecr_type == 'Mechanical':
            engineer = mechanical_engineer
        elif ecr_type == 'Electrical':
            engineer = electrical_engineer
        elif ecr_type == 'Structural':
            engineer = structural_engineer
        elif ecr_type == 'Other':
            engineer = project_lead

        if engineer == None:
            engineer = project_lead

        body_html += '''<tr><td>{}&nbsp;</td><td>{}&nbsp;</td><td>{}&nbsp;</td></tr>\n'''.format(sales_order, item,
                                                                                                 engineer)

    body_html += '''</table>\n'''

    body = MIMEMultipart('alternative')
    body.attach(MIMEText(body_html, 'html'))
    msg.attach(body)

    # print email_string

    try:
        server.sendmail(sender, recievers, msg.as_string())
    # 1/0.
    # server.sendmail(sender, ('Travis.Stuart@Heatcraftrpd.com'), msg.as_string())
    except Exception, e:
        wx.MessageBox('Unable to send email. Error: {}'.format(e), 'An error occurred!', wx.OK | wx.ICON_ERROR)

    server.close()


def send_ecr_closed_email(ecr, order, reciever, sender):
    server = smtplib.SMTP('mailrelay.lennoxintl.com')

    shortcuts = ''

    item_number = ''
    sales_order = ''
    customer = ''
    location = ''
    model = ''

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = reciever
    if order != None:
        msg["Subject"] = 'Closed ECR: {}, SO: {}-{}, RefNo: {}'.format(ecr[0], order[1], order[2], ecr[1])

        order_directory = OrderFileOpeners.get_order_directory(order[1])

        if order_directory:
            shortcuts = '<a href=\"file:///{}\">Open Order Folder</a>'.format(order_directory)

        item_number = order[0]
        sales_order = '{}-{}'.format(order[1], order[2])
        customer = order[5]
        location = '{}, {}'.format(order[8], order[9])
        model = order[11]
    else:
        msg["Subject"] = 'Closed ECR: {}, RefNo: {}'.format(ecr[0], ecr[1])

    msg['Date'] = formatdate(localtime=True)

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
		</table>
		<hr>
		<table border="0">
		<tr><td align=\"right\">ECR&nbsp;ID:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">RefNo:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\" valign=\"top\">Request:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\" valign=\"top\">Resolution:&nbsp;</td><td>{}</td></tr>
		</table>
		'''.format(shortcuts, item_number, sales_order, customer, location, model, ecr[0], ecr[1], ecr[2], ecr[3])

    body = MIMEMultipart('alternative')
    body.attach(MIMEText(body_html, 'html'))
    msg.attach(body)

    # print email_string

    try:
        server.sendmail(sender, reciever, msg.as_string())
    except Exception, e:
        wx.MessageBox('Unable to send email. Error: {}'.format(e), 'An error occurred!', wx.OK | wx.ICON_ERROR)

    server.close()


def send_ecr_soe_email(ecr, order, reciever, sender):
    server = smtplib.SMTP('mailrelay.lennoxintl.com')

    shortcuts = ''

    item_number = ''
    sales_order = ''
    customer = ''
    location = ''
    model = ''

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = reciever
    if order != None:
        msg["Subject"] = 'Closed ECR: {}, SO: {}-{}, RefNo: {}'.format(ecr[0], order[1], order[2], ecr[1])

        order_directory = OrderFileOpeners.get_order_directory(order[1])

        if order_directory:
            shortcuts = '<a href=\"file:///{}\">Open Order Folder</a>'.format(order_directory)

        item_number = order[0]
        sales_order = '{}-{}'.format(order[1], order[2])
        customer = order[5]
        location = '{}, {}'.format(order[8], order[9])
        model = order[11]
    else:
        msg["Subject"] = 'Closed ECR: {}, RefNo: {}'.format(ecr[0], ecr[1])

    msg['Date'] = formatdate(localtime=True)

    # size="3"
    body_html = '''<style type=\"text/css\">td{{font-family:Arial; color:black; font-size:12pt;}}</style>
		<font face=\"arial\">
		You have been marked as the source of error for this ECR<br><br>
		{}
		<hr>
		<table border="0">
		<tr><td align=\"right\">Item&nbsp;Number:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Sales&nbsp;Order:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Customer:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Location:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">Model:&nbsp;</td><td>{}</td></tr>
		</table>
		<hr>
		<table border="0">
		<tr><td align=\"right\">ECR&nbsp;ID:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\">RefNo:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\" valign=\"top\">Request:&nbsp;</td><td>{}</td></tr>
		<tr><td align=\"right\" valign=\"top\">Resolution:&nbsp;</td><td>{}</td></tr>
		</table>
		'''.format(shortcuts, item_number, sales_order, customer, location, model, ecr[0], ecr[1], ecr[2], ecr[3])

    body = MIMEMultipart('alternative')
    body.attach(MIMEText(body_html, 'html'))
    msg.attach(body)

    # print email_string

    try:
        server.sendmail(sender, reciever, msg.as_string())
    except Exception, e:
        wx.MessageBox('Unable to send email. Error: {}'.format(e), 'An error occurred!', wx.OK | wx.ICON_ERROR)

    server.close()


def update_useful_info_panel(query_result):
    if validate_reference_entry() == True:
        # if the entry number is in the DB, write its data to the 'usefull information' screen
        if query_result != None:
            material = query_result[21]
            if material in ('CDA', 'CA', 'DSS'):
                wx.MessageBox(
                    'This product should be supported with full 3D CAD models. As such, the BOM and other engineering documentation should be accurate. Please ensure the product is being built per the drawings and consult with your lead-man or supervisor before proceeding.',
                    'Warning', wx.OK | wx.ICON_WARNING)

            ctrl(General.app.new_ecr_dialog, 'label:item_number').SetLabel(str(query_result[0]))
            ctrl(General.app.new_ecr_dialog, 'label:sales_order').SetLabel(str(query_result[1]))
            ctrl(General.app.new_ecr_dialog, 'label:line_up').SetLabel(str(query_result[2]))
            ctrl(General.app.new_ecr_dialog, 'label:serial').SetLabel(str(query_result[3]))
            ctrl(General.app.new_ecr_dialog, 'label:quote').SetLabel(str(query_result[4]))
            ctrl(General.app.new_ecr_dialog, 'label:model').SetLabel(str(query_result[11]))

            ctrl(General.app.new_ecr_dialog, 'label:customer').SetLabel(str(query_result[5]))
            ctrl(General.app.new_ecr_dialog, 'label:store_name').SetLabel(str(query_result[6]))
            ctrl(General.app.new_ecr_dialog, 'label:store_id').SetLabel(str(query_result[7]))
            ctrl(General.app.new_ecr_dialog, 'label:city').SetLabel(str(query_result[8]))
            ctrl(General.app.new_ecr_dialog, 'label:state').SetLabel(str(query_result[9]))
            ctrl(General.app.new_ecr_dialog, 'label:country').SetLabel(str(query_result[10]))

            # ctrl(General.app.new_ecr_dialog, 'label:mechanical').SetLabel(str(query_result[13]))
            # ctrl(General.app.new_ecr_dialog, 'label:electrical').SetLabel(str(query_result[14]))
            # ctrl(General.app.new_ecr_dialog, 'label:structural').SetLabel(str(query_result[15]))
            # ctrl(General.app.new_ecr_dialog, 'label:program').SetLabel(str(query_result[16]))

            if query_result[17] == True:
                ctrl(General.app.new_ecr_dialog, 'label:order_status').SetLabel('Canceled')
            else:
                ctrl(General.app.new_ecr_dialog, 'label:order_status').SetLabel('Valid')
            ctrl(General.app.new_ecr_dialog, 'label:date_released').SetLabel(
                General.format_date_nicely(str(query_result[18]))[:8])
            ctrl(General.app.new_ecr_dialog, 'label:date_produced').SetLabel(
                General.format_date_nicely(str(query_result[19]))[:8])
            ctrl(General.app.new_ecr_dialog, 'label:date_shipped').SetLabel(
                General.format_date_nicely(str(query_result[20]))[:8])

    else:
        print 'validation failed :('
        clear_useful_info_panel()


def validate_reference_entry():
    # uppercase user's entry
    text_box = ctrl(General.app.new_ecr_dialog, 'text:reference_number')
    selection = text_box.GetSelection()
    value = text_box.GetValue().upper()
    text_box.ChangeValue(value)
    text_box.SetSelection(*selection)

    # validate reference num by entry length and other defining characteristics
    entry = ctrl(General.app.new_ecr_dialog, 'text:reference_number').GetValue()
    if (len(entry) == 7) and (entry[0] == '0'):
        return True
    elif (len(entry) == 9) and (entry[0:3] == 'KW0'):
        return True
    elif len(entry) == 10:
        return True
    elif len(entry) == 6:
        return True
    elif '-' in entry:
        return True
    elif len(entry) == 8 and entry[0] == '5':
        return True
    elif len(entry) == 8 and entry[0] == '2':
        return True
    else:
        return False
