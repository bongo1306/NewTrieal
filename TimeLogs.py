import wx		#wxWidgets used as the GUI
from wx.html import HtmlEasyPrinting
from wx import xrc		#allows the loading and access of xrc file (xml) that describes GUI
import wx.grid as gridlib
from wxPython.calendar import *
ctrl = xrc.XRCCTRL		#define a shortined function name (just for convienience)

import sched #for updating the timer every second or so

import sys
import os
import time
import datetime as dt
from threading import Thread #for timer updating in the background

import Database
import General



def on_activate_open_modify_dialog(event):
	time_log_id = event.GetItem().GetText()
	General.app.init_modify_time_log_dialog(time_log_id)


def on_click_save_time_log(event):
	id = General.app.modify_time_log_dialog.GetTitle().split(' ')[-1]
	item_number = ctrl(General.app.modify_time_log_dialog, 'text:item_number').GetValue()
	hours_worked = ctrl(General.app.modify_time_log_dialog, 'text:hours_worked').GetValue()
	when_logged = ctrl(General.app.modify_time_log_dialog, 'calendar:when_logged').GetDate()
	when_logged = dt.date(when_logged.GetYear(), when_logged.GetMonth()+1, when_logged.GetDay())
	tags = ctrl(General.app.modify_time_log_dialog, 'text:tags').GetValue()
	
	sql = "UPDATE time_logs SET "
	sql += "item='{}', ".format(item_number)
	sql += "hours={}, ".format(hours_worked)
	sql += "when_logged='{} 12:00:00', ".format(when_logged)
	sql += "tags='{}' ".format(tags)
	sql += "WHERE id={}".format(id)
	
	try:
		cursor = Database.connection.cursor()
		cursor.execute(sql)
		Database.connection.commit()
		
		refresh_log_list()
		General.app.modify_time_log_dialog.Destroy()
		
	except Exception as e:
		print e
		wx.MessageBox(str(e), 'Error', wx.OK | wx.ICON_WARNING)
		


def on_click_log_time(event):
	cursor = Database.connection.cursor()
	
	hour_field = ctrl(General.app.main_frame, 'text:hours')
	item_field = ctrl(General.app.main_frame, 'text:item_number_to_log')
	
	#simple field validation
	if item_field.GetValue() == '':
		wx.MessageBox('Enter a valid item number before logging time.', 'Hint', wx.OK | wx.ICON_WARNING)
		return
	if hour_field.GetValue() == '':
		wx.MessageBox('Enter hours worked before logging time.', 'Hint', wx.OK | wx.ICON_WARNING)
		return
	try:
		float(hour_field.GetValue())
	except:
		wx.MessageBox('The hours value you entered does not appear to be a number...', 'Hint', wx.OK | wx.ICON_WARNING)
		return

	#check if item number is valid
	if cursor.execute("SELECT TOP 1 item FROM orders WHERE item = \'{}\'".format(item_field.GetValue())).fetchone() == None:
		wx.MessageBox('The item number you entered was not found in the orders table.\nIt may be a typo or the orders table is not up to date.', 'Hint', wx.OK | wx.ICON_WARNING)
		return

	item = item_field.GetValue()
	employee = General.app.current_user
	hours = hour_field.GetValue()
	when_logged = str(dt.datetime.today())[:19]
	
	new_id = cursor.execute("SELECT MAX(id) FROM time_logs").fetchone()[0] + 1
	
	sql = 'INSERT INTO time_logs (id, item, employee, hours, when_logged) VALUES ('
	sql += '\'{}\', '.format(new_id)
	sql += '\'{}\', '.format(item)
	sql += '\'{}\', '.format(employee)
	sql += '\'{}\', '.format(hours)
	sql += '\'{}\')'.format(when_logged)

	cursor.execute(sql)
	Database.connection.commit()
	
	item_field.SetValue('')
	hour_field.SetValue('')
	
	#turn off timer if it's on
	General.app.timer_start_time = None
	ctrl(General.app.main_frame, 'toggle:timer').SetLabel('Start Timer')
	ctrl(General.app.main_frame, 'toggle:timer').SetValue(False)
	
	refresh_log_list()
	

def on_click_timer(event):
	hour_field = ctrl(General.app.main_frame, 'text:hours')
	
	if hour_field.GetValue().strip() == '':
		hour_field.SetValue('0')
		
	if ctrl(General.app.main_frame, 'toggle:timer').GetLabel() == 'Start Timer':
		print 'Starting timer'
		if General.app.timer_thread != None:
			General.app.timer_thread.join()
			print 'finished joining old thread'
		
		#try to pickup timer where it left off...
		try:
			General.app.timer_start_time = dt.datetime.now()-dt.timedelta(seconds=float(hour_field.GetValue())*3600.)
		except:
			General.app.timer_start_time = dt.datetime.now()
			hour_field.SetValue('0')
			
		General.app.timer_schedule.enter(.1, 1, update_timer, (General.app.timer_schedule, ))
	
		General.app.timer_thread = Thread(target=General.app.timer_schedule.run)
		#General.app.timer_thread.setDaemon(True)
		General.app.timer_thread.start()
		
		ctrl(General.app.main_frame, 'toggle:timer').SetLabel('Stop Timer')
		ctrl(General.app.main_frame, 'toggle:timer').SetValue(True)
	else:
		#clicking the timer button toggles the stopwatch effect
		General.app.timer_start_time = None
		ctrl(General.app.main_frame, 'toggle:timer').SetLabel('Start Timer')
		ctrl(General.app.main_frame, 'toggle:timer').SetValue(False)



def on_select_log(event):
	item = event.GetEventObject()
	ctrl(General.app.main_frame, 'text:item_number_to_log').SetValue(item.GetItem(item.GetFirstSelected(), 1).GetText())
	


def update_timer(timer_schedule):
	if General.app.timer_start_time != None:
		#print (dt.datetime.now() - General.app.timer_start_time).seconds
		print 'workin the thread'
		
		try:
			ctrl(General.app.main_frame, 'text:hours').SetValue('{:.4f}'.format((dt.datetime.now() - General.app.timer_start_time).seconds/3600.))
		except:
			return
		
		timer_schedule.enter(1, 1, update_timer, (timer_schedule, ))
	else:
		print 'this thread should be dying'
		
		#kill the schedule!
		#actually.... not really needed... just don't enter a new schedule entry
		return


def refresh_log_list(event=None):
	cursor = Database.connection.cursor()
	
	log_list = ctrl(General.app.main_frame, 'list:my_time_logs')
	
	#clear out the list
	log_list.DeleteAllItems()

	#create columns in list if they're not already there
	column_names = Database.get_table_column_names('time_logs', presentable=True)
	
	if log_list.GetColumn(0) == None:
		for index, column_name in enumerate(column_names):
			log_list.InsertColumn(index, column_name)
			
		log_list.InsertColumn(index+1, 'Total hours for this Item')

	#get logs for logged in user
	time_logs = cursor.execute("SELECT TOP 200 id, item, employee, hours, when_logged, tags FROM time_logs WHERE employee = \'{}\' ORDER BY when_logged DESC".format(General.app.current_user)).fetchall()

	#populate list with records
	for time_log_index, time_log in enumerate(time_logs):
		log_list.InsertStringItem(sys.maxint, '#')
		
		
		
		for column_index, column_value in enumerate(time_log):
			if column_names[column_index] == 'When Logged':
				if column_value != None: column_value = General.format_date_nicely(column_value)

			if column_names[column_index] == 'Hours':
				if column_value != None: column_value = '{:.4f}'.format(column_value).rstrip('0').rstrip('.')

			if column_value != None:
				log_list.SetStringItem(time_log_index, column_index, str(column_value))

		item = time_log[1]
		total_hours_for_item = sum(zip(*cursor.execute("SELECT hours FROM time_logs WHERE employee = \'{}\' and item = \'{}\'".format(General.app.current_user, item)).fetchall())[0])
		log_list.SetStringItem(time_log_index, column_index+1, '{:.4f}'.format(total_hours_for_item).rstrip('0').rstrip('.'))


	columns_to_hide = ['id', 'employee']
	#columns_to_hide = []

	for column_index, column_name in enumerate(column_names):
		if column_name.lower().replace(' ', '_') in columns_to_hide:
			log_list.SetColumnWidth(column_index, 0)
		else:
			log_list.SetColumnWidth(column_index, wx.LIST_AUTOSIZE_USEHEADER)
	
	#total time col
	log_list.SetColumnWidth(column_index+1, wx.LIST_AUTOSIZE_USEHEADER)
	
	
	
