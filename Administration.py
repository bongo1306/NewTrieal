#!/usr/bin/env python
# -*- coding: utf8 -*-

import wx		#wxWidgets used as the GUI
from wx.html import HtmlEasyPrinting
from wx import xrc		#allows the loading and access of xrc file (xml) that describes GUI
import wx.grid as gridlib
from wx.calendar import *
ctrl = xrc.XRCCTRL		#define a shortined function name (just for convienience)

import pyodbc #for connecting to dbworks database

import csv #for reading in exported filemaker data

import os
import xlwt
import datetime

import Database
import General
import Ecrs

def on_select_admin(event):
	name = event.GetEventObject().GetStringSelection()
	
	cursor = Database.connection.cursor()
	cursor.execute("UPDATE administration SET ecr_admin_name = '{}' where Production_Plant = \'{}\' ".format(name.replace("'", "''"),Ecrs.Prod_Plant))
	Database.connection.commit()
	
	#reset status bar to reflect newly assigned admin
	ctrl(General.app.main_frame, 'statusbar:main').SetStatusText('ECR administrator: {}'.format(name))


def on_click_backup_database(event):
	default_file_name = "eng04_sql_backup {}".format(str(datetime.date.today()))
	
	#prompt user to choose where to save
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

	cursor = Database.connection.cursor()

	#save to excel
	workbook = xlwt.Workbook()
	
	tables = list(zip(*cursor.execute('SELECT * FROM information_schema.tables').fetchall())[2])
	tables.sort()
	for table in tables:
		print table
		worksheet = workbook.add_sheet(table)
		
		#write out headers
		column_names = zip(*cursor.execute("SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_Name='{}' ORDER by ORDINAL_POSITION".format(table)).fetchall())[3]
		for index, column_name in enumerate(column_names):
			worksheet.write(0, index, column_name)

		#write out data
		table_data = cursor.execute("SELECT * FROM {}".format(table))
		for row_index, row in enumerate(table_data):
			for col_index, col in enumerate(row):
				worksheet.write(row_index+1, col_index, col)

	workbook.save(save_path)
	
	wx.MessageBox('Backup completed.', 'Info', wx.OK | wx.ICON_INFORMATION)



