#!/usr/bin/env python
# -*- coding: utf8 -*-
version = '9.8'

# extend Python's functionality by importing modules
import sys
import os
import traceback  # if an error occurs in the program, traceback can give us a run up to what caused the problem
import operator
import json

# wxWidgets used as the GUI
import wx
from wx import xrc  # allows the loading and access of xrc file (xml) that describes GUI
import wx.grid as gridlib  # an excel like table widget
import wx.lib.scrolledpanel as scrolled  # for the scrollable search panel
from wx.calendar import *
import wx.richtext as rt

ctrl = xrc.XRCCTRL  # define a shortined function name (just for convienience)

import psutil  # for killing the splash screen process
from subprocess import Popen  # for opening documents in their native program
import xlwt  # for writing data to excel files
import ConfigParser  # for reading local config data (*.cfg)

import datetime as dt
import time

import pyodbc  # for access to SQL server
# import sqlite3 as lite	#for manipulating the test database

import sched  # for updating the timer every second or so
from threading import Thread  # so a slow querry won't make the gui lag (ONLY FOR READS NOT WRITES!)
from functools import partial  # allows us to pass more arguments to a function "Bind"ed to a GUI event

# bring in our own modules
import General
import Database
import Ecrs
import Revisions
# import TimeLogs
import Search
import Reports
import Administration
import OrderFileOpeners
import TweakedGrid

import BetterListCtrl


def check_for_updates():
    try:
        with open(os.path.join(General.updates_dir, "releases.json")) as file:
            releases = json.load(file)

            latest_version = releases[0]['version']

            if version != latest_version:
                return True
            else:
                return False
    except Exception as e:
        print 'Failed update check:', e


def open_software_update_frame():
    SoftwareUpdateFrame(None)


class SoftwareUpdateFrame(wx.Frame):
    def __init__(self, parent):
        # load frame XRC description
        pre = wx.PreFrame()
        #res = xrc.XmlResource.Get()
        res = xrc.XmlResource(General.resource_path('interface.xrc'))
        res.LoadOnFrame(pre, parent, "frame:software_update")
        self.PostCreate(pre)
        #self.SetIcon(wx.Icon(General.resource_path('ECRev.ico'), wx.BITMAP_TYPE_ICO))

        # read in update text data
        with open(os.path.join(General.updates_dir, "releases.json")) as file:
            releases = json.load(file)

        latest_version = releases[0]['version']
        self.install_filename = releases[0]['installer filename']

        it_is_mandatory = False

        # build up what changed text
        changes_panel = ctrl(self, 'panel:changes')
        richtext_ctrl = rt.RichTextCtrl(changes_panel, style=wx.VSCROLL | wx.HSCROLL | wx.TE_READONLY)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(richtext_ctrl, 1, wx.EXPAND)
        changes_panel.SetSizer(sizer)
        changes_panel.Layout()

        for release in releases:
            if float(version) < float(release['version']):

                richtext_ctrl.BeginBold()
                richtext_ctrl.WriteText('v{} - {}'.format(release['version'], release['release date']))
                richtext_ctrl.EndBold()
                richtext_ctrl.Newline()

                richtext_ctrl.BeginStandardBullet('*', 50, 30)

                for change in release['changes']:
                    richtext_ctrl.WriteText('{}'.format(change))
                    richtext_ctrl.Newline()

                richtext_ctrl.EndStandardBullet()

                if release['mandatory'] == True:
                    it_is_mandatory = True

        # bindings
        self.Bind(wx.EVT_CLOSE, self.on_close_frame)
        self.Bind(wx.EVT_BUTTON, self.on_click_cancel, id=xrc.XRCID('button:not_now'))
        self.Bind(wx.EVT_BUTTON, self.on_click_update, id=xrc.XRCID('button:update'))

        # misc
        ctrl(self, 'label:intro_version').SetLabel(
            "A new version of the ECRev software was found on the server: v{}".format(latest_version))
        ctrl(self, 'button:update').SetFocus()

        if it_is_mandatory == False:
            ctrl(self, 'button:not_now').Enable()
            ctrl(self, 'label:mandatory').Hide()

        self.Show()

    def on_click_cancel(self, event):
        self.Close()

    def on_click_update(self, event):
        General.app.login_frame.Destroy()
        print 'copy install file over, open it, and close this program'
        # create a dialog to show log of what's goin on
        dialog = wx.Dialog(self, id=wx.ID_ANY, title=u"Starting Update...", pos=wx.DefaultPosition, size=wx.DefaultSize,
                           style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        dialog.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
        dialog.SetFont(wx.Font(10, 70, 90, 90, False, wx.EmptyString))
        bSizer53 = wx.BoxSizer(wx.VERTICAL)
        dialog.m_panel22 = wx.Panel(dialog, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer54 = wx.BoxSizer(wx.VERTICAL)
        bSizer55 = wx.BoxSizer(wx.VERTICAL)
        dialog.text_notice = wx.TextCtrl(dialog.m_panel22, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                         wx.Size(350, 120), wx.TE_DONTWRAP | wx.TE_MULTILINE | wx.TE_READONLY)
        bSizer55.Add(dialog.text_notice, 1, wx.ALL | wx.EXPAND, 5)
        bSizer54.Add(bSizer55, 1, wx.EXPAND, 5)
        dialog.m_panel22.SetSizer(bSizer54)
        dialog.m_panel22.Layout()
        bSizer54.Fit(dialog.m_panel22)
        bSizer53.Add(dialog.m_panel22, 1, wx.EXPAND | wx.ALL, 5)
        dialog.SetSizer(bSizer53)
        dialog.Layout()
        bSizer53.Fit(dialog)
        dialog.Centre(wx.BOTH)
        dialog.Show()

        dialog.text_notice.AppendText('Opening install file... ')
        wx.Yield()

        try:
            source_filepath = os.path.join(General.updates_dir, self.install_filename)
            os.startfile(source_filepath)

        except Exception as e:
            dialog.text_notice.AppendText('[FAIL]\n')
            dialog.text_notice.AppendText('ERROR: {}\n'.format(e))
            dialog.text_notice.AppendText('\nSoftware update failed.\n')
            wx.Yield()
            return

        dialog.text_notice.AppendText('[OK]\n')
        dialog.text_notice.AppendText('Shutting down this program...')
        wx.Yield()
        self.Close()

        """try:
            try:
                gn.splash_frame.Close()
            except Exception as e:
                print 'splash', e

            try:
                gn.login_frame.Close()
            except Exception as e:
                print 'login', e

            try:
                gn.main_frame.Close()
            except Exception as e:
                print 'main', e

            self.Close()

        except Exception as e:
            dialog.text_notice.AppendText('[FAIL]\n')
            dialog.text_notice.AppendText('ERROR: {}\n'.format(e))
            dialog.text_notice.AppendText('\nSoftware update failed.\n')
            wx.Yield()
            return"""

    def on_close_frame(self, event):
        print 'called on_close_frame'
        self.Destroy()

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
def sort_list(event):
    list = event.GetEventObject()
    sort_column_index = event.GetColumn()
    selected_entry = list.GetFirstSelected()

    list.Hide()
    
    headers = []
    column_widths = []
    for col in range(list.GetColumnCount()):
        headers.append(list.GetColumn(col).GetText())
        column_widths.append(list.GetColumnWidth(col))
        
    entries = []
    for row in range(list.GetItemCount()):
        fields = []
        for col in range(list.GetColumnCount()):
            value = list.GetItem(row, col).GetText()
            
            #try to convert string numbers to legit numbers
            try:
                if value[0] == '0' and value[1] != '.':
                    pass #it's not a 'real' number... probaly and old format item number
                else:
                    if '.' in value:
                        try: value = float(value)
                        except: pass
                    else:
                        try: value = int(value)
            except:
                        except: pass
                pass
            
            fields.append(value)


        if row == selected_entry:
            is_selected = True
        else:
            is_selected = False
            
        color = list.GetItemBackgroundColour(row)
        
        fields.extend((is_selected, color))
        
        entries.append(fields)
    
    #↓↑
    
    if u'↓' in headers[sort_column_index]:
        headers[sort_column_index] = headers[sort_column_index].replace(u'↓', u'↑')
        entries.sort(key = operator.itemgetter(sort_column_index), reverse=False)
        
    elif u'↑' in headers[sort_column_index]:
        headers[sort_column_index] = headers[sort_column_index].replace(u'↑', u'↓')
        entries.sort(key = operator.itemgetter(sort_column_index), reverse=True)
        
    else:
        headers[sort_column_index] = u'{} {}'.format(headers[sort_column_index], u'↑')
        entries.sort(key = operator.itemgetter(sort_column_index), reverse=False)

    for header_index, header in enumerate(headers):
        if header_index != sort_column_index:
            headers[header_index] = headers[header_index].replace(u'↑', u'').replace(u'↓', u'').strip()
        
    #rebuild list
    list.DeleteAllColumns()
    list.DeleteAllItems()

    for header_index, header in enumerate(headers):
        list.InsertColumn(header_index, header)

    for entry_index, entry in enumerate(entries):
        list.InsertStringItem(sys.maxint, '#')
        for field_index, field in enumerate(entry[:-2]):
            list.SetStringItem(entry_index, field_index, str(field))
            
        is_selected = entries[entry_index][-2]
        if is_selected:
            list.Select(entry_index)
            list.EnsureVisible(entry_index)
        
        color = entries[entry_index][-1]
        list.SetItemBackgroundColour(entry_index, color)
    
    #return column widths to their previous values
    for column_widths_index, column_width in enumerate(column_widths):
        list.SetColumnWidth(column_widths_index, column_width)
        
    #except for the column we are sorting... as the arrow could make the column not wide enough to display
    list.SetColumnWidth(sort_column_index, wx.LIST_AUTOSIZE_USEHEADER)
    
    list.Show()
'''


class ECRevApp(wx.App):
    def OnInit(self):
        # load the file that describes our GUI
        self.res = xrc.XmlResource(General.resource_path('interface.xrc'))

        self.current_user = None
        self.search_panel = None

        self.main_frame = None
        self.login_frame = None
        self.search_ecrs_dialog = None
        self.new_ecr_dialog = None

        return True

    def quit_app(self, event):
        event.GetEventObject().Destroy()

        # clean up some loose ends
        try:
            # bring that thread back home to mom
            if General.app.timer_thread != None:
                General.app.timer_thread.join()
                print 'joined last timer thread'
        except Exception as e:
            print e

        try:
            Database.connection.close()
        except Exception as e:
            print e

        try:
            General.app.Destroy()
        except Exception as e:
            print e

        print 'Attempting to end program'
        # sys.exit('ECRev Program Exited') #doesn't actually kill the process everytime for some reason...
        os._exit(0)  # but this does

    def do_nothing(self, evt):
        print 'on events pit'

    def init_assign_ecr_dialog(self, ecr_id):
        self.assign_ecr_dialog = self.res.LoadDialog(None, 'dialog:assign_ecr')
        self.assign_ecr_dialog.SetTitle('Assign ECR: {}'.format(ecr_id))

        cursor = Database.connection.cursor()
        people = list(zip(*cursor.execute(
            'SELECT name FROM employees WHERE department = \'Design Engineering\' OR department = \'Applications Engineering\' OR gets_assignments = 1 ORDER BY name ASC').fetchall())[0])
        people.insert(0, '')
        ctrl(self.assign_ecr_dialog, 'choice:name').AppendItems(people)

        self.assign_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_select_assign_ecr, id=xrc.XRCID('choice:name'))
        self.assign_ecr_dialog.ShowModal()

    '''
    def init_modify_time_log_dialog(self, time_log_id):
        self.modify_time_log_dialog = self.res.LoadDialog(None, 'dialog:modify_time_log')
        self.modify_time_log_dialog.SetTitle('Modify Time Log: {}'.format(time_log_id))
        
        cursor = Database.connection.cursor()
        time_log = cursor.execute("SELECT TOP 1 item, hours, when_logged, tags FROM time_logs WHERE id = {}".format(time_log_id)).fetchall()[0]
        
        ctrl(self.modify_time_log_dialog, 'text:item_number').SetValue(str(time_log[0]))
        ctrl(self.modify_time_log_dialog, 'text:hours_worked').SetValue(str(time_log[1]))
        ctrl(self.modify_time_log_dialog, 'text:hours_worked').SetFocus()
        
        when_logged = time.strptime(str(time_log[2]), "%Y-%m-%d %H:%M:%S") #to python time object
        ctrl(self.modify_time_log_dialog, 'calendar:when_logged').SetDate(wx.DateTimeFromDMY(when_logged.tm_mday, when_logged.tm_mon-1, when_logged.tm_year))

        if time_log[3]:
            ctrl(self.modify_time_log_dialog, 'text:tags').SetValue(str(time_log[3]))

        self.modify_time_log_dialog.Bind(wx.EVT_BUTTON, TimeLogs.on_click_save_time_log, id=xrc.XRCID('button:save'))		
        self.modify_time_log_dialog.ShowModal()
    '''

    def init_login_frame(self):
        if self.login_frame != None:
            self.login_frame.Destroy()

        self.login_frame = self.res.LoadFrame(None, 'frame:login')
        # self.login_frame.SetSize((950, 650))
        # self.login_frame.Maximize()

        # add usernames from DB to choice box
        cursor = Database.connection.cursor()
        cursor.execute('SELECT name FROM employees WHERE activated = 1 ORDER BY name ASC')
        ctrl(self.login_frame, 'choice:name').AppendItems(zip(*cursor.fetchall())[0])

        self.login_frame.Bind(wx.EVT_CLOSE, self.quit_app)
        self.login_frame.Bind(wx.EVT_BUTTON, on_click_login, id=xrc.XRCID('button:log_in'))
        self.login_frame.Bind(wx.EVT_CHOICE, on_click_login, id=xrc.XRCID('choice:Plant'))
        #self.login_frame.Bind(wx.EVT_CHECKBOX, on_click_login, id=xrc.XRCID('m_checkBoxRemPass'))

        ###set temp defaults
        # ctrl(self.login_frame, 'choice:name').SetStringSelection(str('Stuart, Travis'))
        # ctrl(self.login_frame, 'text:password').SetValue(str('ts7587'))
        # ctrl(self.login_frame, 'choice:name').SetStringSelection(str('Keith, Brian'))
        # ctrl(self.login_frame, 'text:password').SetValue(str('bk475'))
        # ctrl(self.login_frame, 'choice:name').SetStringSelection(str('Williams, Richard'))
        # ctrl(self.login_frame, 'text:password').SetValue(str('Sup3rman42'))

        # default the login name to the last entered name
        login_name = ''
        remember_password = False
        password = ''
        plant = ''
        config = ConfigParser.ConfigParser()
        config.read('ECRev.cfg')
        if config.read('ECRev.cfg'):
            login_name = config.get('Application', 'login_name')
            remember_password = config.get('Application', 'remember_password')
            password = config.get('Application', 'password')
            plant = config.get('Application', 'plant')
            #print remember_password
            #print password

        if login_name != '':
            ctrl(self.login_frame, 'choice:name').SetStringSelection(login_name)

        if remember_password == 'True':
            ctrl(self.login_frame, 'm_checkBoxRemPass').SetValue(True)
            ctrl(self.login_frame, 'text:password').SetValue(password)
            ctrl(self.login_frame, 'choice:plant').SetStringSelection(plant)


        # put focus on password box
        ctrl(self.login_frame, 'text:password').SetFocus()

        # kill off the spash screen if it's still up
        for proc in psutil.process_iter():
            if proc.name == 'ECRev.exe':  # 'SplashScreenStarter.exe'
                proc.kill()

        # os.system("ps -C SplashScreenStarter -o pid=|xargs kill -9");

        self.login_frame.Show()

    def init_main_frame(self):
        start_time = int(round(time.time() * 1000))
        print '1: {}'.format(int(round(time.time() * 1000)) - start_time)
        start_time = int(round(time.time() * 1000))

        if self.main_frame != None:
            self.main_frame.Destroy()

        self.main_frame = self.res.LoadFrame(None, 'frame:main')
        self.main_frame.SetSize((950, 650))
        self.main_frame.Maximize()

        print '2: {}'.format(int(round(time.time() * 1000)) - start_time)
        start_time = int(round(time.time() * 1000))

        # add buttons to toolbar
        toolbar = ctrl(self.main_frame, 'toolBar')
        toolbar.AddLabelTool(id=9990, label='Log Out',
                             bitmap=wx.Bitmap(General.resource_path('icons\system-log-out.png')))
        self.main_frame.Bind(wx.EVT_TOOL, on_click_logout, id=9990)
        # toolbar.AddLabelTool(id=9991, label='Preferences', bitmap=wx.Bitmap(General.resource_path('icons\preferences-desktop.png')))
        # toolbar.AddLabelTool(id=9992, label='About', bitmap=wx.Bitmap(General.resource_path('icons\internet-news-reader.png')))
        ctrl(self.main_frame, 'toolBar').Realize()

        self.main_frame.Bind(wx.EVT_CLOSE, self.quit_app)

        print '3: {}'.format(int(round(time.time() * 1000)) - start_time)
        start_time = int(round(time.time() * 1000))

        # set window's title with user's name reordered as first then last name
        reordered_name = self.current_user.replace(' ', '')  # remove any spaces
        reordered_name = reordered_name.split(',')[1] + ' ' + reordered_name.split(',')[0]
        self.main_frame.SetTitle('ECRev v{} - Logged in as {} in {} Database'.format(version, reordered_name, Ecrs.Prod_Plant))
        ctrl(self.main_frame, 'm_textCtrlHeader').AppendText('You are in {} Database'.format(Ecrs.Prod_Plant))

        # initially populate some common lists
        Ecrs.refresh_my_ecrs_list(limit=15)
        print '4: {}'.format(int(round(time.time() * 1000)) - start_time)
        start_time = int(round(time.time() * 1000))

        Ecrs.refresh_open_ecrs_list(limit=15)
        print '5: {}'.format(int(round(time.time() * 1000)) - start_time)
        start_time = int(round(time.time() * 1000))

        Ecrs.refresh_closed_ecrs_list(limit=15)
        print '6: {}'.format(int(round(time.time() * 1000)) - start_time)
        start_time = int(round(time.time() * 1000))

        Ecrs.refresh_my_assigned_ecrs_list(limit=15)
        print '7: {}'.format(int(round(time.time() * 1000)) - start_time)
        start_time = int(round(time.time() * 1000))

        Ecrs.refresh_committee_ecrs_list()

        ##Ecrs.reset_search_fields()
        Ecrs.populate_ecr_order_panel(None)  # clear out info panel
        Ecrs.populate_ecr_panel(None)  # clear out info panel

        cursor = Database.connection.cursor()

        # hide some tabs based on user's department
        department = cursor.execute(
            'SELECT TOP 1 department FROM employees WHERE name = \'{}\''.format(General.app.current_user)).fetchone()[0]

        notebook = ctrl(self.main_frame, 'notebook:main')
        tabs_to_remove = []

        for index in range(notebook.GetPageCount()):
            if (department != 'Design Engineering') and (department != 'Applications Engineering'):
                if notebook.GetPageText(index).strip() == 'Track Time':
                    tabs_to_remove.append(index)
                    continue
                if notebook.GetPageText(index).strip() == 'Reports':
                    tabs_to_remove.append(index)
                    continue
                if notebook.GetPageText(index).strip() == 'Database':
                    tabs_to_remove.append(index)
                    continue
                if notebook.GetPageText(index).strip() == 'Administration':
                    tabs_to_remove.append(index)
                    continue

            if notebook.GetPageText(index).strip() == 'Notifications':
                tabs_to_remove.append(index)
                continue
                # if notebook.GetPageText(index).strip() == 'Search':
                #	tabs_to_remove.append(index)
                #	continue

        # let mike jones see reports... this shouldn't be done this way but this is a quick fix
        if self.current_user == "Jones, Mike":
            try:
                for index in range(notebook.GetPageCount()):
                    if notebook.GetPageText(index).strip() == 'Reports':
                        tabs_to_remove.remove(index)
            except Exception as e:
                print e

        removed_count = 0
        for index in tabs_to_remove:
            notebook.RemovePage(index - removed_count)
            removed_count += 1

        # now remove some from the ecrs tab
        notebook = ctrl(self.main_frame, 'notebook:ecrs')
        tabs_to_remove = []

        for index in range(notebook.GetPageCount()):
            if (department != 'Design Engineering') and (department != 'Applications Engineering'):
                if notebook.GetPageText(index).strip() == 'My Assigned ECRs':
                    tabs_to_remove.append(index)
                    continue

        removed_count = 0
        for index in tabs_to_remove:
            notebook.RemovePage(index - removed_count)
            removed_count += 1

        # say who the admin is in status bar
        admin_name = cursor.execute("SELECT ecr_admin_name FROM administration WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchone()[0]
        ctrl(self.main_frame, 'statusbar:main').SetStatusText('ECR administrator: {}'.format(admin_name))

        # show export for committee button if authorized
        can_approve_first, can_approve_second = cursor.execute(
            "SELECT can_approve_first, can_approve_second FROM employees WHERE name = '{}'".format(
                self.current_user)).fetchone()
        if can_approve_first or can_approve_second:
            ctrl(self.main_frame, 'button:export_for_committee').Enable()

        else:
            ctrl(self.main_frame, 'button:approve_selected_ecrs').Disable()

        self.init_ecrs_tab()
        self.init_revisions_tab()
        # self.init_time_logs_tab()
        self.init_search_tab()
        self.init_reports_tab()
        self.init_administration_tab()

        # self.main_frame.Bind(wx.EVT_BUTTON, self.OnSubmit, id=xrc.XRCID('button:submit_ecr'))

        ###
        self.main_frame.Show()

    def init_close_ecr_dialog(self, close_ecr_id):
        if close_ecr_id == '':
            return

        self.close_ecr_dialog = self.res.LoadDialog(None, 'dialog:edit_ecr')
        self.close_ecr_dialog.SetTitle('Close ECR: {}'.format(close_ecr_id))
        # self.close_ecr_dialog.SetSize((750, 550))

        #Hide Attach Documents Stuff
        ctrl(self.close_ecr_dialog, 'button:Attach').Hide()
        ctrl(self.close_ecr_dialog, 'text:m_AttachList').Hide()
        ctrl(self.close_ecr_dialog, 'text:m_AttachListPaths').Hide()

        # Bind Do_Nothing Event upon mousewheel scroll in order to not change users Dropdowns selection accidently
        ctrl(self.close_ecr_dialog, 'choice:ecr_reason').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.close_ecr_dialog, 'choice:ecr_document').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.close_ecr_dialog, 'choice:ecr_component').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.close_ecr_dialog, 'choice:ecr_sub_system').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.close_ecr_dialog, 'choice:who_errored').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.close_ecr_dialog, 'choice:stage').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)

        # Hide workflow and Reopen Button for now until it is ready for release
        ctrl(self.close_ecr_dialog, 'm_panelWorkflow').Hide()
        ctrl(self.close_ecr_dialog, 'm_buttonAssign').Hide()

        ctrl(self.close_ecr_dialog, 'm_buttonAssign').Disable()
        ctrl(self.close_ecr_dialog, 'm_buttonReopen').Hide()

        cursor = Database.connection.cursor()

        # Hide Assign Workflow Button if user in Systems Plant
        if Ecrs.Prod_Plant == 'Systems':
            ctrl(self.close_ecr_dialog, 'm_buttonAssign').Hide()
            ctrl(self.close_ecr_dialog, 'm_panelWorkflow').Hide()

        try:
            workflow_exists = cursor.execute('Select top 1 step_no from Ecrev_Status where Ecrev_no =?',close_ecr_id).fetchone()[0]
            if workflow_exists:
                workflow = True
        except:
            workflow = False

        if workflow:
            workflow_info = cursor.execute('Select Assigned_to, Step_description, current_Status from Ecrev_Status where Ecrev_no = ?',close_ecr_id).fetchall()
            print "Yass"
            ctrl(General.app.close_ecr_dialog, 'm_textStep1').SetValue(workflow_info[0][1])
            ctrl(General.app.close_ecr_dialog, 'm_textStep2').SetValue(workflow_info[1][1])
            ctrl(General.app.close_ecr_dialog, 'm_textStep3').SetValue(workflow_info[2][1])
            ctrl(General.app.close_ecr_dialog, 'm_textStep4').SetValue(workflow_info[3][1])
            ctrl(General.app.close_ecr_dialog, 'm_textStep5').SetValue(workflow_info[4][1])

            ctrl(General.app.close_ecr_dialog, 'm_textCtrlWho1').SetValue(workflow_info[0][0])
            ctrl(General.app.close_ecr_dialog, 'm_textCtrlWho2').SetValue(workflow_info[1][0])
            ctrl(General.app.close_ecr_dialog, 'm_textCtrlWho3').SetValue(workflow_info[2][0])
            ctrl(General.app.close_ecr_dialog, 'm_textCtrlWho4').SetValue(workflow_info[3][0])
            ctrl(General.app.close_ecr_dialog, 'm_textCtrlWho5').SetValue(workflow_info[4][0])

            if workflow_info[0][2] == 'Completed':
                ctrl(General.app.close_ecr_dialog, 'm_checkStep1').SetValue(True)
            ctrl(General.app.close_ecr_dialog, 'm_checkStep1').Disable()

            if workflow_info[1][2] == 'Completed':
                ctrl(General.app.close_ecr_dialog, 'm_checkStep2').SetValue(True)
            ctrl(General.app.close_ecr_dialog, 'm_checkStep2').Disable()

            if workflow_info[2][2] == 'Completed':
                ctrl(General.app.close_ecr_dialog, 'm_checkStep3').SetValue(True)
            ctrl(General.app.close_ecr_dialog, 'm_checkStep3').Disable()

            if workflow_info[3][2] == 'Completed':
                ctrl(General.app.close_ecr_dialog, 'm_checkStep4').SetValue(True)
            ctrl(General.app.close_ecr_dialog, 'm_checkStep4').Disable()

            if workflow_info[4][2] == 'Completed':
                ctrl(General.app.close_ecr_dialog, 'm_checkStep5').SetValue(True)
            ctrl(General.app.close_ecr_dialog, 'm_checkStep5').Disable()
        else:
            ctrl(General.app.close_ecr_dialog, 'm_checkStep1').Disable()
            ctrl(General.app.close_ecr_dialog, 'm_checkStep2').Disable()
            ctrl(General.app.close_ecr_dialog, 'm_checkStep3').Disable()
            ctrl(General.app.close_ecr_dialog, 'm_checkStep4').Disable()
            ctrl(General.app.close_ecr_dialog, 'm_checkStep5').Disable()



        # show committee panel if authorized
        can_approve_first = cursor.execute(
            "SELECT can_approve_first FROM employees WHERE name = '{}'".format(self.current_user)).fetchone()[0]
        if can_approve_first:
            ctrl(self.close_ecr_dialog, 'spin:priority').Enable()
            ctrl(self.close_ecr_dialog, 'choice:stage').Enable()
        # ctrl(self.close_ecr_dialog, 'button:approve_1').Enable()

        can_approve_second = cursor.execute(
            "SELECT can_approve_second FROM employees WHERE name = '{}'".format(self.current_user)).fetchone()[0]
        if can_approve_second:
            # ctrl(self.close_ecr_dialog, 'panel:committee').Enable()
            # ctrl(self.close_ecr_dialog, 'button:approve_1').Show()
            ctrl(self.close_ecr_dialog, 'spin:priority').Enable()
            ctrl(self.close_ecr_dialog, 'choice:stage').Enable()
        # ctrl(self.close_ecr_dialog, 'button:approve_2').Enable()

        if not can_approve_first and not can_approve_second:
            ctrl(self.close_ecr_dialog, 'button:approve_1').Hide()
            ctrl(self.close_ecr_dialog, 'button:approve_2').Hide()
        # ctrl(self.close_ecr_dialog, 'panel:committee').Hide()

        ctrl(self.close_ecr_dialog, 'choice:stage').AppendItems(
            ['New Request, needs reviewing', 'Reviewed, engineering in process', 'Change Complete, pending approval',
             'Approved', 'Prototype Stage'])

        # add document options to choice box
        cursor.execute(
            "SELECT document FROM ecr_document_choices where Production_Plant = \'{}\'".format(Ecrs.Prod_Plant))
        ctrl(self.close_ecr_dialog, 'choice:ecr_document').AppendItems(zip(*cursor.fetchall())[0])

        ##self.ecr_type = 'Mechanical'

        self.close_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Mechanical'),
                                   id=xrc.XRCID('radio:mechanical'))
        self.close_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Electrical'),
                                   id=xrc.XRCID('radio:electrical'))
        self.close_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Structural'),
                                   id=xrc.XRCID('radio:structural'))
        self.close_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Other'),
                                   id=xrc.XRCID('radio:other'))
        # self.close_ecr_dialog.Bind(wx.EVT_TEXT, Ecrs.check_reference_field, id=xrc.XRCID('text:reference_number'))
        # self.close_ecr_dialog.Bind(wx.EVT_TEXT, Ecrs.on_text_ecr_description, id=xrc.XRCID('text:description'))
        # self.close_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_select_ecr_reason, id=xrc.XRCID('choice:ecr_reason'))
        # self.close_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_submit_ecr, id=xrc.XRCID('button:submit_ecr'))
        # self.close_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_attatch_document, id=xrc.XRCID('button:attach_document'))

        self.close_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_close_ecr, id=xrc.XRCID('button:modify_or_close_ecr'))

        self.close_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_approve_1_for_close, id=xrc.XRCID('button:approve_1'))
        self.close_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_approve_2_for_close, id=xrc.XRCID('button:approve_2'))

        self.close_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_choice_set_severity_default,
                                   id=xrc.XRCID('choice:ecr_document'))
        self.close_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_choice_set_severity_default,
                                   id=xrc.XRCID('choice:ecr_component'))
        self.close_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_choice_set_severity_default,
                                   id=xrc.XRCID('choice:ecr_sub_system'))

        # fill in the fields of the ecr we are modifing
        ecr = cursor.execute(
            "SELECT reference_number, document, reason, type, request, resolution, when_needed, who_errored, priority, who_approved_first, who_approved_second, approval_stage, item, component, sub_system, severity, Units_Affected FROM ecrs WHERE id = \'{}\'".format(
                close_ecr_id)).fetchone()
        reference_number, document, reason, type, request, resolution, when_needed, who_errored, priority, who_approved_first, who_approved_second, approval_stage, item, component, sub_system, severity, Units_Affected = ecr

        # add component options
        if Ecrs.Prod_Plant == 'Systems':
            if type == 'Other':
                components = list(zip(*cursor.execute("SELECT DISTINCT component FROM ecr.components WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0])
            else:
                components = list(zip(*cursor.execute("SELECT DISTINCT component FROM ecr.components WHERE Production_Plant = \'{}\' AND discipline= \'{}\'".format(Ecrs.Prod_Plant, type)).fetchall())[0])

            components.insert(0, '')
            ctrl(self.close_ecr_dialog, 'choice:ecr_component').AppendItems(components)

            if component:
                ctrl(self.close_ecr_dialog, 'choice:ecr_component').Insert(component, 1)
                ctrl(self.close_ecr_dialog, 'choice:ecr_component').SetStringSelection(component)

        # add sub_system options
            if type == 'Other':
                sub_systems = list(zip(*cursor.execute("SELECT DISTINCT sub_system FROM ecr.sub_systems WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0])
            else:
                sub_systems = list(zip(*cursor.execute("SELECT DISTINCT sub_system FROM ecr.sub_systems WHERE Production_Plant = \'{}\' AND discipline=\'{}\'".format(Ecrs.Prod_Plant, type)).fetchall())[0])

            sub_systems.insert(0, '')
            ctrl(self.close_ecr_dialog, 'choice:ecr_sub_system').AppendItems(sub_systems)

            if sub_system:
                ctrl(self.close_ecr_dialog, 'choice:ecr_sub_system').Insert(sub_system, 1)
                ctrl(self.close_ecr_dialog, 'choice:ecr_sub_system').SetStringSelection(sub_system)
        else:
            if type == 'Other':
                components = list(zip(*cursor.execute("SELECT DISTINCT component FROM ecr.components WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0])
            else:
                components = list(zip(*cursor.execute(
                    "SELECT DISTINCT component FROM ecr.components WHERE Production_Plant = \'{}\' AND discipline=\'{}\'".format(Ecrs.Prod_Plant, type)).fetchall())[0])

            components.insert(0, '')
            ctrl(self.close_ecr_dialog, 'choice:ecr_component').AppendItems(components)

            if component:
                ctrl(self.close_ecr_dialog, 'choice:ecr_component').Insert(component, 1)
                ctrl(self.close_ecr_dialog, 'choice:ecr_component').SetStringSelection(component)

                # add sub_system options
            if type == 'Other':
                sub_systems = list(
                    zip(*cursor.execute("SELECT DISTINCT sub_system FROM ecr.sub_systems WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0])
            else:
                sub_systems = list(zip(*cursor.execute(
                    "SELECT DISTINCT sub_system FROM ecr.sub_systems WHERE Production_Plant = \'{}\' AND discipline=\'{}\'".format(Ecrs.Prod_Plant, type)).fetchall())[
                                       0])

            sub_systems.insert(0, '')
            ctrl(self.close_ecr_dialog, 'choice:ecr_sub_system').AppendItems(sub_systems)

            if sub_system:
                ctrl(self.close_ecr_dialog, 'choice:ecr_sub_system').Insert(sub_system, 1)
                ctrl(self.close_ecr_dialog, 'choice:ecr_sub_system').SetStringSelection(sub_system)

            """components_list_cases = ['Air block', 'Air Deflector', 'Base', 'Brackets', 'Breaker', 'Bumper/retainer',
                                     'Coil',
                                     'Controller', 'Deck pans', 'Door/frame', 'End Assy', 'Fan', 'Glass', 'Horse Head',
                                     'Kick plates', 'Lights', 'Other', 'Painted part', 'Piping', 'Pnl, Foam, Back',
                                     'Pnl, Foam, Cnpy', 'Pnl, Foam, Front', 'PVC', 'Raceway', ' Rack', 'Sensor, Temp',
                                     'Sensor, Pressure', 'Shelf standard', 'Shelves', 'Skin', 'Tub, foam', 'Valve',
                                     'Wire racks']
            sub_system_list_cases = ['Base', 'Coil Piping', 'Controls', 'Doors/frames', 'End', 'Foam', 'Kitting',
                                     'Knock up',
                                     'Lighting', 'Paint', 'Piping', 'Piping Option Pack', 'QA', 'Raceway',
                                     'Sheet Metal',
                                     'Subassy', 'Trimming', 'Wiring']

            # components_list_cases.insert(0, '')
            ctrl(self.close_ecr_dialog, 'choice:ecr_component').AppendItems(components_list_cases)
            # ctrl(self.modify_ecr_dialog, 'choice:ecr_component').Insert(components_list_cases, 1)
            # ctrl(self.modify_ecr_dialog, 'choice:ecr_component').SetStringSelection(components_list_cases)

            # sub_system_list_cases.insert(0, '')
            ctrl(self.close_ecr_dialog, 'choice:ecr_sub_system').AppendItems(sub_system_list_cases)
            # ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').Insert(sub_system_list_cases, 1)
            # ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').SetStringSelection(sub_system_list_cases)
            """

        # add severity options
        ctrl(self.close_ecr_dialog, 'choice:ecr_severity').AppendItems(['High', 'Medium', 'Low'])

        severity = float(severity)

        if severity == 1.0:
            ctrl(self.close_ecr_dialog, 'choice:ecr_severity').SetStringSelection('High')
        elif severity == 0.5:
            ctrl(self.close_ecr_dialog, 'choice:ecr_severity').SetStringSelection('Medium')
        elif severity == 0.1:
            ctrl(self.close_ecr_dialog, 'choice:ecr_severity').SetStringSelection('Low')
        else:
            ctrl(self.close_ecr_dialog, 'choice:ecr_severity').SetStringSelection('High')

        # add reason for request options to choice box
        ##cursor.execute("SELECT reason FROM ecr_reason_choices ORDER BY reason ASC")
        ##ctrl(self.modify_ecr_dialog, 'choice:ecr_reason').AppendItems(zip(*cursor.fetchall())[0])
        if Ecrs.Prod_Plant == 'Systems':
            reasons = list(zip(*cursor.execute("SELECT code FROM secondary_ecr_reason_codes ORDER BY code ASC").fetchall())[0])

            ctrl(self.close_ecr_dialog, 'choice:ecr_reason').AppendItems(reasons)

        # support old ECR reason codes if ECR originally submitted under them
            if reason not in reasons:
                cursor.execute("SELECT reason FROM ecr_reason_choices where Production_Plant = \'{}\' ORDER BY reason ASC".format(Ecrs.Prod_Plant))
                ctrl(self.close_ecr_dialog, 'choice:ecr_reason').AppendItems(zip(*cursor.fetchall())[0])

        # uncheck send similar notices if it's a BOM Rec
            if reason == 'BOM Reconciliation':
                ctrl(self.close_ecr_dialog, 'checkbox:similar_ecrs').SetValue(False)
        else:
            reasons = zip(*cursor.execute(
                "SELECT reason FROM ecr_reason_choices where Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0]
            ctrl(self.close_ecr_dialog, 'choice:ecr_reason').AppendItems((reasons))

        # no actually, uncheck send similar notices by default until this SAP transition mess is over
        ctrl(self.close_ecr_dialog, 'checkbox:similar_ecrs').SetValue(False)

        # find items that could be similarly affected by the issue in this ECR
        similar_items = Ecrs.get_similar_items(item)
        if similar_items:
            ctrl(self.close_ecr_dialog, 'text:similar').SetValue(similar_items[0])

        # hide committee panel if not the right type of ECR reason
        if reason not in Ecrs.reasons_needing_approval:
           ctrl(self.close_ecr_dialog, 'panel:committee').Hide()
        try:
            ctrl(self.close_ecr_dialog, 'choice:stage').SetStringSelection(approval_stage)
        except Exception as e:
            print 'Error, failed to set approval stage:', e

        ctrl(self.close_ecr_dialog, 'text:reference_number').SetValue(str(ecr[0]))

        if ecr[3] == 'Mechanical':
            ctrl(self.close_ecr_dialog, 'radio:mechanical').SetValue(True)
            self.ecr_type = 'Mechanical'
        if ecr[3] == 'Electrical':
            ctrl(self.close_ecr_dialog, 'radio:electrical').SetValue(True)
            self.ecr_type = 'Electrical'
        if ecr[3] == 'Structural':
            ctrl(self.close_ecr_dialog, 'radio:structural').SetValue(True)
            self.ecr_type = 'Structural'
        if ecr[3] == 'Other':
            ctrl(self.close_ecr_dialog, 'radio:other').SetValue(True)
            self.ecr_type = 'Other'
        if ecr[3] == '*':
            ctrl(self.close_ecr_dialog, 'radio:other').SetValue(True)
            self.ecr_type = '*'

        try:
            ctrl(self.close_ecr_dialog, 'choice:ecr_reason').SetStringSelection(ecr[2])
            ctrl(self.close_ecr_dialog, 'choice:ecr_reason').Focus()
        except:
            pass
        try:
            ctrl(self.close_ecr_dialog, 'choice:ecr_document').SetStringSelection(ecr[1])
        except:
            pass

        if ecr[4] != None: ctrl(self.close_ecr_dialog, 'text:description').SetValue(ecr[4])
        if ecr[5] != None: ctrl(self.close_ecr_dialog, 'text:resolution').SetValue(ecr[5])
        if ecr[16] != None: ctrl(self.close_ecr_dialog, 'm_NoUnitsAffected').SetValue(ecr[16])

        # paste in what was typed in the revisions for this ecr if there are any
        if ctrl(self.close_ecr_dialog, 'text:resolution').GetValue() == '':
            default_resolution = ''
            revisions = cursor.execute(
                "SELECT document, description FROM revisions WHERE related_ecr='{}'".format(close_ecr_id)).fetchall()
            for revision in revisions:
                default_resolution += '{}: {}\n'.format(revision[0], revision[1])

            ctrl(self.close_ecr_dialog, 'text:resolution').SetValue(default_resolution[:-1])

        need_by_date = time.strptime(str(ecr[6]), "%Y-%m-%d %H:%M:%S")  # to python time object
        ctrl(self.close_ecr_dialog, 'calendar:ecr_need_by').SetDate(
            wx.DateTimeFromDMY(need_by_date.tm_mday, need_by_date.tm_mon - 1, need_by_date.tm_year))

        engineers = list(zip(*cursor.execute(
            "SELECT name FROM employees WHERE department LIKE '%Engineering%' ORDER BY name ASC").fetchall())[0])
        engineers.insert(0, '')
        ctrl(self.close_ecr_dialog, 'choice:who_errored').AppendItems(engineers)

        if ecr[7] != '' and ecr[7] != None:
            ctrl(self.close_ecr_dialog, 'choice:who_errored').SetStringSelection(ecr[7])

        ctrl(self.close_ecr_dialog, 'button:modify_or_close_ecr').SetLabel('Close Out')

        # committee stuff
        ctrl(self.close_ecr_dialog, 'spin:priority').SetValue(priority)
        if who_approved_first:
            ctrl(self.close_ecr_dialog, 'label:who_approved_first').SetLabel(who_approved_first)
            ctrl(self.close_ecr_dialog, 'button:approve_1').SetLabel('Unapprove')

        if who_approved_second:
            ctrl(self.close_ecr_dialog, 'label:who_approved_second').SetLabel(who_approved_second)
            ctrl(self.close_ecr_dialog, 'button:approve_2').SetLabel('Unapprove')

        #ctrl(self.close_ecr_dialog, 'text:resolution').GetParent().GetSizer().Layout()
        self.close_ecr_dialog.Layout()
        self.close_ecr_dialog.Fit()

        ctrl(self.close_ecr_dialog, 'panel:committee').Layout()

        self.close_ecr_dialog.ShowModal()

    def init_administration_tab(self):
        # add engineering names to choice box
        cursor = Database.connection.cursor()
        cursor.execute(
            'SELECT name FROM employees WHERE activated = 1 AND department = \'Design Engineering\' ORDER BY name ASC')
        ctrl(self.main_frame, 'choice:admin').AppendItems(zip(*cursor.fetchall())[0])

        # set current admin as current selection
        ctrl(self.main_frame, 'choice:admin').SetStringSelection(
            cursor.execute("SELECT ecr_admin_name FROM administration").fetchone()[0])

        self.main_frame.Bind(wx.EVT_CHOICE, Administration.on_select_admin, id=xrc.XRCID('choice:admin'))
        self.main_frame.Bind(wx.EVT_BUTTON, Administration.on_click_backup_database,
                             id=xrc.XRCID('button:backup_database'))
        # Bind Do_Nothing Event upon mousewheel scroll in order to not change users Dropdowns selection accidently
        ctrl(self.main_frame, 'choice:admin').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)

    def init_ecrs_tab(self):
        cursor = Database.connection.cursor()

        self.main_frame.Bind(wx.EVT_BUTTON, open_folder, id=xrc.XRCID('button:open_folder'))
        self.main_frame.Bind(wx.EVT_BUTTON, open_bom, id=xrc.XRCID('button:open_bom'))
        self.main_frame.Bind(wx.EVT_BUTTON, open_piping, id=xrc.XRCID('button:open_piping'))
        self.main_frame.Bind(wx.EVT_BUTTON, open_wiring, id=xrc.XRCID('button:open_wiring'))
        self.main_frame.Bind(wx.EVT_BUTTON, open_dataplate, id=xrc.XRCID('button:open_dataplate'))
        self.main_frame.Bind(wx.EVT_BUTTON, open_workbook, id=xrc.XRCID('button:open_workbook'))
        self.main_frame.Bind(wx.EVT_BUTTON, open_legend, id=xrc.XRCID('button:open_legend'))

        if Ecrs.Prod_Plant == "Cases":
            ctrl(General.app.main_frame, 'button:open_folder').Hide()
            ctrl(General.app.main_frame, 'button:open_bom').Hide()
            ctrl(General.app.main_frame, 'button:open_piping').Hide()
            ctrl(General.app.main_frame, 'button:open_wiring').Hide()
            ctrl(General.app.main_frame, 'button:open_dataplate').Hide()
            ctrl(General.app.main_frame, 'button:open_workbook').Hide()
            ctrl(General.app.main_frame, 'button:open_legend').Hide()

        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_add_revisions_with_ecr,
                             id=xrc.XRCID('button:add_revisions_with_ecr'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_open_new_ecr_form, id=xrc.XRCID('button:create_new_ecr'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_open_search_ecrs_form, id=xrc.XRCID('button:search_ecrs'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_open_email_ecr_form, id=xrc.XRCID('button:email_ecr'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_open_duplicate_ecr_form, id=xrc.XRCID('button:duplicate_ecr'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_open_attachments, id=xrc.XRCID('button:open_attachment'))

        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_claim_ecr, id=xrc.XRCID('button:claim'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_open_modify_ecr_form, id=xrc.XRCID('button:modify'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_open_assign_ecr_form, id=xrc.XRCID('button:assign'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_open_close_ecr_form, id=xrc.XRCID('button:close'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_print_ecr, id=xrc.XRCID('button:print'))

        Ecrs.hide_things_based_on_user_department()

        ctrl(self.main_frame, 'text:ecr_panel_description').SetBackgroundColour(
            ctrl(self.main_frame, 'notebook:main').GetBackgroundColour())
        ctrl(self.main_frame, 'text:ecr_panel_resolution').SetBackgroundColour(
            ctrl(self.main_frame, 'notebook:main').GetBackgroundColour())

        ###tab: New ECR form
        '''
        #add reason for request options to choice box
        cursor.execute("SELECT reason FROM ecr_reason_choices ORDER BY reason ASC")
        ctrl(self.main_frame, 'choice:ecr_reason').AppendItems(zip(*cursor.fetchall())[0])

        #add document options to choice box
        cursor.execute("SELECT document FROM ecr_document_choices WHERE type=\'Mechanical\' OR type=\'*\'")
        ctrl(self.main_frame, 'choice:ecr_document').AppendItems(zip(*cursor.fetchall())[0])

        self.main_frame.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Mechanical'), id=xrc.XRCID('radio:mechanical'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Electrical'), id=xrc.XRCID('radio:electrical'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Structural'), id=xrc.XRCID('radio:structural'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Other'), id=xrc.XRCID('radio:other'))
        self.main_frame.Bind(wx.EVT_TEXT, Ecrs.check_reference_field, id=xrc.XRCID('text:reference_number'))
        self.main_frame.Bind(wx.EVT_CHOICE, Ecrs.on_select_ecr_reason, id=xrc.XRCID('choice:ecr_reason'))
        '''

        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.export_for_approval, id=xrc.XRCID('button:export_for_committee'))

        ###tab: My ECRs
        self.main_frame.Bind(wx.EVT_LIST_ITEM_SELECTED, Ecrs.on_select_ecr_item, id=xrc.XRCID('list:my_ecrs'))
        ###self.main_frame.Bind(wx.EVT_LIST_COL_CLICK, sort_list, id=xrc.XRCID('list:my_ecrs'))
        # self.main_frame.Bind(wx.EVT_BUTTON, partial(Ecrs.make_dialog_email_ecr, tab='my_ecrs'), id=xrc.XRCID('button:my_ecrs_email_admin'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.refresh_my_ecrs_list, id=xrc.XRCID('button:my_ecrs_refresh'))
        ctrl(self.main_frame, 'list:my_ecrs').printer_paper_type = wx.PAPER_11X17

        # item = event.GetEventObject()
        # ctrl(self.main_frame, 'list:my_ecrs').GetItem(ctrl(self.main_frame, 'list:my_ecrs').GetFirstSelected(), 1).GetText()
        # ctrl(self.main_frame, 'list:my_ecrs').getColumnText(ctrl(self.main_frame, 'list:my_ecrs').currentItem, 1)

        ###tab: Open ECRs
        self.main_frame.Bind(wx.EVT_LIST_ITEM_SELECTED, Ecrs.on_select_ecr_item, id=xrc.XRCID('list:open_ecrs'))
        ###self.main_frame.Bind(wx.EVT_LIST_COL_CLICK, sort_list, id=xrc.XRCID('list:open_ecrs'))
        # self.main_frame.Bind(wx.EVT_BUTTON, partial(Ecrs.make_dialog_email_ecr, tab='open_ecrs'), id=xrc.XRCID('button:open_ecrs_email_admin'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.refresh_open_ecrs_list, id=xrc.XRCID('button:open_ecrs_refresh'))
        ctrl(self.main_frame, 'list:open_ecrs').printer_paper_type = wx.PAPER_11X17

        ###tab: Closed ECRs
        self.main_frame.Bind(wx.EVT_LIST_ITEM_SELECTED, Ecrs.on_select_ecr_item, id=xrc.XRCID('list:closed_ecrs'))
        ###self.main_frame.Bind(wx.EVT_LIST_COL_CLICK, sort_list, id=xrc.XRCID('list:closed_ecrs'))
        # self.main_frame.Bind(wx.EVT_BUTTON, partial(Ecrs.make_dialog_email_ecr, tab='closed_ecrs'), id=xrc.XRCID('button:closed_ecrs_email_admin'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.refresh_closed_ecrs_list, id=xrc.XRCID('button:closed_ecrs_refresh'))
        ctrl(self.main_frame, 'list:closed_ecrs').printer_paper_type = wx.PAPER_11X17

        ###tab: My Assigned ECRs
        self.main_frame.Bind(wx.EVT_LIST_ITEM_SELECTED, Ecrs.on_select_ecr_item, id=xrc.XRCID('list:my_assigned_ecrs'))
        ###self.main_frame.Bind(wx.EVT_LIST_COL_CLICK, sort_list, id=xrc.XRCID('list:my_assigned_ecrs'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.refresh_my_assigned_ecrs_list,
                             id=xrc.XRCID('button:my_assigned_ecrs_refresh'))
        ctrl(self.main_frame, 'list:my_assigned_ecrs').printer_paper_type = wx.PAPER_11X17

        ###tab: Committee ECRs
        self.main_frame.Bind(wx.EVT_LIST_ITEM_SELECTED, Ecrs.on_select_ecr_item, id=xrc.XRCID('list:committee_ecrs'))
        ###self.main_frame.Bind(wx.EVT_LIST_COL_CLICK, sort_list, id=xrc.XRCID('list:committee_ecrs'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.on_click_approve_selected_ers,
                             id=xrc.XRCID('button:approve_selected_ecrs'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.refresh_committee_ecrs_list,
                             id=xrc.XRCID('button:committee_ecrs_refresh'))
        ctrl(self.main_frame, 'list:committee_ecrs').printer_paper_type = wx.PAPER_11X17

        ###tab: Search ECRs
        ###self.main_frame.Bind(wx.EVT_LIST_ITEM_SELECTED, Ecrs.on_select_ecr_item, id=xrc.XRCID('list:results'))
        ###self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.export_search_results, id=xrc.XRCID('button:export'))

        self.main_frame.Bind(wx.EVT_LIST_ITEM_ACTIVATED, Ecrs.on_activated_ecr, id=xrc.XRCID('list:my_ecrs'))
        self.main_frame.Bind(wx.EVT_LIST_ITEM_ACTIVATED, Ecrs.on_activated_ecr, id=xrc.XRCID('list:open_ecrs'))
        self.main_frame.Bind(wx.EVT_LIST_ITEM_ACTIVATED, Ecrs.on_activated_ecr, id=xrc.XRCID('list:closed_ecrs'))
        self.main_frame.Bind(wx.EVT_LIST_ITEM_ACTIVATED, Ecrs.on_activated_ecr, id=xrc.XRCID('list:my_assigned_ecrs'))
        self.main_frame.Bind(wx.EVT_LIST_ITEM_ACTIVATED, Ecrs.on_activated_ecr, id=xrc.XRCID('list:committee_ecrs'))


        '''
        #the search panel is initialialy hidden as described by the GUI XRC. This is because
        # for some reason the window gets force to tall otherwise. So gotta call Show() on it.
        ctrl(self.main_frame, 'scrolled_window:search').SetScrollRate(1, 10)
        ctrl(self.main_frame, 'scrolled_window:search').Show()
        ctrl(self.main_frame, 'scrolled_window:search').GetParent().Layout()
        
        #populate some search field choices
        ctrl(self.main_frame, 'combo:search_value3').AppendItems(zip(*cursor.execute("SELECT document FROM ecr_document_choices").fetchall())[0])
        ctrl(self.main_frame, 'combo:search_value4').AppendItems(zip(*cursor.execute("SELECT reason FROM ecr_reason_choices").fetchall())[0])
        ctrl(self.main_frame, 'combo:search_value5').AppendItems(zip(*cursor.execute("SELECT department FROM departments").fetchall())[0])
        ctrl(self.main_frame, 'combo:search_value6').AppendItems(zip(*cursor.execute("SELECT name FROM employees ORDER BY name ASC").fetchall())[0])
        ctrl(self.main_frame, 'combo:search_value7').AppendItems(zip(*cursor.execute("SELECT type FROM ecr_types").fetchall())[0])
        ctrl(self.main_frame, 'combo:search_value10').AppendItems(zip(*cursor.execute("SELECT name FROM employees ORDER BY name ASC").fetchall())[0])
        ctrl(self.main_frame, 'combo:search_value11').AppendItems(zip(*cursor.execute("SELECT name FROM employees ORDER BY name ASC").fetchall())[0])
        ctrl(self.main_frame, 'combo:search_value12').AppendItems(zip(*cursor.execute("SELECT name FROM employees ORDER BY name ASC").fetchall())[0])
        
        self.main_frame.Bind(wx.EVT_LIST_ITEM_SELECTED, Ecrs.on_select_ecr_item, id=xrc.XRCID('list:results'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.search_ecrs, id=xrc.XRCID('button:search'))
        self.main_frame.Bind(wx.EVT_BUTTON, Ecrs.export_search_results, id=xrc.XRCID('button:export'))
        

        for i in range(11):
            self.main_frame.Bind(wx.EVT_CHOICE, partial(Ecrs.search_condition_selected, index=i), id=xrc.XRCID('choice:search_condition'+str(i)))
            self.main_frame.Bind(wx.EVT_TEXT, partial(Ecrs.search_value_entered, index=i), id=xrc.XRCID('combo:search_value'+str(i)))
        '''

        ###dialog:email ecr
        self.email_ecr_dialog = None  # self.res.LoadDialog(None, 'dialog:email_ecr')

    def init_email_ecr_dialog(self):
        self.email_ecr_dialog = self.res.LoadDialog(None, 'dialog:email_ecr')

        cursor = Database.connection.cursor()

        # get ecr admin
        admin_name = cursor.execute("SELECT ecr_admin_name FROM administration WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchone()[0]
        ##admin_email = cursor.execute("SELECT email FROM employees WHERE name = \'{}\'".format(admin_name)).fetchone()[0]
        ##user_email = cursor.execute("SELECT email FROM employees WHERE name = \'{}\'".format(self.current_user)).fetchone()[0]

        # add names from DB to recipient choice box (default to ecr admin)
        cursor.execute("SELECT name FROM employees ORDER BY name ASC")
        ctrl(self.email_ecr_dialog, 'choice:to').AppendItems(zip(*cursor.fetchall())[0])
        ctrl(self.email_ecr_dialog, 'choice:to').SetStringSelection(admin_name)

        ecr_id = ctrl(self.main_frame, 'label:ecr_panel_id').GetLabel()
        if ecr_id == '':
            return

        # get more data on the ecr's order
        ecr = cursor.execute("SELECT * FROM ecrs WHERE id = \'{}\'".format(ecr_id)).fetchone()
        # order = Database.get_order_data_from_ref(ecr[2])
        order = cursor.execute("SELECT TOP 1 * FROM {} WHERE item = \'{}\'".format(Ecrs.table_used, ecr[3])).fetchone()
        ecr_details = ''
        if order != None:
            ecr_details = "ECR ID: " + str(ecr[0]) + \
                          "\nItem:   " + str(order[0]) + \
                          "\nSales:  " + str(order[1]) + "-" + str(order[2]) + \
                          "\n-----" + \
                          "\nRequest:\n	" + str(ecr[8]) + \
                          "\n\nResolution:\n	" + str(ecr[9]) + \
                          "\n-----" + \
                          "\nModel: " + str(order[11]) + \
                          "\nCust:  " + str(order[5]) + \
                          "\nStore: " + str(order[6])

        ###set temp defaults
        # ctrl(self.email_ecr_dialog, 'label:admin_email').SetLabel(admin_email +" (ECR Admin)")
        ctrl(self.email_ecr_dialog, 'label:user').SetLabel(self.current_user)
        ctrl(self.email_ecr_dialog, 'text:email_ecr_details').SetValue(ecr_details)
        ctrl(self.email_ecr_dialog, 'text:email_message').SetFocus()

        self.email_ecr_dialog.Bind(wx.EVT_BUTTON, partial(destroy_dialog, dialog=self.email_ecr_dialog),
                                   id=xrc.XRCID('button:cancel_ecr_email'))

        self.email_ecr_dialog.ShowModal()
        self.email_ecr_dialog.Destroy()

    def init_new_ecr_dialog(self, duplicate_ecr_id=None):
        self.new_ecr_dialog = self.res.LoadDialog(None, 'dialog:new_ecr')
        self.new_ecr_dialog.SetSize((850, 650))

        # Bind Do_Nothing Event upon mousewheel scroll in order to not change users Dropdowns selection accidently
        ctrl(self.new_ecr_dialog, 'choice:ecr_reason').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.new_ecr_dialog, 'choice:ecr_document').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        #ctrl(self.modify_ecr_dialog, 'choice:ecr_component').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        #ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        #ctrl(self.modify_ecr_dialog, 'choice:who_errored').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        #ctrl(self.modify_ecr_dialog, 'choice:stage').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)

        # DBworks connection for checking part numbers entered in ecr description
        try:
            ##1/0.
            self.dbworks_connection = pyodbc.connect(
                'DSN=DBWorks;UID=rmiller;APP=ECRev;WSID=SEN30;DATABASE=DBWorks_Kysor;Trusted_Connection=Yes')
            self.dbworks_cursor = self.dbworks_connection.cursor()
        except:
            self.dbworks_connection = None
            self.dbworks_cursor = None
            print 'ERROR: Could not connect to DBWorks database.'

        self.list_of_checked_words_entered = [(None, None)]

        cursor = Database.connection.cursor()

        if Ecrs.Prod_Plant == 'Systems':

            # add reason for request options to choice box
            ##reasons = zip(*cursor.execute("SELECT reason FROM ecr_reason_choices ORDER BY reason ASC").fetchall())[0]
            reasons = zip(*cursor.execute("SELECT code FROM secondary_ecr_reason_codes ORDER BY code ASC").fetchall())[0]
            user_department = cursor.execute('SELECT TOP 1 department FROM employees WHERE name = \'{}\''.format(General.app.current_user)).fetchone()[0]
            ##reasons_white_list = cursor.execute('SELECT TOP 1 can_choose_which_ecr_reasons FROM departments WHERE department = \'{}\''.format(user_department)).fetchone()[0]
            ##reasons_white_list = reasons_white_list.split('|')
            reasons_white_list = zip(*cursor.execute("SELECT code FROM departments_ecr_reason_codes WHERE department = '{}'".format(user_department)).fetchall())[0]

            for reason in reasons:
                if (reason in reasons_white_list):
                    ctrl(self.new_ecr_dialog, 'choice:ecr_reason').AppendItems((reason,))

        # ctrl(self.new_ecr_dialog, 'choice:ecr_reason').AppendItems(zip(*cursor.fetchall())[0])
        else:
            reasons = zip(*cursor.execute(
                "SELECT reason FROM ecr_reason_choices where Production_Plant = \'{}\' ".format(
                    Ecrs.Prod_Plant)).fetchall())[0]
            ctrl(self.new_ecr_dialog, 'choice:ecr_reason').AppendItems((reasons))

        # add document options to choice box
        if Ecrs.Prod_Plant == 'Systems':
            if duplicate_ecr_id == None:
                cursor.execute(
                    "SELECT document FROM ecr_document_choices WHERE type=\'Mechanical\' OR type=\'*\' AND Production_Plant = \'{}\'".format(
                        Ecrs.Prod_Plant))
            else:
                cursor.execute("SELECT document FROM ecr_document_choices where Production_Plant = \'{}\' ".format(
                    Ecrs.Prod_Plant))
            ctrl(self.new_ecr_dialog, 'choice:ecr_document').AppendItems(zip(*cursor.fetchall())[0])
        else:
            cursor.execute(
                "SELECT document FROM ecr_document_choices where Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant))
            ctrl(self.new_ecr_dialog, 'choice:ecr_document').AppendItems(zip(*cursor.fetchall())[0])

        self.ecr_type = 'Mechanical'

        self.new_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Mechanical'),
                                 id=xrc.XRCID('radio:mechanical'))
        self.new_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Electrical'),
                                 id=xrc.XRCID('radio:electrical'))
        self.new_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Structural'),
                                 id=xrc.XRCID('radio:structural'))
        self.new_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Other'),
                                 id=xrc.XRCID('radio:other'))
        self.new_ecr_dialog.Bind(wx.EVT_TEXT, Ecrs.check_reference_field, id=xrc.XRCID('text:reference_number'))
        self.new_ecr_dialog.Bind(wx.EVT_TEXT, Ecrs.on_text_ecr_description, id=xrc.XRCID('text:description'))
        self.new_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_select_ecr_reason, id=xrc.XRCID('choice:ecr_reason'))
        self.new_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_submit_ecr, id=xrc.XRCID('button:submit_ecr'))
        self.new_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_attatch_document, id=xrc.XRCID('button:attach_document'))

        # fill in some of the fields if user is duplicating an ecr
        if duplicate_ecr_id != None:
            ecr = cursor.execute(
                "SELECT reference_number, type, reason, document, request, resolution FROM ecrs WHERE id = \'{}\'".format(
                    duplicate_ecr_id)).fetchone()
            ctrl(self.new_ecr_dialog, 'text:reference_number').SetValue(str(ecr[0]))

            if ecr[1] == 'Mechanical': ctrl(self.new_ecr_dialog, 'radio:mechanical').SetValue(True)
            if ecr[1] == 'Electrical': ctrl(self.new_ecr_dialog, 'radio:electrical').SetValue(True)
            if ecr[1] == 'Structural': ctrl(self.new_ecr_dialog, 'radio:structural').SetValue(True)
            if ecr[1] == 'Other': ctrl(self.new_ecr_dialog, 'radio:other').SetValue(True)
            if ecr[1] == '*': ctrl(self.new_ecr_dialog, 'radio:other').SetValue(True)

            try:
                ctrl(self.new_ecr_dialog, 'choice:ecr_reason').SetStringSelection(ecr[2])
                ctrl(self.new_ecr_dialog, 'choice:ecr_reason').Focus()
            except:
                pass
            try:
                ctrl(self.new_ecr_dialog, 'choice:ecr_document').SetStringSelection(ecr[3])
            except:
                pass

            ctrl(self.new_ecr_dialog, 'text:description').SetValue(ecr[4])

            # refigure the need by date
            # cursor.execute("SELECT lead_time FROM ecr_reason_choices WHERE reason=\'{}\'".format(ecr[2]))
            lead_time_date = dt.datetime.today() + dt.timedelta(cursor.execute(
                "SELECT lead_time FROM ecr_reason_choices WHERE reason=\'{}\' AND Production_Plant = \'{}\'".format(
                    ecr[2], Ecrs.Prod_Plant)).fetchone()[0])
            ctrl(self.new_ecr_dialog, 'calendar:ecr_need_by').SetDate(
                wx.DateTimeFromDMY(lead_time_date.day, lead_time_date.month - 1, lead_time_date.year))

        ctrl(self.new_ecr_dialog, 'text:mentioned_parts').SetBackgroundColour(
            ctrl(self.main_frame, 'notebook:main').GetBackgroundColour())

        self.new_ecr_dialog.ShowModal()
        self.new_ecr_dialog = None

    def init_workflow_dialog(self):
        self.workflow_dialog = self.res.LoadDialog(None, 'dialog:workflow')

        self.wq = cursor.execute('Select Question from Workflow_Questions').fetchall()
        self.wqa = cursor.execute('Select Answers from Workflow_Questions').fetchall()
        self.C1 = []
        self.C2 = []
        self.C3 = []
        self.C4 = []
        self.C5 = []
        self.C6 = []
        self.C7 = []
        self.C8 = []
        self.C9 = []
        self.C10 = []

        for i in range(len(self.wq)):
            if i == 0:
                ctrl(self.workflow_dialog, 'WQST1').Show()
                ctrl(self.workflow_dialog, 'WQST1').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice1').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice1').Show()
                self.my_string = str(self.wqa[i][0])
                self.C1 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice1').SetItems(self.C1)

            elif i == 1:
                ctrl(self.workflow_dialog, 'WQST2').Show()
                ctrl(self.workflow_dialog, 'WQST2').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice2').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice2').Show()
                self.my_string = str(self.wqa[i][0])
                self.C2 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice2').SetItems(self.C2)

            elif i == 2:
                ctrl(self.workflow_dialog, 'WQST3').Show()
                ctrl(self.workflow_dialog, 'WQST3').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice3').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice3').Show()
                self.my_string = str(self.wqa[i][0])
                self.C3 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice3').SetItems(self.C3)

            elif i == 3:
                ctrl(self.workflow_dialog, 'WQST4').Show()
                ctrl(self.workflow_dialog, 'WQST4').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice4').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice4').Show()
                self.my_string = str(self.wqa[i][0])
                self.C4 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice4').SetItems(self.C4)

            elif i == 4:
                ctrl(self.workflow_dialog, 'WQST5').Show()
                ctrl(self.workflow_dialog, 'WQST5').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice5').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice5').Show()
                self.my_string = str(self.wqa[i][0])
                self.C5 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice5').SetItems(self.C5)

            elif i == 5:
                ctrl(self.workflow_dialog, 'WQST6').Show()
                ctrl(self.workflow_dialog, 'WQST6').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice6').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice6').Show()
                self.my_string = str(self.wqa[i][0])
                self.C6 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice6').SetItems(self.C6)

            elif i == 6:
                ctrl(self.workflow_dialog, 'WQST7').Show()
                ctrl(self.workflow_dialog, 'WQST7').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice7').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice7').Show()
                self.my_string = str(self.wqa[i][0])
                self.C7 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice7').SetItems(self.C7)

            elif i == 7:
                ctrl(self.workflow_dialog, 'WQST8').Show()
                ctrl(self.workflow_dialog, 'WQST8').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice8').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice8').Show()
                self.my_string = str(self.wqa[i][0])
                self.C8 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice8').SetItems(self.C8)

            elif i == 8:
                ctrl(self.workflow_dialog, 'WQST9').Show()
                ctrl(self.workflow_dialog, 'WQST9').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice9').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice9').Show()
                self.my_string = str(self.wqa[i][0])
                self.C9 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice9').SetItems(self.C9)

            elif i == 9:
                ctrl(self.workflow_dialog, 'WQST10').Show()
                ctrl(self.workflow_dialog, 'WQST10').SetLabel(str(self.wq[i][0]))
                ctrl(self.workflow_dialog, 'WQChoice10').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
                ctrl(self.workflow_dialog, 'WQChoice10').Show()
                self.my_string = str(self.wqa[i][0])
                self.C10 = [self.x.strip() for self.x in self.my_string.split(',')]
                ctrl(self.workflow_dialog, 'WQChoice10').SetItems(self.C10)


        # Bind Assign Workflow Button to an event to pop up Workflow Dialog
        self.workflow_dialog.Bind(wx.EVT_BUTTON, Ecrs.get_workflow_steps, id=xrc.XRCID('m_buttonOK'))
        self.workflow_dialog.ShowModal()


    def init_modify_ecr_dialog(self, modify_ecr_id):
        if modify_ecr_id == '':
            return

        self.modify_ecr_dialog = self.res.LoadDialog(None, 'dialog:edit_ecr')
        self.modify_ecr_dialog.SetTitle('Modify ECR: {}'.format(modify_ecr_id))
        #ctrl(self.modify_ecr_dialog, 'panel:committee').Enable()
        # self.modify_ecr_dialog.SetSize((750, 550))

        # Bind Do_Nothing Event upon mousewheel scroll in order to not change users Dropdowns selection accidently
        ctrl(self.modify_ecr_dialog, 'choice:ecr_reason').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.modify_ecr_dialog, 'choice:ecr_document').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.modify_ecr_dialog, 'choice:ecr_component').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.modify_ecr_dialog, 'choice:who_errored').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)
        ctrl(self.modify_ecr_dialog, 'choice:stage').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)

        cursor = Database.connection.cursor()

        # enable Reopen this ECR button if ECR status is closed and to only original initiator

        ecr_status = cursor.execute('select top 1 status from ecrs where id = {}'.format(modify_ecr_id)).fetchone()[0]
        ecr_initiator = cursor.execute('select top 1 who_requested from ecrs where id = {}'.format(modify_ecr_id)).fetchone()[0]
        print ecr_status, ecr_initiator
        if ecr_status == 'Open' or (ecr_initiator != General.app.current_user):
            ctrl(self.modify_ecr_dialog, 'm_buttonReopen').Disable()

        #Bind Assign Workflow Button to an event to pop up Workflow Dialog
        self.modify_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_assign_workflow, id=xrc.XRCID('m_buttonAssign'))

        #Hide workflow for now until it is ready for release
        ctrl(self.modify_ecr_dialog, 'm_panelWorkflow').Hide()
        ctrl(self.modify_ecr_dialog, 'm_buttonAssign').Hide()

        #Hide Assign Workflow Button if user in Systems Plant
        if Ecrs.Prod_Plant == 'Systems':
            ctrl(self.modify_ecr_dialog, 'm_buttonAssign').Hide()
            ctrl(self.modify_ecr_dialog, 'm_panelWorkflow').Hide()

        #Disable Workflow Assign button if it has already been assigned
        try:
            workflow_exists = cursor.execute('Select top 1 step_no from Ecrev_Status where Ecrev_no =?',modify_ecr_id).fetchone()[0]
            if workflow_exists:
                workflow = True
        except:
            workflow = False

        if workflow:
            ctrl(self.modify_ecr_dialog, 'm_buttonAssign').Disable()
            workflow_info = cursor.execute('Select Assigned_to, Step_description, current_Status from Ecrev_Status where Ecrev_no = ?',modify_ecr_id).fetchall()

            ctrl(General.app.modify_ecr_dialog, 'm_textStep1').SetValue(workflow_info[0][1])
            ctrl(General.app.modify_ecr_dialog, 'm_textStep2').SetValue(workflow_info[1][1])
            ctrl(General.app.modify_ecr_dialog, 'm_textStep3').SetValue(workflow_info[2][1])
            ctrl(General.app.modify_ecr_dialog, 'm_textStep4').SetValue(workflow_info[3][1])
            ctrl(General.app.modify_ecr_dialog, 'm_textStep5').SetValue(workflow_info[4][1])

            ctrl(General.app.modify_ecr_dialog, 'm_textCtrlWho1').SetValue(workflow_info[0][0])
            ctrl(General.app.modify_ecr_dialog, 'm_textCtrlWho2').SetValue(workflow_info[1][0])
            ctrl(General.app.modify_ecr_dialog, 'm_textCtrlWho3').SetValue(workflow_info[2][0])
            ctrl(General.app.modify_ecr_dialog, 'm_textCtrlWho4').SetValue(workflow_info[3][0])
            ctrl(General.app.modify_ecr_dialog, 'm_textCtrlWho5').SetValue(workflow_info[4][0])

            if workflow_info[0][2] == 'Completed':
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep1').SetValue(True)
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep1').Disable()

            if workflow_info[1][2] == 'Completed':
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep2').SetValue(True)
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep2').Disable()

            if workflow_info[2][2] == 'Completed':
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep3').SetValue(True)
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep3').Disable()

            if workflow_info[3][2] == 'Completed':
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep4').SetValue(True)
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep4').Disable()

            if workflow_info[4][2] == 'Completed':
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep5').SetValue(True)
                ctrl(General.app.modify_ecr_dialog, 'm_checkStep5').Disable()


        #Bind Workflow step status checkbox event
        self.modify_ecr_dialog.Bind(wx.EVT_CHECKBOX, Ecrs.assign_ecrev_workflow_next_step, id=xrc.XRCID('m_checkStep1'))
        self.modify_ecr_dialog.Bind(wx.EVT_CHECKBOX, Ecrs.assign_ecrev_workflow_next_step, id=xrc.XRCID('m_checkStep2'))
        self.modify_ecr_dialog.Bind(wx.EVT_CHECKBOX, Ecrs.assign_ecrev_workflow_next_step, id=xrc.XRCID('m_checkStep3'))
        self.modify_ecr_dialog.Bind(wx.EVT_CHECKBOX, Ecrs.assign_ecrev_workflow_next_step, id=xrc.XRCID('m_checkStep4'))
        self.modify_ecr_dialog.Bind(wx.EVT_CHECKBOX, Ecrs.assign_ecrev_workflow_next_step, id=xrc.XRCID('m_checkStep5'))


        # show committee panel if authorized
        can_approve_first = cursor.execute(
            "SELECT can_approve_first FROM employees WHERE name = '{}'".format(self.current_user)).fetchone()[0]
        if can_approve_first:
            ctrl(self.modify_ecr_dialog, 'spin:priority').Enable()
            ctrl(self.modify_ecr_dialog, 'choice:stage').Enable()
        # ctrl(self.modify_ecr_dialog, 'button:approve_1').Enable()

        can_approve_second = cursor.execute(
            "SELECT can_approve_second FROM employees WHERE name = '{}'".format(self.current_user)).fetchone()[0]
        if can_approve_second:
            # ctrl(self.modify_ecr_dialog, 'panel:committee').Enable()
            # ctrl(self.modify_ecr_dialog, 'button:approve_1').Show()
            ctrl(self.modify_ecr_dialog, 'spin:priority').Enable()
            ctrl(self.modify_ecr_dialog, 'choice:stage').Enable()
        # ctrl(self.modify_ecr_dialog, 'button:approve_2').Enable()

        if not can_approve_first and not can_approve_second:
            ctrl(self.modify_ecr_dialog, 'button:approve_1').Hide()
            ctrl(self.modify_ecr_dialog, 'button:approve_2').Hide()
        # ctrl(self.modify_ecr_dialog, 'panel:committee').Hide()

        ctrl(self.modify_ecr_dialog, 'choice:stage').AppendItems(
            ['New Request, needs reviewing', 'Reviewed, engineering in process', 'Change Complete, pending approval',
             'Approved', 'Prototype Stage'])

        # add document options to choice box
        cursor.execute(
            "SELECT document FROM ecr_document_choices where Production_Plant = \'{}\'".format(Ecrs.Prod_Plant))
        ctrl(self.modify_ecr_dialog, 'choice:ecr_document').AppendItems(zip(*cursor.fetchall())[0])

        ##self.ecr_type = 'Mechanical'

        self.modify_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Mechanical'),
                                    id=xrc.XRCID('radio:mechanical'))
        self.modify_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Electrical'),
                                    id=xrc.XRCID('radio:electrical'))
        self.modify_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Structural'),
                                    id=xrc.XRCID('radio:structural'))
        self.modify_ecr_dialog.Bind(wx.EVT_RADIOBUTTON, partial(Ecrs.radio_button_selected, type='Other'),
                                    id=xrc.XRCID('radio:other'))
        # self.modify_ecr_dialog.Bind(wx.EVT_TEXT, Ecrs.check_reference_field, id=xrc.XRCID('text:reference_number'))
        # self.modify_ecr_dialog.Bind(wx.EVT_TEXT, Ecrs.on_text_ecr_description, id=xrc.XRCID('text:description'))
        # self.modify_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_select_ecr_reason, id=xrc.XRCID('choice:ecr_reason'))
        # self.modify_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_submit_ecr, id=xrc.XRCID('button:submit_ecr'))
        self.modify_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_attach_document, id=xrc.XRCID('button:Attach'))

        self.modify_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_modify_ecr, id=xrc.XRCID('button:modify_or_close_ecr'))

        self.modify_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_approve_1_for_modify, id=xrc.XRCID('button:approve_1'))
        self.modify_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_approve_2_for_modify, id=xrc.XRCID('button:approve_2'))
        self.modify_ecr_dialog.Bind(wx.EVT_BUTTON, Ecrs.on_click_Reopen_Ecr, id=xrc.XRCID('m_buttonReopen'))

        self.modify_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_choice_set_severity_default,
                                    id=xrc.XRCID('choice:ecr_document'))
        self.modify_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_choice_set_severity_default,
                                    id=xrc.XRCID('choice:ecr_component'))
        self.modify_ecr_dialog.Bind(wx.EVT_CHOICE, Ecrs.on_choice_set_severity_default,
                                    id=xrc.XRCID('choice:ecr_sub_system'))

        # fill in the fields of the ecr we are modifing
        ecr = cursor.execute(
            "SELECT reference_number, document, reason, type, request, resolution, when_needed, who_errored, priority, who_approved_first, who_approved_second, approval_stage, item, component, sub_system, severity, Units_Affected FROM ecrs WHERE id = \'{}\'".format(
                modify_ecr_id)).fetchone()
        reference_number, document, reason, type, request, resolution, when_needed, who_errored, priority, who_approved_first, who_approved_second, approval_stage, item, component, sub_system, severity, Units_Affected = ecr

        if Ecrs.Prod_Plant == 'Systems':
            if type == 'Other':
                components = list(zip(*cursor.execute("SELECT DISTINCT component FROM ecr.components WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0])
            else:
                components = list(zip(*cursor.execute("SELECT DISTINCT component FROM ecr.components WHERE Production_Plant = \'{}\' AND discipline= \'{}\'".format(Ecrs.Prod_Plant, type)).fetchall())[0])

            components.insert(0, '')
            ctrl(self.modify_ecr_dialog, 'choice:ecr_component').AppendItems(components)

            if component:
                ctrl(self.modify_ecr_dialog, 'choice:ecr_component').Insert(component, 1)
                ctrl(self.modify_ecr_dialog, 'choice:ecr_component').SetStringSelection(component)

        # add sub_system options
            if type == 'Other':
                sub_systems = list(zip(*cursor.execute("SELECT DISTINCT sub_system FROM ecr.sub_systems WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0])
            else:
                sub_systems = list(zip(*cursor.execute("SELECT DISTINCT sub_system FROM ecr.sub_systems WHERE Production_Plant = \'{}\' AND discipline=\'{}\'".format(Ecrs.Prod_Plant, type)).fetchall())[0])

            sub_systems.insert(0, '')
            ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').AppendItems(sub_systems)

            if sub_system:
                ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').Insert(sub_system, 1)
                ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').SetStringSelection(sub_system)
        else:
            if type == 'Other':
                components = list(zip(*cursor.execute("SELECT DISTINCT component FROM ecr.components WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0])
            else:
                components = list(zip(*cursor.execute(
                    "SELECT DISTINCT component FROM ecr.components WHERE Production_Plant = \'{}\' AND discipline=\'{}\'".format(Ecrs.Prod_Plant, type)).fetchall())[0])

            components.insert(0, '')
            ctrl(self.modify_ecr_dialog, 'choice:ecr_component').AppendItems(components)

            if component:
                ctrl(self.modify_ecr_dialog, 'choice:ecr_component').Insert(component, 1)
                ctrl(self.modify_ecr_dialog, 'choice:ecr_component').SetStringSelection(component)

                # add sub_system options
            if type == 'Other':
                sub_systems = list(
                    zip(*cursor.execute("SELECT DISTINCT sub_system FROM ecr.sub_systems WHERE Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0])
            else:
                sub_systems = list(zip(*cursor.execute(
                    "SELECT DISTINCT sub_system FROM ecr.sub_systems WHERE Production_Plant = \'{}\' AND discipline=\'{}\'".format(Ecrs.Prod_Plant, type)).fetchall())[
                                       0])

            sub_systems.insert(0, '')
            ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').AppendItems(sub_systems)

            if sub_system:
                ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').Insert(sub_system, 1)
                ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').SetStringSelection(sub_system)

            """components_list_cases = ['Air block', 'Air Deflector', 'Base', 'Brackets', 'Breaker', 'Bumper/retainer',
                                     'Coil',
                                     'Controller', 'Deck pans', 'Door/frame', 'End Assy', 'Fan', 'Glass', 'Horse Head',
                                     'Kick plates', 'Lights', 'Other', 'Painted part', 'Piping', 'Pnl, Foam, Back',
                                     'Pnl, Foam, Cnpy', 'Pnl, Foam, Front', 'PVC', 'Raceway', ' Rack', 'Sensor, Temp',
                                     'Sensor, Pressure', 'Shelf standard', 'Shelves', 'Skin', 'Tub, foam', 'Valve',
                                     'Wire racks']
            sub_system_list_cases = ['Base', 'Coil Piping', 'Controls', 'Doors/frames', 'End', 'Foam', 'Kitting',
                                     'Knock up',
                                     'Lighting', 'Paint', 'Piping', 'Piping Option Pack', 'QA', 'Raceway',
                                     'Sheet Metal',
                                     'Subassy', 'Trimming', 'Wiring']

            # components_list_cases.insert(0, '')
            ctrl(self.modify_ecr_dialog, 'choice:ecr_component').AppendItems(components_list_cases)
            # ctrl(self.modify_ecr_dialog, 'choice:ecr_component').Insert(components_list_cases, 1)
            #ctrl(self.modify_ecr_dialog, 'choice:ecr_component').SetStringSelection(components_list_cases)

            # sub_system_list_cases.insert(0, '')
            ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').AppendItems(sub_system_list_cases)
            # ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').Insert(sub_system_list_cases, 1)
            #ctrl(self.modify_ecr_dialog, 'choice:ecr_sub_system').SetStringSelection(sub_system_list_cases)"""

        # add severity options
        ctrl(self.modify_ecr_dialog, 'choice:ecr_severity').AppendItems(['High', 'Medium', 'Low'])

        severity = float(severity)

        if severity == 1.0:
            ctrl(self.modify_ecr_dialog, 'choice:ecr_severity').SetStringSelection('High')
        elif severity == 0.5:
            ctrl(self.modify_ecr_dialog, 'choice:ecr_severity').SetStringSelection('Medium')
        elif severity == 0.1:
            ctrl(self.modify_ecr_dialog, 'choice:ecr_severity').SetStringSelection('Low')
        else:
            ctrl(self.modify_ecr_dialog, 'choice:ecr_severity').SetStringSelection('High')

        # add reason for request options to choice box
        ##cursor.execute("SELECT reason FROM ecr_reason_choices ORDER BY reason ASC")
        ##ctrl(self.modify_ecr_dialog, 'choice:ecr_reason').AppendItems(zip(*cursor.fetchall())[0])
        if Ecrs.Prod_Plant == 'Systems':
            reasons = list(
                zip(*cursor.execute("SELECT code FROM secondary_ecr_reason_codes ORDER BY code ASC").fetchall())[0])

            ctrl(self.modify_ecr_dialog, 'choice:ecr_reason').AppendItems(reasons)

            # support old ECR reason codes if ECR originally submitted under them
            if reason not in reasons:
                cursor.execute(
                    "SELECT reason FROM ecr_reason_choices where Production_Plant = \'{}\' ORDER BY reason ASC".format(
                        Ecrs.Prod_Plant))
                ctrl(self.modify_ecr_dialog, 'choice:ecr_reason').AppendItems(zip(*cursor.fetchall())[0])
        else:
            reasons = cursor.execute(
                "SELECT reason FROM ecr_reason_choices where Production_Plant = \'{}\' ORDER BY reason ASC".format(
                    Ecrs.Prod_Plant))
            ctrl(self.modify_ecr_dialog, 'choice:ecr_reason').AppendItems(zip(*cursor.fetchall())[0])
        
        # hide committee panel if not the right type of ECR reason
        if reason not in Ecrs.reasons_needing_approval:
            ctrl(self.modify_ecr_dialog, 'panel:committee').Hide()
        try:
            ctrl(self.modify_ecr_dialog, 'choice:stage').SetStringSelection(approval_stage)
        except Exception as e:
            print 'Error, failed to set approval stage:', e

        ctrl(self.modify_ecr_dialog, 'text:reference_number').SetValue(str(ecr[0]))

        if ecr[3] == 'Mechanical':
            ctrl(self.modify_ecr_dialog, 'radio:mechanical').SetValue(True)
            self.ecr_type = 'Mechanical'
        if ecr[3] == 'Electrical':
            ctrl(self.modify_ecr_dialog, 'radio:electrical').SetValue(True)
            self.ecr_type = 'Electrical'
        if ecr[3] == 'Structural':
            ctrl(self.modify_ecr_dialog, 'radio:structural').SetValue(True)
            self.ecr_type = 'Structural'
        if ecr[3] == 'Other':
            ctrl(self.modify_ecr_dialog, 'radio:other').SetValue(True)
            self.ecr_type = 'Other'
        if ecr[3] == '*':
            ctrl(self.modify_ecr_dialog, 'radio:other').SetValue(True)
            self.ecr_type = '*'

        try:
            ctrl(self.modify_ecr_dialog, 'choice:ecr_reason').SetStringSelection(ecr[2])
            ctrl(self.modify_ecr_dialog, 'choice:ecr_reason').Focus()
        except:
            pass
        try:
            ctrl(self.modify_ecr_dialog, 'choice:ecr_document').SetStringSelection(ecr[1])
        except:
            pass

        if ecr[4] != None: ctrl(self.modify_ecr_dialog, 'text:description').SetValue(ecr[4])
        if ecr[5] != None: ctrl(self.modify_ecr_dialog, 'text:resolution').SetValue(ecr[5])
        if ecr[16] != None: ctrl(self.modify_ecr_dialog, 'm_NoUnitsAffected').SetValue(ecr[16])
        ctrl(self.modify_ecr_dialog, 'text:m_AttachList').SetValue('')

        need_by_date = time.strptime(str(ecr[6]), "%Y-%m-%d %H:%M:%S")  # to python time object
        ctrl(self.modify_ecr_dialog, 'calendar:ecr_need_by').SetDate(
            wx.DateTimeFromDMY(need_by_date.tm_mday, need_by_date.tm_mon - 1, need_by_date.tm_year))

        engineers = list(zip(*cursor.execute(
            "SELECT name FROM employees WHERE department LIKE '%Engineering%' ORDER BY name ASC").fetchall())[0])
        engineers.insert(0, '')
        ctrl(self.modify_ecr_dialog, 'choice:who_errored').AppendItems(engineers)

        if ecr[7] != '' and ecr[7] != None:
            ctrl(self.modify_ecr_dialog, 'choice:who_errored').SetStringSelection(ecr[7])

        # hide similar items panel
        ctrl(self.modify_ecr_dialog, 'panel:similar_ecrs').Hide()

        # committee stuff
        ctrl(self.modify_ecr_dialog, 'spin:priority').SetValue(priority)
        if who_approved_first:
            ctrl(self.modify_ecr_dialog, 'label:who_approved_first').SetLabel(who_approved_first)
            ctrl(self.modify_ecr_dialog, 'button:approve_1').SetLabel('Unapprove')

        if who_approved_second:
            ctrl(self.modify_ecr_dialog, 'label:who_approved_second').SetLabel(who_approved_second)
            ctrl(self.modify_ecr_dialog, 'button:approve_2').SetLabel('Unapprove')

        ctrl(self.modify_ecr_dialog, 'button:modify_or_close_ecr').SetLabel('Save Changes')


        # ctrl(self.modify_ecr_dialog, 'panel:main').Layout()
        # ctrl(self.modify_ecr_dialog, 'panel:main').Refresh()
        # ctrl(self.modify_ecr_dialog, 'panel:main').Update()
        ctrl(self.modify_ecr_dialog, 'panel:committee').Layout()

        self.modify_ecr_dialog.ShowModal()

    def init_revisions_tab(self):
        self.revision_lists = []

        # deletethis#self.main_frame.Bind(wx.EVT_BUTTON, on_click_select_revision_document, id=xrc.XRCID('button:select_document'))
        self.main_frame.Bind(wx.EVT_TEXT, Revisions.on_entry_revision_items, id=xrc.XRCID('text:revision_items'))
        self.main_frame.Bind(wx.EVT_BUTTON, Revisions.on_click_add_new_revision,
                             id=xrc.XRCID('button:add_new_revision'))
        self.main_frame.Bind(wx.EVT_BUTTON, Revisions.on_click_print_revisions, id=xrc.XRCID('button:print_revisions'))

    def init_search_tab(self):
        self.table_search_criteria = None

        # so we can search a table with more criteria
        self.joining_tables = []
        self.joining_tables.append(('ecrs', Ecrs.table_used, 'item', 'item'))
        self.joining_tables.append(('ecrs', 'revisions', 'id', 'related_ecr'))
        self.joining_tables.append(('revisions', Ecrs.table_used, 'item', 'item'))
        self.joining_tables.append(('revisions', 'ecrs', 'related_ecr', 'id'))
        self.joining_tables.append((Ecrs.table_used, 'revisions', 'item', 'item'))
        self.joining_tables.append((Ecrs.table_used, 'ecrs', 'item', 'item'))
        # self.joining_tables.append(('time_logs', Ecrs.table_used, 'item', 'item'))

        # Bind Do_Nothing Event upon mousewheel scroll in order to not change users Dropdowns selection accidently
        ctrl(self.main_frame, 'choice:which_table').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)


        cursor = Database.connection.cursor()
        # tables = list(zip(*cursor.execute(r"SELECT name FROM sqlite_master WHERE type in ('table', 'view') AND name NOT LIKE 'sqlite_%' UNION ALL SELECT name FROM sqlite_temp_master WHERE type IN ('table', 'view') ORDER BY 1").fetchall())[0])
        tables = list(zip(*cursor.execute('SELECT * FROM information_schema.tables ORDER BY table_name').fetchall())[2])

        table_white_list = ['ecrs', Ecrs.table_used, 'revisions']

        '''
        table_black_list = []
        if self.current_user != 'Stuart, Travis':
            table_black_list.append('administration')
            table_black_list.append('departments')
            table_black_list.append('ecr_document_choices')
            table_black_list.append('ecr_reason_choices')
            table_black_list.append('engineering_focuses')
            table_black_list.append('revision_document_choices')
            table_black_list.append('revision_reason_choices')
            table_black_list.append('employees')
            table_black_list.append('custom_resource_allotments')
            table_black_list.append('family_hour_estimates')
            table_black_list.append('unit_selections')
            table_black_list.append('compressors')
            table_black_list.append('item_responsibilities')
            table_black_list.append('item_responsibilities2')
            table_black_list.append('orders_sandbox')
            table_black_list.append('std_family_hour_estimates')
            table_black_list.append('email_lists')
            table_black_list.append('project_time_logs')
            table_black_list.append('projects')
            table_black_list.append('tasks')
            table_black_list.append('employee_owner')
            table_black_list.append('primary_ecr_reason_codes')
            table_black_list.append('secondary_ecr_reason_codes')
            table_black_list.append('departments_ecr_reason_codes')
            table_black_list.append('orders_sandbox')
            table_black_list.append('ecr_reason_code_changes')
        table_black_list.append('sysdiagrams')
        #table_black_list.append('ecrs')
        #^make these based on department!
        
        for table in table_black_list:
            tables.remove(table)
        '''

        if self.current_user != 'Stuart, Travis':
            tables = table_white_list

        ctrl(self.main_frame, 'choice:which_table').AppendItems(tables)

        self.main_frame.Bind(wx.EVT_CHOICE, Search.on_select_table, id=xrc.XRCID('choice:which_table'))
        self.main_frame.Bind(wx.EVT_BUTTON, Search.on_click_begin_search, id=xrc.XRCID('button:search'))
        self.main_frame.Bind(wx.EVT_LIST_ITEM_SELECTED, Search.on_select_result, id=xrc.XRCID('list:search_results'))
        ###self.main_frame.Bind(wx.EVT_LIST_COL_CLICK, sort_list, id=xrc.XRCID('list:search_results'))
        self.main_frame.Bind(wx.EVT_BUTTON, Search.on_click_open_how_to_search, id=xrc.XRCID('button:how_to_search'))
        ###self.main_frame.Bind(wx.EVT_BUTTON, Search.export_search_results, id=xrc.XRCID('button:export_search_results'))
        self.main_frame.Bind(wx.EVT_BUTTON, Search.export_search_results,
                             id=xrc.XRCID('button:export_search_results'))
        self.main_frame.Bind(wx.EVT_BUTTON, ctrl(self.main_frame, 'list:search_results').print_list,
                             id=xrc.XRCID('button:print_search_results'))
        ctrl(self.main_frame, 'list:search_results').printer_paper_type = wx.PAPER_11X17

    def init_reports_tab(self):
        ##self.main_frame.Bind(wx.EVT_BUTTON, Reports.plot_demo, id=xrc.XRCID('button:plot_time'))

        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:today'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:this_week'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:this_month'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:this_quarter'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:this_year'))

        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:yesterday'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:last_week'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:last_month'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:last_quarter'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:last_year'))

        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:all_time'))
        self.main_frame.Bind(wx.EVT_RADIOBUTTON, Reports.on_click_radio, id=xrc.XRCID('radio:custom_range'))

        self.main_frame.Bind(wx.EVT_DATE_CHANGED, Reports.on_change_date, id=xrc.XRCID('date:report_start'))
        self.main_frame.Bind(wx.EVT_DATE_CHANGED, Reports.on_change_date, id=xrc.XRCID('date:report_end'))
        self.main_frame.Bind(wx.EVT_BUTTON, Reports.Advanced_Report, id=xrc.XRCID('m_buttonAdReport'))

        self.main_frame.Bind(wx.EVT_LISTBOX, Reports.on_select_report, id=xrc.XRCID('listbox:reports'))

        self.main_frame.Bind(wx.EVT_PAINT, Reports.on_paint_window, id=xrc.XRCID('frame:main'))

        # default to this month as the data range
        today = dt.date.today()
        self.start_date = dt.date(today.year, today.month, 1)
        self.end_date = dt.date(today.year, today.month, today.day)

        self.plotting_panel = None  # Reports.DesignEngineeringHours(ctrl(self.main_frame, 'panel:plot'), self.start_date, self.end_date)
        self.report_name = None

        # add some reports to select from
        ctrl(self.main_frame, 'listbox:reports').Append("DE's Logged Hours")
        ctrl(self.main_frame, 'listbox:reports').Append("ECR Reasons")
        ctrl(self.main_frame, 'listbox:reports').Append("ECR Documents")
        ctrl(self.main_frame, 'listbox:reports').Append("ECRs Closed On Time")
        ctrl(self.main_frame, 'listbox:reports').Append("ECRs by Product Family")
        ctrl(self.main_frame, 'listbox:reports').Append("ECRs by Customer")
        ctrl(self.main_frame, 'listbox:reports').Append("Hours Logged by Product Family")

    def init_search_ecrs_dialog(self):
        if self.search_ecrs_dialog == None:
            self.search_ecrs_dialog = self.res.LoadDialog(None, 'dialog:search_ecrs')

            self.search_ecrs_dialog.Bind(wx.EVT_CLOSE, Ecrs.hide_search_ecrs_dialog)

            Ecrs.reset_search_fields()

            cursor = Database.connection.cursor()

            # populate some search field choices
            ctrl(self.search_ecrs_dialog, 'combo:search_value3').AppendItems(
                zip(*cursor.execute("SELECT document FROM ecr_document_choices where Production_Plant = \'{}\'".format(
                    Ecrs.Prod_Plant)).fetchall())[0])
            ctrl(self.search_ecrs_dialog, 'combo:search_value4').AppendItems(
                zip(*cursor.execute("SELECT reason FROM ecr_reason_choices where Production_Plant = \'{}\' ".format(
                    Ecrs.Prod_Plant)).fetchall())[0])
            ctrl(self.search_ecrs_dialog, 'combo:search_value5').AppendItems(
                zip(*cursor.execute("SELECT department FROM departments").fetchall())[0])
            ctrl(self.search_ecrs_dialog, 'combo:search_value6').AppendItems(
                zip(*cursor.execute("SELECT name FROM employees ORDER BY name ASC").fetchall())[0])
            ctrl(self.search_ecrs_dialog, 'combo:search_value7').AppendItems(
                zip(*cursor.execute("SELECT focus FROM engineering_focuses").fetchall())[0])
            ctrl(self.search_ecrs_dialog, 'combo:search_value10').AppendItems(
                zip(*cursor.execute("SELECT name FROM employees ORDER BY name ASC").fetchall())[0])
            ctrl(self.search_ecrs_dialog, 'combo:search_value11').AppendItems(
                zip(*cursor.execute("SELECT name FROM employees ORDER BY name ASC").fetchall())[0])
            ctrl(self.search_ecrs_dialog, 'combo:search_value12').AppendItems(
                zip(*cursor.execute("SELECT name FROM employees ORDER BY name ASC").fetchall())[0])

            ctrl(self.search_ecrs_dialog, 'choice:sort_by').AppendItems(
                Database.get_table_column_names('ecrs', presentable=True))
            ctrl(self.search_ecrs_dialog, 'choice:sort_by').SetStringSelection('When requested')

            self.search_ecrs_dialog.Bind(wx.EVT_LIST_ITEM_SELECTED, Ecrs.on_select_ecr_item,
                                         id=xrc.XRCID('list:results'))
            self.search_ecrs_dialog.Bind(wx.EVT_BUTTON, Ecrs.search_ecrs, id=xrc.XRCID('button:search'))
            self.search_ecrs_dialog.Bind(wx.EVT_BUTTON, Ecrs.export_search_results, id=xrc.XRCID('button:export'))

            for i in range(11):
                self.search_ecrs_dialog.Bind(wx.EVT_CHOICE, partial(Ecrs.search_condition_selected, index=i),
                                             id=xrc.XRCID('choice:search_condition' + str(i)))
                self.search_ecrs_dialog.Bind(wx.EVT_TEXT, partial(Ecrs.search_value_entered, index=i),
                                             id=xrc.XRCID('combo:search_value' + str(i)))

        self.search_ecrs_dialog.Show()

    def init_submit_revision_dialog(self, item_entries):
        self.new_revision_dialog = self.res.LoadDialog(None, 'dialog:new_revision')
        ###self.new_revision_dialog.Bind(wx.EVT_BUTTON, on_click_select_revision_document, id=xrc.XRCID('button:select_document'))
        self.new_revision_dialog.Bind(wx.EVT_BUTTON, Revisions.on_click_submit_revision,
                                      id=xrc.XRCID('button:submit_revision'))
        self.new_revision_dialog.Bind(wx.EVT_CHOICE, Revisions.on_select_revision_reason,
                                      id=xrc.XRCID('choice:revision_reasons'))

        # Bind Do_Nothing Event upon mousewheel scroll in order to not change users Dropdowns selection accidently
        ctrl(self.new_revision_dialog, 'choice:revision_reasons').Bind(wx.EVT_MOUSEWHEEL, self.do_nothing)

        item_label = 'For Item Numbers:  '
        for item in item_entries:
            item_label += '{}, '.format(item)
        item_label = item_label[:-2]

        related_ecr = ctrl(self.main_frame, 'text:related_ecr').GetValue()
        if related_ecr != '':
            item_label += '  (related to ECR ID: {})'.format(related_ecr)

        # populate rev reason codes drop down... BUT don't put in BOM Rec as an option unless
        # the ECR associated with the change was coded as BOM Rec
        cursor = Database.connection.cursor()
        if Ecrs.Prod_Plant == 'Systems':

            reasons = list(zip(*cursor.execute("SELECT reason FROM revision_reason_choices where Production_Plant = \'{}\' ORDER BY reason ASC".format(Ecrs.Prod_Plant)).fetchall())[0])
            try:
                ecr_reason = cursor.execute("SELECT reason FROM ecrs WHERE id={}".format(related_ecr)).fetchone()[0]
            except:
                ecr_reason = None
            if ecr_reason != 'BOM Reconciliation':
                try:
                    reasons.remove('BOM Reconciliation')
                except Exception as e:
                    print e

            ctrl(self.new_revision_dialog, 'choice:revision_reasons').AppendItems(reasons)

        else:
            reasons = zip(*cursor.execute("SELECT reason FROM revision_reason_choices where Production_Plant = \'{}\' ORDER BY reason ASC".format(Ecrs.Prod_Plant)).fetchall())[0]
            ctrl(self.new_revision_dialog, 'choice:revision_reasons').AppendItems(reasons)

        ctrl(self.new_revision_dialog, 'label:item_numbers').SetLabel(item_label)

        table_panel = ctrl(self.new_revision_dialog, 'panel:table')

        table = TweakedGrid.TweakedGrid(table_panel)
        table.Bind(gridlib.EVT_GRID_CELL_LEFT_CLICK, Revisions.on_click_table_cell)

        table.CreateGrid(25, 2)
        # table.SetMargins(0, 0)
        # table.EnableScrolling(False, False)
        table.SetRowLabelSize(0)
        table.SetColLabelValue(0, 'Documents')
        table.SetColLabelValue(1, 'Descriptions of change')

        table.SetCellValue(0, 0, ' (click to select document) ')
        table.AutoSize()

        table.EnableDragRowSize(False)

        table.Bind(wx.EVT_SIZE, Revisions.on_size_document_table)

        # for row in range(1, 25):
        #	for col in range(2):
        #		table.SetCellValue(row, col,"cell (%d,%d)" % (row, col))

        for row in range(25):
            table.SetReadOnly(row, 0)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(table, 1, wx.EXPAND)
        table_panel.SetSizer(sizer)
        # table_panel.SetAutoLayout(True)
        # table_panel.Layout()
        # print table.GetSize()[0]
        # table_panel.Refresh()

        # print '.GetSize(): {}'.format(table_panel.GetSize())
        # print table.GetColSize(1)
        # table.SetColSize(1, table.GetSize()[0] - table.GetColSize(0) +50)
        # table.SetColSize(1, table.GetSize()[0])
        # print table.GetColSize(1)

        # sizer.Layout()

        self.new_revision_dialog.ShowModal()

    def init_revision_document_selection_dialog(self):
        self.revision_document_selection_dialog = self.res.LoadDialog(None, 'dialog:select_revision_document')

        self.selected_revision_document = None
        tree = ctrl(self.revision_document_selection_dialog, 'tree:revision_documents')
        root = tree.AddRoot("Root")
        tree.SetPyData(root, None)

        # generate doc tree from database
        cursor = Database.connection.cursor()
        if Ecrs.Prod_Plant == 'Systems':
            raw_tree_data = zip(*cursor.execute("SELECT document FROM revision_document_choices where Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0]
        else:
            raw_tree_data = zip(*cursor.execute("SELECT document FROM revision_document_choices where Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0]

        level_values = [None, None, None, None, None]
        branch_pointers = [root, root, root, root, root]

        for doc in raw_tree_data:
            for branch_level, branch_name in enumerate(doc.split('>')):
                if branch_name != level_values[branch_level]:
                    level_values[branch_level] = branch_name
                    if branch_level == len(doc.split('>')) - 1:
                        child = tree.AppendItem(branch_pointers[branch_level], branch_name + '               ')
                    else:
                        child = tree.AppendItem(branch_pointers[branch_level], branch_name + '')
                    tree.SetPyData(child, None)
                    branch_pointers[branch_level + 1] = child

        # tree.Expand(root)

        tree.Bind(wx.EVT_LEFT_DOWN, Revisions.on_click_document_tree)
        tree.Bind(wx.EVT_TREE_SEL_CHANGED, Revisions.on_selection_change_document_tree)

        self.revision_document_selection_dialog.ShowModal()
        return self.selected_revision_document

    '''
    def init_time_logs_tab(self):
        #cursor = Database.connection.cursor()
        
        TimeLogs.refresh_log_list()
        
        self.timer_start_time = None
        self.timer_schedule = sched.scheduler(time.time, time.sleep)
        self.timer_thread = None
        
        #self.main_frame.Bind(wx.EVT_BUTTON, TimeLogs.on_click_timer, id=xrc.XRCID('button:timer'))
        self.main_frame.Bind(wx.EVT_TOGGLEBUTTON, TimeLogs.on_click_timer, id=xrc.XRCID('toggle:timer'))
        self.main_frame.Bind(wx.EVT_BUTTON, TimeLogs.on_click_log_time, id=xrc.XRCID('button:log_time'))
        self.main_frame.Bind(wx.EVT_BUTTON, TimeLogs.refresh_log_list, id=xrc.XRCID('button:refresh_time'))
        self.main_frame.Bind(wx.EVT_LIST_ITEM_SELECTED, TimeLogs.on_select_log, id=xrc.XRCID('list:my_time_logs'))
        
        self.main_frame.Bind(wx.EVT_LIST_ITEM_ACTIVATED, TimeLogs.on_activate_open_modify_dialog, id=xrc.XRCID('list:my_time_logs'))
    '''


def destroy_dialog(event, dialog):
    dialog.Destroy()


def email_ecr(to_who, from_who, message, ecr_id):
    pass


def on_click_login(event):
    # get entered user and password values from form fields
    selected_user = ctrl(General.app.login_frame, 'choice:name').GetStringSelection()
    entered_password = ctrl(General.app.login_frame, 'text:password').GetValue()
    selected_plant = ctrl(General.app.login_frame, 'choice:plant').GetStringSelection()
    remember_password = ctrl(General.app.login_frame, 'm_checkBoxRemPass').GetValue()

    if selected_plant == "Cases":
        Ecrs.table_used = 'orders_cases'
        Ecrs.Prod_Plant = 'Cases'
    elif selected_plant == 'Systems':
        Ecrs.table_used = 'orders'
        Ecrs.Prod_Plant = 'Systems'
    else:
        wx.MessageBox("You must select a plant.", 'Login failed')
        # Ecrs.table_used = ''
        # Ecrs.Prod_Plant = 'NULL'

    print Ecrs.table_used

    # field validation
    if selected_user == '':
        wx.MessageBox('You must select a user.', 'Login failed')
        return
    # if entered_password == '':
    #	wx.MessageBox('You must enter a password.', 'Login failed')
    #	return

    # get password from selected user name
    cursor = Database.connection.cursor()
    cursor.execute("SELECT TOP 1 password FROM employees WHERE name = \'{}\'".format(selected_user.replace("'", "''")))
    real_password = cursor.fetchone()[0]

    # check credentials
    if entered_password.upper() == real_password.upper():
        # log in
        General.app.current_user = selected_user.replace("'", "''")

        # set the default login name in config file to the one just entered
        login_name = ''
        config = ConfigParser.ConfigParser()
        config.read('ECRev.cfg')
        config.set('Application', 'login_name', selected_user)
        config.set('Application', 'password', entered_password)
        config.set('Application', 'plant', selected_plant)
        config.set('Application', 'remember_password', remember_password )


        with open('ECRev.cfg', 'w+') as configfile:
            config.write(configfile)

        ##make ecr form choice boxes defaulted to logged in user
        ##ctrl(General.app.main_frame, 'choice:ecr_requestor').SetStringSelection(str(General.app.current_user))

        ##default user's department choice box
        ##cursor.execute("SELECT department FROM employees WHERE name = \'{}\'".format(General.app.current_user))
        ##ctrl(General.app.main_frame, 'choice:department').SetStringSelection(str(cursor.fetchone()[0]))

        # clear out password entry on dialog
        # ctrl(General.app.login_frame, 'text:password').SetValue('')



        # set window's title with user's name reordered as first then last name
        # reordered_name = selected_user.replace(' ', '') #remove any spaces
        # reordered_name = reordered_name.split(',')[1] +' '+ reordered_name.split(',')[0]
        # General.app.main_frame.SetTitle('Zookeeper - Logged in as {}'.format(reordered_name))


        # General.app.main_frame.Show()
        General.app.init_main_frame()
        General.app.login_frame.Destroy()
        General.app.login_frame = None
    else:
        wx.MessageBox('Invalid password for user %s.' % selected_user, 'Login failed')


def on_click_logout(event):
    # General.app.init_login_dialog()
    General.app.init_login_frame()
    General.app.main_frame.Destroy()
    General.app.main_frame = None


# General.app.main_frame.Hide()

# wx.MessageBox('made it this far', 'derp')



def open_bom(event):
    event.GetEventObject().Disable()
    sales_order = ctrl(General.app.main_frame, 'label:order_panel_sales_order').GetLabel()[:6]
    item_number = ctrl(General.app.main_frame, 'label:order_panel_item_number').GetLabel()

    if not OrderFileOpeners.open_bom(sales_order, item_number):
        print 'Failed to open BOM for item number \'{}\''.format(item_number)
    event.GetEventObject().Enable()


def open_dataplate(event):
    event.GetEventObject().Disable()
    sales_order = ctrl(General.app.main_frame, 'label:order_panel_sales_order').GetLabel()[:6]
    item_number = ctrl(General.app.main_frame, 'label:order_panel_item_number').GetLabel()

    if not OrderFileOpeners.open_dataplate(sales_order, item_number):
        print 'Failed to open dataplate for item number \'{}\''.format(item_number)
    event.GetEventObject().Enable()


def open_folder(event):
    event.GetEventObject().Disable()
    sales_order = ctrl(General.app.main_frame, 'label:order_panel_sales_order').GetLabel()

    if not OrderFileOpeners.open_folder(sales_order):
        print 'Failed to open folder for sales order \'{}\''.format(sales_order)
    event.GetEventObject().Enable()


def open_piping(event):
    event.GetEventObject().Disable()
    sales_order = ctrl(General.app.main_frame, 'label:order_panel_sales_order').GetLabel()[:6]
    item_number = ctrl(General.app.main_frame, 'label:order_panel_item_number').GetLabel()

    if not OrderFileOpeners.open_piping(sales_order, item_number):
        print 'Failed to open piping diagram for item number \'{}\''.format(item_number)
    event.GetEventObject().Enable()


def open_wiring(event):
    event.GetEventObject().Disable()
    sales_order = ctrl(General.app.main_frame, 'label:order_panel_sales_order').GetLabel()[:6]
    item_number = ctrl(General.app.main_frame, 'label:order_panel_item_number').GetLabel()

    if not OrderFileOpeners.open_wiring(sales_order, item_number):
        print 'Failed to open wiring diagram for item number \'{}\''.format(item_number)
    event.GetEventObject().Enable()


def open_workbook(event):
    event.GetEventObject().Disable()
    sales_order = ctrl(General.app.main_frame, 'label:order_panel_sales_order').GetLabel()[:6]
    item_number = ctrl(General.app.main_frame, 'label:order_panel_item_number').GetLabel()

    if not OrderFileOpeners.open_workbook(sales_order, item_number):
        print 'Failed to open workbook for item number \'{}\''.format(item_number)
    event.GetEventObject().Enable()


def open_legend(event):
    event.GetEventObject().Disable()
    sales_order = ctrl(General.app.main_frame, 'label:order_panel_sales_order').GetLabel()[:6]
    item_number = ctrl(General.app.main_frame, 'label:order_panel_item_number').GetLabel()

    if not OrderFileOpeners.open_legend(sales_order, item_number):
        print 'Failed to open legend for item number \'{}\''.format(item_number)
    event.GetEventObject().Enable()


if __name__ == '__main__':
    print 'starting app'

    try:
        General.app = ECRevApp(False)
        Database.connection = Database.connect_to_database()
        cursor = Database.connection.cursor()
        General.attachment_directory = cursor.execute("SELECT TOP 1 attachment_directory FROM administration").fetchone()[0]
        if Database.connection:
            if check_for_updates() == True:
                wx.CallAfter(open_software_update_frame)
            General.app.init_login_frame()
            General.app.MainLoop()

    except Exception as e:
        print 'Failed update check:', e


        #if General.check_for_updates(version) == False:
             #Database.connection = Database.connect_to_database()

             #set attachment directory
             #config = ConfigParser.ConfigParser()
             #config.read('ECRev.cfg')
             #General.attachment_directory = config.get('Application', 'attachment_directory')

    except Exception as e:
        # kill off the spash screen if it's still up
        for proc in psutil.process_iter():
            if proc.name == 'ECRev.exe':  # 'SplashScreenStarter.exe'
                proc.kill()

        if e[0] == 'IM002':
            help_message = "Your DSN pointing to the ECR database is probably not set up correctly.\nRunning 'S:\Everyone\Management Software\SetupDSN.bat' may fix it.\n\n"

        if e[0] == '28000':
            help_message = "It appears that your username is not authorized to access the eng04_sql database\nRequest permission from IT\n\n"

        wx.MessageBox('{}{}\n\n{}'.format(help_message, e, traceback.format_exc()), 'An error occurred!',
                      wx.OK | wx.ICON_ERROR)
        print 'An error occurred!'

    # clean up some loose ends
    try:
        if General.app.timer_thread != None:
            General.app.timer_thread.join()
            print 'joined last timer thread'
    except Exception as e:
        print e

    try:
        Database.connection.close()
    except Exception as e:
        print e

    try:
        General.app.Destroy()
    except Exception as e:
        print e

    print 'Attempting to end program (outside of app)'
    # sys.exit('ECRev Program Exited') #doesn't actually kill the process everytime for some reason...
    os._exit(0)
