import datetime as dt
import time
import os
import sys
import wx
from subprocess import Popen
import shlex
import psutil  # for killing the splash screen process
import json
import Database

import ConfigParser
import shutil  # for coping files (like app update files)

# global variable hold reference to wxGUI
app = None

# attachment_directory = r'S:\Everyone\Management Software\ECRev\ecr attachments'
attachment_directory = None

updates_dir = r"C:\Users\sdb25\Desktop\Mydocs\New Trieal"


def format_date_nicely(date_string):
    date_string = str(date_string)
    if date_string == '' or date_string == None:
        return

    dt_object = time.strptime(date_string, "%Y-%m-%d %H:%M:%S")  # to python time object
    return time.strftime("%m/%d/%y   %I:%M %p", dt_object)


def resource_path(relative):
    try:
        return os.path.join(sys._MEIPASS, relative)
    except:
        return relative


def check_for_updates(current_version):
    try:
        with open(os.path.join(updates_dir, "releases.json")) as file:
            releases = json.load(file)

            latest_version = releases[0]['version']
            print latest_version

            if current_version != latest_version:
                wx.MessageBox("A new version of ECRev is found on Server. You must update to continue",
                              "Software Update Available", wx.OK | wx.ICON_INFORMATION)
                install_filename = releases[0]['installer filename']
                source_filepath = os.path.join(updates_dir, install_filename)
                os.startfile(source_filepath)
                return True
            else:
                return False

    except Exception as e:
        print 'Failed update check:', e
    """   
	try:
		sys._MEIPASS
	except:
		print 'Not checking for updates because this is the dev version'
		return False
	
	config = ConfigParser.ConfigParser()
	if config.read('ECRev.cfg'):
		check_for_updates = config.get('Application', 'check_for_updates')
		#directory_to_check_for_updates = config.get('Application', 'directory_to_check_for_updates')
		
		cursor = Database.connection.cursor()
		directory_to_check_for_updates = cursor.execute("SELECT TOP 1 updates_directory FROM administration").fetchone()[0]
		
		if check_for_updates == 'True':
			
			try:
				os.listdir(directory_to_check_for_updates)
			except Exception as e:
				#kill off the spash screen if it's still up
				for proc in psutil.process_iter():
					if proc.name == 'ECRev.exe':
						proc.kill()
				wx.MessageBox("ECRev requires a connection to the S: drive to function.\nMount the drive manually by navigating to a folder on the S: drive\nor contact IT to get set up.\n\n{}".format(e), 'ERROR', wx.OK | wx.ICON_ERROR)

			
			latest_version = float(current_version)
			latest_file = None
			
			for file_name in os.listdir(directory_to_check_for_updates):
				#print 'file_name[-4:]={}'.format(file_name[-4:])
				#raw_input('press any key')
				###if file_name[-4:] == '.app':
				if file_name[-4:] == '.exe':
					if file_name != 'ECRev.exe':
						#print 'file_name[:4]={}'.format(file_name[:4])
						#raw_input('press any key')
						app_version = float(file_name[:4])/10.
						#print 'current_version: {}, found_version: {}'.format(float(current_version), app_version)

						if app_version > latest_version:
							latest_version = app_version
							latest_file = file_name
			
			if latest_file:
				updating_frame = app.res.LoadFrame(None, 'frame:updating')
				updating_frame.Show()
				updating_frame.Refresh()
				updating_frame.Update()
				wx.Yield()
				
				#kill off the spash screen if it's still up
				for proc in psutil.process_iter():
					if proc.name == 'ECRev.exe':
						proc.kill()
				
				print 'copying latest file: {}'.format(latest_file)
				shutil.copyfile(os.path.join(directory_to_check_for_updates, latest_file), os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), latest_file))
				
				
				#i can't currently automatically relaunch the app because of a bug in pyinstaller...
				# see: http://www.mail-archive.com/pyinstaller@googlegroups.com/msg04773.html
				#print 'file is copied... relaunch splash'
				#Popen(os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "ECRev.exe"), shell=False)
				
				updating_frame.Destroy()
				wx.MessageBox('ECRev has successful updated to version {}\n\nYou can now relaunch the application.'.format(latest_version), 'Update Successful', wx.OK | wx.ICON_INFORMATION)
				
				
				
				#Popen(os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "AppUpdater.exe"), shell=False)
				#Popen(os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "LaunchAppUpdater.bat"), shell=False)
				
				#print os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "ECRev.exe")
				

				#print os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "ECRev.exe")
				#Popen(os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "ECRev.exe"), shell=True)

				#Popen(os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "RelaunchECRev.bat"), shell=False)
				

				#print sys.argv[0].split('\\')
				

				#hey = r'\'{}\''.format(os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "AppUpdater.exe"))
				
				#hey = []
				#hey.append(os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "AppUpdater.exe"))
				
				#print hey
				#Popen(hey, shell=True)

				#print 'try again...'
				#print hey
				#Popen('\'{}\''.format(hey), shell=True)


				'''
				try:
					print 'about to popen..'
					Popen(hey)
					
				except:
					print 'well that didnt work'
				'''

				#command_line = '\'{}\' ECRev.app {} ECRev.exe \'{}\''.format(os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "AppUpdater.exe"), latest_file, sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''))
				#wx.MessageBox('Command line:\n{}'.format(command_line), 'FYI', wx.OK | wx.ICON_ERROR)
				#wx.MessageBox('ready? _____________________________', 'FYI', wx.OK | wx.ICON_ERROR)
				
				#args = shlex.split(command_line)
				
				#args = ['\'{}\''.format(os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "AppUpdater.exe")), 'ECRev.app', latest_file, 'ECRev.exe', '\'{}\''.format(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''))]
				#print args
				
				#wx.MessageBox('yoyoyoyo ___________________________________', 'FYI', wx.OK | wx.ICON_ERROR)
				
				#Popen(args[0])
				
				#Popen([os.path.join(sys.argv[0].replace(sys.argv[0].split('\\')[-1], ''), "AppUpdater.exe"), "ECRev.app", latest_file, "ECRev.exe"])
				
		
				#this_cwd_is = os.getcwd()
				#this_exe_dir = sys.argv[0].replace(sys.argv[0].split('\\')[-1], '')
				#assert os.path.isdir(this_exe_dir)
				#os.chdir(this_exe_dir)
				
				#Popen(args, shell=True)
				
				#assert os.path.isdir(this_cwd_is)
				#os.chdir(this_cwd_is)
				
				#wx.MessageBox('did it work? _______________________________________', 'FYI', wx.OK | wx.ICON_ERROR)

				return True
			else:
				return False
			
	else:
		wx.MessageBox('Could not locate \'ECRev.cfg\' in the ECRev folder.\nUnable to read update information.', 'ERROR', wx.OK | wx.ICON_ERROR)
		return False
"""
