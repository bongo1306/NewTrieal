import os
# import datetime
# from operator import attrgetter
# import wx
# import pyodbc
# import win32clipboard
import win32com.client
import traceback
# import ConfigParser
import subprocess  # allows opening of explorer folder


# recursively search head_dir for dir_name
def find_folder(head_dir, dir_name):
    for root, dirs, files in os.walk(head_dir):
        for d in dirs:
            if d.upper() == dir_name.upper():
                return str(os.path.join(root, d))
    return head_dir


def open_bom(sales_order, item_number):
    order_directory = get_order_directory(sales_order)

    if sales_order == '':
        return

    bom_files = []
    correct_bom_file = ''

    for root, dirs, files in os.walk(order_directory):
        for f in files:
            if f.upper().find(sales_order + '-BOM.') != -1:
                correct_bom_file = str(os.path.join(root, f))
            if f.upper().find('BOM') != -1 and f.upper().find('.BOM') == -1 and f.upper().find('BOMRE') == -1:
                bom_files.append(str(os.path.join(root, f)))

    xlApp = win32com.client.dynamic.Dispatch('Excel.Application')

    for bom_file in bom_files:
        try:
            xlBook = xlApp.Workbooks.Open(bom_file)

            sht = xlBook.Worksheets('BOM')
            for cell in range(3, 20):
                if str(sht.Cells(1, cell).Value).find(item_number) != -1:
                    correct_bom_file = bom_file
                    break

            xlBook.Close(SaveChanges=0)
        except:
            pass

    del xlApp

    if correct_bom_file != '':
        os.startfile(correct_bom_file)
    else:
        return False
    # frame.frame_1_statusbar.SetStatusText('Could not locate BOM for that Item.',1)

    '''
	bom_file = ''
	files_found = 0
	
	for root, dirs, files in os.walk(get_order_directoryectory(customer, sales_order)):
		for f in files:
			if f.upper() == sales_order+'-BOM.XLS':
				bom_file = str(os.path.join(root, f))
				files_found += 1

	if files_found == 1:
		#subprocess.Popen('cmd /c \"'+bom_file+'\" && exit')
		#subprocess.call(('start', bom_file))
		#os.system("start "+bom_file)
		os.startfile(bom_file)
		
		#child = subprocess.Popen('cmd START /c \"'+bom_file+'\" && exit', preexec_fn = os.setsid)
		#os.killpg(child.pid, signal.SIGINT)		
		#proc = subprocess.Popen('cmd START /c \"'+bom_file+'\"', stdin=subprocess.PIPE)
		#proc.stdin.write("Something")
		#proc.wait()

		#proc = subprocess.Popen('cmd /c \"'+bom_file+'\"', stderr=subprocess.STDOUT, stdout=file("YourOutput.txt", 'w'))
		#proc = subprocess.Popen('cmd /c \"'+bom_file+'\"')
		#sleep(4)
		
		#proc.wait()
		#proc.kill()
	elif files_found > 1:
		frame.frame_1_statusbar.SetStatusText('Multiple files found, so you pick.',1)
		subprocess.Popen('explorer '+get_order_directoryectory(customer, sales_order))
	else:
		frame.frame_1_statusbar.SetStatusText('No file found. You look for it. It might just be named oddly.',1)
		subprocess.Popen('explorer '+get_order_directoryectory(customer, sales_order))
	'''


def open_piping(sales_order, item_number):
    piping_file = ''
    files_found = 0
    order_directory = get_order_directory(sales_order)

    if order_directory:
        for root, dirs, files in os.walk(order_directory):
            for f in files:
                if f.upper() == item_number + '-PD.DWG':
                    piping_file = str(os.path.join(root, f))
                    files_found += 1

        if files_found == 1:
            # subprocess.Popen('cmd /c \"'+piping_file+'\" && exit')
            os.startfile(piping_file)
            return True
        elif files_found > 1:
            # multiple possible files found, user must pick
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False
        else:
            # no file found, might be named oddly, user must find
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False
    return False


def open_wiring(sales_order, item_number):
    wiring_file = ''
    files_found = 0
    order_directory = get_order_directory(sales_order)

    if order_directory:
        for root, dirs, files in os.walk(order_directory):
            for f in files:
                if f.upper() == item_number + '-WD.DWG':
                    wiring_file = str(os.path.join(root, f))
                    files_found += 1

        if files_found == 1:
            # subprocess.Popen('cmd /c \"'+wiring_file+'\" && exit')
            os.startfile(wiring_file)
            return True
        elif files_found > 1:
            # frame.frame_1_statusbar.SetStatusText('Multiple files found, so you pick.',1)
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False
        else:
            # frame.frame_1_statusbar.SetStatusText('No file found. You look for it. It might just be named oddly.',1)
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False

    return False


def open_dataplate(sales_order, item_number):
    dataplate_file = ''
    files_found = 0

    order_directory = get_order_directory(sales_order)
    if order_directory:
        for root, dirs, files in os.walk(order_directory):
            for f in files:
                if f.upper() == sales_order + '-DP.XLS':
                    dataplate_file = str(os.path.join(root, f))
                    files_found += 1
                if f.upper() == item_number + '-DP.XLS':
                    dataplate_file = str(os.path.join(root, f))
                    files_found += 1

        if files_found == 1:
            # subprocess.Popen('cmd /c \"'+dataplate_file+'\" && exit')
            os.startfile(dataplate_file)
            return True
        elif files_found > 1:
            # frame.frame_1_statusbar.SetStatusText('Multiple files found, so you pick.',1)
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False
        else:
            # frame.frame_1_statusbar.SetStatusText('No file found. You look for it. It might just be named oddly.',1)
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False


def open_workbook(sales_order, item_number):
    workbook_file = ''
    files_found = 0

    order_directory = get_order_directory(sales_order)
    if order_directory:
        for root, dirs, files in os.walk(order_directory):
            for f in files:
                if f.upper() == sales_order + '-WB.XLS':
                    workbook_file = str(os.path.join(root, f))
                    files_found += 1

        if files_found == 1:
            # subprocess.Popen('cmd /c \"'+workbook_file+'\" && exit')
            os.startfile(workbook_file)
        elif files_found > 1:
            # frame.frame_1_statusbar.SetStatusText('Multiple files found, so you pick.',1)
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False
        else:
            # frame.frame_1_statusbar.SetStatusText('No file found. You look for it. It might just be named oddly.',1)
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False

    return False


def open_legend(sales_order, item_number):
    legend_file = ''
    files_found = 0

    order_directory = get_order_directory(sales_order)
    if order_directory:
        for root, dirs, files in os.walk(order_directory):
            for f in files:
                if 'LEGEND' in f.upper():
                    # print f.upper()
                    if f.upper()[-3:] == 'PDF':
                        if files_found == 0:
                            legend_file = str(os.path.join(root, f))
                            files_found += 1
                    else:
                        legend_file = str(os.path.join(root, f))
                        files_found += 1

        if files_found == 1:
            # subprocess.Popen('cmd /c \"'+legend_file+'\" && exit')
            os.startfile(legend_file)
        elif files_found > 1:
            # frame.frame_1_statusbar.SetStatusText('Multiple files found, so you pick.',1)
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False
        else:
            # frame.frame_1_statusbar.SetStatusText('No file found. You look for it. It might just be named oddly.',1)
            subprocess.Popen('explorer \"{}\"'.format(order_directory))
            return False
    return False


'''
This def no longer useable since Mike's new ecr program...
def open_ecr(frame):
	ecrID = frame.label_1_copy_copy.GetLabel()

	if ecrID == "":
		return

	fff = open('requestID.txt', 'w')
	fff.write(ecrID)
	#fff.write('\n')
	#fff.write(os.curdir)
	fff.close()

	#subprocess.Popen('cmd /c \"\\\\kw_engineering\sharepoint$\Everyone\ENGR_REQUESTS\ECRmonitor\ECR_Injector.xlsm\" && exit')
	#subprocess.Popen('cmd /c ECR_Injector.xlsm')
	os.startfile('ECR_Injector.xlsm')
'''


def upload_bom(frame):
    customer = frame.label_10.GetLabel()[11:]
    sales_order = frame.label_8.GetLabel()[14:]
    item_number = frame.REF_NUMBER.GetLabel()
    so_dir = get_so_directory(customer, sales_order)

    if so_dir == '':
        frame.frame_1_statusbar.SetStatusText('Could not locate Order.', 1)
        return

    if customer == '' or sales_order == '':
        return

    bom_files = []
    correct_bom_file = ''

    for root, dirs, files in os.walk(so_dir):
        for f in files:
            if f.upper().find(sales_order + '-BOM.') != -1:
                correct_bom_file = str(os.path.join(root, f))
            if f.upper().find('BOM') != -1 and f.upper().find('.BOM') == -1:
                bom_files.append(str(os.path.join(root, f)))

    xlApp = win32com.client.dynamic.Dispatch('Excel.Application')

    for bom_file in bom_files:
        try:
            xlBook = xlApp.Workbooks.Open(bom_file)

            sht = xlBook.Worksheets('BOM')
            for cell in range(3, 20):
                if str(sht.Cells(1, cell).Value).find(item_number) != -1:
                    correct_bom_file = bom_file
                    break

            xlBook.Close(SaveChanges=0)
        except:
            pass

    del xlApp

    if correct_bom_file != '':
        fff = open('uploadBOMdata.txt', 'w')
        fff.write(item_number)
        fff.write('\n')
        fff.write(correct_bom_file)
        fff.close()

        uploader = os.getcwd() + '\\uploadBOM_Injector.xlsm'
        subprocess.Popen('cmd /c \"' + uploader + '\" && exit')
    # os.startfile('uploadBOM_Injector.xlsm')
    else:
        frame.frame_1_statusbar.SetStatusText('Could not locate BOM for that Item.', 1)


def open_folder(sales_order):
    # clean up and seperate sales order if needed
    sales_order = sales_order.split('-')[0]

    if len(sales_order) == 6:
        bpcs_so = sales_order
        sap_so = None

    elif len(sales_order) == 8:
        bpcs_so = None
        sap_so = sales_order

    elif len(sales_order) == 15:
        bpcs_so = sales_order.split('/')[0]
        sap_so = sales_order.split('/')[1]

    else:
        bpcs_so = None
        sap_so = None

    if sap_so:
        sap_order_folder_path = find_sap_order_folder_path(sap_so)

        if sap_order_folder_path:
            subprocess.Popen('explorer "{}"'.format(sap_order_folder_path))

    if bpcs_so:
        bpcs_order_folder_path = find_bpcs_order_folder_path(bpcs_so)

        if bpcs_order_folder_path:
            subprocess.Popen('explorer "{}"'.format(bpcs_order_folder_path))

    return True


def find_bpcs_order_folder_path(bpcs_so):
    starting_path = r"\\kw_engineering\eng_res\Design_Eng\Orders\Orders_20{}".format(bpcs_so[1:3])

    # plow through three directories deep looking for a folder named that bpcs sales order
    for x in os.listdir(starting_path):
        starting_path_x = os.path.join(starting_path, x)

        if os.path.isdir(starting_path_x):
            for y in os.listdir(starting_path_x):
                starting_path_x_y = os.path.join(starting_path_x, y)

                if os.path.isdir(starting_path_x_y):
                    for z in os.listdir(starting_path_x_y):
                        starting_path_x_y_z = os.path.join(starting_path_x_y, z)

                        if os.path.isdir(starting_path_x_y_z):
                            if z == bpcs_so:
                                return starting_path_x_y_z

    return None


# recursively look for the SAP sales order folder
def find_sap_order_folder_path(sap_so, starting_path=r"\\kw_engineering\eng_res\Design_Eng\Orders\SAP_ORDERS_COLS"):
    if os.path.split(starting_path)[-1] == sap_so:
        return starting_path

    for x in os.listdir(starting_path):
        if x not in sap_so:
            continue

        starting_path_x = os.path.join(starting_path, x)

        if os.path.isdir(starting_path_x):
            return find_sap_order_folder_path(sap_so, starting_path_x)

    return None


# returns sales order directory
def get_order_directory(sales_order):
    # clean up and seperate sales order if needed
    sales_order = sales_order.split('-')[0]

    if len(sales_order) == 6:
        bpcs_so = sales_order
        sap_so = None

    elif len(sales_order) == 8:
        bpcs_so = None
        sap_so = sales_order

    elif len(sales_order) == 15:
        bpcs_so = sales_order.split('/')[0]
        sap_so = sales_order.split('/')[1]

    else:
        bpcs_so = None
        sap_so = None

    if sap_so:
        sap_order_folder_path = find_sap_order_folder_path(sap_so)

        if sap_order_folder_path:
            return sap_order_folder_path

    if bpcs_so:
        bpcs_order_folder_path = find_bpcs_order_folder_path(bpcs_so)

        if bpcs_order_folder_path:
            return bpcs_order_folder_path

    return False
