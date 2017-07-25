# USE_DATABASE_TYPE = 'SQLite'
USE_DATABASE_TYPE = 'MS-Server'

# import sqlite3 as lite	#for manipulating the test database
import pyodbc

connection = None
# cursor = None

import sys
import ConfigParser

import General

import wx
import Ecrs


def connect_to_database():
    # global cursor

    # get the connection string from the config file
    config = ConfigParser.ConfigParser()
    if config.read('ECRev.cfg'):
        connection_string = config.get('Database', 'connection_string')
    else:
        wx.MessageBox('Could not locate \'ECRev.cfg\' in the ECRev folder.\nUnable to read database connection string.',
                      'ERROR', wx.OK | wx.ICON_ERROR)
        return False

    if USE_DATABASE_TYPE == 'SQLite':
        # for prototyping with SQLite

        # return lite.connect(r'S:\Everyone\Management Software\prototype.db3')
        # db_connection = lite.connect(r'S:\Everyone\Management Software\prototype.db3')
        # db_connection = lite.connect(r'C:\Users\tms18\Desktop\local test\prototype.db3')
        # db_connection = lite.connect(r'R:\ECRev.db3')
        # db_connection = lite.connect(r'S:\Everyone\Management Software\ECRev.db3')
        db_connection = lite.connect(r'C:\Users\tms18\Desktop\local test\ECRev.db3')
        cursor = db_connection.cursor()
        cursor.execute("PRAGMA foreign_keys = ON")
        cursor.execute("PRAGMA synchronous = OFF")

        return db_connection

    elif USE_DATABASE_TYPE == 'MS-Server':
        # for real SQL server
        # print connection_string

        db_connection = pyodbc.connect(connection_string)

        # DSN=eng04_sql;APP=ECRev;Trusted_Connection=Yes

        return db_connection


'''
def get_order_data_from_ref(reference_number):
    if reference_number == None: return
    
    try:
        cursor = connection.cursor()
        ##cursor.execute("SELECT * FROM ecrs WHERE who_requested = \'{}\'".format(General.app.current_user)) ### ORDER BY date_requested ASC")
        
        #determine reference type by reference_number length and other defining characteristics
        if (len(reference_number) == 7) and (reference_number[0] == '0'):
            #it might be an item number, search for it in the DB
            return cursor.execute("SELECT * FROM {} WHERE item = \'{}\' LAMIT 1".format(Ecrs.table_used, reference_number)).fetchone()

        elif (len(reference_number) == 9) and (reference_number[0:3] == 'KW0'):
            #it might be an item number with KW prefix, search for it in the DB
            return cursor.execute("SELECT * FROM {} WHERE item = \'{}\' LAMIT 1".format(Ecrs.table_used, reference_number)).fetchone()

        elif len(reference_number) == 10:
            #it might be a serial number, search for it in the DB
            return cursor.execute("SELECT * FROM {} WHERE serial = \'{}\' LAMIT 1".format(Ecrs.table_used, reference_number)).fetchone()

        elif len(reference_number) == 6:
            #it might be a sales order, search for it in the DB
            return cursor.execute("SELECT * FROM {} WHERE sales_order = \'{}\' OR quote = \'{}\' LAMIT 1".format(Ecrs.table_used, reference_number, reference_number)).fetchone()

        elif '-' in reference_number:
            #it might be a sales order with specified line up, search for it in the DB
            sales_order = reference_number.split('-')[0]
            line_up = 1
            if reference_number.split('-')[1] != '':
                line_up = int(float(reference_number.split('-')[1]))
            return cursor.execute("SELECT * FROM {} WHERE sales_order = \'{}\' AND line_up = \'{}\' LAMIT 1".format(Ecrs.table_used, sales_order, line_up)).fetchone()

        else:
            return None
    except:
        #for whatever reason, maybe bad ref#, couldn't get any data...
        print sys.exc_info()
        return None
'''


def get_item_from_ref(reference_number):
    if reference_number == None: return
    item_number = None
    try:
        cursor = connection.cursor()

        # determine reference type by reference_number length and other defining characteristics
        if (len(reference_number) == 7) and (reference_number[0] == '0'):
            # it might be an item number, search for it in the DB
            item_number = cursor.execute(
                "SELECT TOP 1 * FROM {} WHERE item LIKE '%{}%'".format(Ecrs.table_used,reference_number)).fetchone()

        elif (len(reference_number) == 9) and (reference_number[0:3] == 'KW0'):
            # it might be an item number with KW prefix, search for it in the DB
            item_number = cursor.execute(
                "SELECT TOP 1 * FROM {} WHERE item LIKE '%{}%'".format(Ecrs.table_used,reference_number)).fetchone()

        elif (len(reference_number) == 8) and (reference_number[0] == '2'):
            # it might be a production order, search for it in the DB
            item_number = cursor.execute(
                "SELECT TOP 1 * FROM {} WHERE item LIKE '%{}%'".format(Ecrs.table_used,reference_number)).fetchone()

        elif (len(reference_number) == 8) and (reference_number[0] == '5'):
            # it might be a SAP sales order, search for it in the DB
            item_number = cursor.execute(
                "SELECT TOP 1 * FROM {} WHERE sales_order LIKE '%{}%'".format(Ecrs.table_used,reference_number)).fetchone()

        elif len(reference_number) == 10:
            # it might be a serial number, search for it in the DB
            item_number = cursor.execute(
                "SELECT TOP 1 * FROM {} WHERE serial = \'{}\'".format(Ecrs.table_used,reference_number)).fetchone()

        elif len(reference_number) == 6:
            # it might be a sales order, search for it in the DB
            item_number = cursor.execute(
                "SELECT TOP 1 * FROM {} WHERE sales_order LIKE '%{}%' OR quote = \'{}\'".format(Ecrs.table_used,
                    reference_number, reference_number)).fetchone()

        elif '-' in reference_number:
            # it might be a sales order with specified line up, search for it in the DB
            # it might be a sales order with specified line up, search for it in the DB
            sales_order = reference_number.split('-')[0]
            line_up = 1
            if reference_number.split('-')[1] != '':
                line_up = int(float(reference_number.split('-')[1]))
            item_number = cursor.execute(
                "SELECT TOP 1 * FROM {} WHERE sales_order LIKE '%{}%' AND line_up = \'{}\'".format(Ecrs.table_used,
                    sales_order, line_up)).fetchone()

        elif (len(reference_number) == 7) and (reference_number[0] == '8'):
        # it might be an item number from cases starting with 8, search for it in the DB
            item_number = cursor.execute(
                "SELECT TOP 1 * FROM {} WHERE item LIKE '%{}%'".format(Ecrs.table_used,reference_number)).fetchone()

        else:
            item_number = None
    except:
        # for whatever reason, maybe bad ref#, couldn't get any data...
        print sys.exc_info()
        item_number = None

    if item_number != None:
        item_number = item_number[0]

    return item_number


def get_table_column_names(table, presentable=False):
    cursor = connection.cursor()
    if USE_DATABASE_TYPE == 'SQLite':
        column_names = zip(*cursor.execute('PRAGMA table_info({})'.format(table)).fetchall())[1]

    elif USE_DATABASE_TYPE == 'MS-Server':
        column_names = zip(*cursor.execute(
            "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_Name='{}' ORDER by ORDINAL_POSITION".format(
                table)).fetchall())[3]

    if presentable:
        presentable_column_names = []

        for column_name in column_names:
            presentable_name = column_name

            # uppercase the first letter
            presentable_name = presentable_name[0].upper() + presentable_name[1:]

            # uppercase any letter after an underscore and replace underscore with space
            while 1:
                underscore_index = presentable_name.find('_')
                if underscore_index == -1:
                    break

                # uppercase letter after underscore
                presentable_name = presentable_name[:underscore_index + 1] + \
                                   presentable_name[underscore_index + 1].upper() + \
                                   presentable_name[underscore_index + 2:]

                # replace underscore with space
                presentable_name = presentable_name.replace('_', ' ', 1)

            presentable_column_names.append(presentable_name)

        return presentable_column_names
    else:
        return column_names


# global variable to keep track of number of ongoing querry threads
###active_query_threads = 0
# this function will run in it's own thread so that a slow query won't lock_held
# the GUI up. Pass it the SQL query and function you want the results to go to
def query_one(sql, function_to_pass_results):
    ###global active_query_threads
    ###active_query_threads += 1
    threaded_connection = connect_to_database()
    cursor = threaded_connection.cursor()
    cursor.execute(sql)
    function_to_pass_results(cursor.fetchone())

###active_query_threads -= 1
