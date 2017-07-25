import wx  # wxWidgets used as the GUI
from wx.html import HtmlEasyPrinting
from wx import xrc  # allows the loading and access of xrc file (xml) that describes GUI
import wx.grid as gridlib
from wxPython.calendar import *

ctrl = xrc.XRCCTRL  # define a shortined function name (just for convienience)

import sys
import re  # for regular expressions
import os
import xlwt

from functools import partial  # allows us to pass more arguments to a function "Bind"ed to a GUI event

import Database
import General
import TweakedGrid
import Ecrs


def export_search_results(event):
    # prompt user to choose where to save
    save_dialog = wx.FileDialog(General.app.main_frame, message="Export file as ...",
                                defaultDir=os.path.join(os.path.expanduser("~"), "Desktop"),
                                defaultFile="search_results", wildcard="Excel Spreadsheet (*.xls)|*.xls",
                                style=wx.SAVE | wx.OVERWRITE_PROMPT)

    # show the save dialog and get user's input... if not canceled
    if save_dialog.ShowModal() == wx.ID_OK:
        save_path = save_dialog.GetPath()
        save_dialog.Destroy()
    else:
        save_dialog.Destroy()
        return

    # save results data to excel
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Search Results')

    results_list = ctrl(General.app.main_frame, 'list:search_results')
    # current_item = results_list.GetTopItem()

    try:
        # write out headers
        for col in range(results_list.GetColumnCount()):
            worksheet.write(0, col, results_list.GetColumn(col).GetText())

        # write out data
        for row in range(results_list.GetItemCount()):
            for col in range(results_list.GetColumnCount()):
                worksheet.write(row + 1, col, results_list.GetItem(row, col).GetText())

        workbook.save(save_path)

        wx.MessageBox('Export completed.', 'Info', wx.OK | wx.ICON_INFORMATION)

    except Exception as e:
        wx.MessageBox('Export failed!\n\nYou may have exceeded the row limit of an xls file.\n{}'.format(e), 'Error',
                      wx.OK | wx.ICON_ERROR)


def destroy_dialog(event, dialog):
    dialog.Destroy()


def multy_split(s, seps):
    res = [s]
    for sep in seps:
        s, res = res, []
        for seq in s:
            # res += seq.split(sep)
            res += re.split('({})(?i)'.format(sep), seq)
    return res


def on_change_grid_cell(event):
    ctrl(General.app.main_frame, 'text:sql_query').SetValue(generate_sql_query())
    event.Skip()


def on_click_begin_search(event):
    # dangerous_sql_commands = ['INSERT', 'CREATE', 'DROP', 'UPDATE', 'GRANT', 'REVOKE', 'VARIABLES', 'OUTFILE', 'SHUTDOWN', 'RELOAD', 'SUPER', 'PROCESS', 'EXECUTE']
    dangerous_sql_commands = []

    ctrl(General.app.main_frame, 'text:sql_query').SetValue(generate_sql_query())

    sql = ctrl(General.app.main_frame, 'text:sql_query').GetValue()
    for dangerous_sql_command in dangerous_sql_commands:
        if dangerous_sql_command in sql.upper():
            wx.MessageBox('{} is not permitted in the search query.'.format(dangerous_sql_command), 'Warning',
                          wx.OK | wx.ICON_WARNING)
            return

    event.GetEventObject().SetLabel('Searching...')
    find_records_in_table = ctrl(General.app.main_frame, 'choice:which_table').GetStringSelection()

    if find_records_in_table == '':
        return

    results_list = ctrl(General.app.main_frame, 'list:search_results')

    # clear out the list
    results_list.DeleteAllItems()

    column_names = Database.get_table_column_names(find_records_in_table, presentable=False)

    if results_list.GetColumn(0) != None:
        results_list.DeleteAllColumns()

    # populate column names
    for index, column_name in enumerate(column_names):
        results_list.InsertColumn(index, column_name)

    # query the database
    cursor = Database.connection.cursor()
    try:
        records = cursor.execute(sql).fetchall()
    except:
        records = None

    if records != None:
        for index, record in enumerate(records):
            results_list.InsertStringItem(sys.maxint, '#')

            for column_index, column_value in enumerate(record):
                '''
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
				'''
                if column_value != None:
                    results_list.SetStringItem(index, column_index, str(column_value).replace('\n', ' \\ '))

    for column_index in range(len(column_names)):
        results_list.SetColumnWidth(column_index, wx.LIST_AUTOSIZE_USEHEADER)

    event.GetEventObject().SetLabel('Begin Search')


def on_click_open_how_to_search(event):
    how_to_search_frame = General.app.res.LoadFrame(None, 'frame:search_how_to')
    how_to_search_frame.Bind(wx.EVT_BUTTON, partial(destroy_dialog, dialog=how_to_search_frame),
                             id=xrc.XRCID('button:close'))
    how_to_search_frame.Show()


def search_criteria_to_sql(column, criteria):
    # operators = ['=', '==', '<', '<=', '>', '>=', '!=', '<>', 'IN', 'NOT IN', 'BETWEEN', 'IS', 'IS NOT', 'LIKE']

    # operators = ['==', '<=', '>=', '!=', '<>', '=', '<', '>', 'IN', 'NOT IN', 'BETWEEN', 'IS', 'IS NOT', 'LIKE']
    operators = ['==', '<=', '>=', '!=', '<>', '=', '<', '>', 'LIKE ']

    tokens = ['AND ', 'OR ']

    # distinguish betweeen criteria that is a legit number a pesudo number... aka and item number
    # which starts with zero which must be preserved
    try:
        float(criteria)
        if criteria[0] == '0':
            is_number = False
        else:
            is_number = True
    except:
        is_number = False

    # prepend an = sign if no operator specified
    operator_found = False
    for operator in operators:
        if operator in criteria:
            operator_found = True

    if operator_found == False:
        # print 'is_number: {}'.format(is_number)

        if is_number == True:
            criteria = '= ' + criteria
        else:
            # criteria = 'LIKE \'%{}%\''.format(criteria)
            criteria = 'LIKE %{}%'.format(criteria)

    # first seperate by tokens
    split_by_tokens = multy_split(criteria, tokens)
    # print split_by_tokens

    new_criteria_string = ''

    for split_by_token in split_by_tokens:
        if split_by_token.upper().strip() in tokens:
            new_criteria_string += ' {}'.format(split_by_token)

        else:
            seperated_criteria_list = multy_split(split_by_token, operators)
            # print seperated_criteria_list

            for seperated_index, seperated_criteria in enumerate(seperated_criteria_list):
                for operator in operators:
                    if seperated_criteria == operator:
                        if seperated_criteria_list[seperated_index + 1] not in operators:
                            seperated_criteria_list[seperated_index + 1] = seperated_criteria_list[
                                seperated_index + 1].strip()
                            seperated_criteria_list[seperated_index + 2:seperated_index + 2] = "'"
                            seperated_criteria_list[seperated_index + 1:seperated_index + 1] = "'"

            # insert the column name for this particular criteria
            for seperated_index, seperated_criteria in enumerate(reversed(seperated_criteria_list)):
                for operator in operators:
                    if seperated_criteria == operator:
                        if seperated_criteria_list[seperated_index - 1] not in operators:
                            seperated_criteria_list[seperated_index - 2] = ' {} '.format(
                                seperated_criteria_list[seperated_index - 2])
                            seperated_criteria_list[seperated_index - 2:seperated_index - 2] = column

            new_criteria_string += ''.join(seperated_criteria_list)

    new_criteria_string = new_criteria_string.replace(r'""', r'"')

    sql = '{} AND '.format(new_criteria_string)
    return sql


def on_select_result(event):
    if ctrl(General.app.main_frame, 'choice:which_table').GetStringSelection() == 'ecrs':
        # print 'yay'
        item = event.GetEventObject()

        Ecrs.populate_ecr_order_panel(item_number=item.GetItem(item.GetFirstSelected(), 3).GetText())
        Ecrs.populate_ecr_panel(id=item.GetItem(item.GetFirstSelected(), 0).GetText())


def on_select_table(event):
    table_panel = ctrl(General.app.main_frame, 'panel:search_criteria_table')

    # remove the table if there is already one there
    if General.app.table_search_criteria != None:
        General.app.table_search_criteria.Destroy()

    General.app.table_search_criteria = TweakedGrid.TweakedGrid(table_panel)

    # table.Bind(gridlib.EVT_GRID_CELL_LEFT_CLICK, Revisions.on_click_table_cell)



    table_to_search = event.GetEventObject().GetStringSelection()

    columns = list(Database.get_table_column_names(table_to_search, presentable=False))

    for column_index, column in enumerate(columns):
        columns[column_index] = '{}.{}'.format(table_to_search, column)

    for joining_table in General.app.joining_tables:
        if table_to_search == joining_table[0]:
            extend_list = ['']
            # extend_list.append('-{}'.format(joining_table[1]))
            extend_list.extend(list(Database.get_table_column_names(joining_table[1], presentable=False)))
            extend_list.remove(joining_table[3])

            for column_index, column in enumerate(extend_list):
                if column_index > 0:
                    extend_list[column_index] = '{}.{}'.format(joining_table[1], column)

            columns.extend(extend_list)

    # sort columns alphabeticaly <--(lol spelling)
    # columns.sort()


    General.app.table_search_criteria.CreateGrid(len(columns), 2)
    General.app.table_search_criteria.SetRowLabelSize(0)
    General.app.table_search_criteria.SetColLabelValue(0, 'Field')
    General.app.table_search_criteria.SetColLabelValue(1, 'Criteria')

    for column_index, column in enumerate(columns):
        if column != '':
            General.app.table_search_criteria.SetCellValue(column_index, 0, column)
            General.app.table_search_criteria.SetCellValue(column_index, 1, '')
        else:
            General.app.table_search_criteria.SetReadOnly(column_index, 1)
        # General.app.table_search_criteria.SetCellValue(column_index, 0, column)
        # General.app.table_search_criteria.SetCellValue(column_index, 1, 'Fields from {}:'.format(column[1:]))

    # General.app.table_search_criteria.SetCellValue(0, 0,' (click to select document) ')
    General.app.table_search_criteria.AutoSize()

    General.app.table_search_criteria.EnableDragRowSize(False)

    General.app.table_search_criteria.Bind(wx.EVT_SIZE, on_size_criteria_table)
    General.app.table_search_criteria.Bind(wx.grid.EVT_GRID_CELL_CHANGE, on_change_grid_cell)
    # General.app.table_search_criteria.Bind(wx.EVT_CHAR, on_change_grid_cell)


    # for row in range(1, 25):
    #	for col in range(2):
    #		table.SetCellValue(row, col,"cell (%d,%d)" % (row, col))

    for row in range(len(columns)):
        General.app.table_search_criteria.SetReadOnly(row, 0)

    sizer = wx.BoxSizer(wx.VERTICAL)
    sizer.Add(General.app.table_search_criteria, 1, wx.EXPAND)
    table_panel.SetSizer(sizer)

    table_panel.Layout()


def generate_sql_query():
    print 'generating sql query'
    find_records_in_table = ctrl(General.app.main_frame, 'choice:which_table').GetStringSelection()

    sql = 'SELECT '

    # limit the records pulled if desired
    if ctrl(General.app.main_frame, 'choice:search_limit').GetStringSelection() != '(no limit)':
        sql += "TOP {} ".format(
            int(ctrl(General.app.main_frame, 'choice:search_limit').GetStringSelection().split(' ')[0]))

    sql += '* FROM {} '.format(find_records_in_table)

    sql_criteria = ''
    tables_to_join = []

    # loop through fields
    for row in range(General.app.table_search_criteria.GetNumberRows()):
        if General.app.table_search_criteria.GetCellValue(row, 1) != '':
            if General.app.table_search_criteria.GetCellValue(row, 0) != '':  # skip over spacer cell (between tables)
                tables_to_join.append(General.app.table_search_criteria.GetCellValue(row, 0).split('.')[0])
                sql_criteria += search_criteria_to_sql(
                    # column = General.app.table_search_criteria.GetCellValue(row, 0).split(':')[-1].strip(),
                    column=General.app.table_search_criteria.GetCellValue(row, 0),
                    criteria=General.app.table_search_criteria.GetCellValue(row, 1))

    # remove duplicates from list
    tables_to_join = list(set(tables_to_join))

    for table_to_join in tables_to_join:
        if table_to_join != find_records_in_table:
            for joining_table in General.app.joining_tables:
                if (find_records_in_table == joining_table[0]) and (table_to_join == joining_table[1]):
                    sql += 'INNER JOIN {} ON {}.{} = {}.{} '.format(joining_table[1], joining_table[0],
                                                                    joining_table[2], joining_table[1],
                                                                    joining_table[3])
                    break

    # INNER JOIN orders_cases ON ecrs.item = orders_cases.item
    sql += 'WHERE {}'.format(sql_criteria[:-4])

    ##limit the records pulled if desired
    # if ctrl(General.app.main_frame, 'choice:search_limit').GetStringSelection() != '(no limit)':
    #	sql += "LIMIT {}".format(int(ctrl(General.app.main_frame, 'choice:search_limit').GetStringSelection().split(' ')[0]))

    return sql


def on_size_criteria_table(event):
    table = event.GetEventObject()
    table.SetColSize(1, table.GetSize()[0] - table.GetColSize(0) - wx.SystemSettings.GetMetric(wx.SYS_VSCROLL_X))
    # table.SetColSize(1, table.GetSize()[0] - table.GetColSize(0))

    # print wx.SystemSettings.GetMetric(wx.SYS_VSCROLL_X)
    # print wx.SystemSettings.GetMetric(wx.SYS_VSCROLL_ARROW_X)
    event.Skip()
