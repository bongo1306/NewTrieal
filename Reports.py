
import wx  # wxWidgets used as the GUI
from wx.html import HtmlEasyPrinting
from wx import xrc  # allows the loading and access of xrc file (xml) that describes GUI
import wx.grid as gridlib
from wx.calendar import *
ctrl = xrc.XRCCTRL  # define a shortined function name (just for convienience)

# from datetime import date, timedelta
import os
from win32com.client import Dispatch
import datetime as dt
from dateutil.relativedelta import relativedelta
import time
import math

import traceback


import matplotlib
matplotlib.use('WXAgg')
###Nevermind the below... i changed it back because it made no difference
# I modified the default backend for matplotlib in C:\Python27\Lib\site-packages\matplotlib\mpl-data\matplotlibrc
# so that i wouldn't need to import the entire matplotlib module and change it after the fact...
###print matplotlib.matplotlib_fname()

# from scipy.interpolate import interp1d
# from numpy import linspace
# from numpy import interp

from pylab import figure, show, FormatStrFormatter, ylim
# from matplotlib.dates import MONDAY, SATURDAY, SUNDAY
# from matplotlib.finance import quotes_historical_yahoo
# from matplotlib.dates import MonthLocator, WeekdayLocator, DateFormatter, AutoDateFormatter, AutoDateLocator
import matplotlib.dates as mdates


####import numpy as num
##from numpy import arange, polyfit, asarray

import Database
import General
import Ecrs



def utc_mktime(utc_tuple):
    """Returns number of seconds elapsed since epoch
    Note that no timezone are taken into consideration.
    utc tuple must be: (year, month, day, hour, minute, second)
    """

    if len(utc_tuple) == 6:
        utc_tuple += (0, 0, 0)
    return time.mktime(utc_tuple) - time.mktime((1970, 1, 1, 0, 0, 0, 0, 0, 0))

def datetime_to_timestamp(dt):
    """Converts a datetime object to UTC timestamp"""
    return int(utc_mktime(dt.timetuple()))


def get_week_days():
	week = dt.date.today().isocalendar()[1] - 1
	year = dt.date.today().year

	d = dt.date(year, 1, 1)
	if (d.weekday() > 3):
		d = d + dt.timedelta(7 - d.weekday())
	else:
		d = d - dt.timedelta(d.weekday())
	dlt = dt.timedelta(days=(week - 1) * 7)
	return (d + dlt + dt.timedelta(days=6)).day


class PlotPanel(wx.Panel):
    """The PlotPanel has a Figure and a Canvas. OnSize events simply set a
    flag, and the actual resizing of the figure is triggered by an Idle event."""

    def __init__(self, parent, color=None, dpi=None, **kwargs):
        from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg
        from matplotlib.figure import Figure
        # initialize Panel
        if 'id' not in kwargs.keys():
            kwargs['id'] = wx.ID_ANY
        if 'style' not in kwargs.keys():
            kwargs['style'] = wx.NO_FULL_REPAINT_ON_RESIZE
        wx.Panel.__init__(self, parent, **kwargs)

        # initialize matplotlib stuff
        self.figure = Figure(None, dpi)
        self.figure.subplots_adjust(left=0.1 - .02, right=0.9 + .02, top=0.9, bottom=0)
        self.canvas = FigureCanvasWxAgg(self, -1, self.figure)
        self.SetColor(color)

        self._SetSize()
        self.draw()

        self._resizeflag = False

        self.Bind(wx.EVT_IDLE, self._onIdle)
        self.Bind(wx.EVT_SIZE, self._onSize)

    def SetColor(self, rgbtuple=None):
        """Set figure and canvas colors to be the same."""
        if rgbtuple is None:
            rgbtuple = wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNFACE).Get()
        clr = [c / 255. for c in rgbtuple]
        self.figure.set_facecolor(clr)
        self.figure.set_edgecolor(clr)
        self.canvas.SetBackgroundColour(wx.Colour(*rgbtuple))

    def _onSize(self, event):
        self._resizeflag = True
        print 'hit onsize'

    def _onIdle(self, evt):
        # print 'idling'
        if self._resizeflag:
            self._resizeflag = False
            self._SetSize()

    def _SetSize(self):
        pixels = tuple(self.parent.GetClientSize())
        self.SetSize(pixels)
        self.canvas.SetSize(pixels)
        self.figure.set_size_inches(float(pixels[0]) / self.figure.get_dpi(),
                                    float(pixels[1]) / self.figure.get_dpi())

        # if plot has a legend, scale it depending on the figure height
        if hasattr(self, 'subplot'):
            self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8,
                                prop={'size': self.figure.get_figheight() * 1.4})

    def draw(self):
        pass  # abstract, to be overridden by child classes


# 1/0... collect data in init... then just replot in draw()

class DesignEngineeringHours(PlotPanel):
    def __init__(self, parent, start_date, end_date, **kwargs):
        self.parent = parent

        self.start_date = start_date
        self.end_date = end_date

        cursor = Database.connection.cursor()
        design_engineers = list(zip(*cursor.execute(
            'SELECT name FROM employees WHERE department = \'Design Engineering\' ORDER BY name ASC').fetchall())[0])

        self.employee_data = []
        for employee_index, employee in enumerate(design_engineers):
            # print 'SELECT when_logged, hours FROM time_logs WHERE employee=\'{}\' AND when_logged>\'{}\' AND when_logged<\'{}\' ORDER BY when_logged ASC'.format(employee.replace("'", "''"), self.start_date, '{} 23:59:59'.format(self.end_date))

            self.employee_data.append((
                employee,
                # cursor.execute("SELECT when_logged, hours FROM time_logs WHERE employee=\'{}\' AND when_logged>\'{}\' AND when_logged<\'{}\' ORDER BY when_logged ASC".format(employee, self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
                cursor.execute(
                    'SELECT when_logged, hours FROM time_logs WHERE employee=\'{}\' AND when_logged>\'{}\' AND when_logged<\'{}\' ORDER BY when_logged ASC'.format(
                        employee.replace("'", "''"), self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
            )

        # print self.employee_data[-1]

        self.lowest_start_date = end_date
        for data1 in zip(*self.employee_data)[1]:
            for data2 in data1:
                when_logged = dt.datetime.strptime(str(data2[0]), "%Y-%m-%d %H:%M:%S").date()
                if when_logged < self.lowest_start_date:
                    self.lowest_start_date = when_logged
        self.lowest_start_date = self.lowest_start_date - dt.timedelta(days=1)

        # initiate plotter
        PlotPanel.__init__(self, parent, **kwargs)
        self.SetColor((255, 255, 255))

    def draw(self):
        """Draw data."""
        if not hasattr(self, 'subplot'):
            self.subplot = self.figure.add_subplot(111)

        total_hours = 0

        for index, data in enumerate(self.employee_data):
            x = [dt.datetime.strptime(str(p[0]), "%Y-%m-%d %H:%M:%S") for p in data[1]]

            # skip employee if they have no time logs
            if not x:
                continue

            hours_sum = 0
            y = []
            for p in data[1]:
                hours_sum += p[1]
                y.append(hours_sum)

            employee = data[0]

            total_hours += hours_sum

            self.subplot.plot_date(x, y, '-', label='{}: {:.1f}'.format(employee.split(',')[0], hours_sum))

            annotation_x_index = int((len(x) - index) * .92)
            annotation_y_index = int((len(y) - index) * .92)

            if annotation_x_index < 0: annotation_x_index = len(x) - 1
            if annotation_y_index < 0: annotation_y_index = len(y) - 1

            self.subplot.annotate(employee.split(',')[0], xy=(x[annotation_x_index], y[annotation_y_index]),
                                  xytext=(-15, 15),
                                  textcoords='offset points', ha='center', va='bottom',
                                  bbox=dict(boxstyle='round,pad=0.2', fc='yellow', alpha=0.5),
                                  arrowprops=dict(arrowstyle='->', connectionstyle='arc3', color='black'))

        # plot SWAG
        '''
        swag_days = (self.end_date - self.lowest_start_date).days
        if swag_days < 1: swag_days = 1
        swag_x = [self.end_date - dt.timedelta(days=x) for x in range(0, swag_days)]
        swag_x = swag_x[::-1]
        swag_y = []
        for swag_index, swag_date in enumerate(swag_x):
            swag_y.append(swag_index*4.873)
            
        self.subplot.plot_date(swag_x, swag_y, '--', label='SWAG: {:.1f}'.format(swag_days*4.873))
        '''

        ticks = get_ticks(self.start_date, self.end_date)

        self.subplot.xaxis.set_major_locator(ticks[0])
        self.subplot.xaxis.set_major_formatter(ticks[1])

        if ticks[2] != None:
            self.subplot.xaxis.set_minor_locator(ticks[2])
            self.subplot.xaxis.set_minor_formatter(ticks[3])

        self.subplot.xaxis.set_label_text('Time Period')
        self.subplot.yaxis.set_label_text('Hours Logged')
        self.subplot.set_title(r"Design Engineers' Logged Order Hours   ({:.1f} total)".format(total_hours))
        self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8,
                            prop={'size': self.figure.get_figheight() * 1.4})

        self.subplot.autoscale_view()
        self.subplot.grid(True)

        # gives x-axis tick labels a slight rotation for better fitting
        self.figure.autofmt_xdate()


class EcrReasons(PlotPanel):
    def __init__(self, parent, start_date, end_date, **kwargs):
        self.parent = parent

        self.start_date = start_date
        self.end_date = end_date

        cursor = Database.connection.cursor()
        ecr_reasons = list(zip(*cursor.execute("SELECT reason FROM ecr_reason_choices where Production_Plant = \'{}\' ".format(Ecrs.Prod_Plant)).fetchall())[0])

        self.reason_data = []
        for reason_index, reason in enumerate(ecr_reasons):
            self.reason_data.append((
                reason,
                cursor.execute(
                    'SELECT when_closed FROM ecrs WHERE reason=\'{}\' AND when_closed>\'{}\' AND when_closed<\'{}\' AND Production_Plant = \'{}\' ORDER BY when_closed ASC'.format(
                        reason, self.start_date, '{} 23:59:59'.format(self.end_date),Ecrs.Prod_Plant)).fetchall())
            )

        # print self.employee_data[-1]

        # initiate plotter
        PlotPanel.__init__(self, parent, **kwargs)
        self.SetColor((255, 255, 255))

    def draw(self):
        """Draw data."""
        if not hasattr(self, 'subplot'):
            self.subplot = self.figure.add_subplot(111)

        total_reasons = 0

        for index, data in enumerate(self.reason_data):
            x = [dt.datetime.strptime(str(p[0]), "%Y-%m-%d %H:%M:%S") for p in data[1]]

            # skip this reason if not ecrs closed with it
            if not x:
                continue

            reason_sum = 0
            y = []
            for p in data[1]:
                reason_sum += 1
                y.append(reason_sum)

            total_reasons += reason_sum

            reason = data[0]

            # self.subplot.plot_date(x, y, '-', label='{}: {}'.format(reason, reason_sum))
            self.subplot.plot_date(x, y, '-', label='{}: {}'.format(reason, reason_sum))

            annotation_x_index = int((len(x) - index) * .92)
            annotation_y_index = int((len(y) - index) * .92)

            if annotation_x_index < 0: annotation_x_index = len(x) - 1
            if annotation_y_index < 0: annotation_y_index = len(y) - 1

            self.subplot.annotate(reason.split(',')[0], xy=(x[annotation_x_index], y[annotation_y_index]),
                                  xytext=(-15, 15),
                                  textcoords='offset points', ha='center', va='bottom',
                                  bbox=dict(boxstyle='round,pad=0.2', fc='yellow', alpha=0.5),
                                  arrowprops=dict(arrowstyle='->', connectionstyle='arc3', color='black'))

        ticks = get_ticks(self.start_date, self.end_date)
        self.subplot.xaxis.set_major_locator(ticks[0])
        self.subplot.xaxis.set_major_formatter(ticks[1])

        if ticks[2] != None:
            self.subplot.xaxis.set_minor_locator(ticks[2])
            self.subplot.xaxis.set_minor_formatter(ticks[3])

        self.subplot.xaxis.set_label_text('Time Period')
        self.subplot.yaxis.set_label_text('ECRs Closed')
        self.subplot.set_title(r"ECRs Closed by Reason   ({} total)".format(total_reasons))

        # self.subplot.legend(bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0.)
        # self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, borderaxespad=.8, prop={'size':10})
        self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8,
                            prop={'size': self.figure.get_figheight() * 1.4})

        self.subplot.autoscale_view()
        self.subplot.grid(True)

        # gives x-axis tick labels a slight rotation for better fitting
        self.figure.autofmt_xdate()


class EcrDocuments(PlotPanel):
    def __init__(self, parent, start_date, end_date, **kwargs):
        self.parent = parent

        self.start_date = start_date
        self.end_date = end_date

        cursor = Database.connection.cursor()

        ecr_documents = list(zip(*cursor.execute('SELECT document FROM ecr_document_choices where Production_Plant = \'{}\''.format(Ecrs.Prod_Plant)).fetchall())[0])

        self.document_data = []
        for document_index, document in enumerate(ecr_documents):
            self.document_data.append((
                document,
                cursor.execute(
                    'SELECT when_closed FROM ecrs WHERE document=\'{}\' AND when_closed>\'{}\' AND when_closed<\'{}\' AND Production_Plant = \'{}\' ORDER BY when_closed ASC'.format(
                        document, self.start_date, '{} 23:59:59'.format(self.end_date), Ecrs.Prod_Plant)).fetchall())
                # cursor.execute('SELECT when_closed FROM ecrs WHERE document=\'{}\' AND when_closed>\'{}\' AND when_closed<\'{}\' AND (resolution LIKE \'%picklist%\' OR resolution LIKE \'%pick list%\' OR request LIKE \'%picklist%\' OR request LIKE \'%pick list%\') ORDER BY when_closed ASC'.format(document, self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
            )

        # print self.employee_data[-1]

        # initiate plotter
        PlotPanel.__init__(self, parent, **kwargs)
        self.SetColor((255, 255, 255))

    def draw(self):
        """Draw data."""
        if not hasattr(self, 'subplot'):
            self.subplot = self.figure.add_subplot(111)

        total_documents = 0

        for index, data in enumerate(self.document_data):
            x = [dt.datetime.strptime(str(p[0]), "%Y-%m-%d %H:%M:%S") for p in data[1]]
            # print x

            # skip this document if not ecrs closed with it
            if not x:
                continue

            document_sum = 0
            y = []
            for p in data[1]:
                document_sum += 1
                y.append(document_sum)

            # y = derivative(x, y)

            total_documents += document_sum

            document = data[0]

            # self.subplot.plot_date(x, y, '-', label='{}: {}'.format(document, document_sum))
            self.subplot.plot_date(x, y, '-', label='{}: {}'.format(document, document_sum))

            '''
            try:
                x, y = derivative(x, y)
                self.subplot.plot_date(x, y, '*', label='{}: {}'.format(document, document_sum))
            except Exception as e:
                print e, traceback.format_exc()
            '''

            annotation_x_index = int((len(x) - index) * .92)
            annotation_y_index = int((len(y) - index) * .92)

            if annotation_x_index < 0: annotation_x_index = len(x) - 1
            if annotation_y_index < 0: annotation_y_index = len(y) - 1

            self.subplot.annotate(document, xy=(x[annotation_x_index], y[annotation_y_index]), xytext=(-15, 15),
                                  textcoords='offset points', ha='center', va='bottom',
                                  bbox=dict(boxstyle='round,pad=0.2', fc='yellow', alpha=0.5),
                                  arrowprops=dict(arrowstyle='->', connectionstyle='arc3', color='black'))

        ticks = get_ticks(self.start_date, self.end_date)
        self.subplot.xaxis.set_major_locator(ticks[0])
        self.subplot.xaxis.set_major_formatter(ticks[1])

        if ticks[2] != None:
            self.subplot.xaxis.set_minor_locator(ticks[2])
            self.subplot.xaxis.set_minor_formatter(ticks[3])

        self.subplot.xaxis.set_label_text('Time Period')
        self.subplot.yaxis.set_label_text('ECRs Closed')
        self.subplot.set_title(r"ECRs Closed by document   ({} total)".format(total_documents))

        # self.subplot.legend(bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0.)
        # self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, borderaxespad=.8, prop={'size':10})
        self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8,
                            prop={'size': self.figure.get_figheight() * 1.4})

        self.subplot.autoscale_view()
        self.subplot.grid(True)

        # gives x-axis tick labels a slight rotation for better fitting
        self.figure.autofmt_xdate()


class EcrsClosedOnTime(PlotPanel):
    def __init__(self, parent, start_date, end_date, **kwargs):
        self.parent = parent

        self.start_date = start_date
        self.end_date = end_date

        cursor = Database.connection.cursor()

        # self.ecr_data = cursor.execute("SELECT when_needed, when_closed FROM ecrs WHERE when_needed>\'{}\' AND when_needed<\'{}\' AND when_closed IS NOT NULL ORDER BY when_closed ASC".format(self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall()
        self.ecr_data = cursor.execute(
            "SELECT when_needed, when_closed FROM ecrs WHERE when_closed>\'{}\' AND when_closed<\'{}\' AND Production_Plant = \'{}\' ORDER BY when_closed ASC".format(
                self.start_date, '{} 23:59:59'.format(self.end_date), Ecrs.Prod_Plant)).fetchall()
        # self.ecr_data = cursor.execute("SELECT when_needed, when_closed FROM ecrs WHERE reason <> 'BOM Reconciliation' AND when_closed>\'{}\' AND when_closed<\'{}\' ORDER BY when_closed ASC".format(self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall()
        #print self.ecr_data

        self.lowest_start_date = end_date
        for ecr in self.ecr_data:
            when_closed_dt = dt.datetime.strptime(str(ecr[1]), "%Y-%m-%d %H:%M:%S").date()
            if when_closed_dt < self.lowest_start_date:
                self.lowest_start_date = when_closed_dt

        print 'got this far'

        # initiate plotter
        PlotPanel.__init__(self, parent, **kwargs)
        self.SetColor((255, 255, 255))

    def draw(self):
        """Draw data."""
        if not hasattr(self, 'subplot'):
            self.subplot = self.figure.add_subplot(111)

        # when closed
        x = [dt.datetime.strptime(str(p[1]), "%Y-%m-%d %H:%M:%S") for p in self.ecr_data]

        when_needed_list = [dt.datetime.strptime(str(p[0]), "%Y-%m-%d %H:%M:%S") for p in self.ecr_data]

        on_time_sum = 0.0
        total_sum = 0.00001

        y = []
        for index, when_closed in enumerate(x):
            total_sum += 1

            if when_closed <= when_needed_list[index]:
                on_time_sum += 1

            y.append((on_time_sum / total_sum) * 100)

        self.subplot.plot_date(x, y, '-')

        # show goal line (93%)
        self.subplot.plot_date([self.lowest_start_date, self.end_date + relativedelta(hours=17)],
                               [.93 * 100, .93 * 100], '-')

        # plot some points at top and bottom so whole range is shown...
        self.subplot.plot_date([self.lowest_start_date, self.end_date], [0, 0], '-')
        self.subplot.plot_date([self.lowest_start_date, self.end_date], [100, 100], '-')

        ticks = get_ticks(self.start_date, self.end_date)
        self.subplot.xaxis.set_major_locator(ticks[0])
        self.subplot.xaxis.set_major_formatter(ticks[1])

        if ticks[2] != None:
            self.subplot.xaxis.set_minor_locator(ticks[2])
            self.subplot.xaxis.set_minor_formatter(ticks[3])

        # self.subplot.axis('tight')
        # self.subplot.yaxis.set_major_locator(FormatStrFormatter('%2.1f'))

        # label_text   = [r"$%i \cdot 10^4$" % int(loc/10**4) for loc in range(1, 10)]
        # label_text   = ['{0:.0f}%'.format(loc) for loc in range(0, 4)]
        # self.subplot.set_yticklabels(label_text)

        self.subplot.xaxis.set_label_text('Time Period')
        self.subplot.yaxis.set_label_text('Percent Closed On Time')
        self.subplot.set_title(
            r"ECRs Closed On Time   ({:.2f}% of {} ECRs)".format(float(on_time_sum / total_sum) * 100, int(total_sum)))

        self.subplot.autoscale_view()
        self.subplot.grid(True)

        # gives x-axis tick labels a slight rotation for better fitting
        self.figure.autofmt_xdate()


class EcrsByProductFamily(PlotPanel):
	def __init__( self, parent, start_date, end_date, **kwargs ):
		self.parent = parent

		self.start_date = start_date
		self.end_date = end_date

		cursor = Database.connection.cursor()

		try:
			families = list(zip(*cursor.execute('SELECT {}.family FROM ecrs INNER JOIN {} ON ecrs.item = {}.item WHERE ecrs.when_closed>\'{}\' AND ecrs.when_closed<\'{}\' AND ecrs.Production_Plant = \'{}\' ORDER BY {}.family ASC'.format(Ecrs.table_used, Ecrs.table_used, Ecrs.table_used, self.start_date, '{} 23:59:59'.format(self.end_date), Ecrs.Prod_Plant, Ecrs.table_used)).fetchall())[0])
		except:
			families = []
		#remove duplicate families from list
		families = list(set(families))
		families.sort()

		self.family_data = []
		for family_index, family in enumerate(families):
			self.family_data.append( (
					family, 
					#cursor.execute('SELECT when_closed FROM ecrs WHERE document=\'{}\' AND when_closed>\'{}\' AND when_closed<\'{}\' ORDER BY when_closed ASC'.format(document, self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
					#cursor.execute('SELECT when_closed FROM ecrs WHERE document=\'{}\' AND when_closed>\'{}\' AND when_closed<\'{}\' AND (resolution LIKE \'%picklist%\' OR resolution LIKE \'%pick list%\' OR request LIKE \'%picklist%\' OR request LIKE \'%pick list%\') ORDER BY when_closed ASC'.format(document, self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
					cursor.execute('SELECT ecrs.when_closed FROM ecrs INNER JOIN {} ON ecrs.item = {}.item WHERE {}.family=\'{}\' AND ecrs.when_closed>\'{}\' AND ecrs.when_closed<\'{}\' ORDER BY ecrs.when_closed ASC'.format(Ecrs.table_used, Ecrs.table_used, Ecrs.table_used, family, self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
				)

		#initiate plotter
		PlotPanel.__init__( self, parent, **kwargs )
		self.SetColor( (255,255,255) )

	def draw(self):
		"""Draw data."""
		if not hasattr(self, 'subplot'):
			self.subplot = self.figure.add_subplot(111)

		total_ecrs = 0

		for index, data in enumerate(self.family_data):
			x = [dt.datetime.strptime(str(p[0]), "%Y-%m-%d %H:%M:%S") for p in data[1]]
			#print x

			#skip this document if not ecrs closed with it
			if not x:
				continue
			
			family_sum = 0
			y = []
			for p in data[1]:
				family_sum += 1
				y.append(family_sum)
				
			total_ecrs += family_sum

			family = data[0]

			#self.subplot.plot_date(x, y, '-', label='{}: {}'.format(document, document_sum))
			self.subplot.plot_date(x, y, '-', label='{}: {}'.format(family, family_sum))
			
			annotation_x_index = int((len(x)-index)*.92)
			annotation_y_index = int((len(y)-index)*.92)

			if annotation_x_index < 0: annotation_x_index = len(x)-1
			if annotation_y_index < 0: annotation_y_index = len(y)-1

			self.subplot.annotate(family, xy=(x[annotation_x_index], y[annotation_y_index]), xytext=(-15,15), 
						textcoords='offset points', ha='center', va='bottom',
						bbox=dict(boxstyle='round,pad=0.2', fc='yellow', alpha=0.5),
						arrowprops=dict(arrowstyle='->', connectionstyle='arc3', color='black'))


		ticks = get_ticks(self.start_date, self.end_date)
		self.subplot.xaxis.set_major_locator(ticks[0])
		self.subplot.xaxis.set_major_formatter(ticks[1])
		
		if ticks[2] != None:
			self.subplot.xaxis.set_minor_locator(ticks[2])
			self.subplot.xaxis.set_minor_formatter(ticks[3])
		
		self.subplot.xaxis.set_label_text('Time Period')
		self.subplot.yaxis.set_label_text('ECRs Closed')
		self.subplot.set_title(r"ECRs by Product Family   ({} total)".format(total_ecrs))
		
		#self.subplot.legend(bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0.)
		#self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, borderaxespad=.8, prop={'size':10})
		#self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8, prop={'size':self.figure.get_figheight() * 1.4})
		self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8, prop={'size':self.figure.get_figheight() * 0.8})

		self.subplot.autoscale_view()
		self.subplot.grid(True)

		#gives x-axis tick labels a slight rotation for better fitting
		self.figure.autofmt_xdate()


class EcrsByCustomer(PlotPanel):
    def __init__(self, parent, start_date, end_date, **kwargs):
        self.parent = parent

        self.start_date = start_date
        self.end_date = end_date

        cursor = Database.connection.cursor()

        try:
            customers = list(zip(*cursor.execute(
                'SELECT {}.customer FROM ecrs INNER JOIN {} ON ecrs.item = {}.item WHERE ecrs.when_closed>\'{}\' AND ecrs.when_closed<\'{}\' AND ecrs.Production_Plant = \'{}\' ORDER BY {}.customer ASC'.format(
                    Ecrs.table_used, Ecrs.table_used, Ecrs.table_used, self.start_date,
                    '{} 23:59:59'.format(self.end_date), Ecrs.Prod_Plant, Ecrs.table_used)).fetchall())[0])
        except:
            customers = []
        # remove duplicate customers from list
        customers = list(set(customers))
        customers.sort()

        self.customer_data = []
        for customer_index, customer in enumerate(customers):
            self.customer_data.append((
                customer,
                # cursor.execute('SELECT when_closed FROM ecrs WHERE document=\'{}\' AND when_closed>\'{}\' AND when_closed<\'{}\' ORDER BY when_closed ASC'.format(document, self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
                # cursor.execute('SELECT when_closed FROM ecrs WHERE document=\'{}\' AND when_closed>\'{}\' AND when_closed<\'{}\' AND (resolution LIKE \'%picklist%\' OR resolution LIKE \'%pick list%\' OR request LIKE \'%picklist%\' OR request LIKE \'%pick list%\') ORDER BY when_closed ASC'.format(document, self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
                cursor.execute(
                    'SELECT ecrs.when_closed FROM ecrs INNER JOIN {} ON ecrs.item = {}.item WHERE {}.customer=\'{}\' AND ecrs.when_closed>\'{}\' AND ecrs.when_closed<\'{}\' ORDER BY ecrs.when_closed ASC'.format(
                        Ecrs.table_used, Ecrs.table_used, Ecrs.table_used, customer, self.start_date,
                        '{} 23:59:59'.format(self.end_date))).fetchall())
            )

        # initiate plotter
        PlotPanel.__init__(self, parent, **kwargs)
        self.SetColor((255, 255, 255))

    def draw(self):
        """Draw data."""
        if not hasattr(self, 'subplot'):
            self.subplot = self.figure.add_subplot(111)

        total_ecrs = 0

        for index, data in enumerate(self.customer_data):
            x = [dt.datetime.strptime(str(p[0]), "%Y-%m-%d %H:%M:%S") for p in data[1]]
            # print x

            # skip this document if not ecrs closed with it
            if not x:
                continue

            customer_sum = 0
            y = []
            for p in data[1]:
                customer_sum += 1
                y.append(customer_sum)

            total_ecrs += customer_sum

            customer = data[0]

            # self.subplot.plot_date(x, y, '-', label='{}: {}'.format(document, document_sum))
            self.subplot.plot_date(x, y, '-', label='{}: {}'.format(customer, customer_sum))

            annotation_x_index = int((len(x) - index) * .92)
            annotation_y_index = int((len(y) - index) * .92)

            if annotation_x_index < 0: annotation_x_index = len(x) - 1
            if annotation_y_index < 0: annotation_y_index = len(y) - 1

            self.subplot.annotate(customer, xy=(x[annotation_x_index], y[annotation_y_index]), xytext=(-15, 15),
                                  textcoords='offset points', ha='center', va='bottom',
                                  bbox=dict(boxstyle='round,pad=0.2', fc='yellow', alpha=0.5),
                                  arrowprops=dict(arrowstyle='->', connectionstyle='arc3', color='black'))

        ticks = get_ticks(self.start_date, self.end_date)
        self.subplot.xaxis.set_major_locator(ticks[0])
        self.subplot.xaxis.set_major_formatter(ticks[1])

        if ticks[2] != None:
            self.subplot.xaxis.set_minor_locator(ticks[2])
            self.subplot.xaxis.set_minor_formatter(ticks[3])

        self.subplot.xaxis.set_label_text('Time Period')
        self.subplot.yaxis.set_label_text('ECRs Closed')
        self.subplot.set_title(r"ECRs by Customer   ({} total)".format(total_ecrs))

        # self.subplot.legend(bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0.)
        # self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, borderaxespad=.8, prop={'size':10})
        # self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8, prop={'size':self.figure.get_figheight() * 1.4})
        self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8,
                            prop={'size': self.figure.get_figheight() * 1})

        self.subplot.autoscale_view()
        self.subplot.grid(True)

        # gives x-axis tick labels a slight rotation for better fitting
        self.figure.autofmt_xdate()


class HoursLoggedByProductFamily(PlotPanel):
    def __init__(self, parent, start_date, end_date, **kwargs):
        self.parent = parent

        self.start_date = start_date
        self.end_date = end_date

        cursor = Database.connection.cursor()

        families = list(zip(*cursor.execute(
            'SELECT {}.family FROM time_logs INNER JOIN {} ON time_logs.item = {}.item WHERE time_logs.when_logged>\'{}\' AND time_logs.when_logged<\'{}\' ORDER BY {}.family ASC'.format(
                Ecrs.table_used, Ecrs.table_used, Ecrs.table_used, self.start_date, '{} 23:59:59'.format(self.end_date),
                Ecrs.table_used)).fetchall())[0])
        # remove duplicate families from list
        families = list(set(families))
        families.sort()
        # print families

        self.family_data = []
        for family_index, family in enumerate(families):
            self.family_data.append((
                family,
                # cursor.execute('SELECT when_closed FROM ecrs WHERE document=\'{}\' AND when_closed>\'{}\' AND when_closed<\'{}\' ORDER BY when_closed ASC'.format(document, self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
                # cursor.execute('SELECT when_closed FROM ecrs WHERE document=\'{}\' AND when_closed>\'{}\' AND when_closed<\'{}\' AND (resolution LIKE \'%picklist%\' OR resolution LIKE \'%pick list%\' OR request LIKE \'%picklist%\' OR request LIKE \'%pick list%\') ORDER BY when_closed ASC'.format(document, self.start_date, '{} 23:59:59'.format(self.end_date))).fetchall())
                cursor.execute(
                    'SELECT time_logs.when_logged, time_logs.hours FROM time_logs INNER JOIN {} ON time_logs.item = {}.item WHERE {}.family=\'{}\' AND time_logs.when_logged>\'{}\' AND time_logs.when_logged<\'{}\' ORDER BY time_logs.when_logged ASC'.format(
                        Ecrs.table_used, Ecrs.table_used, Ecrs.table_used, family, self.start_date,
                        '{} 23:59:59'.format(self.end_date))).fetchall())
            )

        # initiate plotter
        PlotPanel.__init__(self, parent, **kwargs)
        self.SetColor((255, 255, 255))

    def draw(self):
        """Draw data."""
        if not hasattr(self, 'subplot'):
            self.subplot = self.figure.add_subplot(111)

        total_hours = 0

        for index, data in enumerate(self.family_data):
            x = [dt.datetime.strptime(str(p[0]), "%Y-%m-%d %H:%M:%S") for p in data[1]]
            # print x

            # skip this document if not ecrs closed with it
            if not x:
                continue

            hours_sum = 0
            y = []
            for p in data[1]:
                hours_sum += p[1]
                y.append(hours_sum)

            total_hours += hours_sum

            family = data[0]

            # self.subplot.plot_date(x, y, '-', label='{}: {}'.format(document, document_sum))
            self.subplot.plot_date(x, y, '-', label='{}: {:.1f}'.format(family, hours_sum))

            annotation_x_index = int((len(x) - index) * .92)
            annotation_y_index = int((len(y) - index) * .92)

            if annotation_x_index < 0: annotation_x_index = len(x) - 1
            if annotation_y_index < 0: annotation_y_index = len(y) - 1

            self.subplot.annotate(family, xy=(x[annotation_x_index], y[annotation_y_index]), xytext=(-15, 15),
                                  textcoords='offset points', ha='center', va='bottom',
                                  bbox=dict(boxstyle='round,pad=0.2', fc='yellow', alpha=0.5),
                                  arrowprops=dict(arrowstyle='->', connectionstyle='arc3', color='black'))

        ticks = get_ticks(self.start_date, self.end_date)
        self.subplot.xaxis.set_major_locator(ticks[0])
        self.subplot.xaxis.set_major_formatter(ticks[1])

        if ticks[2] != None:
            self.subplot.xaxis.set_minor_locator(ticks[2])
            self.subplot.xaxis.set_minor_formatter(ticks[3])

        self.subplot.xaxis.set_label_text('Time Period')
        self.subplot.yaxis.set_label_text('Hours Logged')
        self.subplot.set_title(r"Hours Logged by Product Family   ({:.1f} total)".format(total_hours))

        # self.subplot.legend(bbox_to_anchor=(1.05, 1), loc=2, borderaxespad=0.)
        # self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, borderaxespad=.8, prop={'size':10})
        # self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8, prop={'size':self.figure.get_figheight() * 1.4})
        self.subplot.legend(bbox_to_anchor=(0, 1), loc=2, fancybox=True, borderaxespad=.8,
                            prop={'size': self.figure.get_figheight() * .9})

        self.subplot.autoscale_view()
        self.subplot.grid(True)

        # gives x-axis tick labels a slight rotation for better fitting
        self.figure.autofmt_xdate()


def on_change_date(event):
    print 'hit on change date!'

    temp_date = ctrl(General.app.main_frame, 'date:report_start').GetValue()
    General.app.start_date = dt.date(temp_date.GetYear(), temp_date.GetMonth() + 1, temp_date.GetDay())
    temp_date = ctrl(General.app.main_frame, 'date:report_end').GetValue()
    General.app.end_date = dt.date(temp_date.GetYear(), temp_date.GetMonth() + 1, temp_date.GetDay())

    # set radio value to "custom range" lolololol
    ctrl(General.app.main_frame, 'radio:custom_range').SetValue(1)

    refresh_plots()


def on_click_radio(event):
    label = event.GetEventObject().GetLabel()
    today = dt.date.today()

    ###
    ###today = dt.date(2012, 11, 2)

    if label == 'Today':
        General.app.start_date = dt.date(today.year, today.month, today.day)
        General.app.end_date = dt.date(today.year, today.month, today.day)
    elif label == 'This Week':
        # General.app.start_date = dt.date(today.year, today.month, get_week_days())
        General.app.start_date = today - dt.timedelta(days=today.weekday()) + dt.timedelta(days=-1, weeks=0)
        General.app.end_date = dt.date(today.year, today.month, today.day)
    elif label == 'This Month':
        General.app.start_date = dt.date(today.year, today.month, 1)
        General.app.end_date = dt.date(today.year, today.month, today.day)
    elif label == 'This Quarter':
        quarter = int(((today.month - 1) / 12. * 4) + 1)
        month_start_of_quarter = ((quarter - 1) * 3 + 1)
        General.app.start_date = dt.date(today.year, month_start_of_quarter, 1)
        General.app.end_date = dt.date(today.year, today.month, today.day)
    elif label == 'This Year':
        General.app.start_date = dt.date(today.year, 1, 1)
        General.app.end_date = dt.date(today.year, today.month, today.day)

    elif label == 'Yesterday':
        General.app.start_date = today - dt.timedelta(days=1)
        General.app.end_date = today - dt.timedelta(days=1)
    elif label == 'Last Week':
        General.app.start_date = today - dt.timedelta(days=today.weekday()) + dt.timedelta(days=-1, weeks=-1)
        General.app.end_date = today - dt.timedelta(days=today.weekday()) + dt.timedelta(days=-2, weeks=0)

    elif label == 'Last Month':
        d = today - relativedelta(months=1)
        General.app.start_date = dt.date(d.year, d.month, 1)
        General.app.end_date = dt.date(today.year, today.month, 1) - relativedelta(days=1)

    elif label == 'Last Quarter':
        quarter = int(((today.month - 1) / 12. * 4) + 1)
        month_start_of_quarter = ((quarter - 1) * 3 + 1)
        General.app.start_date = dt.date(today.year, month_start_of_quarter, 1) + relativedelta(months=-3)
        General.app.end_date = General.app.start_date + relativedelta(months=3, days=-1)
    elif label == 'Last Year':
        General.app.start_date = dt.date(today.year - 1, 1, 1)
        General.app.end_date = dt.date(today.year, 1, 1) + relativedelta(days=-1)


    elif label == 'All time':
        General.app.start_date = dt.date(1900, 1, 1)
        General.app.end_date = dt.date(today.year, today.month, today.day)
    elif label == 'Custom range':
        temp_date = ctrl(General.app.main_frame, 'date:report_start').GetValue()
        General.app.start_date = dt.date(temp_date.GetYear(), temp_date.GetMonth() + 1, temp_date.GetDay())
        temp_date = ctrl(General.app.main_frame, 'date:report_end').GetValue()
        General.app.end_date = dt.date(temp_date.GetYear(), temp_date.GetMonth() + 1, temp_date.GetDay())
    else:
        General.app.start_date = dt.date(1900, 1, 1)
        General.app.end_date = dt.date(today.year, today.month, today.day)

    # set custum range fields to display what range selected
    date_wx_format = wx.DateTimeFromDMY(General.app.start_date.day, General.app.start_date.month - 1,
                                        General.app.start_date.year)
    ctrl(General.app.main_frame, 'date:report_start').SetValue(date_wx_format)

    date_wx_format = wx.DateTimeFromDMY(General.app.end_date.day, General.app.end_date.month - 1,
                                        General.app.end_date.year)
    ctrl(General.app.main_frame, 'date:report_end').SetValue(date_wx_format)

    refresh_plots()


def refresh_plots():
    report_frame = General.app.res.LoadFrame(None, 'frame:generating_report')
    ctrl(report_frame, 'text:report').SetValue(ctrl(report_frame, 'text:report').GetValue() + "querying database...\n")
    report_frame.Show()

    try:
        if General.app.plotting_panel:
            General.app.plotting_panel.Destroy()

        if General.app.report_name == "DE's Logged Hours":
            General.app.plotting_panel = DesignEngineeringHours(ctrl(General.app.main_frame, 'panel:plot'),
                                                                General.app.start_date, General.app.end_date)
        if General.app.report_name == "ECR Reasons":
            General.app.plotting_panel = EcrReasons(ctrl(General.app.main_frame, 'panel:plot'), General.app.start_date,
                                                    General.app.end_date)
        if General.app.report_name == "ECR Documents":
            General.app.plotting_panel = EcrDocuments(ctrl(General.app.main_frame, 'panel:plot'),
                                                      General.app.start_date, General.app.end_date)
        if General.app.report_name == "ECRs Closed On Time":
            General.app.plotting_panel = EcrsClosedOnTime(ctrl(General.app.main_frame, 'panel:plot'),
                                                          General.app.start_date, General.app.end_date)
        if General.app.report_name == "ECRs by Product Family":
            General.app.plotting_panel = EcrsByProductFamily(ctrl(General.app.main_frame, 'panel:plot'),
                                                             General.app.start_date, General.app.end_date)
        if General.app.report_name == "ECRs by Customer":
            General.app.plotting_panel = EcrsByCustomer(ctrl(General.app.main_frame, 'panel:plot'),
                                                        General.app.start_date, General.app.end_date)
        if General.app.report_name == "Hours Logged by Product Family":
            General.app.plotting_panel = HoursLoggedByProductFamily(ctrl(General.app.main_frame, 'panel:plot'),
                                                                    General.app.start_date, General.app.end_date)

    except Exception as e:
        print e
        ctrl(report_frame, 'text:report').SetValue(ctrl(report_frame, 'text:report').GetValue() + str(e))

    report_frame.Destroy()


def on_select_report(event):
    General.app.report_name = event.GetString()
    refresh_plots()


def on_paint_window(event):
    notebook = ctrl(General.app.main_frame, 'notebook:main')

    if notebook.GetPageText(notebook.GetSelection()).strip() == 'Reports':
        print 'attempting to resize plot'
        try:
            General.app.plotting_panel._SetSize()
        except:
            pass

    event.Skip()


def derivative(x_date_values, y_values):
    # intertolate for even distribution
    x_values = []

    for x_date_value in x_date_values:
        x_values.append(datetime_to_timestamp(x_date_value))

    '''
    print x_date_values
    print '@'
    print x_values
    
    print x_values[0]
    print x_values[-1]
    print x_values[-1]-x_values[0]
    print (x_values[-1]-x_values[0])/10.
    '''

    # distributed_x_values = linspace(x_values[0], (x_values[-1]-x_values[0])/10., x_values[-1])
    # distributed_x_values = linspace(0, (x_values[-1]-x_values[0])/10., x_values[-1]-x_values[0])
    distributed_x_values = linspace(0, 50, x_values[-1] - x_values[0])

    ###interp_func = interp1d(x_values, y_values)

    distributed_y_values = []

    '''
    for x in distributed_x_values:
        distributed_y_values.append(interp_func(x))
    '''

    distributed_y_values = interp(distributed_x_values, x_values, y_values)

    distributed_x_date_values = []

    for x in distributed_x_values:
        distributed_x_date_values.append(dt.datetime.utcfromtimestamp(x + x_values[0]))

    print '...'
    for poo in x_date_values:
        print poo
    print '~~~~~~~~~~'
    for poo in distributed_x_date_values:
        print poo
    print '......'

    return (distributed_x_date_values, distributed_y_values)

    print '_____'
    print start_date
    print datetime_to_timestamp(start_date)
    print dt.datetime.utcfromtimestamp(datetime_to_timestamp(start_date))
    print '^^^^^'

    y_derivatives = []

    # x_range = time.mktime((x_values[-1] - x_values[0]).timetuple())
    # print x_range

    x_start = time.mktime(x_values[0].timetuple())
    x_end = time.mktime(x_values[-1].timetuple())

    x_range = (x_end - x_start)
    print x_range

    for index in range(0, len(y_values) - 1):
        # y_derivatives.append( (y_values[index+1] - y_values[index]) / (x_values[index+1] - x_values[index]) )

        x2 = time.mktime(x_values[index + 1].timetuple()) / x_range
        x1 = time.mktime(x_values[index].timetuple()) / x_range

        y_derivatives.append((y_values[index + 1] - y_values[index]) / (x2 - x1))

    y_derivatives.insert(0, y_derivatives[0])

    return y_derivatives


def get_ticks(start, end):
    minor_loc = None
    minor_fmt = None

    delta = end - start

    '''
    if delta <= dt.timedelta(minutes=10):
        print 'minutes=10'
        major_loc = mdates.MinuteLocator()
        major_fmt = mdates.DateFormatter('%I:%M %p')
    elif delta <= dt.timedelta(minutes=30):
        major_loc = mdates.MinuteLocator(byminute=range(0,60,5))
        major_fmt = mdates.DateFormatter('%I:%M %p')
    elif delta <= dt.timedelta(hours=1):
        major_loc = mdates.MinuteLocator(byminute=range(0,60,15))
        major_fmt = mdates.DateFormatter('%I:%M %p')
    '''
    if delta <= dt.timedelta(hours=6):
        major_loc = mdates.HourLocator()
        major_fmt = mdates.DateFormatter('%I:%M %p')
    elif delta <= dt.timedelta(days=1):
        major_loc = mdates.HourLocator(byhour=range(0, 24, 3))
        major_fmt = mdates.DateFormatter('%I:%M %p')
    elif delta <= dt.timedelta(days=3):
        # major_loc = mdates.HourLocator(byhour=range(0,24,6))
        # major_fmt = mdates.DateFormatter('%I:%M %p')
        major_loc = mdates.DayLocator()
        major_fmt = mdates.DateFormatter('%b %d')
    elif delta <= dt.timedelta(weeks=2):
        major_loc = mdates.DayLocator()
        major_fmt = mdates.DateFormatter('%b %d')
    elif delta <= dt.timedelta(weeks=5):
        # major_loc = mdates.WeekdayLocator()
        # major_fmt = mdates.DateFormatter('%b %d')
        major_loc = mdates.DayLocator(interval=3)
        major_fmt = mdates.DateFormatter('%b %d')
    elif delta <= dt.timedelta(weeks=12):
        major_loc = mdates.WeekdayLocator()
        major_fmt = mdates.DateFormatter('%b %d')
    # minor_loc = mdates.DayLocator(interval=3)
    # minor_fmt = mdates.DateFormatter('%d')
    elif delta <= dt.timedelta(weeks=52):
        major_loc = mdates.MonthLocator()
        major_fmt = mdates.DateFormatter('%b')
    else:
        major_loc = mdates.MonthLocator(interval=2)
        major_fmt = mdates.DateFormatter('%b %Y')

    return major_loc, major_fmt, minor_loc, minor_fmt


def last_day_of_month(d):
    return (dt.date(d.year, d.month + 1, 1) - dt.timedelta(1)).day


'''
def plot_demo(event):
    today = dt.date.today()


    #start_date = dt.date( today.year, today.month-1, 1 )
    #end_date = dt.date( today.year, today.month, last_day_of_month(today) )

    start_date = dt.date( today.year-1, today.month-0, today.day )
    end_date = dt.date( today.year, today.month, today.day )

    fig = figure()
    ax = fig.add_subplot(111)

    cursor = Database.connection.cursor()
    
    design_engineers = list(zip(*cursor.execute('SELECT name FROM employees WHERE department = \'Design Engineering\' ORDER BY name ASC').fetchall())[0])
    
    for employee_index, employee in enumerate(design_engineers):
        points = cursor.execute('SELECT when_logged, hours FROM time_logs WHERE employee=\'{}\' AND when_logged>\'{}\' AND when_logged<\'{}\' ORDER BY when_logged ASC'.format(employee, start_date, end_date)).fetchall()

        x = [dt.datetime.strptime(p[0], "%Y-%m-%d %H:%M:%S") for p in points]
        
        #skip employee if they have no time logs
        if not x:
            continue
        
        hours_sum = 0
        y = []
        for p in points:
            hours_sum += p[1]
            y.append(hours_sum)

        ax.plot_date(x, y, '-')
        
        annotation_x_index = int((len(x)-employee_index)*.92)
        annotation_y_index = int((len(y)-employee_index)*.92)

        if annotation_x_index < 0: annotation_x_index = len(x)-1
        if annotation_y_index < 0: annotation_y_index = len(y)-1

        ax.annotate(employee.split(',')[0], xy=(x[annotation_x_index], y[annotation_y_index]), xytext=(-15,15), 
                    textcoords='offset points', ha='center', va='bottom',
                    bbox=dict(boxstyle='round,pad=0.2', fc='yellow', alpha=0.5),
                    arrowprops=dict(arrowstyle='->', connectionstyle='arc3', color='black'))


    ticks = get_ticks(start_date, end_date)
    ax.xaxis.set_major_locator(ticks[0])
    ax.xaxis.set_major_formatter(ticks[1])
    
    if ticks[2] != None:
        ax.xaxis.set_minor_locator(ticks[2])
        ax.xaxis.set_minor_formatter(ticks[3])
    
    ax.xaxis.set_label_text('Time Period')
    ax.yaxis.set_label_text('Hours Logged')
    ax.set_title(r"Design Engineers' Logged Order Hours")
    
    ax.autoscale_view()
    ax.grid(True)

    #gives x-axis tick labels a slight rotation for better fitting
    fig.autofmt_xdate()
    
    print dir(fig)
    
    #fig.plot()
    show()
    
    #import matplotlib.pyplot as plt
    #plt.savefig('common_labels.png', dpi=300)
'''


#def on_click_Advanced_Report(event):
 #   startdate = ctrl(self.main_frame, 'date:report_start')
  #  enddate =

def Advanced_Report(event):
    LoopStopDummy = False
    LoopStop = False

    ResultGrid = []

    Headers = ["Date ID", "Year", "Week", "Period", "Start Day", "End Day","Engg Error ECRs Closed On Time","Total Engg Error ECRs",
               "% of Total Engg Error ECRs Closed on Time","Total ECRs Closed On Time","Total ECRs","% of Total ECRs Closed on Time"]
    ColumnCount = len(Headers)

    Date_Start = ctrl(General.app.main_frame, 'date:report_start').GetValue()
    Date_End = ctrl(General.app.main_frame, 'date:report_end').GetValue()

    Date_Loop_Start = dt.datetime.strptime(str(Date_Start), "%m/%d/%y %H:%M:%S")
    Date_Main_End = dt.datetime.strptime(str(Date_End), "%m/%d/%y %H:%M:%S")
    print Date_Main_End

    if Date_Loop_Start > Date_Main_End:
        wx.MessageBox("Date Start cannot be greater than Date End", "Check Dates", wx.OK | wx.ICON_INFORMATION)
        return

    cursor = Database.connection.cursor()

    Date_NextLoop_Start = Date_Loop_Start

    while LoopStop == False:
        for i in range(0, 7):
            DateNext = Date_Loop_Start + dt.timedelta(i + 1)
            NextWeekDayNo = DateNext.weekday()
            if DateNext == Date_Main_End:
                #Date_Loop_End = DateNext
                LoopStopDummy = True
            else:
                if NextWeekDayNo == 5:
                    Date_Loop_End = DateNext
                    #print Date_Loop_End
                    #print Date_Loop_End.isocalendar()[1]
                if NextWeekDayNo == 6:
                    Date_NextLoop_Start = DateNext
                    #print Date_NextLoop_Start
        if LoopStopDummy == True:
            Date_Loop_End = Date_Main_End

        Year = Date_Loop_Start.year
        Period = Date_Loop_Start.month
        Week = str(Year % 2000) + "/" + str(Date_Loop_Start.isocalendar()[1])
        DateID = str(Year) + "-" + str(Period) + "-" + str(Week)

        ecr_data = cursor.execute("SELECT when_needed, when_closed FROM ecrs WHERE when_closed>\'{}\' AND when_closed<\'{}\' AND Production_Plant = \'{}\' "
            "ORDER BY when_closed ASC".format(str(Date_Loop_Start.strftime("%m/%d/%Y")), '{} 23:59:59'.format(str(Date_Loop_End.strftime("%m/%d/%Y"))),
                                              Ecrs.Prod_Plant)).fetchall()


        # when closed
        x = [dt.datetime.strptime(str(p[1]), "%Y-%m-%d %H:%M:%S") for p in ecr_data]

        when_needed_list = [dt.datetime.strptime(str(p[0]), "%Y-%m-%d %H:%M:%S") for p in ecr_data]

        on_time_sum = 0.0
        prct_on_time = 0

        for index, when_closed in enumerate(x):
            if when_closed <= when_needed_list[index]:
                on_time_sum += 1
        try:
            prct_on_time = (on_time_sum / len(ecr_data)) * 100
        except:
            prct_on_time = 0


        ecr_data_EE = cursor.execute("SELECT when_needed, when_closed FROM ecrs WHERE when_closed>\'{}\' AND when_closed<\'{}\' AND reason = \'Engineering Error\' "
                                     "AND Production_Plant = \'{}\' "
            "ORDER BY when_closed ASC".format(str(Date_Loop_Start.strftime("%m/%d/%Y")), '{} 23:59:59'.format(str(Date_Loop_End.strftime("%m/%d/%Y"))),
                                              Ecrs.Prod_Plant)).fetchall()

        x = [dt.datetime.strptime(str(p[1]), "%Y-%m-%d %H:%M:%S") for p in ecr_data_EE]

        when_needed_list = [dt.datetime.strptime(str(p[0]), "%Y-%m-%d %H:%M:%S") for p in ecr_data_EE]

        on_time_sum_EE = 0.0
        prct_on_time_EE = 0

        for index, when_closed in enumerate(x):
            if when_closed <= when_needed_list[index]:
                on_time_sum_EE += 1
        try:
            prct_on_time_EE = (on_time_sum_EE / len(ecr_data_EE)) * 100
        except:
            prct_on_time_EE = 0


        ResultGrid.append([str(DateID),str(Year),str(Week),str(Period), str(Date_Loop_Start),str(Date_Loop_End), str(on_time_sum_EE),str(len(ecr_data_EE)),str(prct_on_time_EE),
                           str(on_time_sum),str(len(ecr_data)),str(prct_on_time)])


        Date_Loop_Start = Date_NextLoop_Start
        LoopStop = LoopStopDummy

        print Date_Loop_End


    excel = Dispatch('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Add()

    wb.ActiveSheet.Cells(1, 1).Value = 'Transferring data to Excel...'
    wb.ActiveSheet.Columns(1).AutoFit()

    R = len(ResultGrid)
    C = ColumnCount

    # Write the header
    excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1), wb.ActiveSheet.Cells(1, C))
    excel_range.Value = Headers

    # Write the main results
    excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(2, 1), wb.ActiveSheet.Cells(R + 1, C))
    excel_range.Value = ResultGrid

    # Write the footers (totals)
    # excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(R+2, 1),wb.ActiveSheet.Cells(R+2,C))
    # excel_range.Value = self.Footers

    # Autofit the columns
    excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1), wb.ActiveSheet.Cells(1, C))
    excel_range.Font.Bold = True

    # Autofit the columns
    excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1), wb.ActiveSheet.Cells(R + 2, C))
    excel_range.Columns.AutoFit()





