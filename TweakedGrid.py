import wx

class TweakedGrid(wx.grid.Grid):
	""" A Copy&Paste enabled grid class"""
	#def __init__(self, parent, id, pos, size, style):
	#	wx.grid.Grid.__init__(self, parent, id, pos, size, style)
	def __init__(self, parent):
		wx.grid.Grid.__init__(self, parent, -1)
		wx.EVT_KEY_DOWN(self, self.OnKey)
		
		self.SetLabelFont(self.GetParent().GetFont())

	'''
	def selection(self):
		# Show cell selection
		# If selection is cell...
		if self.GetSelectedCells():
			print "Selected cells " + str(self.GetSelectedCells())
		# If selection is block...
		if self.GetSelectionBlockTopLeft():
			print "Selection block top left " + str(self.GetSelectionBlockTopLeft())
		if self.GetSelectionBlockBottomRight():
			print "Selection block bottom right " + str(self.GetSelectionBlockBottomRight())
		
		# If selection is col...
		if self.GetSelectedCols():
			print "Selected cols " + str(self.GetSelectedCols())
		
		# If selection is row...
		if self.GetSelectedRows():
			print "Selected rows " + str(self.GetSelectedRows())
	
	def currentcell(self):
		# Show cursor position
		row = self.GetGridCursorRow()
		col = self.GetGridCursorCol()
		cell = (row, col)
		print "Current cell " + str(cell)
	'''
	
	def OnKey(self, event):
		# If Ctrl+C is pressed...
		if event.ControlDown() and event.GetKeyCode() == 67:
			#print "Ctrl+C"
			##self.selection()
			# Call copy method
			self.copy()
			
		# If Ctrl+V is pressed...
		if event.ControlDown() and event.GetKeyCode() == 86:
			#print "Ctrl+V"
			##self.currentcell()
			# Call paste method
			self.paste()
			
		# If Supr is presed
		if event.GetKeyCode() == 127:
			#print "Supr"
			# Call delete method
			self.delete()
			
		# Skip other Key events
		if event.GetKeyCode():
			event.Skip()
			return

	def copy(self):
		try:
			# Number of rows and cols
			if not self.GetSelectionBlockTopLeft():
				rows = self.GetGridCursorRow()
				cols = self.GetGridCursorCol()
			else:
				rows = self.GetSelectionBlockBottomRight()[0][0] - self.GetSelectionBlockTopLeft()[0][0] + 1
				cols = self.GetSelectionBlockBottomRight()[0][1] - self.GetSelectionBlockTopLeft()[0][1] + 1
			
			# data variable contain text that must be set in the clipboard
			data = ''
			
			# For each cell in selected range append the cell value in the data variable
			# Tabs '\t' for cols and '\r' for rows
			if not self.GetSelectionBlockTopLeft():
				data = data + str(self.GetCellValue(self.GetGridCursorRow(), self.GetGridCursorCol()))
				
			else:
				for r in range(rows):
					for c in range(cols):
						data = data + str(self.GetCellValue(self.GetSelectionBlockTopLeft()[0][0] + r, self.GetSelectionBlockTopLeft()[0][1] + c))
						if c < cols - 1:
							data = data + '\t'
					data = data + '\n'
				
			# Create text data object
			clipboard = wx.TextDataObject()
			# Set data object value
			clipboard.SetText(data)
			# Put the data in the clipboard
			if wx.TheClipboard.Open():
				wx.TheClipboard.SetData(clipboard)
				wx.TheClipboard.Close()
			else:
				wx.MessageBox("Can't open the clipboard", "Error")
		except:
			print 'error: copy method'
				
	def paste(self):
		try:
			clipboard = wx.TextDataObject()
			if wx.TheClipboard.Open():
				wx.TheClipboard.GetData(clipboard)
				wx.TheClipboard.Close()
			else:
				wx.MessageBox("Can't open the clipboard", "Error")
			data = clipboard.GetText()
			table = []
			y = -1
			# Convert text in a array of lines
			for r in data.splitlines():
				y = y +1
				x = -1
				# Convert c in a array of text separated by tab
				for c in r.split('\t'):
					x = x +1
					self.SetCellValue(self.GetGridCursorRow() + y, self.GetGridCursorCol() + x, c)
		except:
			print 'error: paste method'
				
	def delete(self):
		try:
			# Number of rows and cols
			if not self.GetSelectionBlockTopLeft():
				rows = self.GetGridCursorRow()
				cols = self.GetGridCursorCol()
			else:
				rows = self.GetSelectionBlockBottomRight()[0][0] - self.GetSelectionBlockTopLeft()[0][0] + 1
				cols = self.GetSelectionBlockBottomRight()[0][1] - self.GetSelectionBlockTopLeft()[0][1] + 1
				
			# Clear cells contents
			if not self.GetSelectionBlockTopLeft():
				self.SetCellValue(self.GetGridCursorRow(), self.GetGridCursorCol(), '')
			else:
				for r in range(rows):
					for c in range(cols):
						self.SetCellValue(self.GetSelectionBlockTopLeft()[0][0] + r, self.GetSelectionBlockTopLeft()[0][1] + c, '')
		except:
			print 'error: delete method'
