import wx
import wx.grid as gridlib
import wx.lib.mixins.grid as gae

import os
import sys
import operator
import Utils
import Model
from FieldMap import standard_field_map
from Excel import GetExcelReader

class AutoEditGrid( gridlib.Grid, gae.GridAutoEditMixin ):
	def __init__( self, parent, id=wx.ID_ANY, style=0 ):
		gridlib.Grid.__init__( self, parent, id=id, style=style )
		gae.GridAutoEditMixin.__init__(self)
		
#--------------------------------------------------------------------------------
class StartList(wx.Panel):

	def __init__(self, parent):
		wx.Panel.__init__(self, parent)
		
		explanation = wx.StaticText( self, label=u'Click Header name to sort.  To delete, set Bib to blank.' )
		self.importFromExcel = wx.Button( self, label=u'Import from Excel' )
		self.importFromExcel.Bind( wx.EVT_BUTTON, self.onImportFromExcel )
		
		hs = wx.BoxSizer( wx.HORIZONTAL )
		hs.Add( explanation, flag=wx.ALL|wx.ALIGN_CENTRE_VERTICAL, border=4 )
		hs.Add( self.importFromExcel, flag=wx.ALL|wx.ALIGN_RIGHT, border=4 )
 
		self.fieldNames  = Model.RiderInfo.FieldNames
		self.headerNames = Model.RiderInfo.HeaderNames
		
		self.grid = AutoEditGrid( self, style = wx.BORDER_SUNKEN )
		self.grid.DisableDragRowSize()
		self.grid.SetRowLabelSize( 0 )
		self.grid.CreateGrid( 0, len(self.headerNames) )
		
		self.sortCol = 0
		self.grid.Bind( wx.grid.EVT_GRID_LABEL_LEFT_CLICK, self.onColumnHeader )

		# Set specialized editors for appropriate columns.
		for col in xrange(self.grid.GetNumberCols()):
			self.grid.SetColLabelValue( col, self.headerNames[col] )
			attr = gridlib.GridCellAttr()
			if col == 0:
				attr.SetRenderer( gridlib.GridCellNumberRenderer() )
			self.grid.SetColAttr( col, attr )
		
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add( hs, 0, flag=wx.ALL, border = 4 )
		sizer.Add(self.grid, 1, flag=wx.EXPAND|wx.ALL, border = 6)
		self.SetSizer(sizer)
		
	def getGrid( self ):
		return self.grid
		
	def updateGrid( self ):
		race = Model.race
		riderInfoList = sorted( (race.riderInfo.itervalues()) if race else [], key=operator.attrgetter(self.fieldNames[self.sortCol], 'bib') )
		self.grid.BeginBatch()
		Utils.AdjustGridSize( self.grid, rowsRequired = len(riderInfoList) )
		for row, ri in enumerate(riderInfoList):
			for col, field in enumerate(self.fieldNames):
				self.grid.SetCellValue( row, col, unicode(getattr(ri, field)) )
		self.grid.AutoSize()
		self.grid.EndBatch()
		
	def onColumnHeader( self, event ):
		self.sortCol = event.GetCol()
		self.updateGrid()
		
	def onImportFromExcel( self, event ):
		# Get the excel filename.
		openFileDialog = wx.FileDialog(self, "Open Excel file", "", "",
									   "Excel files (*.xls,*.xlsx,*.xlsm)|*.xls;*.xlsx;*.xlsm", wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)

		if openFileDialog.ShowModal() == wx.ID_CANCEL:
			return

		# proceed loading the file chosen by the user
		# this can be done with e.g. wxPython input streams:
		excelFile = openFileDialog.GetPath()
		
		excel = GetExcelReader( excelFile )
		
		# Get the sheet in the excel file.
		sheetName = excel.sheet_names()[0]

		riderInfo = {}
		bibHeader = set(v.lower() for v in ('Bib','BibNum','Bib Num', 'Bib #', 'Bib#'))
		fm = None
		for row in excel.iter_list(sheetName):
			if fm:
				f = fm.finder( row )
				info = {
					'bib': 			f('bib',u''),
					'first_name':	unicode(f('first_name',u'')).strip(),
					'last_name':	unicode(f('last_name',u'')).strip(),
					'license':		unicode(f('license_code',u'')).strip(),
					'team':			unicode(f('team',u'')).strip(),
					'uci_code':		unicode(f('uci_code',u'')).strip(),
				}
				
				info['bib'] = unicode(info['bib']).strip()
				if not info['bib']:	# If missing bib, assume end of input.
					continue
				
				# Check for comma-separated name.
				name = unicode(f('name', u'')).strip()
				if name and not info['first_name'] and not info['last_name']:
					try:
						info['last_name'], info['first_name'] = name.split(',',1)
					except:
						pass
				
				if not info['first_name'] and not info['last_name']:
					continue
				
				# If there is a bib it must be numeric.
				try:
					info['bib'] = int(unicode(info['bib']).strip())
				except ValueError:
					continue
				
				ri = Model.RiderInfo( **info )
				riderInfo[ri.bib] = ri
				
			elif any( unicode(h).strip().lower() in bibHeader for h in row ):
				fm = standard_field_map()
				fm.set_headers( row )
				
			Model.race.setRiderInfo( riderInfo )
			self.updateGrid()
		
	def refresh( self ):
		self.updateGrid()
		
	def commit( self ):
		race = Model.race
		riderInfo = {}
		for r in xrange(self.grid.GetNumberRows()):
			info = {f:self.grid.GetCellValue(r, c).strip() for c, f in enumerate(self.fieldNames)}
			if not info['bib']:
				continue
			ri = Model.RiderInfo( **info )
			riderInfo[ri.bib] = ri
		race.setRiderInfo( riderInfo )
			
########################################################################

class StartListFrame(wx.Frame):
	def __init__(self):
		"""Constructor"""
		wx.Frame.__init__(self, None, title="StartList", size=(800,600) )
		panel = StartList(self)
		panel.refresh()
		self.Show()
 
#----------------------------------------------------------------------
if __name__ == "__main__":
	app = wx.App(False)
	frame = StartListFrame()
	app.MainLoop()
