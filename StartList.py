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
		
		explanation = wx.StaticText( self, label=u'To delete a row, set Bib to blank and switch screens.' )
		
		self.addRows = wx.Button( self, label=u'Add Rows' )
		self.addRows.Bind( wx.EVT_BUTTON, self.onAddRows )
		
		self.importFromExcel = wx.Button( self, label=u'Import from Excel' )
		self.importFromExcel.Bind( wx.EVT_BUTTON, self.onImportFromExcel )
		
		hs = wx.BoxSizer( wx.HORIZONTAL )
		hs.Add( explanation, flag=wx.ALL|wx.ALIGN_CENTRE_VERTICAL, border=4 )
		hs.Add( self.addRows, flag=wx.ALL, border=4 )
		hs.Add( self.importFromExcel, flag=wx.ALL|wx.ALIGN_RIGHT, border=4 )
 
		self.fieldNames  = Model.RiderInfo.FieldNames
		self.headerNames = Model.RiderInfo.HeaderNames
		
		self.iExistingPointsCol = next(c for c, f in enumerate(self.fieldNames) if 'existing' in f)
		
		self.grid = AutoEditGrid( self, style = wx.BORDER_SUNKEN )
		self.grid.DisableDragRowSize()
		self.grid.SetRowLabelSize( 0 )
		self.grid.CreateGrid( 0, len(self.headerNames) )
		
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add( hs, 0, flag=wx.ALL|wx.EXPAND, border = 4 )
		sizer.Add(self.grid, 1, flag=wx.EXPAND|wx.ALL, border = 6)
		self.SetSizer(sizer)
		
		wx.CallAfter( self.refresh )
		
	def getGrid( self ):
		return self.grid
		
	def updateGrid( self ):
		race = Model.race
		riderInfo = race.riderInfo if race else []
		
		self.grid.BeginBatch()
		Utils.AdjustGridSize( self.grid, len(riderInfo), len(self.headerNames) )
		
		# Set specialized editors for appropriate columns.
		for col, name in enumerate(self.headerNames):
			self.grid.SetColLabelValue( col, name )
			attr = gridlib.GridCellAttr()
			if col == 0:
				attr.SetRenderer( gridlib.GridCellNumberRenderer() )
			elif col == 1:
				attr.SetRenderer( gridlib.GridCellFloatRenderer(precision=1) )
			attr.SetFont( Utils.BigFont() )
			self.grid.SetColAttr( col, attr )
		
		for row, ri in enumerate(riderInfo):
			for col, field in enumerate(self.fieldNames):
				self.grid.SetCellValue( row, col, unicode(getattr(ri, field)) )
		self.grid.AutoSize()
		self.grid.EndBatch()
		self.Layout()
		
	def onAddRows( self, event ):
		growSize = 10
		Utils.AdjustGridSize( self.grid, rowsRequired = self.grid.GetNumberRows()+growSize )
		self.Layout()
		self.grid.MakeCellVisible( self.grid.GetNumberRows()-growSize, 0 )
		
	def onImportFromExcel( self, event ):
		dlg = wx.MessageBox(
			u'Import from Excel\n\n'
			u'Reads the first sheet in the file.\n'
			u'Looks for the first row starting with "Bib","BibNum","Bib Num", "Bib #" or "Bib#".\n\n'
			u'Recognizes the following header fields (in any order, case insensitive):\n'
			u'\u2022 Bib|BibNum|Bib Num|Bib #|Bib#: Bib Number\n'
			u'\u2022 Points|Existing Points: Existing points at the start of the race.\n'
			u'\u2022 LastName|Last Name|LName: Last Name\n'
			u'\u2022 FirstName|First Name|FName: First Name\n'
			u'\u2022 Name: in the form "LastName, FirstName".  Used only if no Last Name or First Name\n'
			u'\u2022 Team|Team Name|TeamName|Rider Team|Club|Club Name|ClubName|Rider Club: Team\n'
			u'\u2022 License|Licence: Regional License (not uci code)\n'
			u'\u2022 UCI ID|UCIID: UCI ID.\n'
			u'\u2022 Nat Code|NatCode|NationCode: 3 letter nation code.\n'
			,
			u'Import from Excel',
			wx.OK|wx.CANCEL | wx.ICON_INFORMATION,
		)
		
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
					'uci_id':		unicode(f('uci_id',u'')).strip(),
					'nation_code':		unicode(f('nation_code',u'')).strip(),
					'existing_points':	unicode(f('existing_points',u'0')).strip(),
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
				
				# If there is a bib it must be numeric.
				try:
					info['bib'] = int(unicode(info['bib']).strip())
				except ValueError:
					continue
				
				# If there are existing points they must be numeric.
				try:
					info['existing_points'] = int(info['existing_points'])
				except ValueError:
					info['existing_points'] = 0
				
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
		riderInfo = []
		for r in xrange(self.grid.GetNumberRows()):
			info = {f:self.grid.GetCellValue(r, c).strip() for c, f in enumerate(self.fieldNames)}
			if not info['bib']:
				continue
			riderInfo.append( Model.RiderInfo(**info) )
		race.setRiderInfo( riderInfo )
		Utils.refresh()

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
