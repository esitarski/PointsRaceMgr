import wx
import wx.grid as gridlib
import wx.lib.mixins.grid as gae

import os
import sys
import operator
from html.parser import HTMLParser

import Utils
import Model
from FieldMap import standard_field_map
from Excel import GetExcelReader

class AutoEditGrid( gridlib.Grid, gae.GridAutoEditMixin ):
	def __init__( self, parent, id=wx.ID_ANY, style=0 ):
		gridlib.Grid.__init__( self, parent, id=id, style=style )
		gae.GridAutoEditMixin.__init__(self)

class TableHTMLParser( HTMLParser ):
	def __init__( self, *args, **kwargs ):
		super(TableHTMLParser, self).__init__( *args, **kwargs )
		self.result = []
		self.in_cell = False
		self.colspan = 1
		self.data = []

	def handle_starttag(self, tag, attrs):
		if tag == 'tr':
			self.result.append( [] )
		elif tag in ('td', 'th'):
			self.in_cell = True
			self.colspan = int( dict(attrs).get('colspan',1) )

	def handle_endtag(self, tag):
		if tag in ('td', 'th'):
			self.in_cell = False
			self.result[-1].append( ''.join(self.data) )
			del self.data[:]
			self.result[-1].extend( [''] * (self.colspan-1) )

	def handle_data(self, data):
		if self.in_cell:
			self.data.append( data )
		
def listFromHtml( html ):
	parser = TableHTMLParser()
	parser.feed( html )
	parser.close()
	return parser.result
	
#--------------------------------------------------------------------------------
class StartList(wx.Panel):

	def __init__(self, parent):
		wx.Panel.__init__(self, parent)
		
		explanation = wx.StaticText( self, label=u'To delete a row, set Bib to blank and press Commit.' )
		
		self.commitButton = wx.Button( self, label=u'Commit' )
		self.commitButton.Bind( wx.EVT_BUTTON, self.onCommit )
		
		self.addRows = wx.Button( self, label=u'Add Rows' )
		self.addRows.Bind( wx.EVT_BUTTON, self.onAddRows )
		
		self.importFromExcel = wx.Button( self, label=u'Import from Excel' )
		self.importFromExcel.Bind( wx.EVT_BUTTON, self.onImportFromExcel )
		
		self.pasteFromClipboard = wx.Button( self, id=wx.ID_PASTE, label=u'Paste \U0001F4CB' )
		self.pasteFromClipboard.Bind( wx.EVT_BUTTON, self.onPaste )
		
		hs = wx.BoxSizer( wx.HORIZONTAL )
		hs.Add( explanation, flag=wx.ALL|wx.ALIGN_CENTRE_VERTICAL, border=4 )
		hs.Add( self.commitButton, flag=wx.ALL, border=4 )
		hs.Add( self.addRows, flag=wx.ALL, border=4 )
		hs.Add( self.importFromExcel, flag=wx.ALL, border=4 )
		hs.Add( self.pasteFromClipboard, flag=wx.ALL, border=4 )
 
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
		
		'''
		self.Bind(wx.EVT_MENU, lambda event: self.onPaste)		# Required to process accelerator.
		entries = [wx.AcceleratorEntry()] 
		entries[0].Set(wx.ACCEL_CTRL, ord('V'), wx.ID_PASTE)
		accel = wx.AcceleratorTable(entries)
		self.SetAcceleratorTable(accel)
		'''
				
		self.SetDoubleBuffered( True )
		wx.CallAfter( self.refresh )
		
	def getGrid( self ):
		return self.grid
		
	def updateGrid( self ):
		race = Model.race
		riderInfo = race.riderInfo if race else []
		
		Utils.AdjustGridSize( self.grid, len(riderInfo), len(self.headerNames) )
		
		# Set specialized editors for appropriate columns.
		for col, name in enumerate(self.headerNames):
			self.grid.SetColLabelValue( col, name )
			attr = gridlib.GridCellAttr()
			if col == 0:
				attr.SetRenderer( gridlib.GridCellNumberRenderer() )
			elif col == 1:
				attr.SetRenderer( gridlib.GridCellFloatRenderer(precision=1) )
			if name == 'Bib':
				attr.SetBackgroundColour( wx.Colour(178, 236, 255) )
			attr.SetFont( Utils.BigFont() )
			self.grid.SetColAttr( col, attr )
		
		missingBib = 5000
		for row, ri in enumerate(riderInfo):
			for col, field in enumerate(self.fieldNames):
				v = getattr(ri, field, None)
				if v is None:
					if field == 'bib':
						v = ri.bib = missingBib
						missingBib += 1
				self.grid.SetCellValue( row, col, u'{}'.format(v) )
		self.grid.AutoSize()
		self.Layout()
		
	def onAddRows( self, event ):
		growSize = 10
		Utils.AdjustGridSize( self.grid, rowsRequired = self.grid.GetNumberRows()+growSize )
		self.Layout()
		self.grid.MakeCellVisible( self.grid.GetNumberRows()-growSize, 0 )
	
	bibHeader = set(v.lower() for v in ('Bib','BibNum','Bib Num', 'Bib #', 'Bib#'))
	
	def onPaste( self, event ):
		success = False
		table = None
		
		# Try to get html format.
		html_data = wx.HTMLDataObject()
		if wx.TheClipboard.Open():
			success = wx.TheClipboard.GetData(html_data)
			wx.TheClipboard.Close()
		if success:
			table = listFromHtml( html_data.GetHTML() )
			if not table:
				success = False
				
		# If no success, try tab delimited.
		if not success:
			text_data = wx.TextDataObject()
			if wx.TheClipboard.Open():
				success = wx.TheClipboard.GetData(text_data)
				wx.TheClipboard.Close()
			if success:
				table = []
				for line in text_data.GetText().split('\n'):
					table.append( line.split('\t') )
				if not table:
					success = False
		if not success:
			return

		riderInfo = []
		fm = None
		for row in table:
			if fm:
				f = fm.finder( row )
				info = {
					'bib': 			f('bib',u''),
					'first_name':	u'{}'.format(f('first_name',u'')).strip(),
					'last_name':	u'{}'.format(f('last_name',u'')).strip(),
					'license':		u'{}'.format(f('license_code',u'')).strip(),
					'team':			u'{}'.format(f('team',u'')).strip(),
					'uci_id':		u'{}'.format(f('uci_id',u'')).strip(),
					'nation_code':		u'{}'.format(f('nation_code',u'')).strip(),
					'existing_points':	u'{}'.format(f('existing_points',u'0')).strip(),
				}
				
				info['bib'] = u'{}'.format(info['bib']).strip()
				if not info['bib']:	# If missing bib, assume end of input.
					continue
				
				# Check for comma-separated name.
				name = u'{}'.format(f('name', u'')).strip()
				if name and not info['first_name'] and not info['last_name']:
					try:
						info['last_name'], info['first_name'] = name.split(',',1)
					except:
						pass
				
				# If there is a bib it must be numeric.
				try:
					info['bib'] = int(u'{}'.format(info['bib']).strip())
				except ValueError:
					continue

				# If there are existing points they must be numeric.
				try:
					info['existing_points'] = float(info['existing_points'])
				except ValueError:
					if info['existing_points'] not in Model.Rider.statusNames:
						info['existing_points'] = 0.0
				
				ri = Model.RiderInfo( **info )
				riderInfo.append( ri )
				
			elif any( u'{}'.format(h).strip().lower() in self.bibHeader for h in row ):
				fm = standard_field_map()
				fm.set_headers( row )
				
			Model.race.setRiderInfo( riderInfo )
			self.updateGrid()
		
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

		riderInfo = []
		fm = None
		for row in excel.iter_list(sheetName):
			if fm:
				f = fm.finder( row )
				info = {
					'bib': 			f('bib',u''),
					'first_name':	u'{}'.format(f('first_name',u'')).strip(),
					'last_name':	u'{}'.format(f('last_name',u'')).strip(),
					'license':		u'{}'.format(f('license_code',u'')).strip(),
					'team':			u'{}'.format(f('team',u'')).strip(),
					'uci_id':		u'{}'.format(f('uci_id',u'')).strip(),
					'nation_code':		u'{}'.format(f('nation_code',u'')).strip(),
					'existing_points':	u'{}'.format(f('existing_points',u'0')).strip(),
				}
				
				info['bib'] = u'{}'.format(info['bib']).strip()
				if not info['bib']:	# If missing bib, assume end of input.
					continue
				
				# Check for comma-separated name.
				name = u'{}'.format(f('name', u'')).strip()
				if name and not info['first_name'] and not info['last_name']:
					try:
						info['last_name'], info['first_name'] = name.split(',',1)
					except:
						pass
				
				# If there is a bib it must be numeric.
				try:
					info['bib'] = int(u'{}'.format(info['bib']).strip())
				except ValueError:
					continue
				
				# If there are existing points they must be numeric.
				try:
					info['existing_points'] = float(info['existing_points'])
				except ValueError:
					info['existing_points'] = 0
				
				ri = Model.RiderInfo( **info )
				riderInfo.append( ri )
				
			elif any( u'{}'.format(h).strip().lower() in self.bibHeader for h in row ):
				fm = standard_field_map()
				fm.set_headers( row )
				
			Model.race.setRiderInfo( riderInfo )
			self.updateGrid()
		
	def refresh( self ):
		self.updateGrid()
		
	def commit( self, local=False ):
		self.grid.SaveEditControlValue()
		self.grid.DisableCellEditControl()
		race = Model.race
		riderInfo = []
		for r in range(self.grid.GetNumberRows()):
			info = {f:self.grid.GetCellValue(r, c).strip() for c, f in enumerate(self.fieldNames)}
			if not info['bib']:
				continue
			riderInfo.append( Model.RiderInfo(**info) )
		race.setRiderInfo( riderInfo )
		if not local:
			Utils.refresh()

	def onCommit( self, event ):
		self.commit( True )
		self.refresh()
		
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
