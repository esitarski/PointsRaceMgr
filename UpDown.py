import Utils
import Model
import Sprints
import wx
import re
import wx.grid		as gridlib
import wx.lib.mixins.grid as gae

notNumberRE = re.compile( u'[^0-9]' )

# Columns for the table.
BibExistingPoints = 0
ValExistingPoints = 1

BibUpDown = 3
ValUpDown = 4

BibStatus = 6
ValStatus = 7

ValFinish = 9
BibFinish = 10

EmptyCols = [2, 5, 8]

class UpDownEditor(gridlib.PyGridCellEditor):
	Empty = u''
	
	def __init__(self):
		self._tc = None
		self.startValue = self.Empty
		gridlib.PyGridCellEditor.__init__(self)
		
	def Create( self, parent, id = wx.ID_ANY, evtHandler = None ):
		self._tc = wx.SpinCtrl(parent, id, style = wx.TE_PROCESS_ENTER, min=-160, max=160)
		self.SetControl( self._tc )
		if evtHandler:
			self._tc.PushEventHandler( evtHandler )
	
	def SetSize( self, rect ):
		self._tc.SetDimensions(rect.x, rect.y, rect.width+2, rect.height+2, wx.SIZE_ALLOW_MINUS_ONE )
	
	def BeginEdit( self, row, col, grid ):
		self.startValue = grid.GetTable().GetValue(row, col).strip()
		v = self.startValue
		self._tc.SetValue( int(v or u'0') )
		self._tc.SetFocus()
		
	def EndEdit( self, row, col, grid, value = None ):
		changed = False
		v = self._tc.GetValue()
		if v == 0:
			v = u''
		elif v > 0:
			v = u'+' + unicode(v)
		else:
			v = unicode(v)
		
		if v != self.startValue:
			changed = True
			grid.GetTable().SetValue( row, col, v )
		
		self.startValue = self.Empty
		self._tc.SetValue( 0 )
		
	def Reset( self ):
		self._tc.SetValue( self.startValue )
		
	def Clone( self ):
		return UpDownEditor()

class UpDownGrid( gridlib.Grid, gae.GridAutoEditMixin ):
	def __init__( self, parent, id=wx.ID_ANY, style=0 ):
		gridlib.Grid.__init__( self, parent, id=id, style=style )
		gae.GridAutoEditMixin.__init__(self)
		
class UpDown( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(UpDown, self).__init__( parent, id, style=wx.BORDER_SUNKEN)
		self.SetBackgroundColour( wx.WHITE )
		
		self.inCellChange = False

		self.hbs = wx.BoxSizer(wx.HORIZONTAL)

		self.gridUpDown = UpDownGrid( self )
		labels = [u'Bib', u'Existing\nPoints', u' ', u'Bib', u'Laps\n+/-', u' ', u'Bib', u'Status', u' ', u'Finish\nOrder', u'Bib', ]
		self.gridUpDown.CreateGrid( 200, len(labels) )
		
		for col, colName in enumerate(labels):
			self.gridUpDown.SetColLabelValue( col, colName )
			
			attr = gridlib.GridCellAttr()
						
			if col in (BibUpDown, BibFinish, BibStatus, BibExistingPoints, ValExistingPoints):
				attr.SetEditor( gridlib.GridCellNumberEditor() )
				attr.SetAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
			
			elif col == ValUpDown:
				attr.SetAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
			
			elif col == ValStatus:
				attr.SetEditor( gridlib.GridCellChoiceEditor([' '] + Model.Rider.statusNames[1:], False) )
				attr.SetAlignment( wx.ALIGN_CENTRE, wx.ALIGN_CENTRE )
			
			else:
				attr.SetReadOnly()
				attr.SetAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
				
			if col == ValUpDown:
				attr.SetEditor( UpDownEditor() )
				
			if col in EmptyCols:
				attr.SetBackgroundColour( self.gridUpDown.GetLabelBackgroundColour() )
				
			self.gridUpDown.SetColAttr( col, attr )

		self.gridUpDown.SetRowLabelSize( 0 )
		
		self.gridUpDown.EnableDragColSize( False )
		self.gridUpDown.EnableDragRowSize( False )
		self.gridUpDown.AutoSize()
		
		try:
			mainWin = Utils.getMainWin()
			widestCol = mainWin.sprints.gridWorksheet.GetColSize( 0 )
		except:
			widestCol = 64
		for i in xrange( self.gridUpDown.GetNumberCols() ):
			self.gridUpDown.SetColSize( i, max(widestCol, self.gridUpDown.GetColSize(i)) )

		for col in EmptyCols:
			self.gridUpDown.SetColSize( col, 16 )
		
		self.Bind(gridlib.EVT_GRID_CELL_CHANGE, self.onCellChange)
		self.Bind(gridlib.EVT_GRID_SELECT_CELL, self.onCellEnableEdit)
		self.Bind(gridlib.EVT_GRID_LABEL_LEFT_CLICK, self.onLabelClick)
		self.Bind(gridlib.EVT_GRID_EDITOR_CREATED, self.onGridEditorCreated)
		self.hbs.Add( self.gridUpDown, 1, wx.GROW|wx.ALL, border = 4 )
		
		self.gridUpDown.AutoSizeColumn( 9 )
		
		self.SetSizer(self.hbs)
		self.hbs.SetSizeHints(self)

	def onLeaveWindow( self, evt ):
		pass
		
	def onGridEditorCreated( self, event ):
		editor = event.GetControl()
		editor.Bind( wx.EVT_KILL_FOCUS, self.onKillFocus )
		event.Skip()
		
	def onKillFocus( self, event ):
		grid = event.GetEventObject().GetGrandParent()
		grid.SaveEditControlValue()
		grid.HideCellEditControl()
		event.Skip()
		
	def onLabelClick( self, evt ):
		self.gridUpDown.DisableCellEditControl()
		self.refresh()
		
	def onCellChange( self, evt ):
		r = evt.GetRow()
		c = evt.GetCol()
		value = self.gridUpDown.GetCellValue(r, c)
		value = value.strip()
		neg = True if value and value[0] == u'-' and value != u'-' else False
		if c != ValStatus:
			value = notNumberRE.sub( u'', value )
			if value in (u'-', u'-0', u'0'):
				value = u''
				neg = False
		if c == ValUpDown:
			if neg:
				value = u'-' + value
			else:
				try:
					v = int(value)
					if v > 0:
						value = u'+' + value
				except:
					pass

		self.gridUpDown.SetCellValue(r, c, value)
		self.commit()
		Utils.refreshResults()
	
	def onCellEnableEdit( self, evt ):
		if evt.GetCol() == ValStatus:
			wx.CallAfter( self.gridUpDown.EnableCellEditControl )
		evt.Skip()
	
	def clear( self ):
		for r in xrange( self.gridUpDown.GetNumberRows() ):
			for c in xrange( self.gridUpDown.GetNumberCols() ):
				self.gridUpDown.SetCellValue( r, c, u'' )
		
	def refresh( self ):
		self.clear()
		race = Model.race
		if not race:
			return
			
		riderInfo = {}
		
		for r, (num, updown) in enumerate(race.getUpDown()):
			self.gridUpDown.SetCellValue( r, BibUpDown, unicode(num) )
			prefix = '+' if updown > 0 else u''
			self.gridUpDown.SetCellValue( r, ValUpDown, prefix + unicode(updown) )
		
		for r, (num, s) in enumerate(race.getStatus()):
			self.gridUpDown.SetCellValue( r, BibStatus, unicode(num) )
			self.gridUpDown.SetCellValue( r, ValStatus, Model.Rider.statusNames[s] )
			
		for r in xrange( self.gridUpDown.GetNumberRows() ):
			self.gridUpDown.SetCellValue( r, ValFinish, unicode(r+1) )
			
		orderNum = {order: num for num, order in race.getFinishOrder()}
		for num, order in race.getFinishOrder():
			self.gridUpDown.SetCellValue( order-1, BibFinish, unicode(num) )
		
		for r, (num, p) in enumerate(race.getExistingPoints()):
			self.gridUpDown.SetCellValue( r, BibStatus, unicode(num) )
			self.gridUpDown.SetCellValue( r, ValStatus, unicode(p) )
			
	def commit( self ):
		race = Model.race
		if not race:
			return
		lastFinishVal = 0
		existingPoints = {}
		updowns = {}
		finishOrder = {}
		status = {}
		statusIndex = dict( (n, i) for i, n in enumerate(Model.Rider.statusNames) )
		for r in xrange(self.gridUpDown.GetNumberRows()):
			try:
				existingPoints[int(self.gridUpDown.GetCellValue(r, BibExistingPoints), 10)] = int(self.gridUpDown.GetCellValue(r, ValExistingPoints), 10)
			except ValueError:
				pass
		
			try:
				updowns[int(self.gridUpDown.GetCellValue(r, BibUpDown), 10)] = int(self.gridUpDown.GetCellValue(r, ValUpDown), 10)
			except ValueError:
				pass
				
			try:
				bib = int(self.gridUpDown.GetCellValue(r, BibFinish), 10)
				if bib == 0:
					continue
				finishStr = self.gridUpDown.GetCellValue(r, ValFinish).strip()
				if finishStr:
					finishVal = int(finishStr)
					lastFinishVal = finishVal
				else:
					finishVal = lastFinishVal + 1
				finishOrder[bib] = finishVal
				lastFinishVal = finishVal
			except ValueError:
				pass
				
			try:
				status[int(self.gridUpDown.GetCellValue(r, BibStatus), 10)] = statusIndex[self.gridUpDown.GetCellValue(r, ValStatus)]
			except (ValueError, KeyError):
				pass

		race.setExistingPoints( existingPoints )
		race.setUpDowns( updowns )
		race.setFinishOrder( finishOrder )
		race.setStatus( status )
	
if __name__ == '__main__':
	app = wx.App( False )
	mainWin = wx.Frame(None,title="PointsRaceMgr", size=(600,400))
	Model.setRace( Model.Race() )
	Model.getRace()._populate()
	updown = UpDown(mainWin)
	updown.refresh()
	mainWin.Show()
	app.MainLoop()
