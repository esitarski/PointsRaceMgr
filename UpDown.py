import Utils
import Model
import Sprints
import wx
import re
import wx.grid		as gridlib

notNumberRE = re.compile( u'[^0-9]' )

# Columns for the table.
BibUpDown = 0
ValUpDown = 1

BibStatus = 3
ValStatus = 4

ValFinish = 6
BibFinish = 7

EmptyCols = [2, 5]

class UpDown( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(UpDown, self).__init__( parent, id, style=wx.BORDER_SUNKEN)
		self.SetBackgroundColour( wx.WHITE )
		
		self.inCellChange = False

		self.hbs = wx.BoxSizer(wx.HORIZONTAL)

		self.gridUpDown = gridlib.Grid( self )
		labels = [u'Bib', u'Laps\n+/-', u' ', u'Bib', u'Status', u' ', u'Finish\nOrder', u'Bib']
		self.gridUpDown.CreateGrid( 40, len(labels) )
		
		for col, colName in enumerate(labels):
			self.gridUpDown.SetColLabelValue( col, colName )
			
			attr = gridlib.GridCellAttr()
						
			if col in (BibUpDown, BibFinish, BibStatus):
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
		
		for r in xrange( self.gridUpDown.GetNumberRows() ):
			self.gridUpDown.SetCellValue( r, ValFinish, unicode(r+1) )
			
		orderNum = {order: num for num, order in race.getFinishOrder()}
		for num, order in race.getFinishOrder():
			self.gridUpDown.SetCellValue( order-1, BibFinish, unicode(num) )
			
		for r, (num, s) in enumerate(race.getStatus()):
			self.gridUpDown.SetCellValue( r, BibStatus, unicode(num) )
			self.gridUpDown.SetCellValue( r, ValStatus, Model.Rider.statusNames[s] )
			
	def commit( self ):
		race = Model.race
		if not race:
			return
		lastFinishVal = 0
		updowns = {}
		finishOrder = {}
		status = {}
		statusIndex = dict( (n, i) for i, n in enumerate(Model.Rider.statusNames) )
		for r in xrange(self.gridUpDown.GetNumberRows()):
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
