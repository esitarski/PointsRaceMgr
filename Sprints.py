import wx
import wx.grid		as gridlib
import Model
import Utils
import re

NumCols = 20
NumRows = 5	

notNumberRE = re.compile( '[^0-9]' )

class EnterHandlingGrid( gridlib.Grid ):
	def __init__(self, parent, NextColCheck ):
		gridlib.Grid.__init__(self, parent, -1)
		self.NextColCheck = NextColCheck
		self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)

	def NextColTop( self, c ):
		self.SetGridCursor( 0, c )
		self.MakeCellVisible( 0, c )

	def OnKeyDown(self, evt):
		if evt.GetKeyCode() != wx.WXK_RETURN:
			evt.Skip()
			return
		if evt.ControlDown():   # the edit control needs this key
			evt.Skip()
			return
		self.DisableCellEditControl()
		
		r = self.GetGridCursorRow()
		c = self.GetGridCursorCol()
		
		if c != self.GetNumberCols() - 1 and self.NextColCheck( r, c ):
			self.NextColTop( c + 1 )
			return
			
		if r != self.GetNumberRows() - 1:
			self.SetGridCursor( r + 1, c )

class Sprints( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(Sprints, self).__init__(parent, id, style=wx.BORDER_SUNKEN)
		
		self.SetDoubleBuffered(True)
		self.SetBackgroundColour( wx.WHITE )

		self.hbs = wx.BoxSizer(wx.HORIZONTAL)

		self.gridPoints = gridlib.Grid( self )
		self.gridPoints.CreateGrid( NumRows, 4 )
		self.gridPoints.SetColLabelValue( 0, '' )
		self.gridPoints.SetColLabelValue( 1, "Place" )
		self.gridPoints.SetColLabelAlignment( 1, wx.ALIGN_CENTRE )
		self.gridPoints.SetColLabelValue( 2, "Points" )
		self.gridPoints.SetColLabelAlignment( 2, wx.ALIGN_CENTRE )
		self.gridPoints.SetColLabelValue( 3, '' )
		self.gridPoints.SetDefaultCellAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
		for i, (place, points) in enumerate((('1st',5),('2nd',3),('3rd',2),('4th',1),('5th',0))):
			self.gridPoints.SetCellValue( i, 1, place )
			self.gridPoints.SetCellValue( i, 2, unicode(points) )
			
		self.gridPoints.SetRowLabelSize( 0 )
		self.gridPoints.EnableDragRowSize( False )
		self.gridPoints.EnableDragColSize( False )
		self.gridPoints.AutoSize()
		self.gridPoints.SetColSize( 0, 40 )
		self.gridPoints.SetColSize( 2, 64 )
		self.gridPoints.SetColSize( 3, 16 )
		attr = gridlib.GridCellAttr()
		attr.SetReadOnly()
		self.gridPoints.SetColAttr( 0, attr )
		self.gridPoints.SetColAttr( 1, attr )
		
		self.hbs.Add( self.gridPoints, 0, wx.ALL, border=4 )
		
		self.gridSprint = EnterHandlingGrid( self, self.NextColCheck )
		self.gridSprint.CreateGrid( NumRows, NumCols )
		for i in xrange( NumCols ):
			self.gridSprint.SetColLabelValue( i, "Sp%d" % (i+1) )
		self.gridSprint.SetRowLabelSize( 0 )
		self.gridSprint.SetDefaultCellAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
		self.gridSprint.EnableDragRowSize( False )
		self.gridSprint.EnableDragColSize( False )
		self.gridSprint.AutoSize()
		widestCol = int(self.gridSprint.GetColSize( self.gridSprint.GetNumberCols() - 1 ) * 1.1	)
		for i in xrange( self.gridSprint.GetNumberCols() ):
			self.gridSprint.SetColSize( i, widestCol )
			self.gridSprint.SetColFormatNumber( i )
		self.gridSprint.Bind(wx.EVT_SCROLLWIN, self.onScroll)
		
		self.Bind(gridlib.EVT_GRID_CELL_CHANGE, self.onCellChange)
		self.Bind(gridlib.EVT_GRID_EDITOR_CREATED, self.onGridEditorCreated)
		
		self.gridPoints.Bind(wx.EVT_SCROLLWIN, self.onScroll)
		self.gridSprint.Bind(wx.EVT_SCROLLWIN, self.onScroll)

		self.hbs.Add( self.gridSprint, 1, wx.GROW|wx.ALL, border=4 )
		
		self.SetSizer(self.hbs)
		self.hbs.SetSizeHints(self)

	def onGridEditorCreated( self, event ):
		editor = event.GetControl()
		editor.Bind( wx.EVT_KILL_FOCUS, self.onKillFocus )
		event.Skip()
		
	def onKillFocus( self, event ):
		grid = event.GetEventObject().GetGrandParent()
		grid.SaveEditControlValue()
		grid.HideCellEditControl()
		event.Skip()
		
	def NextColCheck( self, r, c ):
		if c == self.gridSprint.GetNumberCols() - 1:
			return False
		if r == self.gridSprint.GetNumberRows() - 1 or not self.gridPoints.GetCellValue(r+1, 2):
			return True
		
	def onCellChange( self, evt ):
		r = evt.GetRow()
		c = evt.GetCol()
		if evt.GetEventObject() == self.gridSprint:
			value = self.gridSprint.GetCellValue(r, c)
			value = notNumberRE.sub( '', value )
			self.gridSprint.SetCellValue(r, c, value)
		else:
			value = self.gridPoints.GetCellValue(r, c)
			value = notNumberRE.sub( '', value )
			self.gridPoints.SetCellValue(r, c, value)
					
		self.commit()
		self.refresh()
		self.updateShading()
		
	def onScroll(self, evt): 
		grid = evt.GetEventObject()
		orientation = evt.GetOrientation()
		if grid == self.gridSprint:
			if orientation == wx.SB_HORIZONTAL:
				wx.CallAfter(self.alignScrollHorizontal, grid)
		if orientation == wx.SB_VERTICAL:
			wx.CallAfter( self.alignScrollVertical, grid )
		evt.Skip() 

	def alignScrollHorizontal(self, grid): 
		try:
			mainWin = Utils.getMainWin()
			Utils.AlignHorizontalScroll( self.gridSprint, mainWin.worksheet.gridWorksheet )
		except:
			pass

	def alignScrollVertical( self, grid ):
		try:
			mainWin = Utils.getMainWin()
			Utils.AlignVerticalScroll( grid, self.gridPoints )
			Utils.AlignVerticalScroll( grid, self.gridSprint )
		except Exception as e:
			print e
			pass
			
	def clear( self ):
		for r in xrange( self.gridSprint.GetNumberRows() ):
			for c in xrange( self.gridSprint.GetNumberCols() ):
				self.gridSprint.SetCellValue( r, c, u'' )
	
	def updateShading( self ):
		race = Model.race
		if not race:
			return
		numSprints = race.getNumSprints()
		for r in xrange( self.gridSprint.GetNumberRows() ):
			for c in xrange( 0, numSprints ):
				self.gridSprint.SetCellBackgroundColour( r, c, wx.WHITE )
			for c in xrange( numSprints, self.gridSprint.GetNumberCols() ):
				self.gridSprint.SetCellBackgroundColour( r, c, Utils.LightGrey )
		
		# Highlight duplicates
		for c in xrange( 0, self.gridSprint.GetNumberCols() ):
			numCount = {}
			for r in xrange( 0, self.gridSprint.GetNumberRows() ):
				cell = self.gridSprint.GetCellValue(r, c)
				if not cell:
					continue
				num = int(self.gridSprint.GetCellValue(r, c))
				numCount[num] = numCount.get(num, 0) + 1
			for r in xrange( 0, self.gridSprint.GetNumberRows() ):
				cell = self.gridSprint.GetCellValue(r, c)
				if not cell:
					continue
				num = int(self.gridSprint.GetCellValue(r, c))
				if numCount[num] > 1:
					self.gridSprint.SetCellBackgroundColour( r, c, Utils.BadHighlightColour )
					
		self.gridSprint.ForceRefresh()
	
	def refresh( self ):
		self.clear()
		race = Model.race
		if not race:
			return
		for (sprint, place), num in race.sprintResults.iteritems():
			self.gridSprint.SetCellValue( place - 1, sprint - 1, str(num) )
	
		for (place, points) in race.pointsForPlace.iteritems():
			row = place - 1
			self.gridPoints.SetCellValue( row, 1, Utils.ordinal(place) if points >= 0 else u'' )
			self.gridPoints.SetCellValue( row, 2, unicode(points) if points >= 0 else u'' )
	
		self.updateShading()
	
	def commit( self ):
		race = Model.race
		if not race:
			return
		newSprintResults = {}
		for r in xrange( self.gridSprint.GetNumberRows() ):
			place = r + 1
			for c in xrange( self.gridSprint.GetNumberCols() ):
				sprint = c + 1
				cell = self.gridSprint.GetCellValue( r, c )
				try:
					num = int(cell)
					newSprintResults[(sprint, place)] = num
				except ValueError:
					pass
		race.setSprintResults( newSprintResults )
		
		newPointsForPlace = {}
		for r in xrange( self.gridPoints.GetNumberRows() ):
			cell = self.gridPoints.GetCellValue( r, 2 )
			try:
				points = int(cell)
			except ValueError:
				points = -1
			newPointsForPlace[r+1] = points
		race.setPoints( pointsForPlace = newPointsForPlace )
		Utils.refresh( self )
	
if __name__ == '__main__':
	app = wx.PySimpleApp()
	mainWin = wx.Frame(None,title="PointsRaceMgr", size=(600,400))
	Model.setRace( Model.Race() )
	Model.getRace()._populate()
	sprints = Sprints(mainWin)
	sprints.refresh()
	mainWin.Show()
	app.MainLoop()
