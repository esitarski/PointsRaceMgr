import Utils
import Model
import Sprints
import wx
import wx.grid		as gridlib

class Worksheet( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(Worksheet, self).__init__(parent, id, style=wx.BORDER_SUNKEN)

		self.SetBackgroundColour( wx.WHITE )

		self.hbs = wx.BoxSizer(wx.HORIZONTAL)

		self.gridBib = gridlib.Grid( self )
		self.gridBib.CreateGrid( 1, 4 )
		self.gridBib.SetColLabelValue( 0, u"Rank" )
		self.gridBib.SetColLabelValue( 1, u"Bib" )
		self.gridBib.SetColLabelValue( 2, u"Total\nPoints" )
		self.gridBib.SetColLabelValue( 3, u'')
		self.gridBib.SetDefaultCellAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
		self.gridBib.SetRowLabelSize( 0 )
		self.gridBib.EnableDragColSize( False )
		self.gridBib.EnableDragRowSize( False )
		self.gridBib.AutoSize()
		Utils.MakeGridReadOnly( self.gridBib )
		self.gridBib.Bind(wx.EVT_SCROLLWIN, self.onScroll)
		
		self.hbs.Add( self.gridBib, 0, wx.ALL, border = 4 )
		
		self.gridWorksheet = gridlib.Grid( self )
		self.gridWorksheet.CreateGrid( 0, Sprints.NumCols )
		for i in xrange( Sprints.NumCols ):
			self.gridWorksheet.SetColLabelValue( i, u"Sp{}".format(i+1) )
		self.gridWorksheet.SetRowLabelSize( 0 )
		self.gridWorksheet.SetDefaultCellAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
		Utils.MakeGridReadOnly( self.gridWorksheet )
		self.gridWorksheet.EnableDragColSize( False )
		self.gridWorksheet.EnableDragRowSize( False )
		self.gridWorksheet.Bind(wx.EVT_SCROLLWIN, self.onScroll)
		
		self.syncSize()
		
		self.hbs.Add( self.gridWorksheet, 1, wx.GROW|wx.ALL, border = 4 )
		
		self.SetSizer(self.hbs)
		self.hbs.SetSizeHints(self)
		
	def onScroll(self, evt): 
		grid = evt.GetEventObject() 
		orientation = evt.GetOrientation()
		wx.CallAfter( self.alignScrollHorizontal if orientation == wx.SB_HORIZONTAL else self.alignScrollVertical, grid )
		evt.Skip() 

	def alignScrollHorizontal( self, grid ): 
		try:
			mainWin = Utils.getMainWin()
			Utils.AlignHorizontalScroll( self.gridWorksheet, mainWin.sprints.gridSprint )
		except:
			pass

	def alignScrollVertical( self, grid ):
		try:
			mainWin = Utils.getMainWin()
			Utils.AlignVerticalScroll( grid, self.gridBib )
			Utils.AlignVerticalScroll( grid, self.gridWorksheet )
			Utils.AlignVerticalScroll( grid, mainWin.results.gridResults )
		except Exception as e:
			pass
			
	def clear( self ):
		Utils.DeleteAllGridRows( self.gridBib )
		Utils.DeleteAllGridRows( self.gridWorksheet )
	
	def syncSize( self ):
		mainWin = Utils.getMainWin()
		if mainWin:
			for c in xrange( self.gridWorksheet.GetNumberCols() ):
				self.gridWorksheet.SetColSize( c, mainWin.sprints.gridSprint.GetColSize(c) )
				#self.gridWorksheet.SetColFormatNumber( c )
			self.gridWorksheet.SetDefaultCellAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )

			for c in xrange( self.gridBib.GetNumberCols() ):
				self.gridBib.SetColSize( c, mainWin.sprints.gridPoints.GetColSize(c) )
				#self.gridBib.SetColFormatNumber( c )
			self.gridBib.SetDefaultCellAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
		else:
			widestCol = 40
			for c in xrange( self.gridWorksheet.GetNumberCols() ):
				self.gridWorksheet.SetColSize( c, widestCol )
				#self.gridWorksheet.SetColFormatNumber( c )

	def refresh( self ):
		self.clear()
		race = Model.race
		if not race:
			return
		riders = race.getRiders()
		if not riders:
			return
		self.gridBib.InsertRows( 0, len(riders), True )
		self.gridWorksheet.InsertRows( 0, len(riders), True )
		riderToRow = {}
		position = 1
		for r, rider in enumerate(riders):
			if r > 0 and not riders[r-1].tiedWith(rider):
				position += 1
			elif r > 0:			# Highlight ties.
				for ic in xrange(3):
					self.gridBib.SetCellBackgroundColour( r-1, ic, Utils.BadHighlightColour )
					self.gridBib.SetCellBackgroundColour( r,   ic, Utils.BadHighlightColour )
			
			self.gridBib.SetCellValue( r, 0, unicode(position)	if rider.status == Model.Rider.Finisher else
											 Model.Rider.statusNames[rider.status] )
			self.gridBib.SetCellValue( r, 1, unicode(rider.num) )
			self.gridBib.SetCellValue( r, 2, unicode(rider.pointsTotal) )
			riderToRow[rider.num] = r
		
		for (sprint, place), num in race.sprintResults.iteritems():
			if num in riderToRow:
				points = race.getSprintPoints(sprint, place)
				if points > 0:
					self.gridWorksheet.SetCellValue( riderToRow[num], sprint - 1, unicode(points) )
		
		numSprints = race.getNumSprints()
		for r in xrange( self.gridWorksheet.GetNumberRows() ):
			for c in xrange( numSprints, self.gridWorksheet.GetNumberCols() ):
				self.gridWorksheet.SetCellBackgroundColour( r, c, Utils.LightGrey )
		
		self.syncSize()
		
		self.hbs.Layout()
	
	def commit( self ):
		pass
	
if __name__ == '__main__':
	app = wx.App( False )
	mainWin = wx.Frame(None,title="PointsRaceMgr", size=(600,400))
	Model.setRace( Model.Race() )
	Model.getRace()._populate()
	worksheet = Worksheet(mainWin)
	worksheet.refresh()
	mainWin.Show()
	app.MainLoop()
