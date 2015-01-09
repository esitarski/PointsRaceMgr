import Utils
import Model
import Sprints
import wx
import wx.grid		as gridlib

import wx
import wx.grid as gridlib
import Utils

class Results( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(Results, self).__init__(parent, id, style=wx.BORDER_SUNKEN)
		self.SetBackgroundColour( wx.WHITE )

		self.hbs = wx.BoxSizer(wx.HORIZONTAL)

		self.gridResults = gridlib.Grid( self )
		labels = [u'Sprint\nPoints', u'Lap\nPoints', u'Laps\n+/-', u'Finish\nOrder', u'Num\nWins', u'Existing\nPoints']
		self.gridResults.CreateGrid( 0, len(labels) )
		for i, lab in enumerate(labels):
			self.gridResults.SetColLabelValue( i, lab )
		self.gridResults.SetRowLabelSize( 0 )
		self.gridResults.SetDefaultCellAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
		
		self.gridResults.AutoSize()
		
		Utils.MakeGridReadOnly( self.gridResults )
		
		self.gridResults.EnableDragColSize( False )
		self.gridResults.EnableDragRowSize( False )

		self.gridResults.AutoSize()
		
		self.gridResults.Bind( gridlib.EVT_GRID_CELL_LEFT_CLICK, self.onClick )

		self.hbs.Add( self.gridResults, 1, wx.GROW|wx.ALL, border = 4 )
		
		self.SetSizer(self.hbs)
		self.hbs.SetSizeHints(self)

	def clear( self ):
		Utils.DeleteAllGridRows( self.gridResults )
		
	def onClick( self, event ):
		wx.CallAfter( Utils.commitPanes )
		
	def refresh( self ):
		self.clear()
		race = Model.race
		if not race:
			return
			
		headers = []
		fields = []
		
		if race.pointsForLapping == 0:
			headers.append( u'Laps\n+/-' )
			fields.append( 'updown' )
		
		headers.append( u'Sprint\nPoints' )
		fields.append( 'sprintsTotal' )
		
		if race.pointsForLapping != 0:
			headers.append( u'Lap\nPoints' )
			fields.append( 'lapsTotal' )
		
		# Only update the Num Wins column if required for the ranking.
		if race.rankBy == race.RankByDistancePointsNumWins:
			headers.append(u'Num\nWins')
			fields.append('numWins')
		
		if race.existingPoints:
			headers.append( u'Existing\nPoints' )
			fields.append( 'existingPoints' )
			
		# Always add the finish order as the last criteria.
		headers.append( u'Finish\nOrder' )
		fields.append( 'finishOrder' )
		
		riders = race.getRiders()
		
		Utils.AdjustGridSize( self.gridResults, len(riders), len(headers) )
		for i, h in enumerate(headers):
			self.gridResults.SetColLabelValue( i, h )
		Utils.MakeGridReadOnly( self.gridResults )
		
		for r, rider in enumerate(riders):
			for c, field in enumerate(fields):
				v = getattr(rider,field)
				s = unicode(v) if v != 0 else u''
				if s and field == 'updown' and s[:1] != u'-':
					s = '+' + s
				if field == 'finishOrder' and s == u'1000':
					s = u''
				self.gridResults.SetCellValue( r, c, s )
	
	def commit( self ):
		pass
	
if __name__ == '__main__':
	app = wx.PySimpleApp()
	mainWin = wx.Frame(None,title="PointsRaceMgr", size=(600,400))
	Model.setRace( Model.Race() )
	Model.getRace()._populate()
	results = Results(mainWin)
	results.refresh()
	mainWin.Show()
	app.MainLoop()
