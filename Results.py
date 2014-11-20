import Utils
import Model
import Sprints
import wx
import wx.grid		as gridlib

class Results( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(Results, self).__init__(parent, id, style=wx.BORDER_SUNKEN)
		self.SetBackgroundColour( wx.WHITE )

		self.hbs = wx.BoxSizer(wx.HORIZONTAL)

		self.gridResults = gridlib.Grid( self )
		labels = ['Sprint Points\nSubtotal', 'Lap Points\nSubtotal', 'Laps\n+/-', 'Num\nWins']
		self.gridResults.CreateGrid( 0, len(labels) )
		for i, lab in enumerate(labels):
			self.gridResults.SetColLabelValue( i, lab )
		self.gridResults.SetRowLabelSize( 0 )
		self.gridResults.SetDefaultCellAlignment( wx.ALIGN_RIGHT, wx.ALIGN_CENTRE )
		Utils.MakeGridReadOnly( self.gridResults )
		self.gridResults.EnableDragRowSize( False )
		self.gridResults.AutoSize()
		self.gridResults.Bind( gridlib.EVT_GRID_CELL_LEFT_CLICK, self.onClick )

		self.hbs.Add( self.gridResults, 1, wx.GROW|wx.ALL, border = 4 )
		
		self.SetSizer(self.hbs)
		self.hbs.SetSizeHints(self)

	def clear( self ):
		Utils.DeleteAllGridRows( self.gridResults )
		
	def onClick( self ):
		wx.CallAfter( Utils.commitPanes )
		
	def refresh( self ):
		self.clear()
		race = Model.race
		if not race:
			return
		riders = race.getRiders()
		self.gridResults.InsertRows( 0, len(riders), True )
		position = 1
		fields = ['sprintsTotal', 'lapsTotal', 'updown']
		# Only update the Num Wins column if required for the ranking.
		if race.rankBy == race.RankByDistancePointsNumWins:
			fields += ['numWins']
			self.gridResults.SetColLabelValue( len(fields)-1, 'Num\nWins' )
		else:
			self.gridResults.SetColLabelValue( len(fields), '' )
		for r, rider in enumerate(riders):
			for c, field in enumerate(fields):
				v = getattr(rider,field)
				s = unicode(v) if v != 0 else u''
				if s and field == 'updown' and s[0] != '-':
					s = '+' + s
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
