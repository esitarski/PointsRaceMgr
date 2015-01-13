import wx
import Model

class Notes( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(Notes, self).__init__(parent, id, style=wx.BORDER_SUNKEN)
		vs = wx.BoxSizer( wx.VERTICAL )
		vs.Add( wx.StaticText(self, label=u'Notes'), flag=wx.ALL, border=4 )
		self.notes = wx.TextCtrl(self, style=wx.TE_MULTILINE|wx.TE_DONTWRAP )
		vs.Add( self.notes, 1, flag=wx.LEFT|wx.RIGHT|wx.BOTTOM|wx.EXPAND, border=4 )
		self.SetSizer( vs )
		
	def refresh( self ):
		race = Model.race
		if not race:
			return
			
	def commis( self ):
		race = Model.race
		if not race:
			return
		race.notes = self.notes.GetValue()

if __name__ == '__main__':
	app = wx.App( False )
	mainWin = wx.Frame(None,title="PointsRaceMgr", size=(600,400))
	Model.setRace( Model.Race() )
	Model.getRace()._populate()
	notes = Notes(mainWin)
	notes.refresh()
	mainWin.Show()
	app.MainLoop()