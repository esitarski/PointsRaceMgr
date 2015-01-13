import wx
import Model

class Notes( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(Notes, self).__init__(parent, id, style=wx.BORDER_SUNKEN)
		vs = wx.BoxSizer( wx.VERTICAL )
		self.notes = wx.TextCtrl(self, style=wx.TE_MULTILINE|wx.TE_DONTWRAP )
		vs.Add( self.notes, 1, flag=wx.ALL|wx.EXPAND, border=4 )
		self.SetSizer( vs )
		
	def refresh( self ):
		race = Model.race
		if not race:
			return
		self.notes.SetValue( race.notes )
			
	def commit( self ):
		race = Model.race
		if not race:
			return
		notes = self.notes.GetValue()
		if notes != race.notes:
			race.notes = notes
			race.setChanged()
		
class NotesDialog( wx.Dialog ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(NotesDialog, self).__init__(parent, id, style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP | wx.RESIZE_BORDER, size=(400,550), title=u'Race Notes')
		vs = wx.BoxSizer( wx.VERTICAL )
		self.notesPanel = Notes( self )
		vs.Add( self.notesPanel, 1, flag=wx.EXPAND )
		hs = wx.BoxSizer( wx.HORIZONTAL )
		hs.Add( wx.StaticText(self, label=u'Reopen with CTRL+n or from File menu'), flag=wx.ALL|wx.ALIGN_CENTRE_VERTICAL, border = 4 )
		hs.AddStretchSpacer()
		self.okButton = wx.Button( self, id=wx.ID_OK )
		hs.Add( self.okButton, flag=wx.ALL, border=4 )
		vs.Add( hs, flag=wx.EXPAND )
		self.SetSizer( vs )
	
	def onOK( self, event ):
		self.commit()
		wx.CallAfter( self.Show, False )
	
	def refresh( self ):
		self.notesPanel.refresh()
		
	def commit( self ):
		self.notesPanel.commit()

if __name__ == '__main__':
	app = wx.App( False )
	mainWin = wx.Frame(None,title="PointsRaceMgr", size=(600,400))
	Model.setRace( Model.Race() )
	Model.getRace()._populate()
	notes = Notes(mainWin)
	notes.refresh()
	mainWin.Show()
	app.MainLoop()