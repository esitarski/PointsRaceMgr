import wx
import wx.grid		as gridlib
from ReorderableGrid import ReorderableGrid
import Model
import Utils
import re
import copy
import operator

class EventListGrid( ReorderableGrid ):
	def OnRearrangeEnd(self, evt):
		if self._didCopy and Utils.getMainWin():
			wx.CallAfter( Utils.getMainWin().eventList.commitReorder )
		return super(EventListGrid, self).OnRearrangeEnd(evt)

class EventDialog( wx.Dialog ):
	def __init__( self, parent, title="Edit Race Event" ):
		super( EventDialog, self ).__init__( parent, wx.ID_ANY, title=title )
		
		fgs = wx.FlexGridSizer( 2, 2, 4, 4 )
		fgs.AddGrowableCol( 1, 1 )
		
		label = wx.StaticText(self, label=u'Bibs:')
		label.SetFont( Utils.BigFont() )
		fgs.Add( label, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT )
		self.bibs = wx.TextCtrl( self, size=(500,-1) )
		self.bibs.SetFont( Utils.BigFont() )
		fgs.Add( self.bibs, 1, wx.EXPAND )
		
		label = wx.StaticText(self, label=u'Type:')
		label.SetFont( Utils.BigFont() )
		fgs.Add( label, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALIGN_RIGHT )
		choices = [('Sprint' if n == 'Sp' else n) for n, v in Model.RaceEvent.Events]
		self.eventType = wx.Choice( self, choices=choices )
		self.eventType.SetFont( Utils.BigFont() )
		fgs.Add( self.eventType, 0 )
		
		vs = wx.BoxSizer( wx.VERTICAL )
		
		borderSizer = wx.BoxSizer( wx.HORIZONTAL )
		borderSizer.Add( fgs, 1, wx.EXPAND|wx.ALL, border=8)
		vs.Add( borderSizer, 1, wx.EXPAND )
		
		hs = wx.BoxSizer( wx.HORIZONTAL )
		self.ok = wx.Button( self, wx.ID_OK )
		self.ok.SetDefault()
		self.cancel = wx.Button( self, wx.ID_CANCEL )
		hs.Add( self.ok, flag=wx.ALL, border=16 )
		hs.Add( self.cancel, flag=wx.ALL, border=16 )
		
		vs.Add( hs )
		self.SetSizerAndFit( vs )
	
	def refresh( self, event=None ):
		self.e = event or self.e
		self.eventType.SetSelection( next(j for j, (n,v) in enumerate(Model.RaceEvent.Events) if v == self.e.eventType) )
		self.bibs.SetValue( value=u','.join(u'{}'.format(b) for b in event.bibs) )
		self.bibs.SetFocus()
		
	def commit( self ):
		changed = False
		
		eventType = Model.RaceEvent.Events[self.eventType.GetSelection()][1]
		if eventType != self.e.eventType:
			self.e.eventType = eventType
			changed = True
		
		bibs = Model.RaceEvent.getCleanBibs( self.bibs.GetValue() )
		if bibs != self.e.bibs:
			self.e.bibs = bibs
			changed = True
		
		return changed

class EventPopupMenu( wx.Menu ):
	def __init__(self, row=None ):
		self.row = row
		super(EventPopupMenu, self).__init__()
		
		mmi = wx.MenuItem(self, wx.ID_DELETE)
		self.AppendItem(mmi)
		self.Bind(wx.EVT_MENU, self.OnDelete, mmi)

	def OnDelete(self, event):
		events = Model.race.events[:]
		del events[self.row]
		Model.race.setEvents( events )
		if Utils.getMainWin():
			Utils.getMainWin().refresh()

class EventList( wx.Panel ):
	def __init__( self, parent ):
		super(EventList, self).__init__(parent, wx.ID_ANY, style=wx.BORDER_SUNKEN)
		
		self.eventDialog = EventDialog( self )
		
		self.SetDoubleBuffered(True)
		self.SetBackgroundColour( wx.WHITE )

		self.hbs = wx.BoxSizer(wx.VERTICAL)
		
		self.newButton = wx.Button( self, label='New Race Event' )
		self.newButton.Bind( wx.EVT_BUTTON, self.onNewEvent )
		self.newButton.SetFont(wx.FontFromPixelSize( (0,20), wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL ))
		self.hbs.Add( self.newButton, 0, wx.ALL, border=4 )

		self.grid = EventListGrid( self )
		self.grid.Bind( gridlib.EVT_GRID_CELL_LEFT_CLICK, self.onLeftClick )
		self.grid.Bind( gridlib.EVT_GRID_CELL_RIGHT_CLICK, self.onRightClick )

		self.grid.CreateGrid( 0, 5 )
		self.hbs.Add( self.grid, 1, wx.EXPAND )
		self.SetSizer(self.hbs)

	def onNewEvent( self, event ):
		self.eventDialog.refresh( Model.RaceEvent() )
		self.eventDialog.CentreOnScreen()
		if self.eventDialog.ShowModal() == wx.ID_OK:
			race = Model.race
			self.eventDialog.commit()
			race.setEvents( race.events + [self.eventDialog.e] )
			(Utils.getMainWin() if Utils.getMainWin() else self).refresh()
			self.grid.MakeCellVisible( len(race.events)-1, 0 )
			self.grid.ClearSelection()
			self.grid.SelectRow( len(race.events)-1 )
	
	def onLeftClick( self, event ):
		race = Model.race
		self.grid.ClearSelection()
		self.grid.SelectRow( event.GetRow() )
		self.eventDialog.refresh( copy.deepcopy(race.events[event.GetRow()]) )
		self.eventDialog.CentreOnScreen()
		if self.eventDialog.ShowModal() == wx.ID_OK and self.eventDialog.commit():
			race.events[event.GetRow()] = self.eventDialog.e
			race.setEvents( race.events )
			(Utils.getMainWin() if Utils.getMainWin() else self).refresh()
			self.grid.MakeCellVisible( event.GetRow(), 0 )
			self.grid.SelectRow( event.GetRow() )
	
	def onRightClick( self, event ):
		self.PopupMenu(EventPopupMenu(event.GetRow()), event.GetPosition())
	
	def refresh( self ):
		race = Model.race
		events = race.events if race else []
			
		headers = ['Event', 'Bibs',]
		
		self.grid.BeginBatch()
		Utils.AdjustGridSize( self.grid, len(events), len(headers) )
		
		for c, name in enumerate(headers):
			self.grid.SetColLabelValue( c, name )
			attr = gridlib.GridCellAttr()
			attr.SetReadOnly()
			attr.SetAlignment(wx.ALIGN_LEFT, wx.ALIGN_CENTRE)
			attr.SetFont( Utils.BigFont() )
			self.grid.SetColAttr( c, attr )
		
		sprintCount = 0
		for row, e in enumerate(events):
			name = e.eventTypeName
			if e.eventType == Model.RaceEvent.Sprint:
				sprintCount += 1
				name += '{}'.format(sprintCount)
			self.grid.SetCellValue( row, 0, name )
			self.grid.SetCellValue( row, 1, u','.join(u'{}'.format(b) for b in e.bibs) )
		
		self.grid.AutoSize()
		self.grid.EndBatch()
		
		self.Layout()

	def commit( self ):
		pass

	def commitReorder( self ):
		race = Model.race
		events = []
		for r in xrange(self.grid.GetNumberRows()):
			events.append( Model.RaceEvent(self.grid.GetCellValue(r, 0), self.grid.GetCellValue(r, 1)) )
		race.setEvents( events )
		Utils.refresh()
	
if __name__ == '__main__':
	app = wx.App( False )
	Utils.disable_stdout_buffering()
	mainWin = wx.Frame(None,title="EventList", size=(600,400))
	Model.setRace( Model.Race() )
	Model.getRace()._populate()
	rd = EventList(mainWin)
	rd.refresh()
	mainWin.Show()
	app.MainLoop()
