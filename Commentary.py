import wx
import Model
import Utils

class Commentary( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY ):
		super(Commentary, self).__init__(parent, id, style=wx.BORDER_SUNKEN)
		
		self.SetDoubleBuffered(True)
		self.SetBackgroundColour( wx.WHITE )

		self.hbs = wx.BoxSizer(wx.HORIZONTAL)

		self.text = wx.TextCtrl( self, style=wx.TE_MULTILINE|wx.TE_READONLY )
		self.hbs.Add( self.text, 1, wx.EXPAND )
		self.SetSizer(self.hbs)

	def refresh( self ):
		race = Model.race
		riderInfo = {info.bib:info for info in race.riderInfo} if race else {}

		def infoLines( bibs, pointsForPlace=None ):
			lines = []
			pfpText = u''
			for place, bib in enumerate(bibs,1):
				ri = riderInfo.get( bib, None )
				if pointsForPlace and pointsForPlace.get(place, 0):
					pfpText = ' ({:+d} pts)'.format(pointsForPlace.get(place, 0))
				else:
					pfpText = u''
				if ri is not None:
					lines.append( u'    {}{}.  {}: {} {}, {}'.format( place, pfpText, bib, ri.first_name, ri.last_name, ri.team ) )
				else:
					lines.append( u'    {}{}.  {}'.format(place, pfpText, bib) )
			return lines
		
		RaceEvent = Model.RaceEvent
		
		lines = []
		self.sprintCount = 0
		for e in race.events:
			if e.eventType == RaceEvent.Sprint:
				self.sprintCount += 1
				lines.append( u'Sprint {} Result:'.format(self.sprintCount) )
				lines.extend( infoLines(e.bibs[:len(race.pointsForPlace)], race.pointsForPlace) )
			elif e.eventType == RaceEvent.LapUp:
				lines.append( u'Gained a Lap:' )
				lines.extend( infoLines(e.bibs, {p:race.pointsForLapping for p in xrange(1,len(e.bibs)+1)}) )
			elif e.eventType == RaceEvent.LapDown:
				lines.append( u'Lost a Lap:' )
				lines.extend( infoLines(e.bibs, {p:-race.pointsForLapping for p in xrange(1,len(e.bibs)+1)}) )
			elif e.eventType == RaceEvent.Finish:
				lines.append( u'Finish:' )
				self.sprintCount += 1
				if race.doublePointsForLastSprint:
					pointsForPlace = {p:v*2 for p,v in race.pointsForPlace.iteritems()}
				else:
					pointsForPlace = race.pointsForPlace
				lines.extend( infoLines(e.bibs, pointsForPlace) )
			elif e.eventType == RaceEvent.DNF:
				lines.append( u'DNF (Did Not Finish):' )
				lines.extend( infoLines(e.bibs) )
			elif e.eventType == RaceEvent.DNS:
				lines.append( u'DNS (Did Not Start):' )
				lines.extend( infoLines(e.bibs) )
			elif e.eventType == RaceEvent.PUL:
				lines.append( u'PUL (Pulled by Race Officials):' )
				lines.extend( infoLines(e.bibs) )
			elif e.eventType == RaceEvent.DSQ:
				lines.append( u'DSQ (Disqualified)' )
				lines.extend( infoLines(e.bibs) )
			lines.append( u'' )

		self.text.SetValue( u'\n'.join(lines) )

	def commit( self ):
		pass
	
if __name__ == '__main__':
	app = wx.App( False )
	mainWin = wx.Frame(None,title="Commentary", size=(600,400))
	Model.newRace()
	Model.race._populate()
	rd = Commentary(mainWin)
	rd.refresh()
	mainWin.Show()
	app.MainLoop()
