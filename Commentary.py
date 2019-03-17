import wx
import cgi
import sys
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
		
	def getText( self ):
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
					lines.append( u'    {}.{}  {}: {} {}, {}'.format( place, pfpText, bib, ri.first_name, ri.last_name, ri.team ) )
				else:
					lines.append( u'    {}.{}  {}'.format(place, pfpText, bib) )
			return lines
		
		RaceEvent = Model.RaceEvent
		
		pointsForGainedLap = {p:race.pointsForLapping for p in range(1,201)}
		pointsForLostLap   = {p:-race.pointsForLapping for p in range(1,201)}
		
		lines = []
		self.sprintCount = 0
		for e in race.events:
			if e.eventType == RaceEvent.Sprint:
				self.sprintCount += 1
				lines.append( u'Sprint {} Result:'.format(self.sprintCount) )
				lines.extend( infoLines(e.bibs[:len(race.pointsForPlace)], race.pointsForPlace) )
			elif e.eventType == RaceEvent.LapUp:
				lines.append( u'Gained a Lap:' )
				lines.extend( infoLines(e.bibs, pointsForGainedLap) )
			elif e.eventType == RaceEvent.LapDown:
				lines.append( u'Lost a Lap:' )
				lines.extend( infoLines(e.bibs, pointsForLostLap) )
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
		
		return u'\n'.join(lines)

	def toHtml( self, html ):
		text = self.getText().replace(u'.', u'')
		if not text:
			return u''
		lines = []
		inList = False
		html.write( u'<dl>' )
		for line in text.split(u'\n'):
			if not line:
				continue
			if line[:1] != u' ':
				if inList:
					html.write(u'</ol>\n')
					html.write(u'</dd>\n')
					inList = False
				html.write( u'<dd>\n' )
				html.write( cgi.escape(line) )
				html.write( u'<ol>' )
				inList = True
				continue
			line = line.strip()
			html.write( u'<li>{}</li>\n'.format(line.split(' ',1)[1].strip()) )
		html.write(u'</ol>\n')
		html.write(u'</dd>\n')
		html.write(u'</dl>\n')
		
	def refresh( self ):
		self.text.SetValue( self.getText() )

	def commit( self ):
		pass
	
if __name__ == '__main__':
	app = wx.App( False )
	mainWin = wx.Frame(None,title="Commentary", size=(600,400))
	Model.newRace()
	Model.race._populate()
	rd = Commentary(mainWin)
	rd.refresh()
	rd.toHtml( sys.stdout )
	mainWin.Show()
	app.MainLoop()
