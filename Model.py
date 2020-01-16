import re
import random
import operator
import datetime
import sys
from collections import namedtuple

#------------------------------------------------------------------------------
# Define a global current race.
race = None

def getRace():
	global race
	return race

def newRace():
	global race
	race = Race()
	return race

def setRace( r ):
	global race
	race = r

def fixBibsNML( bibs, bibsNML, isFinish=False ):
	bibsNMLSet = set( bibsNML )
	bibsNew = [b for b in bibs if b not in bibsNMLSet]
	if isFinish:
		bibsNew.extend( bibsNML )
	return bibsNew
	
#------------------------------------------------------------------------------------------------------------------
class RiderInfo(object):
	FieldNames  = ('bib', 'existing_points', 'last_name', 'first_name', 'team', 'team_code', 'license', 'nation_code', 'uci_id')
	HeaderNames = ('Bib', 'Existing\nPoints', 'Last Name', 'First Name', 'Team', 'Team Code', 'License', 'Nat Code', 'UCI ID')

	def __init__( self, bib, last_name=u'', first_name=u'', team=u'', team_code=u'', license=u'', uci_id=u'', nation_code=u'', existing_points=0.0 ):
		self.bib = int(bib)
		self.last_name = last_name
		self.first_name = first_name
		self.team = team
		self.team_code = team_code
		self.license = license
		self.uci_id = uci_id
		self.nation_code = nation_code
		self.existing_points = float(existing_points or '0.0')

	def __eq__( self, ri ):
		return all(getattr(self,a) == getattr(ri,a) for a in self.FieldNames)
		
	def __repr__( self ):
		return 'RiderInfo({})'.format(u','.join('{}="{}"'.format(a,getattr(self,a)) for a in self.FieldNames))

class Rider(object):
	Finisher  = 0
	DNF       = 1
	PUL   	  = 2
	DNS       = 3
	DSQ 	  = 4
	statusNames = ['Finisher', 'DNF', 'PUL', 'DNS', 'DSQ']
	statusSortSeq = { 'Finisher':1,	Finisher:1,
					  'PUL':2,		PUL:2,
					  'DNF':3,		DNF:3,
					  'DNS':4,		DNS:4,
					  'DSQ':5,		DSQ:5,
	}
	
	def __init__( self, num ):
		self.num = num
		self.reset()
		
	def reset( self ):
		self.pointsTotal = 0
		self.sprintPlacePoints = {}
		self.sprintsTotal = 0
		self.lapsTotal = 0
		self.updown = 0
		self.numWins = 0
		self.existingPoints = 0
		self.finishOrder = 1000
		self.status = Rider.Finisher
		
	def addSprintResult( self, sprint, place ):
		points = race.getSprintPoints(sprint, place)
		if points > 0:
			self.pointsTotal += points
			self.sprintsTotal += points
			self.sprintPlacePoints[sprint] = (place, points)
		
		if place == 1:
			self.numWins += 1
	
	def addFinishOrder( self, finishOrder ):
		self.finishOrder = finishOrder
	
	def addUpDown( self, updown ):
		assert updown == -1 or updown == 1
		self.updown += updown
		self.pointsTotal += race.pointsForLapping * updown
		self.lapsTotal += race.pointsForLapping * updown

	def addExistingPoints( self, existingPoints ):
		if existingPoints == int(existingPoints):
			existingPoints = int(existingPoints)
		self.existingPoints += existingPoints
		self.pointsTotal += existingPoints
		
	def getKey( self ):
		if   race.rankBy == race.RankByPoints:
			return (Rider.statusSortSeq[self.status], -self.pointsTotal, self.finishOrder, self.num)
		elif race.rankBy == race.RankByLapsPoints:
			return (Rider.statusSortSeq[self.status], -self.updown, -self.pointsTotal, self.finishOrder, self.num)
		else:	# race.RankByLapsPointsNumWins
			return (Rider.statusSortSeq[self.status], -self.updown, -self.pointsTotal, -self.numWins, self.finishOrder, self.num)

	def tiedWith( s, r ):
		return s.getKey()[:-1] == r.getKey()[:-1]
	
	def __repr__( self ):
		return u"Rider( {}, {}, {}, {}, {} )".format(
			self.num, self.pointsTotal, self.updown, self.numWins,
			self.statusNames[self.status]
		)
		
class GetRank( object ):
	def __init__( self ):
		self.rankLast, self.rrLast = None, None
	
	def __call__( self, rank, rr ):
		if rr.status != Rider.Finisher:
			return Rider.statusNames[rr.status]
		elif self.rrLast and self.rrLast.tiedWith(rr):
			return u'{}'.format(self.rankLast)
		else:
			self.rankLast, self.rrLast = rank, rr
			return u'{}'.format(rank)
		
class RaceEvent(object):
	DNS, DNF, PUL, DSQ, LapUp, LapDown, Sprint, Finish, Break, Chase, OTB, NML = tuple( range(12) )
	
	Events = (
		('Sp', Sprint),
		('+ Lap', LapUp),
		('- Lap', LapDown),
		('DNF', DNF),
		('Finish', Finish),
		('PUL', PUL),
		('DNS', DNS),
		('DSQ', DSQ),
	)

	States = (
		('Break', Break),
		('Chase', Chase),
		('OTB', OTB),
		('NML', NML),
	)
	EventName = {v:n for n,v in Events}
	EventName.update( {v:n for n,v in States} )
	
	@staticmethod
	def getCleanBibs( bibs ):
		if not isinstance(bibs, list):
			try:
				bibs = [int(f) for f in re.sub(r'[^\d]', ' ', bibs).split()]
			except:
				bibs = []
				
		seen = set()
		nonDupBibs = []
		for b in bibs:
			if b > 0 and b not in seen:
				seen.add( b )
				nonDupBibs.append( b )

		return nonDupBibs
	
	def __init__( self, eventType=Sprint, bibs=[] ):
		if not isinstance(eventType, int):
			for n, v in self.Events:
				if eventType.startswith(n):
					eventType = v
					break
			else:
				eventType = self.Sprint
		
		self.eventType = eventType		
		self.bibs = RaceEvent.getCleanBibs( bibs )

	def isState( self ):
		return self.eventType >= self.Break
	
	@property
	def eventTypeName( self ):
		return self.EventName[self.eventType]
			
	def __eq__( s, t ):
		return s.eventType == t.eventType and s.bibs == t.bibs
		
	def __repr__( self ):
		return 'RaceEvent( eventType={}, bibs=[{}] )'.format(self.eventType, ','.join('{}'.format(b) for b in self.bibs))
		
class Race(object):
	RankByPoints = 0
	RankByLapsPoints = 1
	RankByLapsPointsNumWins = 2

	pointsForPlaceDefault = {
		1 : 5,
		2 : 3,
		3 : 2,
		4 : 1,
	}
	startLaps = 0
	
	def __init__( self ):
		self.reset()

	def reset( self ):
		self.name = '<RaceName>'
		self.communique = u''
		self.category = '<Category>'
		self.notes = u''
		self.sprintEvery = 10
		self.courseLength = 250.0
		self.courseLengthUnit = 0	# 0 = Meters, 1 = Km
		self.laps = 160
		self.rankBy = Race.RankByPoints		# 0 = Points only, 1 = Distance, then points, 2 = 
		self.date = datetime.date.today()
		self.pointsForLapping = 20
		self.doublePointsForLastSprint = False
		self.snowball = False
		self.pointsForPlace = Race.pointsForPlaceDefault.copy()

		self.events = []
		self.riders = {}
		self.riderInfo = []
		
		self.sprintCount = 0

		self.isChangedFlag = True
	
	def getDistance( self ):	# Always return in km
		return self.courseLength * self.laps / (1000.0 if self.courseLengthUnit == 0 else 1.0)
	
	def newNext( self ):
		self.events = []
		self.riderInfo = []
		self.isChangedFlag = True
	
	def getDistanceStr( self ):
		d = self.courseLength * self.laps
		if d - int(d) < 0.001:
			return '{:,}'.format(int(d)) + ['m','km'][self.courseLengthUnit]
		else:
			return '{:,.2f}'.format(d) + ['m','km'][self.courseLengthUnit]
	
	def setattr( self, attr, v ):
		if getattr(self, attr, None) != v:
			setattr( self, attr, v )
			self.setChanged()
			return True
		else:
			return False
	
	def getNumSprints( self ):
		try:
			numSprints = max(0, self.laps - self.startLaps) // self.sprintEvery
		except:
			numSprints = 0
		return numSprints
	
	def getSprintCount( self ):
		Sprint = RaceEvent.Sprint
		return sum( 1 for e in self.events if e.eventType == Sprint )
	
	def getSprintLabel( self, sprint ):
		if self.doublePointsForLastSprint and sprint == self.getNumSprints():
			return u'Sp{} \u00D72'.format(sprint)
		return u'Sp{}'.format(sprint)
	
	def getMaxPlace( self ):
		maxPlace = 2
		for place, points in self.pointsForPlace.items():
			if points >= 0:
				maxPlace = max( maxPlace, place )
		return maxPlace
	
	def getRider( self, bib ):
		try:
			return self.riders[bib]
		except KeyError:
			self.riders[bib] = Rider( bib )
			return self.riders[bib]
			
	def getSprintPoints( self, sprint, place ):
		points = self.pointsForPlace.get(place,0)
		if self.snowball:
			points *= sprint
		if self.doublePointsForLastSprint and sprint == self.getNumSprints():
			points *= 2
		return points
	
	def processEvents( self ):
		self.riders = {}
		
		for info in self.riderInfo:
			self.getRider(info.bib).addExistingPoints( info.existing_points )
		
		numSprints = self.getNumSprints()
		self.sprintCount = 0
		for iEvent, e in enumerate(self.events):
			# Ensure the eventType matches the number of sprints.
			if e.eventType == RaceEvent.Finish and self.sprintCount != numSprints-1:
				e.eventType = RaceEvent.Sprint
			elif e.eventType == RaceEvent.Sprint and self.sprintCount == numSprints-1:
				e.eventType = RaceEvent.Finish
				
			if e.eventType == RaceEvent.Sprint:
				self.sprintCount += 1
				for place, b in enumerate(e.bibs, 1):
					self.getRider(b).addSprintResult(self.sprintCount, place)
			elif e.eventType == RaceEvent.LapUp:
				for b in e.bibs:
					self.getRider(b).addUpDown(1)
			elif e.eventType == RaceEvent.LapDown:
				for b in e.bibs:
					self.getRider(b).addUpDown(-1)
			elif e.eventType == RaceEvent.Finish:
				self.sprintCount += 1
				if iEvent != len(self.events)-1 and self.events[iEvent+1].eventType == RaceEvent.NML:
					bibs = fixBibsNML( e.bibs, self.events[iEvent+1].bibs, True )
				else:
					bibs = e.bibs
				for place, b in enumerate(bibs, 1):
					r = self.getRider(b)
					r.addSprintResult(self.sprintCount, place)
					r.addFinishOrder(place)
			elif e.eventType == RaceEvent.DNF:
				for b in e.bibs:
					self.getRider(b).status = Rider.DNF
			elif e.eventType == RaceEvent.DNS:
				for b in e.bibs:
					self.getRider(b).status = Rider.DNS
			elif e.eventType == RaceEvent.PUL:
				for b in e.bibs:
					self.getRider(b).status = Rider.PUL
			elif e.eventType == RaceEvent.DSQ:
				for b in e.bibs:
					self.getRider(b).status = Rider.DSQ
			
	def isChanged( self ):
		return self.isChangedFlag

	def setChanged( self, changed = True ):
		self.isChangedFlag = changed
		#traceback.print_stack()
	
	def getRiders( self ):
		self.processEvents()
		return sorted( self.riders.values(), key=operator.methodcaller('getKey') )
		
	def setRiderInfo( self, riderInfo ):
		self.isChangedFlag = (
			len(self.riderInfo) != len(riderInfo) or
			any(a != b for a, b in zip(self.riderInfo, riderInfo))
		)
		self.riderInfo = riderInfo

	def setEvents( self, events ):
		self.isChangedFlag = (
			len(self.events) != len(events) or
			any(e1 != e2 for e1, e2 in zip(self.events, events))
		)
		self.events = events
		
	def _populate( self ):
		self.reset()
		self.events.append( RaceEvent(RaceEvent.DNS, [41,42]) )
		random.seed( 0xed )
		bibs = list( range(10,34) )
		self.events.append( RaceEvent(RaceEvent.LapUp, bibs=[13]) )
		self.events.append( RaceEvent(RaceEvent.LapDown, bibs=[14]) )
		for lap in range(50,-1,-10):
			random.shuffle( bibs )
			self.events.append( RaceEvent(bibs=bibs[:5]) )
		self.events.append( RaceEvent(RaceEvent.DNF, [51,52]) )
		self.events.append( RaceEvent(RaceEvent.DSQ, [61,62]) )
		random.shuffle( bibs )
		self.events.append( RaceEvent(RaceEvent.Finish, bibs) )
		self.setChanged()

if __name__ == '__main__':
	r = newRace()
	r._populate()
