import random
import datetime
import sys

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

#------------------------------------------------------------------------------------------------------------------

class Rider(object):
	# Rider Status.
	Finisher  = 0
	DNF       = 1
	Pulled    = 2
	DNS       = 3
	DQ 		  = 4
	NP		  = 5
	statusNames = ['Finisher', 'DNF', 'PUL', 'DNS', 'DQ', 'NP']
	statusSortSeq = { 'Finisher':1,	Finisher:1,
					  'PUL':2,		Pulled:2,
					  'DNF':3,		DNF:3,
					  'DNS':4,		DNS:4,
					  'DQ':5,		DQ:5,
					  'NP':6,		NP:6,
	}

	def __init__( self, num ):
		self.num = num
		self.pointsTotal = 0
		self.sprintsTotal = 0
		self.lapsTotal = 0
		self.updown = 0
		self.numWins = 0
		self.status = Rider.Finisher
		self.finishOrder = 1000
		
	def addSprintResult( self, sprint, place ):
		points = race.getSprintPoints(sprint, place)
		if points > 0:
			self.pointsTotal += points
			self.sprintsTotal += points
		
		if place == 1:
			self.numWins += 1
	
	def addFinishOrder( self, finishOrder ):
		self.finishOrder = finishOrder
	
	def addUpDown( self, updown ):
		self.updown = updown
		self.pointsTotal += race.pointsForLapping * updown
		self.lapsTotal += race.pointsForLapping * updown

	def getKey( self ):
		if   race.rankBy == race.RankByPoints:
			return (Rider.statusSortSeq[self.status], -self.pointsTotal, self.finishOrder, self.num)
		elif race.rankBy == race.RankByDistancePoints:
			return (Rider.statusSortSeq[self.status], -self.updown, -self.pointsTotal, self.finishOrder, self.num)
		else:	# race.RankByDistancePointsNumWins
			return (Rider.statusSortSeq[self.status], -self.updown, -self.pointsTotal, -self.numWins, self.finishOrder, self.num)

	def tiedWith( s, r ):
		return cmp( s.getKey()[:-1], r.getKey()[:-1] ) == 0
	
	def __repr__( self ):
		return u"Rider( {}, {}, {}, {}, {} )".format(
			self.num, self.pointsTotal, self.updown, self.numWins,
			self.statusNames[self.status]
		)
		
class Race(object):
	RankByPoints = 0
	RankByDistancePoints = 1
	RankByDistancePointsNumWins = 2

	pointsForLapping = 20
	doublePointsForLastSprint = False
	snowball = False
	pointsForPlace = {
		1 : 5,
		2 : 3,
		3 : 2,
		4 : 1,
		5 : 0
	}
		
	sprintResults = {}		# Results from each sprint.
	updowns = {}			# Laps up/down
	finishOrder = {}		# Results from final sprint.
	riderStatus = {}		# Status (Finisher, DNF, etc.)

	def __init__( self ):
		self.reset()

	def reset( self ):
		self.name = '<RaceName>'
		self.category = '<Category>'
		self.sprintEvery = 10
		self.courseLength = 250.0
		self.courseLengthUnit = 0	# 0 = Meters, 1 = Km
		self.laps = 160
		self.rankBy = Race.RankByPoints		# 0 = Points only, 1 = Distance, then points, 2 = 
		self.date = datetime.date.today()
		self.pointsForLapping = 20
		self.doublePointsForLastSprint = False
		self.snowball = False
		self.pointsForPlace = Race.pointsForPlace.copy()

		self.sprintResults = {}
		self.updowns = {}
		self.finishOrder = {}
		self.riderStatus = {}

		self.isChangedFlag = True
	
	def getDistance( self ):	# Always return in km
		return self.courseLength * self.laps / (1000.0 if self.courseLengthUnit == 0 else 1.0)
	
	def getDistanceStr( self ):
		d = self.getDistance()
		if d - int(d) < 0.001:
			return '%d' % int(d)
		else:
			return '%.2f' % d
	
	def setattr( self, attr, v ):
		if getattr(self, attr, None) != v:
			setattr( self, attr, v )
			self.setChanged()
			return True
		else:
			return False
	
	def getNumSprints( self ):
		try:
			numSprints = self.laps // self.sprintEvery
		except:
			numSprints = 0
		return numSprints
	
	def getMaxPlace( self ):
		maxPlace = 2
		for place, points in self.pointsForPlace.iteritems():
			if points >= 0:
				maxPlace = max( maxPlace, place )
		return maxPlace
	
	def clearSprintResults( self ):
		self.sprintResults = {}
	
	def addSprintResult( self, sprint, num, place ):
		self.sprintResults[(sprint, place)] = num

	def addUpDown( self, num, updown ):
		self.updowns[num] = updown
		
	def clear( self ):
		self.sprintResults = {}
		self.updowns = {}
		self.riderStatus = {}
		
	def setSprintResults( self, sprintResults ):
		if self.sprintResults != sprintResults:
			self.sprintResults = sprintResults
			self.setChanged()
	
	def setUpDowns( self, updowns ):
		if self.updowns != updowns:
			self.updowns = updowns
			self.setChanged()
			
	def setFinishOrder( self, finishOrder ):
		if self.finishOrder != finishOrder:
			self.finishOrder = finishOrder
			self.setChanged()
			
	def setStatus( self, status ):
		status = dict( (n, s) for n, s in status.iteritems() if s != Rider.Finisher )
		if status != self.riderStatus:
			self.riderStatus = status
			self.setChanged()
	
	def setPoints( self, pointsForPlace = None, pointsForLapping = None ):
		if pointsForPlace is not None and self.pointsForPlace != pointsForPlace:
			self.pointsForPlace = pointsForPlace
			# Normalize the pointsForPlace if it contains negative entries.
			minEmpty = len(self.pointsForPlace)
			for place, points in self.pointsForPlace.iteritems():
				if points < 0:
					minEmpty = min( minEmpty, place )
			if minEmpty == 1:
				self.pointsForPlace[2] = 0
				minEmpty = 2
			for place, points in self.pointsForPlace.iteritems():
				if place > minEmpty:
					self.pointsForPlace[place] = -1
			self.setChanged()
			
		if pointsForLapping is not None and self.pointsForLapping != pointsForLapping:
			self.pointsForLapping = pointsForLapping
			self.setChanged()
	
	def isChanged( self ):
		return self.isChangedFlag

	def setChanged( self, changed = True ):
		self.isChangedFlag = changed
		#traceback.print_stack()
	
	def getUpDown( self ):
		riders = self.getRiders()
		riderIndex = { (r.num, i) for i, r in enumerate(riders) }
		ud = [(rider, updown) for rider, updown in self.updowns.iteritems()]
		ud.sort( lambda x, y: cmp(riderIndex[x[0]], riderIndex[y[0]]) if x[0] in riderIndex and y[0] in riderIndex \
								else cmp(x[0], y[0]) )
		return ud
	
	def getFinishOrder( self ):
		finishOrder = list( self.finishOrder.iteritems() )
		finishOrder.sort( key = lambda x: x[1] )
		return finishOrder
	
	def getStatus( self ):
		status = list( self.riderStatus.iteritems() )
		status.sort( key = lambda x: (Rider.statusSortSeq[x[1]], x[0]) )
		return status
			
	def getMaxSprints( self ):
		try:
			return max( sprint for sprint, place in self.sprintResults.iterkeys() )
		except ValueError:
			pass
		return 0
		
	def getSprintPoints( self, sprint, place ):
		return self.pointsForPlace.get(place, 0) * \
				(sprint if self.snowball else 1) * \
				(2 if self.doublePointsForLastSprint and sprint == self.getNumSprints() else 1)
	
	def getRiders( self ):
		riders = {}
		def getRider( num ):
			try:
				rider = riders[num]
			except KeyError:
				rider = Rider( num )
				rider.status = self.riderStatus.get( num, Rider.Finisher )
				riders[num] = rider
			return rider

		# Add the sprint results.
		for (sprint, place), num in self.sprintResults.iteritems():
			getRider(num).addSprintResult( sprint, place )
		
		# Add the finish order.
		for num, place in self.finishOrder.iteritems():
			getRider(num).addFinishOrder( place )
		
		# Add the up/down laps.
		for num, updown in self.updowns.iteritems():
			getRider(num).addUpDown( updown )
		
		# Add the rider status.
		for num, status in self.riderStatus.iteritems():
			getRider( num )
		
		ridersRet = [r for r in riders.itervalues() or r.updown != 0]
		ridersRet.sort( key = lambda x: x.getKey() )
		return ridersRet

	def __contains__( self, sprintNum ):
		return sprintNum in self.sprintResults

	def __getitem__( self, sprintNum ):
		return self.sprintResults[sprintNum]

	def _populate( self ):
		self.reset()
		random.seed( 1010101 )
		for s in xrange(15):
			place = {}
			for p in xrange(5,0,-1):
				while True:
					num = random.randint(1,10)
					if num not in place:
						break
				place[num] = p
			for num, place in place.iteritems():
				self.addSprintResult( s + 1, num, place )
		self.updowns[5] = -2
		self.updowns[6] = 1
		self.setChanged()

if __name__ == '__main__':
	r = newRace()
	r._populate()
