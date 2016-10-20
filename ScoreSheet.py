import wx
import wx.lib.masked.numctrl as NC
import  wx.lib.intctrl as IC
import wx.lib.agw.fourwaysplitter as FWS
from wx.lib.wordwrap import wordwrap
import sys
import os
import re
import datetime
import xlwt
import webbrowser
import cPickle as pickle
import subprocess
from optparse import OptionParser

import Utils
import Model
import Version
from Sprints import Sprints
from UpDown import UpDown
from Worksheet import Worksheet
from Results import Results
from Notes import NotesDialog

class ScoreSheet( wx.Panel ):
	def __init__( self, parent, id = wx.ID_ANY, size=(200,200) ):
		super( ScoreSheet, self ).__init__(parent, id, size=size)
		
		self.SetBackgroundColour( wx.WHITE )
		
		self.inRefresh = False
		
		#-----------------------------------------------------------------------
		self.vbs = wx.BoxSizer( wx.VERTICAL )
		
		self.gbs = wx.GridBagSizer( 4, 4 )
		
		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, label=u'Race Name:' )
		self.gbs.Add( label, pos=(0, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.TextCtrl( self, style=wx.TE_PROCESS_ENTER )
		ctrl.Bind(wx.EVT_TEXT, self.onChange)
		self.gbs.Add( ctrl, pos=(0, 1), span=(1,5), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND )
		self.nameLabel = label
		self.nameCtrl = ctrl
		
		hs = wx.BoxSizer( wx.HORIZONTAL )
		label = wx.StaticText( self, label=u'Date:' )
		hs.Add( label, flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.DatePickerCtrl( self, style = wx.DP_DROPDOWN | wx.DP_SHOWCENTURY, size=(132,-1) )
		ctrl.Bind( wx.EVT_DATE_CHANGED, self.onChange )
		hs.Add( ctrl, flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.dateLabel = label
		self.dateCtrl = ctrl
		
		label = wx.StaticText( self, label=u'Communiqu\u00E9:' )
		hs.Add( label, flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.TextCtrl( self, style=wx.TE_PROCESS_ENTER )
		ctrl.Bind(wx.EVT_TEXT, self.onChange)
		hs.Add( ctrl, flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.communiqueLabel = label
		self.communiqueCtrl = ctrl
		
		self.gbs.Add( hs, pos=(0, 6), span=(1, 3) )
		
		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, label=u'Category:' )
		self.gbs.Add( label, pos=(1, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.TextCtrl( self, style=wx.TE_PROCESS_ENTER )
		ctrl.Bind(wx.EVT_TEXT, self.onChange)
		self.gbs.Add( ctrl, pos=(1, 1), span=(1,5), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND )
		self.categoryLabel = label
		self.categoryCtrl = ctrl
		
		label = wx.StaticText( self, label=u'Rank By:', style = wx.ALIGN_RIGHT )
		self.gbs.Add( label, pos=(1, 6), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.Choice( self, choices=[
				u'Points then Finish Order',
				u'Laps Completed, Points then Finish Order',
				u'Laps Completed, Points, Num Wins then Finish Order'
			]
		)
		ctrl.SetSelection( 0 )
		self.Bind(wx.EVT_CHOICE, self.onRankByChange, ctrl)
		self.gbs.Add( ctrl, pos=(1, 7), span=(1, 2), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.rankByLabel = label
		self.rankByCtrl = ctrl

		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, label=u'Laps:' )
		self.gbs.Add( label, pos=(2, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = IC.IntCtrl( self, min=1, max=300, value=1, limited=True, style=wx.ALIGN_RIGHT, size=(32,-1) )
		ctrl.Bind(IC.EVT_INT, self.onLapsChange)
		self.gbs.Add( ctrl, pos=(2, 1), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.lapsLabel = label
		self.lapsCtrl = ctrl
		
		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, label=u'Start Laps:' )
		self.gbs.Add( label, pos=(2, 3), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = IC.IntCtrl( self, min=0, max=300, value=0, limited=True, style=wx.ALIGN_RIGHT, size=(32,-1) )
		ctrl.Bind(IC.EVT_INT, self.onStartLapsChange)
		self.gbs.Add( ctrl, pos=(2, 4), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.startLapsLabel = label
		self.startLapsCtrl = ctrl
		
		ctrl = wx.StaticText( self, -1, u'Distance: 10.0km' )
		self.gbs.Add( ctrl, pos=(2, 5), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL )
		self.distanceCtrl = ctrl

		label = wx.StaticText( self, label=u'Number of Sprints:', style = wx.ALIGN_RIGHT )
		self.gbs.Add( label, pos=(2, 6), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.StaticText( self )
		self.gbs.Add( ctrl, pos=(2, 7), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.numSprintsLabel = label
		self.numSprintsCtrl = ctrl

		label = wx.CheckBox( self, label=u'Snowball Points' )
		self.gbs.Add( label, pos=(2, 8), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = label
		ctrl.Bind( wx.EVT_CHECKBOX, self.onSnowballChange )
		self.snowballLabel = label
		self.snowballCtrl = ctrl

		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, label=u'Sprint Every:' )
		self.gbs.Add( label, pos=(3, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = IC.IntCtrl( self, min=1, max=300, value=1, limited=True, style=wx.ALIGN_RIGHT, size=(32,-1) )
		ctrl.Bind(IC.EVT_INT, self.onSprintEveryChange)
		self.gbs.Add( ctrl, pos=(3, 1), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		unitsLabel = wx.StaticText( self, -1, 'laps' )
		self.gbs.Add( unitsLabel, pos=(3, 2), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL )
		self.sprintEveryLabel = label
		self.sprintEveryCtrl = ctrl
		self.sprintEveryUnitsLabel = unitsLabel
		
		label = wx.StaticText( self, label=u'Course Length:' )
		self.gbs.Add( label, pos=(3, 3), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		
		ctrl = NC.NumCtrl( self, min = 0, integerWidth = 3, fractionWidth = 2, style=wx.ALIGN_RIGHT, size=(32,-1), useFixedWidthFont = False )
		ctrl.SetAllowNegative(False)
		ctrl.Bind(wx.EVT_TEXT, self.onCourseLengthChange)
		self.gbs.Add( ctrl, pos=(3, 4), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		
		unitCtrl = wx.Choice( self, choices=[u'm', u'km'] )
		unitCtrl.SetSelection( 0 )
		self.Bind(wx.EVT_CHOICE, self.onCourseLengthUnitChange, unitCtrl)
		self.gbs.Add( unitCtrl, pos=(3, 5), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL )
		self.courseLengthLabel = label
		self.courseLengthCtrl = ctrl
		self.courseLengthUnitCtrl = unitCtrl

		label = wx.StaticText( self, label=u'Lap Gain/Lose Points:' )
		self.gbs.Add( label, pos=(3, 6), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = IC.IntCtrl( self, min=0, max=100, value=0, limited=True, style=wx.ALIGN_RIGHT, size=(32,-1) )
		ctrl.Bind(IC.EVT_INT, self.onChange)
		self.gbs.Add( ctrl, pos=(3, 7), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.pointsForLappingLabel = label
		self.pointsForLappingCtrl = ctrl

		label = wx.CheckBox( self, label=u'Double Points for Last Sprint' )
		self.gbs.Add( label, pos=(3, 8), flag=wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = label
		ctrl.Bind( wx.EVT_CHECKBOX, self.onDoublePointsForLastSprintChange )
		self.doublePointsForLastSprintLabel = label
		self.doublePointsForLastSprintCtrl = ctrl
		
		branding = wx.HyperlinkCtrl( self, id=wx.ID_ANY, label=u"Powered by CrossMgr", url=u"http://www.sites.google.com/site/crossmgrsoftware/" )
		branding.SetBackgroundColour( wx.WHITE )
		self.gbs.Add( branding, pos=(3, 9), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )

		#-----------------------------------------------------------------------------------------------------------
		self.vbs.Add( self.gbs, flag = wx.ALL, border = 4 )
		
		# Manage the display with a 4-way splitter.
		sty = wx.SP_LIVE_UPDATE | wx.SP_3DBORDER
		self.splitter = FWS.FourWaySplitter( self, agwStyle=sty )
		self.splitter.SetHSplit( 5800 )
		self.splitter.SetVSplit( 4000 )
		self.splitter.SetBackgroundColour( wx.Colour(176,196,222) )		# Light Steel Blue

		self.sprints = Sprints( self.splitter )
		self.splitter.AppendWindow( self.sprints )
		
		self.updown = UpDown( self.splitter )
		self.splitter.AppendWindow( self.updown )
		
		self.worksheet = Worksheet( self.splitter )
		self.splitter.AppendWindow( self.worksheet )
		
		self.results = Results( self.splitter )
		self.splitter.AppendWindow( self.results )
		
		self.vbs.Add( self.splitter, 1, wx.GROW )
		self.SetSizer( self.vbs )
		
		self.ConfigurePointsRace()
		self.refresh()
		Utils.setMainWin( self )
	
	def showNotes( self ):
		self.notesDialog.refresh()
		width, height = self.notesDialog.GetSizeTuple()
		screenWidth, screenHeight = wx.GetDisplaySize()
		self.notesDialog.MoveXY( screenWidth-width, screenHeight-height-40 )
		self.notesDialog.Show( True )
	
	def getDirName( self ):
		return Utils.getDirName()

	#--------------------------------------------------------------------------------------------

	def configurePointsRaceOptions( self ):
		self.rankByCtrl.SetSelection( Model.Race.RankByPoints )
		self.snowballCtrl.SetValue( False )
		self.doublePointsForLastSprintCtrl.SetValue( True )
		self.startLapsCtrl.SetValue( 0 )
		self.pointsForLappingCtrl.SetValue( 20 )
		self.lapsCtrl.SetValue( 120 )
		self.sprintEveryCtrl.SetValue( 10 )
		
		self.commit()
		self.refresh()

	def ConfigurePointsRace( self ):
		Model.race.pointsForPlace = {
			1 : 5,
			2 : 3,
			3 : 2,
			4 : 1,
			5 : 0
		}
		self.configurePointsRaceOptions()
		
	def ConfigureMadison( self ):
		self.ConfigurePointsRace()
	
	def ConfigurePointALapRace( self ):
		Model.race.pointsForPlace = {
			1 : 1,
			2 : 0,
			3 : -1,
			4 : -1,
			5 : -1
		}
		self.configurePointsRaceOptions()
		self.doublePointsForLastSprintCtrl.SetValue( False )
		self.pointsForLappingCtrl.SetValue( 0 )
		self.commit()
		self.refresh()
	
	def ConfigureTempoRace( self ):
		Model.race.pointsForPlace = {
			1 : 2,
			2 : 1,
			3 : 0,
			4 : -1,
			5 : -1
		}
		self.configurePointsRaceOptions()
		self.doublePointsForLastSprintCtrl.SetValue( False )
		self.pointsForLappingCtrl.SetValue( 4 )
		self.lapsCtrl.SetValue( 4*10 )
		self.startLapsCtrl.SetValue( 5 )
		self.sprintEveryCtrl.SetValue( 1 )
		self.commit()
		self.refresh()
	
	def ConfigureSnowballRace( self ):
		Model.race.pointsForPlace = {
			1 : 1,
			2 : 0,
			3 : -1,
			4 : -1,
			5 : -1
		}
		self.snowballCtrl.SetValue( True )
		self.rankByCtrl.SetSelection( Model.Race.RankByPoints )
		self.commit()
		self.refresh()
		
	def ConfigureCriteriumRace( self ):
		self.ConfigurePointsRace()
		self.rankByCtrl.SetSelection( Model.Race.RankByLapsPointsNumWins )
		self.pointsForLappingCtrl.SetValue( 0 )
		self.doublePointsForLastSprintCtrl.SetValue( False )
		self.commit()
		self.refresh()

	def onChange( self, event ):
		self.commit()

	def onCourseLengthUnitChange( self, event ):
		race = Model.race
		race.courseLengthUnit = self.courseLengthUnitCtrl.GetCurrentSelection()
		self.updateDependentFields()
	
	def onRankByChange( self, event ):
		race = Model.race
		race.rankBy = self.rankByCtrl.GetCurrentSelection()
		self.updateDependentFields()
		self.refreshResults()
	
	def onCourseLengthChange( self, event ):
		race = Model.race
		race.courseLength = self.courseLengthCtrl.GetValue()
		self.updateDependentFields()
		
	def onLapsChange( self, event ):
		race = Model.race
		race.laps = self.lapsCtrl.GetValue()
		self.updateDependentFields()
		self.refreshResults()
		
	def onStartLapsChange( self, event ):
		race = Model.race
		race.startLaps = self.startLapsCtrl.GetValue()
		self.updateDependentFields()
		self.refreshResults()
		
	def onSprintEveryChange( self, event ):
		race = Model.race
		race.sprintEvery = self.sprintEveryCtrl.GetValue()
		self.updateDependentFields()
		self.refreshResults()
			
	def onSnowballChange( self, event ):
		race = Model.race
		race.snowball = self.snowballCtrl.GetValue()
		self.refreshResults()
		
	def onDoublePointsForLastSprintChange( self, event ):
		race = Model.race
		race.doublePointsForLastSprint = self.doublePointsForLastSprintCtrl.GetValue()
		self.updateDependentFields()
		self.refreshResults()
	
	#--------------------------------------------------------------------------------------

	def updateDependentFields( self ):
		race = Model.race
		if not race:
			return

		self.distanceCtrl.SetLabel( u'Distance: {}km'.format(race.getDistanceStr()) )
		self.numSprintsCtrl.SetLabel( unicode(race.getNumSprints()) )
		self.sprints.updateShading()
		self.gbs.Layout()

	def commit( self ):
		if self.inRefresh:	# Don't commit anything while we are refreshing.
			return
		race = Model.race
		if not race:
			return

		for field in [	'name', 'category', 'communique', 'laps', 'startLaps', 'sprintEvery', 'courseLength',
						'doublePointsForLastSprint', 'pointsForLapping', 'snowball']:
			v = getattr(self, field + 'Ctrl').GetValue()
			race.setattr( field, v )
			
		for field in ['rankBy', 'courseLengthUnit']:
			v = getattr(self, field + 'Ctrl').GetCurrentSelection()
			race.setattr( field, v )
			
		for field in ['date']:
			dt = getattr(self, field + 'Ctrl').GetValue()
			v = datetime.date( dt.GetYear(), dt.GetMonth() + 1, dt.GetDay() )	# Adjust for 0-based month.
			race.setattr( field, v )
			
		self.updateDependentFields()
	
	def commitPanes( self ):
		self.commit()
		self.sprints.commit()
		self.updown.commit()
		self.refreshResults()
	
	def refreshResults( self ):
		self.worksheet.refresh()
		self.results.refresh()
	
	def refresh( self, pane = None ):
		self.inRefresh = True
		race = Model.race
		
		for field in [	'name', 'category', 'communique', 'laps', 'startLaps', 'sprintEvery', 'courseLength',
						'doublePointsForLastSprint', 'pointsForLapping', 'snowball']:
			getattr(self, field + 'Ctrl').SetValue( getattr(race, field) )
		
		for field in ['rankBy', 'courseLengthUnit']:
			getattr(self, field + 'Ctrl').SetSelection( getattr(race, field) )
			
		for field in ['date']:
			d = getattr( race, field )
			dt = wx.DateTime()
			dt.Set(d.day, d.month - 1, d.year, 0, 0, 0, 0)	# Adjust to 0-based month
			getattr(self, field + 'Ctrl').SetValue( dt )
					
		self.updateDependentFields()
				
		for p in [self.sprints, self.worksheet, self.updown, self.results]:
			if p != pane:
				p.refresh()
				
		self.inRefresh = False

if __name__ == '__main__':
	Utils.disable_stdout_buffering()
	app = wx.App( False )
	mainWin = wx.Frame(None,title="PointsRaceMan", size=(1000,600))

	dataDir = Utils.getHomeDir()
	os.chdir( dataDir )
	redirectFileName = os.path.join(dataDir, 'PointsRaceMgr.log')
	
	scoreSheet = ScoreSheet( mainWin )
	mainWin.Show()

	# Start processing events.
	app.MainLoop()
	

