import wx
import wx.lib.masked.numctrl as NC
import  wx.lib.intctrl as IC
from wx.lib.wordwrap import wordwrap
import sys
import os
import re
import datetime
import copy
import bisect
import json
import xlwt
import webbrowser
import traceback
import cPickle as pickle
from optparse import OptionParser

import Utils
import Model
from Sprints import Sprints
from UpDown import UpDown
from Worksheet import Worksheet
from Results import Results
from ToExcelSheet import ToExcelSheet

from Version import AppVerName

try:
    from agw import fourwaysplitter as FWS
except ImportError: # if it's not there locally, try the wxPython lib.
    import wx.lib.agw.fourwaysplitter as FWS

try:
	import wx.lib.agw.advancedsplash as AS
	def ShowSplashScreen():
		bitmap = wx.Bitmap( os.path.join(Utils.getImageFolder(), 'TrackSprint.jpg'), wx.BITMAP_TYPE_JPEG )
		estyle = AS.AS_TIMEOUT | AS.AS_CENTER_ON_PARENT
		shadow = wx.WHITE
		try:
			frame = AS.AdvancedSplash(Utils.getMainWin(), bitmap=bitmap, timeout=3000,
									  extrastyle=estyle, shadowcolour=shadow)
		except:
			try:
				frame = AS.AdvancedSplash(Utils.getMainWin(), bitmap=bitmap, timeout=3000,
										  shadowcolour=shadow)
			except:
				pass
								  
except ImportError:
	def ShowSplashScreen():
		pass

class MainWin( wx.Frame ):
	def __init__( self, parent, id = wx.ID_ANY, title='', size=(200,200) ):
		wx.Frame.__init__(self, parent, id, title, size=size)

		self.SetBackgroundColour( wx.WHITE )
		
		# Add code to configure file history.
		self.filehistory = wx.FileHistory(8)
		self.config = wx.Config(appName="PointsRaceMgr",
								vendorName="Edward.Sitarski@gmail.com",
								style=wx.CONFIG_USE_LOCAL_FILE)
		self.filehistory.Load(self.config)
		
		self.fileName = None
		self.inRefresh = False	# Flag to indicate we are doing a refresh.

		Utils.setMainWin( self )
		
		# Configure the main menu.
		self.menuBar = wx.MenuBar(wx.MB_DOCKABLE)

		#-----------------------------------------------------------------------
		self.fileMenu = wx.Menu()

		self.fileMenu.Append( wx.ID_NEW , "&New...", "Create a new race" )
		self.Bind(wx.EVT_MENU, self.menuNew, id=wx.ID_NEW )

		self.fileMenu.Append( wx.ID_OPEN , "&Open...", "Open a race" )
		self.Bind(wx.EVT_MENU, self.menuOpen, id=wx.ID_OPEN )

		self.fileMenu.Append( wx.ID_SAVE , "&Save\tCtrl+S", "Save a race" )
		self.Bind(wx.EVT_MENU, self.menuSave, id=wx.ID_SAVE )

		self.fileMenu.Append( wx.ID_SAVEAS , "Save &As...", "Save a race under a different name" )
		self.Bind(wx.EVT_MENU, self.menuSaveAs, id=wx.ID_SAVEAS )

		self.fileMenu.AppendSeparator()
		
		idCur = wx.NewId()
		idExportToExcel = idCur
		self.fileMenu.Append( idCur , "&Export to Excel...\tCtrl+E", "Export as an Excel Spreadsheet" )
		self.Bind(wx.EVT_MENU, self.menuExportToExcel, id=idCur )

		self.fileMenu.AppendSeparator()
		
		recent = wx.Menu()
		self.fileMenu.AppendMenu(wx.ID_ANY, "&Recent Files", recent)
		self.filehistory.UseMenu( recent )
		self.filehistory.AddFilesToMenu()
		
		self.fileMenu.AppendSeparator()

		self.fileMenu.Append( wx.ID_EXIT , "E&xit", "Exit PointsRaceMgr" )
		self.Bind(wx.EVT_MENU, self.menuExit, id=wx.ID_EXIT )
		
		self.Bind(wx.EVT_MENU_RANGE, self.menuFileHistory, id=wx.ID_FILE1, id2=wx.ID_FILE9)
		
		self.menuBar.Append( self.fileMenu, "&File" )

		#-----------------------------------------------------------------------
		self.vbs = wx.BoxSizer( wx.VERTICAL )
		
		self.gbs = wx.GridBagSizer( 4, 4 )
		
		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, -1, 'Race Name:' )
		self.gbs.Add( label, pos=(0, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.TextCtrl( self, -1, style=wx.TE_PROCESS_ENTER )
		ctrl.Bind(wx.EVT_TEXT_ENTER, self.onChange)
		self.gbs.Add( ctrl, pos=(0, 1), span=(1,5), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND )
		self.nameLabel = label
		self.nameCtrl = ctrl
		
		label = wx.StaticText( self, -1, 'Date:' )
		self.gbs.Add( label, pos=(0, 6), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.DatePickerCtrl( self, style = wx.DP_DROPDOWN | wx.DP_SHOWCENTURY )
		ctrl.Bind( wx.EVT_DATE_CHANGED, self.onChange )
		self.gbs.Add( ctrl, pos=(0, 7), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.dateLabel = label
		self.dateCtrl = ctrl

		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, -1, 'Category:' )
		self.gbs.Add( label, pos=(1, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.TextCtrl( self, -1, style=wx.TE_PROCESS_ENTER )
		ctrl.Bind(wx.EVT_TEXT_ENTER, self.onChange)
		self.gbs.Add( ctrl, pos=(1, 1), span=(1,5), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND )
		self.categoryLabel = label
		self.categoryCtrl = ctrl
		
		label = wx.StaticText( self, -1, 'Rank By:', style = wx.ALIGN_RIGHT )
		self.gbs.Add( label, pos=(1, 6), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.Choice( self, choices=['Points Only', 'Distance, then Points'] )
		ctrl.SetSelection( 0 )
		self.Bind(wx.EVT_CHOICE, self.onRankByChange, ctrl)
		self.gbs.Add( ctrl, pos=(1, 7), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.rankByLabel = label
		self.rankByCtrl = ctrl

		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, -1, 'Laps:' )
		self.gbs.Add( label, pos=(2, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = IC.IntCtrl( self, min=1, max=300, value=1, limited=True, style=wx.ALIGN_RIGHT, size=(32,-1) )
		ctrl.Bind(IC.EVT_INT, self.onLapsChange)
		self.gbs.Add( ctrl, pos=(2, 1), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.lapsLabel = label
		self.lapsCtrl = ctrl
		
		label = wx.StaticText( self, -1, 'Distance:' )
		self.gbs.Add( label, pos=(2, 3), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.StaticText( self, -1, '10.0' )
		self.gbs.Add( ctrl, pos=(2, 4), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL )
		unitsLabel = wx.StaticText( self, -1, 'km' )
		self.gbs.Add( unitsLabel, pos=(2, 5), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.distanceLabel = label
		self.distanceCtrl = ctrl

		label = wx.StaticText( self, -1, 'Num Sprints:', style = wx.ALIGN_RIGHT )
		self.gbs.Add( label, pos=(2, 6), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.StaticText( self, -1, '' )
		self.gbs.Add( ctrl, pos=(2, 7), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.numSprintsLabel = label
		self.numSprintsCtrl = ctrl

		label = wx.CheckBox( self, -1, 'Snowball Points' )
		self.gbs.Add( label, pos=(2, 8), span=(1, 2), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = label
		ctrl.Bind( wx.EVT_CHECKBOX, self.onSnowballChange )
		self.snowballLabel = label
		self.snowballCtrl = ctrl

		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, -1, 'Sprint Every:' )
		self.gbs.Add( label, pos=(3, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = IC.IntCtrl( self, min=1, max=300, value=1, limited=True, style=wx.ALIGN_RIGHT, size=(32,-1) )
		ctrl.Bind(IC.EVT_INT, self.onSprintEveryChange)
		self.gbs.Add( ctrl, pos=(3, 1), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		unitsLabel = wx.StaticText( self, -1, 'laps' )
		self.gbs.Add( unitsLabel, pos=(3, 2), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL )
		self.sprintEveryLabel = label
		self.sprintEveryCtrl = ctrl
		self.sprintEveryUnitsLabel = unitsLabel
		
		label = wx.StaticText( self, -1, 'Course Length:' )
		self.gbs.Add( label, pos=(3, 3), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		
		ctrl = NC.NumCtrl( self, min = 0, integerWidth = 3, fractionWidth = 2, style=wx.ALIGN_RIGHT, size=(32,-1) )
		ctrl.SetAllowNegative(False)
		ctrl.Bind(wx.EVT_TEXT, self.onCourseLengthChange)
		self.gbs.Add( ctrl, pos=(3, 4), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		
		unitCtrl = wx.Choice( self, choices=['m', 'km'] )
		unitCtrl.SetSelection( 0 )
		self.Bind(wx.EVT_CHOICE, self.onCourseLengthUnitChange, unitCtrl)
		self.gbs.Add( unitCtrl, pos=(3, 5), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL )
		self.courseLengthLabel = label
		self.courseLengthCtrl = ctrl
		self.courseLengthUnitCtrl = unitCtrl

		label = wx.StaticText( self, -1, 'Points for Lapping:' )
		self.gbs.Add( label, pos=(3, 6), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = IC.IntCtrl( self, min=0, max=100, value=0, limited=True, style=wx.ALIGN_RIGHT, size=(32,-1) )
		ctrl.Bind(IC.EVT_INT, self.onChange)
		self.gbs.Add( ctrl, pos=(3, 7), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.pointsForLappingLabel = label
		self.pointsForLappingCtrl = ctrl

		label = wx.CheckBox( self, -1, 'Double Points for Last Sprint' )
		self.gbs.Add( label, pos=(3, 8), span=(1, 2), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = label
		ctrl.Bind( wx.EVT_CHECKBOX, self.onDoublePointsForLastSprintChange )
		self.doublePointsForLastSprintLabel = label
		self.doublePointsForLastSprintCtrl = ctrl

		#-----------------------------------------------------------------------------------------------------------
		self.vbs.Add( self.gbs, flag = wx.ALL, border = 4 )
		
		# Manage the display with a 4-way splitter.
		#sty = wx.SP_LIVE_UPDATE | wx.SP_3DBORDER
		sty = 0
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
		
		#-----------------------------------------------------------------------
		self.configureMenu = wx.Menu()
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, "&Points Race", "Configure Points Race" )
		self.Bind(wx.EVT_MENU, self.menuConfigurePointsRace, id=idCur )
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, "Point-a-&Lap", "Configure Point-a-Lap Race" )
		self.Bind(wx.EVT_MENU, self.menuConfigurePointALapRace, id=idCur )

		idCur = wx.NewId()
		self.configureMenu.Append( idCur, "&Tempo", "Configure Tempo Points Race" )
		self.Bind(wx.EVT_MENU, self.menuConfigureTempoRace, id=idCur )

		idCur = wx.NewId()
		self.configureMenu.Append( idCur, "&Snowball", "Configure Snowball Points Race" )
		self.Bind(wx.EVT_MENU, self.menuConfigureSnowballRace, id=idCur )

		self.menuBar.Append( self.configureMenu, "&ConfigureRace" )
		#-----------------------------------------------------------------------
		self.helpMenu = wx.Menu()

		#self.helpMenu.Append( wx.ID_HELP, "&Help...", "Help about PointsRaceMgr..." )
		#self.Bind(wx.EVT_MENU, self.menuHelp, id=wx.ID_HELP )
		
		self.helpMenu.Append( wx.ID_ABOUT , "&About...", "About PointsRaceMgr..." )
		self.Bind(wx.EVT_MENU, self.menuAbout, id=wx.ID_ABOUT )

		self.menuBar.Append( self.helpMenu, "&Help" )

		#------------------------------------------------------------------------------
		self.SetMenuBar( self.menuBar )
		#------------------------------------------------------------------------------
		self.Bind(wx.EVT_CLOSE, self.onCloseWindow)
		
		self.vbs.Add( self.splitter, 1, wx.GROW )
		self.SetSizer( self.vbs )
		
		accel_tbl = wx.AcceleratorTable([
				(wx.ACCEL_CTRL, ord('S'), wx.ID_SAVE),
				(wx.ACCEL_CTRL, ord('E'), idExportToExcel),
		])
		self.SetAcceleratorTable( accel_tbl )
		
		Model.newRace()
		self.refresh()
		Model.race.setChanged( False )

	def getDirName( self ):
		return Utils.getDirName()

	#--------------------------------------------------------------------------------------------

	def menuHelp( self, event ):
		pass
		
	#--------------------------------------------------------------------------------------------
	
	def configurePointsRace( self ):
		self.rankByCtrl.SetSelection( 0 )
		self.snowballCtrl.SetValue( False )
		self.doublePointsForLastSprintCtrl.SetValue( False )
		self.commit()
		self.refresh()

	def menuConfigurePointsRace( self, event ):
		Model.race.pointsForPlace = {
			1 : 5,
			2 : 3,
			3 : 2,
			4 : 1,
			5 : 0
		}
		self.configurePointsRace()
	
	def menuConfigurePointALapRace( self, event ):
		Model.race.pointsForPlace = {
			1 : 1,
			2 : 0,
			3 : -1,
			4 : -1,
			5 : -1
		}
		self.configurePointsRace()
	
	def menuConfigureTempoRace( self, event ):
		Model.race.pointsForPlace = {
			1 : 2,
			2 : 1,
			3 : 0,
			4 : -1,
			5 : -1
		}
		self.configurePointsRace()
	
	def menuConfigureSnowballRace( self, event ):
		Model.race.pointsForPlace = {
			1 : 1,
			2 : 0,
			3 : -1,
			4 : -1,
			5 : -1
		}
		self.snowballCtrl.SetValue( True )
		self.rankByCtrl.SetSelection( 0 )
		self.commit()
		self.refresh()
	
	def onChange( self, event ):
		self.commit()

	def onCourseLengthUnitChange( self, event ):
		race = Model.race
		race.setattr('courseLengthUnit', self.courseLengthUnitCtrl.GetCurrentSelection())
		self.updateDependentFields()
	
	def onRankByChange( self, event ):
		race = Model.race
		race.setattr('rankBy', self.rankByCtrl.GetCurrentSelection())
		self.updateDependentFields()
		self.worksheet.refresh()
		self.results.refresh()
	
	def onCourseLengthChange( self, event ):
		race = Model.race
		race.setattr('courseLength', self.courseLengthCtrl.GetValue())
		self.updateDependentFields()
		
	def onLapsChange( self, event ):
		race = Model.race
		race.setattr('laps', self.lapsCtrl.GetValue())
		self.updateDependentFields()
		self.worksheet.refresh()
		self.results.refresh()
		
	def onSprintEveryChange( self, event ):
		race = Model.race
		race.setattr('sprintEvery', self.sprintEveryCtrl.GetValue())
		self.updateDependentFields()
		self.worksheet.refresh()
		self.results.refresh()
			
	def onSnowballChange( self, event ):
		race = Model.race
		race.setattr('snowball', self.snowballCtrl.GetValue())
		self.worksheet.refresh()
		self.results.refresh()
		
	def onDoublePointsForLastSprintChange( self, event ):
		race = Model.race
		race.setattr('doublePointsForLastSprint', self.doublePointsForLastSprintCtrl.GetValue())
		self.updateDependentFields()
		self.worksheet.refresh()
		self.results.refresh()
	
	def menuExportToExcel( self, event ):
		self.commit()
		if not self.fileName:
			if not Utils.MessageOKCancel( self, 'You must save first.\n\nSave now?', 'Save Now', iconMask = wx.ICON_QUESTION):
				return
			if not self.menuSaveAs( event ):
				return

		xlFName = self.fileName[:-4] + '.xls'
		dlg = wx.DirDialog( self, 'Folder to write "%s"' % os.path.basename(xlFName),
						style=wx.DD_DEFAULT_STYLE, defaultPath=os.path.dirname(xlFName) )
		ret = dlg.ShowModal()
		dName = dlg.GetPath()
		dlg.Destroy()
		if ret != wx.ID_OK:
			return

		race = Model.race
			
		xlFName = os.path.join( dName, os.path.basename(xlFName) )
		
		if os.path.exists(xlFName) and \
		   not Utils.MessageOKCancel( self, 'Excel File Exists:\n\n%s\n\nReplace it?' % xlFName,
											'Excel File Exists', iconMask = wx.ICON_QUESTION):
			return

		wb = xlwt.Workbook()
		sheetCur = wb.add_sheet( re.sub('[:\\/?*\[\]]', ' ', '%s - %s' % (race.name, race.category)) )
		ToExcelSheet( sheetCur )

		try:
			wb.save( xlFName )
			try:
				webbrowser.open( xlFName )
			except:
				pass
			#Utils.MessageOK(self, 'Excel file written to:\n\n   %s' % xlFName, 'Excel Write', iconMask=wx.ICON_INFORMATION)
		except IOError:
			Utils.MessageOK(self,
						'Cannot write "%s".\n\nCheck if this spreadsheet is open.\nIf so, close it, and try again.' % xlFName,
						'Excel File Error', iconMask=wx.ICON_ERROR )

	#--------------------------------------------------------------------------------------------
	
	def onCloseWindow( self, event ):
		self.commit()
		race = Model.race
		if race.isChanged():
			if not self.fileName:
				ret = Utils.MessageYesNoCancel(self, 'Close:\n\nUnsaved changes!\nSave to a file?', 'Missing filename', iconMask = wx.ICON_QUESTION)
				if ret == wx.ID_YES:
					self.menuSaveAs()
					return
				elif ret == wx.ID_CANCEL:
					return
			else:
				ret = Utils.MessageYesNoCancel(self, 'Close:\n\nUnsaved changes!\nSave changes before Exit?', 'Unsaved Changes', iconMask = wx.ICON_QUESTION)
				if ret == wx.ID_YES:
					self.writeRace()
				elif ret == wx.ID_CANCEL:
					return
		sys.exit()

	def writeRace( self ):
		race = Model.race
		if race:
			if not self.fileName:
				if Utils.MessageOKCancel(self, 'WriteRace:\n\nMissing filename.\nSave to a file?', 'Missing filename', iconMask = wx.ICON_QUESTION): 
					self.menuSaveAs()
			if self.fileName:
				race.setChanged( False )
				pickle.dump( race, open(self.fileName, 'wb'), 2 )
				self.updateRecentFiles()
				self.updateDependentFields()

	def menuNew( self, event ):
		if Model.race.isChanged():
			ret = Utils.MessageYesNoCancel( self, 'NewRace:\n\nYou have unsaved changes.\n\nSave now?', 'Unsaved changes', iconMask = wx.ICON_QUESTION)
			if ret == wx.ID_YES:
				self.menuSave()
			elif ret == wx.ID_NO:
				pass
			elif ret == wx.ID_CANCEL:
				return
		self.fileName = ''
		Model.newRace()
		self.refresh()
	
	def updateRecentFiles( self ):
		self.filehistory.AddFileToHistory(self.fileName)
		self.filehistory.Save(self.config)
		self.config.Flush()

	def openRace( self, fileName ):
		if not fileName:
			return
		if Model.race.isChanged():
			ret = Utils.MessageYesNoCancel( self, 'OpenRace:\n\nYou have unsaved changes.\n\nSave now?', 'Unsaved changes', iconMask = wx.ICON_QUESTION)
			if ret == wx.ID_YES:
				self.menuSave()
			elif ret == wx.ID_NO:
				pass
			elif ret == wx.ID_CANCEL:
				return

		try:
			race = pickle.load( open(fileName, 'rb') )
			# Check a few fields to confirm we have the right file.
			a = getattr( race, 'sprintEvery' )
			a = getattr( race, 'courseLengthUnit' )
			Model.race = race
			self.fileName = fileName
			self.updateRecentFiles()
			self.refresh()

		except IOError:
			Utils.MessageOK(self, 'Cannot open file "%s".' % fileName, 'Cannot Open File', iconMask=wx.ICON_ERROR )
		except AttributeError:
			Utils.MessageOK(self, 'Bad race file "%s".' % fileName, 'Cannot Open File', iconMask=wx.ICON_ERROR )

	def menuOpen( self, event ):
		dlg = wx.FileDialog( self, message="Choose a Race file",
							defaultFile = '',
							wildcard = 'PointsRaceMgr files (*.tp5)|*.tp5',
							style=wx.OPEN | wx.CHANGE_DIR )
		if dlg.ShowModal() == wx.ID_OK:
			self.openRace( dlg.GetPath() )
		dlg.Destroy()
		
	def menuSave( self, event = None ):
		if not self.fileName:
			self.menuSaveAs( event )
		else:
			self.writeRace()
		
	def menuSaveAs( self, event = None ):
		dlg = wx.FileDialog( self, message="Save a Race File",
							defaultFile = '',
							wildcard = 'PointsRaceMgr files (*.tp5)|*.tp5',
							style=wx.SAVE | wx.FD_OVERWRITE_PROMPT | wx.CHANGE_DIR )
		success = False
		if dlg.ShowModal() == wx.ID_OK:
			self.fileName = dlg.GetPath()
			self.writeRace()
			success = True
		dlg.Destroy()
		return success

	def menuFileHistory( self, event ):
		fileNum = event.GetId() - wx.ID_FILE1
		fileName = self.filehistory.GetHistoryFile(fileNum)
		self.filehistory.AddFileToHistory(fileName)  # move up the list
		self.openRace( fileName )
		
	def menuExit(self, event):
		self.onCloseWindow( event )

	def menuAbout( self, event ):
		# First we create and fill the info object
		info = wx.AboutDialogInfo()
		info.Name = AppVerName
		info.Version = ''
		info.Copyright = "(C) 2011"
		info.Description = wordwrap(
			"Manage a points race - Track or Criterium.\n\n"
			"* Click on ConfigureRace and choose the race type (Points, Points-a-Lap, Tempo or Snowball)\n"
			"* Enter other Specific Race Information at the top\n"
			"* Enter Sprint Results under the top 'Sp' columns\n"
			"* Enter Laps Up/Down on the top right\n"
			"* Ranking by UCI rules is automatically updated in the lower left half of the screen\n"
			"    * The lower center shows the sprint points per rider\n"
			"    * The lower right shows subtotals, laps up/down and wins\n"
			"* Export results to Excel for final editing and publication\n\n"
			"If ranking by 'Points Only', riders are ranked by:\n"
			"  1.  Most Points\n"
			"  2.  If a tie, check Most Sprint Wins\n"
			"  3.  If a tie, check Finish Order (if known in recorded finishers)\n\n"
			"If ranking by 'Distance, then Points', riders are ranked by:\n"
			"  1.  Most Distance Covered\n"
			"  2.  If a tie, check Most Points\n"
			"  3.  If a tie, check Most Sprint Wins\n"
			"  4.  If a tie, check Finish Order (if known in recorded finishers)\n\n"
			"If PointsRaceMgr does not have enough information to rank some riders, it will give those riders a tie.\n"
			"You will then need to check the finish order to determine the final ranking and adjust it in the Excel output.\n\n"
			"PointsRaceMgr does not have sufficient information to rank riders with negative points as it does not know "
			"the status of all riders in the finish."
			"",
			600, wx.ClientDC(self))
		info.WebSite = ("http://sites.google.com/site/crossmgrsoftware", "CrossMgr home page")
		info.Developers = [
					"Edward Sitarski (edward.sitarski@gmail.com)",
					]

		licenseText = "User Beware!\n\n" \
			"This program is experimental, under development and may have bugs.\n" \
			"Feedback is sincerely appreciated.\n\n" \
			"CRITICALLY IMPORTANT MESSAGE:\n" \
			"This program is not warrented for any use whatsoever.\n" \
			"It may not produce correct results, it might lose your data.\n" \
			"The authors of this program assume no reponsibility or liability for data loss or erronious results produced by this program.\n\n" \
			"Use entirely at your own risk." \
			"Always use a paper manual backup."
		info.License = wordwrap(licenseText, 500, wx.ClientDC(self))

		wx.AboutBox(info)

	#--------------------------------------------------------------------------------------

	def updateDependentFields( self ):
		race = Model.race
		if not race:
			return

		self.SetTitle( '%s%s - %s by Edward Sitarski (edward.sitarski@gmail.com)' % (race.name, ', ' + self.fileName if self.fileName else '', AppVerName) )

		self.distanceCtrl.SetLabel( race.getDistanceStr() )		
		self.numSprintsCtrl.SetLabel( str(race.getNumSprints()) )
		self.sprints.updateShading()
		self.gbs.Layout()

	def commit( self ):
		if self.inRefresh:	# Don't commit anything while we are refreshing.
			return
		race = Model.race
		if not race:
			return

		for field in ['name', 'category', 'laps', 'sprintEvery', 'courseLength', 'doublePointsForLastSprint', 'pointsForLapping', 'snowball']:
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
	
	def refresh( self, pane = None ):
		self.inRefresh = True
		race = Model.race
		
		for field in ['name', 'category', 'laps', 'sprintEvery', 'courseLength', 'doublePointsForLastSprint', 'pointsForLapping']:
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

def MainLoop():
	parser = OptionParser( usage = "usage: %prog [options] [RaceFile.tp5]" )
	parser.add_option("-f", "--file", dest="filename", help="race file", metavar="RaceFile.tp5")
	parser.add_option("-q", "--quiet", action="store_false", dest="verbose", default=True, help='hide splash screen')
	parser.add_option("-r", "--regular", action="store_true", dest="regular", default=False, help='regular size')
	(options, args) = parser.parse_args()

	app = wx.PySimpleApp()
	mainWin = MainWin( None, title=AppVerName, size=(800,600) )
	if not options.regular:
		mainWin.Maximize( True )
	mainWin.Show()

	# Set the upper left icon.
	try:
		icon = wx.Icon( os.path.join(Utils.getImageFolder(), 'PointsRaceMgr16x16.ico'), wx.BITMAP_TYPE_ICO )
		mainWin.SetIcon( icon )
	except:
		pass

	if options.verbose:
		ShowSplashScreen()
	
	# Try a specified filename.
	fileName = options.filename
	
	# If nothing, try a positional argument.
	if not fileName and args:
		fileName = args[0]
	
	# Try to load a race.
	if fileName:
		try:
			mainWin.openRace( fileName )
		except (IndexError, AttributeError, ValueError):
			pass

	# Start processing events.
	app.MainLoop()

if __name__ == '__main__':
	MainLoop()
	