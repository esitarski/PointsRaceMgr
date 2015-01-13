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
from optparse import OptionParser

import Utils
import Model
import Version
from Sprints import Sprints
from UpDown import UpDown
from Worksheet import Worksheet
from Results import Results
from Printing import PointsMgrPrintout
from ToExcelSheet import ToExcelSheet
from Notes import NotesDialog

from Version import AppVerName

def ShowSplashScreen():
	bitmap = wx.Bitmap( os.path.join(Utils.getImageFolder(), 'TrackSprint.jpg'), wx.BITMAP_TYPE_JPEG )
	showSeconds = 2.5
	frame = wx.SplashScreen(bitmap, wx.SPLASH_CENTRE_ON_SCREEN|wx.SPLASH_TIMEOUT, int(showSeconds*1000), None)

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
		
		# Default print options.
		self.printData = wx.PrintData()
		self.printData.SetPaperId(wx.PAPER_LETTER)
		self.printData.SetPrintMode(wx.PRINT_MODE_PRINTER)
		# self.printData.SetOrientation(wx.LANDSCAPE)
		
		self.notesDialog = NotesDialog( self )

		Utils.setMainWin( self )
		
		# Configure the main menu.
		self.menuBar = wx.MenuBar(wx.MB_DOCKABLE)

		#-----------------------------------------------------------------------
		self.fileMenu = wx.Menu()

		self.fileMenu.Append( wx.ID_NEW , "&New...", "Create a new race" )
		self.Bind(wx.EVT_MENU, self.menuNew, id=wx.ID_NEW )

		self.fileMenu.Append( wx.ID_OPEN , "&Open...", "Open a race" )
		self.Bind(wx.EVT_MENU, self.menuOpen, id=wx.ID_OPEN )

		self.fileMenu.Append( wx.ID_SAVE , "&Save\tCtrl+S", "Save the race" )
		self.Bind(wx.EVT_MENU, self.menuSave, id=wx.ID_SAVE )

		self.fileMenu.Append( wx.ID_SAVEAS , "Save &As...", "Save the race under a different name" )
		self.Bind(wx.EVT_MENU, self.menuSaveAs, id=wx.ID_SAVEAS )

		self.fileMenu.AppendSeparator()
		self.fileMenu.Append( wx.ID_PAGE_SETUP , "Page &Setup...", "Setup the print page" )
		self.Bind(wx.EVT_MENU, self.menuPageSetup, id=wx.ID_PAGE_SETUP )

		self.fileMenu.Append( wx.ID_PREVIEW , "Print P&review...\tCtrl+R", "Preview the current page on screen" )
		self.Bind(wx.EVT_MENU, self.menuPrintPreview, id=wx.ID_PREVIEW )

		self.fileMenu.Append( wx.ID_PRINT , "&Print...\tCtrl+P", "Print the current page to a printer" )
		self.Bind(wx.EVT_MENU, self.menuPrint, id=wx.ID_PRINT )
		
		self.fileMenu.AppendSeparator()
		
		idCur = wx.NewId()
		idExportToExcel = idCur
		self.fileMenu.Append( idCur , "&Export to Excel...\tCtrl+E", "Export as an Excel Spreadsheet" )
		self.Bind(wx.EVT_MENU, self.menuExportToExcel, id=idCur )

		self.fileMenu.AppendSeparator()
		idCur = wx.NewId()
		idNotes = idCur
		self.fileMenu.Append( idCur, '&Notes...\tCtrl+N', "Notes")
		self.Bind(wx.EVT_MENU, self.menuNotes, id=idCur )
		
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
		label = wx.StaticText( self, label=u'Race Name:' )
		self.gbs.Add( label, pos=(0, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.TextCtrl( self, style=wx.TE_PROCESS_ENTER )
		ctrl.Bind(wx.EVT_TEXT_ENTER, self.onChange)
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
		ctrl.Bind(wx.EVT_TEXT_ENTER, self.onChange)
		hs.Add( ctrl, flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.communiqueLabel = label
		self.communiqueCtrl = ctrl
		
		self.gbs.Add( hs, pos=(0, 6), span=(1, 3) )
		
		#--------------------------------------------------------------------------------------------------------------
		label = wx.StaticText( self, label=u'Category:' )
		self.gbs.Add( label, pos=(1, 0), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.TextCtrl( self, style=wx.TE_PROCESS_ENTER )
		ctrl.Bind(wx.EVT_TEXT_ENTER, self.onChange)
		self.gbs.Add( ctrl, pos=(1, 1), span=(1,5), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL|wx.EXPAND )
		self.categoryLabel = label
		self.categoryCtrl = ctrl
		
		label = wx.StaticText( self, label=u'Rank By:', style = wx.ALIGN_RIGHT )
		self.gbs.Add( label, pos=(1, 6), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.Choice( self, choices=[
				u'Points then Finish Order',
				u'Distance, Points then Finish Order',
				u'Distance, Points, Num Wins then Finish Order'
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
		
		label = wx.StaticText( self, label=u'Distance:' )
		self.gbs.Add( label, pos=(2, 3), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL|wx.LEFT, border = 16 )
		ctrl = wx.StaticText( self, -1, '10.0' )
		self.gbs.Add( ctrl, pos=(2, 4), flag=wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL )
		unitsLabel = wx.StaticText( self, -1, 'km' )
		self.gbs.Add( unitsLabel, pos=(2, 5), flag=wx.ALIGN_LEFT|wx.ALIGN_CENTER_VERTICAL )
		self.distanceLabel = label
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
		
		#-----------------------------------------------------------------------
		self.configureMenu = wx.Menu()
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Points Race", "Configure Points Race" )
		self.Bind(wx.EVT_MENU, self.menuConfigurePointsRace, id=idCur )
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Madison", "Configure Madison" )
		self.Bind(wx.EVT_MENU, self.menuConfigureMadison, id=idCur )
		
		self.configureMenu.AppendSeparator()
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"Point-a-&Lap", "Configure Point-a-Lap Race" )
		self.Bind(wx.EVT_MENU, self.menuConfigurePointALapRace, id=idCur )

		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Tempo", "Configure Tempo Points Race" )
		self.Bind(wx.EVT_MENU, self.menuConfigureTempoRace, id=idCur )

		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Snowball", "Configure Snowball Points Race" )
		self.Bind(wx.EVT_MENU, self.menuConfigureSnowballRace, id=idCur )
		
		self.configureMenu.AppendSeparator()
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Criterium", "Configure Criterium Race" )
		self.Bind(wx.EVT_MENU, self.menuConfigureCriteriumRace, id=idCur )
		
		self.menuBar.Append( self.configureMenu, u"&ConfigureRace" )
		#-----------------------------------------------------------------------
		self.helpMenu = wx.Menu()

		#self.helpMenu.Append( wx.ID_HELP, u"&Help...", "Help about PointsRaceMgr..." )
		#self.Bind(wx.EVT_MENU, self.menuHelp, id=wx.ID_HELP )
		
		self.helpMenu.Append( wx.ID_ABOUT , u"&Help...", "About PointsRaceMgr..." )
		self.Bind(wx.EVT_MENU, self.menuAbout, id=wx.ID_ABOUT )

		self.menuBar.Append( self.helpMenu, u"&Help" )

		#------------------------------------------------------------------------------
		self.SetMenuBar( self.menuBar )
		#------------------------------------------------------------------------------
		self.Bind(wx.EVT_CLOSE, self.onCloseWindow)
		
		self.vbs.Add( self.splitter, 1, wx.GROW )
		self.SetSizer( self.vbs )
		
		Model.newRace()
		self.refresh()
		self.menuConfigurePointsRace()
		Model.race.setChanged( False )
		
		wx.CallAfter( self.showNotes )
	
	def showNotes( self ):
		self.notesDialog.refresh()
		width, height = self.notesDialog.GetSizeTuple()
		screenWidth, screenHeight = wx.GetDisplaySize()
		self.notesDialog.MoveXY( screenWidth-width, screenHeight-height-40 )
		self.notesDialog.Show( True )
	
	def menuPageSetup( self, event ):
		psdd = wx.PageSetupDialogData(self.printData)
		psdd.CalculatePaperSizeFromId()
		dlg = wx.PageSetupDialog(self, psdd)
		dlg.ShowModal()

		# this makes a copy of the wx.PrintData instead of just saving
		# a reference to the one inside the PrintDialogData that will
		# be destroyed when the dialog is destroyed
		self.printData = wx.PrintData( dlg.GetPageSetupData().GetPrintData() )
		dlg.Destroy()

	def menuPrintPreview( self, event ):
		self.commit()
		printout = PointsMgrPrintout()
		printout2 = PointsMgrPrintout()
		
		data = wx.PrintDialogData(self.printData)
		self.preview = wx.PrintPreview(printout, printout2, data)

		self.preview.SetZoom( 110 )
		if not self.preview.Ok():
			return

		pfrm = wx.PreviewFrame(self.preview, self, "Print preview")

		pfrm.Initialize()
		pfrm.SetPosition(self.GetPosition())
		screenWidth, screenHeight = wx.GetDisplaySize()
		pfrm.SetSize((screenWidth/2, screenHeight * 0.9))
		pfrm.Show(True)

	def menuPrint( self, event ):
		self.commit()
		printout = PointsRaceMgrPrintout()
		
		pdd = wx.PrintDialogData(self.printData)
		pdd.SetAllPages( 1 )
		pdd.EnablePageNumbers( 0 )
		pdd.EnableHelp( 0 )
		
		printer = wx.Printer(pdd)

		if not printer.Print(self, printout, True):
			if printer.GetLastError() == wx.PRINTER_ERROR:
				Utils.MessageOK(self, "There was a printer problem.\nCheck your printer setup.", "Printer Error",iconMask=wx.ICON_ERROR)
		else:
			self.printData = wx.PrintData( printer.GetPrintDialogData().GetPrintData() )

		printout.Destroy()

	def getDirName( self ):
		return Utils.getDirName()

	#--------------------------------------------------------------------------------------------

	def menuHelp( self, event ):
		pass
		
	#--------------------------------------------------------------------------------------------
	
	def configurePointsRace( self ):
		self.rankByCtrl.SetSelection( Model.Race.RankByPoints )
		self.snowballCtrl.SetValue( False )
		self.doublePointsForLastSprintCtrl.SetValue( False )
		self.pointsForLappingCtrl.SetValue( 20 )
		self.lapsCtrl.SetValue( 120 )
		self.sprintEveryCtrl.SetValue( 10 )
		self.commit()
		self.refresh()

	def menuConfigurePointsRace( self, event = None ):
		Model.race.pointsForPlace = {
			1 : 5,
			2 : 3,
			3 : 2,
			4 : 1,
			5 : 0
		}
		self.configurePointsRace()
		
	def menuConfigureMadison( self, event ):
		Model.race.pointsForPlace = {
			1 : 5,
			2 : 3,
			3 : 2,
			4 : 1,
			5 : 0
		}
		self.rankByCtrl.SetSelection( Model.Race.RankByDistancePoints )
		self.snowballCtrl.SetValue( False )
		self.doublePointsForLastSprintCtrl.SetValue( False )
		self.pointsForLappingCtrl.SetValue( 0 )
		self.lapsCtrl.SetValue( 100 )
		self.sprintEveryCtrl.SetValue( 20 )
		self.commit()
		self.refresh()
	
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
		self.rankByCtrl.SetSelection( Model.Race.RankByPoints )
		self.commit()
		self.refresh()

		
	def menuConfigureCriteriumRace( self, event ):
		self.menuConfigurePointsRace()
		self.rankByCtrl.SetSelection( Model.Race.RankByDistancePointsNumWins )
		self.pointsForLappingCtrl.SetValue( 0 )
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
	
	def menuNotes( self, event ):
		self.showNotes()
	
	def menuExportToExcel( self, event ):
		self.commit()
		self.refresh()
		if not self.fileName:
			if not Utils.MessageOKCancel( self, u'You must save first.\n\nSave now?', u'Save Now'):
				return
			if not self.menuSaveAs( event ):
				return

		xlFName = self.fileName[:-4] + '.xls'
		dlg = wx.DirDialog( self, u'Folder to write "{}"'.format(os.path.basename(xlFName)),
						style=wx.DD_DEFAULT_STYLE, defaultPath=os.path.dirname(xlFName) )
		ret = dlg.ShowModal()
		dName = dlg.GetPath()
		dlg.Destroy()
		if ret != wx.ID_OK:
			return

		race = Model.race
			
		xlFName = os.path.join( dName, os.path.basename(xlFName) )
		
		wb = xlwt.Workbook()
		sheetName = re.sub('[:\\/?*\[\]]', ' ', '%s' % race.category)[:31]
		sheetCur = wb.add_sheet( sheetName )
		ToExcelSheet( sheetCur )

		try:
			wb.save( xlFName )
			try:
				webbrowser.open( xlFName )
			except:
				pass
			#Utils.MessageOK(self, 'Excel file written to:\n\n   %s' % xlFName, 'Excel Write', iconMask=wx.ICON_INFORMATION)
		except Exception as e:
			Utils.MessageOK(self,
						u'Cannot write "{}"\n\n{}\n\nCheck if this spreadsheet is open.\nIf so, close it, and try again.'.format(xlFName,e),
						'Excel File Error', iconMask=wx.ICON_ERROR )

	#--------------------------------------------------------------------------------------------
	
	def onCloseWindow( self, event ):
		self.commit()
		race = Model.race
		if race.isChanged():
			if not self.fileName:
				ret = Utils.MessageYesNoCancel(self, u'Close:\n\nUnsaved changes!\nSave to a file?', u'Missing filename')
				if ret == wx.ID_YES:
					if not self.menuSaveAs():
						events.StopPropagation()
						return
				elif ret == wx.ID_CANCEL:
					events.StopPropagation()
					return
			else:
				ret = Utils.MessageYesNoCancel(self, u'Close:\n\nUnsaved changes!\nSave changes before Exit?', u'Unsaved Changes')
				if ret == wx.ID_YES:
					self.writeRace()
				elif ret == wx.ID_CANCEL:
					events.StopPropagation()
					return
		wx.Exit()

	def writeRaceValidFileName( self ):
		race = Model.race
		if not race:
			return
		with open(self.fileName, 'wb') as f:
			race.setChanged( False )
			pickle.dump( race, f, 2 )
		self.updateRecentFiles()
		self.updateDependentFields()
		
	def writeRace( self ):
		race = Model.race
		if not race:
			return
		if not self.fileName:
			if Utils.MessageOKCancel(self, u'WriteRace:\n\nMissing filename.\nSave to a file?', u'Missing filename'): 
				wx.CallAfter( self.menuSaveAs )
			return
			
		try:
			self.writeRaceValidFileName()
		except:
			Utils.MessageOK( self, u'WriteRace:\n\nError writing to file.\n\nRace NOT saved.\n\nTry "File|Save As..." again.', iconMask = wx.ICON_ERROR )

	def menuNew( self, event ):
		if Model.race.isChanged():
			ret = Utils.MessageYesNoCancel( self, u'NewRace:\n\nYou have unsaved changes.\n\nSave now?', u'Unsaved changes')
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
			ret = Utils.MessageYesNoCancel( self, u'OpenRace:\n\nYou have unsaved changes.\n\nSave now?', u'Unsaved changes')
			if ret == wx.ID_YES:
				self.menuSave()
			elif ret == wx.ID_NO:
				pass
			elif ret == wx.ID_CANCEL:
				return

		try:
			race = pickle.load( open(fileName, 'rb') )
			# Check a few fields to confirm we have the right file.
			a = race.sprintEvery
			a = race.courseLengthUnit
			Model.race = race
			self.fileName = fileName
			self.updateRecentFiles()
			self.refresh()

		except IOError:
			Utils.MessageOK(self, u'Cannot open file "{}".'.format(fileName), 'Cannot Open File', iconMask=wx.ICON_ERROR )
		except AttributeError:
			Utils.MessageOK(self, u'Bad race file "{}".'.format(fileName), u'Cannot Open File', iconMask=wx.ICON_ERROR )

	def menuOpen( self, event ):
		dlg = wx.FileDialog( self, message=u"Choose a Race file",
							defaultFile = '',
							wildcard = u'PointsRaceMgr files (*.tp5)|*.tp5',
							style=wx.OPEN | wx.CHANGE_DIR )
		if dlg.ShowModal() == wx.ID_OK:
			self.openRace( dlg.GetPath() )
		dlg.Destroy()
		
	def menuSave( self, event = None ):
		self.commit()
		self.refresh()
		if not self.fileName:
			self.menuSaveAs( event )
		else:
			self.writeRace()
		
	def menuSaveAs( self, event = None ):
		race = Model.race
		if not race:
			return False
			
		dlg = wx.FileDialog( self, message=u"Save a Race File",
							defaultFile = '',
							wildcard = u'PointsRaceMgr files (*.tp5)|*.tp5',
							style=wx.SAVE | wx.FD_OVERWRITE_PROMPT | wx.CHANGE_DIR )
		ret = dlg.ShowModal()
		if ret != wx.ID_OK:
			dlg.Destroy()
			return False
			
		self.fileName = dlg.GetPath()
		dlg.Destroy()
		
		try:
			self.writeRaceValidFileName()
			return True
		except:
			Utils.MessageOK( self, u'WriteRace:\n\nError writing to file.\n\nRace NOT saved.\n\nTry "File|Save As..." again.', iconMask = wx.ICON_ERROR )
			return False

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
		info.SetCopyright( "(C) 2011-{}".format( datetime.datetime.now().year ) )
		info.Description = wordwrap( unicode(
			"Manage a points race - Track or Criterium.\n\n"
			"* Click on ConfigureRace and choose a basic Race Format\n"
			"  (or customize your own race).\n"
			"* Enter other Specific Race Information at the top\n"
			"* Enter Sprint Results in the upper 'Sp1, Sp2, ...' columns\n"
			"* Enter Laps Up/Down on the top-right table\n"
			"* Enter the order the riders finish in the Finish Order column.\n"
			"  (this is only necessary if riders are still tied by procedure below)\n"
			"* Enter Rider DQ/Pull in the Status column\n"
			"* Correct Ranking by is automatically updated in the lower left half of the screen\n"
			"    * The lower center shows the sprint points per rider\n"
			"    * The lower right shows subtotals, laps up/down and wins (if applicable)\n"
			"* Export results to Excel for final editing and publication\n\n"
			"If ranking by 'Points then Finish Order' (eg. Points Race), riders are ranked by:\n"
			"  1.  Most Points\n"
			"  2.  If a tie, by Finish Order (if known in last sprint)\n"
			"  3.  If still a tie, by Finish Order as specified in the 'Finish Order' column\n\n"
			"If ranking by 'Distance, Points then Finish Order' (eg. Madison), riders are ranked by:\n"
			"  1.  Most Distance Covered (as adjusted by Laps +-)\n"
			"  2.  If a tie, by Most Points\n"
			"  3.  If a tie, by Finish Order (if known in last sprint)\n"
			"  4.  If still a tie, by Finish Order as specified in the 'Finish Order' column\n\n"
			"If ranking by 'DDistance, Points, Num Wins then Finish Order' (eg. Criterium with Points), riders are ranked by:\n"
			"  1.  Most Distance Covered (as adjusted by Laps +-)\n"
			"  2.  If a tie, by Most Points\n"
			"  3.  If still a tie, by Most Sprint Wins\n"
			"  4.  If still a tie, by Finish Order (if known in last sprint)\n"
			"  5.  If still a tie, by Finish Order as specified in the 'Finish Order' column\n\n"
			"When there is a tie, enter the 'Finish Order'.\n"
			"PointsRaceMgr will use the Finish Order to break ties.\n"
			"\n"
			"If you are scoring the final Points race in an Omnium, use the 'Existing Points' to add points awarded for each rider.  "
			"These will be added to the poits total and listed in the results."
			""),
			600, wx.ClientDC(self))
		info.WebSite = ("http://sites.google.com/site/crossmgrsoftware", "CrossMgr home page")
		info.Developers = [
			"Edward Sitarski (edward.sitarski@gmail.com)",
		]

		licenseText = unicode("User Beware!\n\n" \
			"This program is experimental, under development and may have bugs.\n" \
			"Feedback is sincerely appreciated.\n\n" \
			"CRITICALLY IMPORTANT MESSAGE:\n" \
			"This program is not warrented for any use whatsoever.\n" \
			"It may not produce correct results, it might lose your data.\n" \
			"The authors of this program assume no reponsibility or liability for data loss or erronious results produced by this program.\n\n" \
			"Use entirely at your own risk." \
			"Always use a paper manual backup."
		)
		info.License = wordwrap(licenseText, 600, wx.ClientDC(self))

		wx.AboutBox(info)

	#--------------------------------------------------------------------------------------

	def updateDependentFields( self ):
		race = Model.race
		if not race:
			return

		self.SetTitle( u'{}{} - {} by Edward Sitarski (edward.sitarski@gmail.com)'.format(race.name, ', ' + self.fileName if self.fileName else '', AppVerName) )

		self.distanceCtrl.SetLabel( race.getDistanceStr() )		
		self.numSprintsCtrl.SetLabel( unicode(race.getNumSprints()) )
		self.sprints.updateShading()
		self.gbs.Layout()

	def commit( self ):
		if self.inRefresh:	# Don't commit anything while we are refreshing.
			return
		race = Model.race
		if not race:
			return

		for field in [	'name', 'category', 'communique', 'laps', 'sprintEvery', 'courseLength',
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
			
		self.notesDialog.commit()
		self.updateDependentFields()
	
	def commitPanes( self ):
		self.commit()
		self.sprints.commit()
		self.updown.commit()
		self.refreshResults()
	
	def refreshResults( self ):
		self.worksheet.refresh()
		self.results.refresh()
		self.notesDialog.refresh()
	
	def refresh( self, pane = None ):
		self.inRefresh = True
		race = Model.race
		
		for field in [	'name', 'category', 'communique', 'laps', 'sprintEvery', 'courseLength',
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

def MainLoop():
	parser = OptionParser( usage = "usage: %prog [options] [RaceFile.tp5]" )
	parser.add_option("-f", "--file", dest="filename", help="race file", metavar="RaceFile.tp5")
	parser.add_option("-q", "--quiet", action="store_false", dest="verbose", default=True, help='hide splash screen')
	parser.add_option("-r", "--regular", action="store_true", dest="regular", default=False, help='regular size')
	(options, args) = parser.parse_args()

	app = wx.App( False )
	
	dataDir = Utils.getHomeDir()
	os.chdir( dataDir )
	redirectFileName = os.path.join(dataDir, 'PointsRaceMgr.log')
	
	if __name__ == '__main__':
		Utils.disable_stdout_buffering()
	else:
		try:
			logSize = os.path.getsize( redirectFileName )
			if logSize > 1000000:
				os.remove( redirectFileName )
		except:
			pass
	
		try:
			app.RedirectStdio( redirectFileName )
		except:
			pass
			
	Utils.writeLog( 'start: {}'.format(Version.AppVerName) )
	
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
	
