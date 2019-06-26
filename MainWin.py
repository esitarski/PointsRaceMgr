import wx
import wx.adv
from wx.lib.wordwrap import wordwrap
import wx.lib.agw.flatnotebook as fnb

import sys
import cgi
import six
import os
import io
import re
import datetime
import xlwt
import webbrowser
import six.moves.cPickle as pickle
import subprocess
import traceback
from optparse import OptionParser

import Utils
from Utils import tag
import Model
import Version
from EventList import EventList
from Configure import Configure
from RankDetails import RankDetails
from RankSummary import RankSummary
from StartList import StartList
from Commentary import Commentary
from ToPrintout import ToPrintout, ToHtml, ToExcel
from ExportGrid import ExportGrid

from Version import AppVerName

def ShowSplashScreen():
	bitmap = wx.Bitmap( os.path.join(Utils.getImageFolder(), 'TrackSprint.jpg'), wx.BITMAP_TYPE_JPEG )
	showSeconds = 2.5
	frame = wx.adv.SplashScreen(bitmap, wx.adv.SPLASH_CENTRE_ON_SCREEN|wx.adv.SPLASH_TIMEOUT, int(showSeconds*1000), None)

class MainWin( wx.Frame ):
	def __init__( self, parent, id = wx.ID_ANY, title='', size=(200,200) ):
		wx.Frame.__init__(self, parent, id, title, size=size)

		Model.newRace()
		
		self.SetBackgroundColour( wx.WHITE )
		
		# Add code to configure file history.
		self.filehistory = wx.FileHistory(8)
		self.config = wx.Config(appName="PointsRaceMgr",
								vendorName="Edward.Sitarski@gmail.com",
		)
		self.filehistory.Load(self.config)
		
		self.fileName = None
		self.inRefresh = False	# Flag to indicate we are doing a refresh.
		
		# Default print options.
		self.printData = wx.PrintData()
		self.printData.SetPaperId(wx.PAPER_LETTER)
		self.printData.SetPrintMode(wx.PRINT_MODE_PRINTER)
		# self.printData.SetOrientation(wx.LANDSCAPE)
		
		self.splitter = wx.SplitterWindow( self )

		self.eventList = EventList( self.splitter )
		mainPanel = wx.Panel( self.splitter )
		
		self.configure = Configure( mainPanel )
		
		#------------------------------------------------------------------------------
		# Configure the notebook.

		sty = wx.BORDER_SUNKEN
		self.notebook = fnb.FlatNotebook(mainPanel, wx.ID_ANY, agwStyle=fnb.FNB_VC8|fnb.FNB_NO_X_BUTTON)
		self.notebook.Bind( fnb.EVT_FLATNOTEBOOK_PAGE_CHANGED, self.onPageChanging )
		
		# Add all the pages to the notebook.
		self.pages = []

		def addPage( page, name ):
			self.notebook.AddPage( page, name )
			self.pages.append( page )
			
		self.attrClassName = [
			[ 'rankDetails',	RankDetails,	'Detail',   {'labelClickCallback':self.onRankDetailsColSelect} ],
			[ 'rankSummary',	RankSummary,	'Summary',	{} ],
			[ 'startList',		StartList,		'StartList',{} ],
			[ 'commentary',		Commentary,		'Commentary',{} ],
		]
		
		for i, (a, c, n, kw) in enumerate(self.attrClassName):
			kw.update( {'parent': self.notebook} )
			setattr( self, a, c(**kw) )
			addPage( getattr(self, a), n )
			
		self.notebook.SetSelection( 0 )
		
		vsPanel = wx.BoxSizer( wx.VERTICAL )
		vsPanel.Add( self.configure, 0, wx.EXPAND )
		vsPanel.Add( self.notebook, 1, wx.EXPAND )
		mainPanel.SetSizer( vsPanel )
		
		#------------------------------------------------------------------------------
		# Configure the main menu.
		
		self.menuBar = wx.MenuBar(wx.MB_DOCKABLE)

		#-----------------------------------------------------------------------
		self.fileMenu = wx.Menu()

		self.fileMenu.Append( wx.ID_NEW , "&New...", "Create a new race" )
		self.Bind(wx.EVT_MENU, self.menuNew, id=wx.ID_NEW )

		idNewNext = idCur = wx.NewId()
		self.fileMenu.Append( idCur , "&New Next...", "Create a new race from the Current Race" )
		self.Bind(wx.EVT_MENU, self.menuNewNext, id=idCur )
		
		self.fileMenu.Append( wx.ID_OPEN , "&Open...", "Open a race" )
		self.Bind(wx.EVT_MENU, self.menuOpen, id=wx.ID_OPEN )

		self.fileMenu.Append( wx.ID_SAVE , "&Save\tCtrl+S", "Save the race" )
		self.Bind(wx.EVT_MENU, self.menuSave, id=wx.ID_SAVE )

		self.fileMenu.Append( wx.ID_SAVEAS , "Save &As...", "Save the race under a different name" )
		self.Bind(wx.EVT_MENU, self.menuSaveAs, id=wx.ID_SAVEAS )

		self.fileMenu.AppendSeparator()
		
		idExportToExcel = idCur = wx.NewId()
		self.fileMenu.Append( idCur , "&Export to HTML...\tCtrl+H", "Export as an HTML Web Page" )
		self.Bind(wx.EVT_MENU, self.menuExportToHtml, id=idCur )

		idExportToExcel = idCur = wx.NewId()
		self.fileMenu.Append( idCur , "&Export to Excel...\tCtrl+E", "Export as an Excel Spreadsheet" )
		self.Bind(wx.EVT_MENU, self.menuExportToExcel, id=idCur )

		self.fileMenu.AppendSeparator()
		
		recent = wx.Menu()
		self.fileMenu.Append(wx.ID_ANY, "&Recent Files", recent)
		self.filehistory.UseMenu( recent )
		self.filehistory.AddFilesToMenu()
		
		self.fileMenu.AppendSeparator()

		self.fileMenu.Append( wx.ID_EXIT , "E&xit", "Exit PointsRaceMgr" )
		self.Bind(wx.EVT_MENU, self.menuExit, id=wx.ID_EXIT )
		
		self.Bind(wx.EVT_MENU_RANGE, self.menuFileHistory, id=wx.ID_FILE1, id2=wx.ID_FILE9)
		
		self.menuBar.Append( self.fileMenu, "&File" )

		self.configureMenu = wx.Menu()
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Points Race", "Configure Points Race" )
		self.Bind(wx.EVT_MENU, lambda e: self.configure.ConfigurePointsRace(), id=idCur )
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Madison", "Configure Madison" )
		self.Bind(wx.EVT_MENU, lambda e: self.configure.ConfigureMadison(), id=idCur )
		
		self.configureMenu.AppendSeparator()
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Tempo", "Configure UCI Tempo Points Race" )
		self.Bind(wx.EVT_MENU, lambda e: self.configure.ConfigureTempoRace(), id=idCur )

		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Tempo Top 2", "Configure Tempo Points Race Top 2" )
		self.Bind(wx.EVT_MENU, lambda e: self.configure.ConfigureTempoTop2Race(), id=idCur )

		self.configureMenu.AppendSeparator()
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Snowball", "Configure Snowball Points Race" )
		self.Bind(wx.EVT_MENU, lambda e: self.configure.ConfigureSnowballRace(), id=idCur )
		
		self.configureMenu.AppendSeparator()
		
		idCur = wx.NewId()
		self.configureMenu.Append( idCur, u"&Criterium", "Configure Criterium Race" )
		self.Bind(wx.EVT_MENU, lambda e: self.configure.ConfigureCriteriumRace(), id=idCur )
		
		self.menuBar.Append( self.configureMenu, u"&ConfigureRace" )
		#-----------------------------------------------------------------------
		self.helpMenu = wx.Menu()

		self.helpMenu.Append( wx.ID_HELP , u"&Help...", "Help..." )
		self.Bind(wx.EVT_MENU, self.menuHelp, id=wx.ID_HELP )

		self.helpMenu.Append( wx.ID_ABOUT , u"&About...", "About PointsRaceMgr..." )
		self.Bind(wx.EVT_MENU, self.menuAbout, id=wx.ID_ABOUT )

		self.menuBar.Append( self.helpMenu, u"&Help" )

		self.SetMenuBar( self.menuBar )
		#------------------------------------------------------------------------------
		self.Bind(wx.EVT_CLOSE, self.onCloseWindow)
		
		self.vbs = wx.BoxSizer( wx.VERTICAL )
		self.vbs.Add( self.splitter, 1, wx.EXPAND )
		self.SetSizer( self.vbs )
		
		self.splitter.SplitVertically( self.eventList, mainPanel )
		wx.CallAfter( self.splitter.SetSashPosition, 300 )
		
		Utils.setMainWin( self )
		Model.newRace()
		self.refresh()
		self.configure.ConfigurePointsRace()
		Model.race.setChanged( False )		
		
		wx.CallAfter( self.eventList.newButton.SetFocus )
	
	def callPageRefresh( self, i ):
		try:
			self.pages[i].refresh()
		except (AttributeError, IndexError) as e:
			pass

	def setTitle( self ):
		race = Model.race
		if self.fileName:
			title = u'{}: {}{} - {}'.format(race.category, '*' if race.isChanged() else '', self.fileName, Version.AppVerName)
		else:
			title = u'{}: {}'.format( race.category, Version.AppVerName )
		self.SetTitle( title )
	
	def callPageCommit( self, i ):
		try:
			self.pages[i].commit()
			self.setTitle()
		except IndexError as e:
			pass
			
	def onRankDetailsColSelect( self, grid, col ):
		race = Model.race
		label = grid.GetColLabelValue(col)
		labels = [label]
		if label == race.getSprintLabel(race.getNumSprints()):
			labels.append( u'Finish' )
		
		egrid = self.eventList.grid
		egrid.ClearSelection()
		for row in range(egrid.GetNumberRows()):
			v = egrid.GetCellValue(row, 0)
			if any(v.startswith(lab) for lab in labels):
				egrid.SelectRow( row, True )
		
	def onPageChanging( self, event ):
		notebook = event.GetEventObject()
		self.callPageCommit( event.GetOldSelection() )
		self.callPageRefresh( event.GetSelection() )
		try:
			Utils.writeLog( u'page: {}\n'.format(notebook.GetPage(event.GetSelection()).__class__.__name__) )
		except IndexError:
			pass
		event.Skip()	# Required to properly repaint the screen.

	def getDirName( self ):
		return Utils.getDirName()

	#--------------------------------------------------------------------------------------------

	def menuExportToHtml( self, event ):
		self.commit()
		self.refresh()
		if not self.fileName:
			if not Utils.MessageOKCancel( self, u'You must save first.\n\nSave now?', u'Save Now'):
				return
			if not self.menuSaveAs( event ):
				return

		htmlFName = os.path.splitext(self.fileName)[0] + '.html'
		
		race = Model.race

		try:
			with io.open( htmlFName, 'w', encoding='utf8' ) as html:
				def write( v ):
					html.write( u'{}'.format(v) )
				
				with tag(html, 'html'):
					with tag(html, 'head'):
						with tag(html, 'title'):
							write( race.name.replace('\n', ' ') )
						with tag(html, 'meta', dict(charset="UTF-8")):
							pass
						with tag(html, 'meta', dict(name='author', content="Edward Sitarski")):
							pass
						with tag(html, 'meta', dict(name='copyright', content="Edward Sitarski, 2013-{}".format(datetime.datetime.now().strftime('%Y')))):
							pass
						with tag(html, 'meta', dict(name='generator', content="PointsRaceMgr")):
							pass
						with tag(html, 'style', dict( type="text/css")):
							write( u'''
body{ font-family: sans-serif; }

h1{ font-size: 250%; }
h2{ font-size: 200%; }

#idRaceName {
	font-size: 200%;
	font-weight: bold;
}
#idImgHeader { box-shadow: 4px 4px 4px #888888; }
.smallfont { font-size: 80%; }
.bigfont { font-size: 120%; }
.hidden { display: none; }

table.results td.numeric { text-align: right; }
table.details td.numeric { text-align: right; }

table.results {
	font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;
	border-collapse:collapse;
}
table.results td, table.results th {
	font-size:1em;
	padding:3px 7px 2px 7px;
	text-align: left;
}
table.results th {
	font-size:1.1em;
	text-align:left;
	padding-top:5px;
	padding-bottom:4px;
	background-color:#7FE57F;
	color:#000000;
	vertical-align:bottom;
}
table.results tr.odd {
	color:#000000;
	background-color:#EAF2D3;
}

table.details {
	font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;
	border-collapse:collapse;
}
table.details td, table.details th {
	font-size:1em;
	padding:3px 7px 2px 7px;
	text-align: left;
}
table.details th {
	font-size:1.1em;
	text-align:left;
	padding-top:5px;
	padding-bottom:4px;
	background-color:#7FE57F;
	color:#000000;
	vertical-align:bottom;
}

table.details td.rightAlign, table.details th.rightAlign {
	text-align:right;
}

table.details td.leftAlign, table.details th.leftAlign {
	text-align:left;
}

table.details td.leftBorder { border-left: 1pt solid #CCC; }
table.details td.rightBorder { border-right: 1pt solid #CCC; }
table.details td.topBorder { border-top: 1pt solid #CCC; }
table.details td.bottomBorder { border-bottom: 1pt solid #CCC; }

.smallFont {
	font-size: 75%;
}

table.results td.leftBorder, table.results th.leftBorder
{
	border-left:1px solid #98bf21;
}

table.results tr:hover
{
	color:#000000;
	background-color:#FFFFCC;
}
table.results tr.odd:hover
{
	color:#000000;
	background-color:#FFFFCC;
}


table.results td {
	border-top:1px solid #98bf21;
}

table.results td.noborder {
	border-top:0px solid #98bf21;
}

table.results td.rightAlign, table.results th.rightAlign {
	text-align:right;
}

table.results td.leftAlign, table.results th.leftAlign {
	text-align:left;
}

.topAlign {
	vertical-align:top;
}

table.results th.centerAlign, table.results td.centerAlign {
	text-align:center;
}

.ignored {
	color: #999;
	font-style: italic;
}

table.points tr.odd {
	color:#000000;
	background-color:#EAF2D3;
}

.rank {
	color: #999;
	font-style: italic;
}

.points-cell {
	text-align: right;
	padding:3px 7px 2px 7px;
}

hr { clear: both; }

@media print {
	.noprint { display: none; }
	.title { page-break-after: avoid; }
}
''')
					with tag(html, 'body'):
						with tag(html, 'h1'):
							write( u'{}: {}'.format(cgi.escape(race.name), race.date.strftime('%Y-%m-%d')) )
							if race.communique:
								write( u': Communiqu\u00E9 {}'.format(cgi.escape(race.communique)) )
						with tag(html, 'h2'):
							write( u'Category: {}'.format(cgi.escape(race.category)) )
						
						d = race.courseLength*race.laps
						if d == int(d):
							d = '{:,d}'.format(int(d))
						else:
							d = '{:,.1f}'.format(float(d))
						
						with tag(html, 'h3'):
							s =  [
								u'Laps: {}'.format(race.laps),
								u'Sprint Every: {} laps'.format(race.sprintEvery),
								u'Distance: {}{}'.format( d, ['m','km'][race.courseLengthUnit] ),
							]
							write( u',  '.join(s) )
						write( '<br/>' )
						write( '<hr/>' )
						write( '<br/>' )
						ToHtml( html )
						write( '<br/>' )
						write( 'Powered by ' )
						with tag(html, 'a', {
								'href':'http://sites.google.com/site/crossmgrsoftware/',
								'target':'_blank'
							} ):
							write( 'CrossMgr' )

		except Exception as e:
			traceback.print_exc()
			Utils.MessageOK(self,
						u'Cannot write "{}"\n\n{}\n\nCheck if this file is open.\nIf so, close it, and try again.'.format(htmlFName,e),
						'HTML File Error', iconMask=wx.ICON_ERROR )
		
		try:
			webbrowser.open( htmlFName )
		except:
			pass
		#Utils.MessageOK(self, 'Excel file written to:\n\n   {}'.format(htmlFName), 'Excel Write', iconMask=wx.ICON_INFORMATION)

	def menuExportToExcel( self, event ):
		self.commit()
		self.refresh()
		if not self.fileName:
			if not Utils.MessageOKCancel( self, u'You must save first.\n\nSave now?', u'Save Now'):
				return
			if not self.menuSaveAs( event ):
				return

		xlFName = os.path.splitext(self.fileName)[0] + '.xls'
		
		wb = xlwt.Workbook()
		ToExcel( wb )

		try:
			wb.save( xlFName )
			try:
				webbrowser.open( xlFName )
			except:
				pass
			#Utils.MessageOK(self, 'Excel file written to:\n\n   {}'.format(xlFName), 'Excel Write', iconMask=wx.ICON_INFORMATION)
		except Exception as e:
			traceback.print_exc()
			Utils.MessageOK(self,
						u'Cannot write "{}"\n\n{}\n\nCheck if this spreadsheet is open.\nIf so, close it, and try again.'.format(xlFName,e),
						'Excel File Error', iconMask=wx.ICON_ERROR )

	#--------------------------------------------------------------------------------------------
	
	def onCloseWindow( self, event ):
		self.commit()
		self.refresh()
		race = Model.race
		if race.isChanged():
			if not self.fileName:
				ret = Utils.MessageYesNoCancel(self, u'Close:\n\nUnsaved changes!\nSave to a file?', u'Missing filename')
				if ret == wx.ID_YES:
					if not self.menuSaveAs():
						event.StopPropagation()
						return
				elif ret == wx.ID_CANCEL:
					event.StopPropagation()
					return
			else:
				ret = Utils.MessageYesNoCancel(self, u'Close:\n\nUnsaved changes!\nSave changes before Exit?', u'Unsaved Changes')
				if ret == wx.ID_YES:
					self.writeRace()
				elif ret == wx.ID_CANCEL:
					event.StopPropagation()
					return
		wx.Exit()

	def writeRaceValidFileName( self ):
		race = Model.race
		if not race:
			return
		with io.open(self.fileName, 'wb') as f:
			race.setChanged( False )
			pickle.dump( race, f, 2 )
		self.updateRecentFiles()
		self.setTitle()
		
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
		except Exception as e:
			Utils.MessageOK( self, u'WriteRace:\n\n{}\n\nError writing to file.\n\nRace NOT saved.\n\nTry "File|Save As..." again.'.format(e), iconMask = wx.ICON_ERROR )

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
	
	def menuNewNext( self, event ):
		if Model.race.isChanged():
			ret = Utils.MessageYesNoCancel( self, u'NewRace:\n\nYou have unsaved changes.\n\nSave now?', u'Unsaved changes')
			if ret == wx.ID_YES:
				self.menuSave()
			elif ret == wx.ID_NO:
				pass
			elif ret == wx.ID_CANCEL:
				return
		if self.fileName:
			self.fileName = os.path.splitext(self.fileName)[0]
			for i in range(1,100):
				s = '-{}'.format(i)
				if self.fileName.endswith(s):
					self.fileName[:-len(s)] + '-{}'.format(i+1)
					break
			else:
				self.fileName += '-1'
			self.fileName += '.tp5'
		else:
			self.fileName = ''
		Model.race.newNext()
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
			if six.PY2:
				race = pickle.load( io.open(fileName, 'rb') )
			else:
				race = pickle.load( io.open(fileName, 'rb'), fix_imports=True, encoding="utf-8", )
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
		if not self.fileName:
			self.menuSaveAs( event )
		else:
			self.writeRace()
		
	def menuSaveAs( self, event = None ):
		race = Model.race
		if not race:
			return False
			
		self.commit()
		dlg = wx.FileDialog( self, message=u"Save a Race File",
							defaultFile = '',
							wildcard = u'PointsRaceMgr files (*.tp5)|*.tp5',
							style=wx.FD_SAVE | wx.FD_CHANGE_DIR )
		while 1:
			ret = dlg.ShowModal()
			if ret != wx.ID_OK:
				dlg.Destroy()
				return False
				
			fileName = os.path.splitext(dlg.GetPath())[0] + '.tp5'
			
			if os.path.exists(fileName):
				if Utils.MessageOKCancel( self, u'File Exists.\n\nOverwrite?', iconMask=wx.ICON_WARNING ):
					break
			else:
				break	
		
		dlg.Destroy()	
		self.fileName = fileName	
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
		
	def menuHelp(self, event):
		message = u'{}'.format(
			"Manage a Points Race: Track or Criterium.\n\n"
			"Click on ConfigureRace menu and choose a standard Race Format "
			" (or customize your own format).\n"
			"Configure all other race information at the top.\n"
			"\n"
			"While the race is underway, press the 'New Race Event' button.\n"
			"When the screen pops up, enter the Bibs involved in the Race Event (space or comma separated)."
			"Then, set the Event Type which can be a Sprint (default), +/- Laps, DNF, Finish, DNS or DSQ.\n"
			"Use 'Finish' to enter the finish order.  This will automatically count as the last Sprint.\n"
			"\n"
			"You can rearrange the sequence of Race Events by dragging-and-dropping rows in Column 1.\n"
			"Delete unwanted Race Events by right-clicking on the Event Type in the list.\n"
			"\n"
			"Up-to-date results are always shown on the Details and Summary screens.\n"
			"\n"
			"Use the 'Start List' screen to enter rider information (you can also import it from Excel).\n"
			"Use the 'Existing Points' field for scoring an Omnimum.\n"
			"\n"
			"In the 'Details' screen, clicking on the Sprint columns header will instantly find the Sprint Event corresponding\nto that sprint.\n"
			"\n"
			"It is possible to enter Race Events without using a mouse.\n"
			"Press Enter to trigger 'New Race Event', then enter the Bibs, then press Tab and Space to change the Event type,\nthen press Enter to save.\n"
			"\n"
			"If ranking by 'Points, then Finish Order' (eg. Points Race, Madison), riders are ranked by:\n"
			"  1.  Most Points\n"
			"  2.  If a tie, by Finish Order\n"
			"If ranking by 'Laps Completed, Points, Num Wins, then Finish Order' (eg. UCI Criterium with Points), riders are ranked by:\n"
			"  1.  Most Laps Completed (as specified by +/- Laps)\n"
			"  2.  If a tie, by Most Points\n"
			"  3.  If still a tie, by Most Sprint Wins\n"
			"  4.  If still a tie, by Finish Order\n\n"
			"If ranking by 'Laps Completed, Points, then Finish Order', riders are ranked by:\n"
			"  1.  Most Laps Completed (as specified by +/- Laps)\n"
			"  2.  If a tie, by Most Points\n"
			"  3.  If still a tie, by Finish Order\n\n"
			"")
		dlg = wx.MessageDialog(self, message, "PointsRaceMgr Help", wx.OK | wx.ICON_INFORMATION)
		dlg.ShowModal()
		dlg.Destroy()

	def menuAbout( self, event ):
		self.commit()
		self.refresh()
		
		# First we create and fill the info object
		info = wx.AboutDialogInfo()
		info.Name = AppVerName
		info.Version = ''
		info.SetCopyright( u"(C) 2011-{}".format( datetime.datetime.now().year ) )
		info.Description = wordwrap( (
			u"Manage a Points Race: Track or Criterium.\n\n"
			u"For details, see Help"
			u""),
			600, wx.ClientDC(self))
		info.WebSite = ("http://sites.google.com/site/crossmgrsoftware", "CrossMgr home page")
		info.Developers = [
			u"Edward Sitarski (edward.sitarski@gmail.com)",
		]

		licenseText = ( u"User Beware!\n\n"
			u"This program is experimental, under development and may have bugs.\n"
			u"Feedback is sincerely appreciated.\n\n"
			u"CRITICALLY IMPORTANT MESSAGE:\n"
			u"This program is not warrented for any use whatsoever.\n"
			u"It may not produce correct results, it might lose your data.\n"
			u"The authors of this program assume no reponsibility or liability for data loss or erronious results produced by this program.\n\n"
			u"Use entirely at your own risk."
			u"Always use a paper manual backup."
		)
		info.License = wordwrap(licenseText, 600, wx.ClientDC(self))

		wx.AboutBox(info)

	#--------------------------------------------------------------------------------------
	def commit( self ):
		self.configure.commit()
		self.eventList.commit()
		self.setTitle()
	
	def refreshCurrentPage( self ):
		self.callPageRefresh( self.notebook.GetSelection() )

	def refresh( self, includeConfigure=True ):
		self.setTitle()
		if includeConfigure:
			self.configure.refresh()
		self.eventList.refresh()
		self.refreshCurrentPage()

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
	
	'''
	def my_handler( type, value, traceback ):
		print 'my_handler'
		print type, value, trackback
	sys.excepthook = my_handler
	'''
	
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
	
	Model.newRace()
	
	Model.race._populate()
	
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
	
