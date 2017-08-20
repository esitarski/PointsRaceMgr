#-----------------------------------------------------------------------
# Set translation locale.
#
import wx
locale = wx.Locale()

from Version import AppVerName
import gettext
initTranslationCalled = False
def initTranslation():
	global initTranslationCalled
	if not initTranslationCalled:
		gettext.install(AppVerName.split(None, 1), './locale', unicode=True)
		initTranslationCalled = True
		
initTranslation()

def BigFont():
	return wx.Font( (0,16), wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL )

try:
	from win32com.shell import shell, shellcon
except ImportError:
	pass
	
import os
import re
import sys
import math
import wx.grid		as gridlib
import unicodedata
import platform
import datetime
import string

import Model
try:
	from win32com.shell import shell, shellcon
except ImportError:
	pass
	
def removeDiacritic(input):
	'''
	Accept a unicode string, and return a normal string (bytes in Python 3)
	without any diacritical marks.
	'''
	if type(input) == str:
		return input
	else:
		return unicodedata.normalize('NFKD', input).encode('ASCII', 'ignore')
	
validFilenameChars = set( c for c in ("-_.() %s%s" % (string.ascii_letters, string.digits)) )
def RemoveDisallowedFilenameChars( filename ):
	cleanedFilename = unicodedata.normalize('NFKD', unicode(filename)).encode('ASCII', 'ignore')
	cleanedFilename = cleanedFilename.replace( '/', '_' )
	return ''.join(c for c in cleanedFilename if c in validFilenameChars)

def ordinal( value ):
	try:
		value = int(value)
	except ValueError:
		return value

	if (value % 100)//10 != 1:
		return "{}{}".format(value, ['th','st','nd','rd','th','th','th','th','th','th'][value%10])
	return "{}{}".format(value, "th")

GoodHighlightColour = wx.Colour( 0, 255, 0 )
BadHighlightColour = wx.Colour( 255, 255, 0 )
LightGrey = wx.Colour( 238, 238, 238 )
	
def MessageOK( parent, message, title = '', iconMask = 0):
	dlg = wx.MessageDialog(parent, message, title, wx.OK | iconMask)
	dlg.ShowModal()
	dlg.Destroy()
	return True
	
def MessageOKCancel( parent, message, title = '', iconMask = 0):
	dlg = wx.MessageDialog(parent, message, title, wx.OK | wx.CANCEL | iconMask )
	response = dlg.ShowModal()
	dlg.Destroy()
	return True if response == wx.ID_OK else False
	
def MessageYesNoCancel( parent, message, title = '', iconMask = 0 ):
	dlg = wx.MessageDialog(parent, message, title, wx.YES_NO | wx.CANCEL | iconMask )
	response = dlg.ShowModal()
	dlg.Destroy()
	return response
	
def SetValue( st, value ):
	if st.GetValue() != value:
		st.SetValue( value )

def SetLabel( st, label ):
	if st.GetLabel() != label:
		st.SetLabel( label )

def MakeGridReadOnly( grid ):
	for c in xrange(grid.GetNumberCols()):
		attr = gridlib.GridCellAttr()
		attr.SetReadOnly()
		grid.SetColAttr( c, attr )

def SetRowBackgroundColour( grid, row, colour ):
	for c in xrange(grid.GetNumberCols()):
		grid.SetCellBackgroundColour( row, c, colour )
		
def DeleteAllGridRows( grid ):
	if grid.GetNumberRows() > 0:
		grid.DeleteRows( 0, grid.GetNumberRows(), True )
		
def SwapGridRows( grid, r, rTarget ):
	if r != rTarget and 0 <= r < grid.GetNumberRows() and 0 <= rTarget < grid.GetNumberRows():
		for c in xrange(grid.GetNumberCols()):
			vSave = grid.GetCellValue( rTarget, c )
			grid.SetCellValue( rTarget, c, grid.GetCellValue(r,c) )
			grid.SetCellValue( r, c, vSave )
		
def AdjustGridSize( grid, rowsRequired = None, colsRequired = None ):
	# print 'AdjustGridSize: rowsRequired=', rowsRequired, ' colsRequired=', colsRequired

	if rowsRequired is not None:
		rowsRequired = int(rowsRequired)
		d = grid.GetNumberRows() - rowsRequired
		if d > 0:
			grid.DeleteRows( rowsRequired, d )
		elif d < 0:
			grid.AppendRows( -d )

	if colsRequired is not None:
		colsRequired = int(colsRequired)
		d = grid.GetNumberCols() - colsRequired
		if d > 0:
			grid.DeleteCols( colsRequired, d )
		elif d < 0:
			grid.AppendCols( -d )

class GridSuspendUpdate( object ):
	def __init__( self, grid ):
		self.grid = grid

	def __enter__( self ):
		self.grid.BeginBatch()
		
	def __exit__( self, *args, **kwargs ):
		self.grid.EndBatch()

def formatTime( secs ):
	if secs is None:
		secs = 0
	secs = int(secs + 0.5)
	hours = int(secs / (60*60));
	minutes = int( (secs / 60) % 60 )
	secs = secs % 60
	if hours > 0:
		return "%d:%02d:%02d" % (hours, minutes, secs)
	else:
		return "%02d:%02d" % (minutes, secs)

def formatDate( date ):
	y, m, d = date.split('-')
	d = datetime.date( int(y,10), int(m,10), int(d,10) )
	return d.strftime( '%B %d, %Y' )
		
def StrToSeconds( str = '' ):
	secs = 0
	for f in str.split(':'):
		secs = secs * 60 + int(f, 10)
	return secs
	
def SecondsToStr( secs = 0 ):
	secs = int(secs+0.5)
	return '%02d:%02d:%02d' % (secs / (60*60), (secs / 60)%60, secs % 60)

def SecondsToMMSS( secs = 0 ):
	secs = int(secs+0.5)
	return '%02d:%02d' % ((secs / 60)%60, secs % 60)
	
def disable_stdout_buffering():
	fileno = sys.stdout.fileno()
	temp_fd = os.dup(fileno)
	sys.stdout.close()
	os.dup2(temp_fd, fileno)
	os.close(temp_fd)
	sys.stdout = os.fdopen(fileno, "w", 0)
		
def getHomeDir():
	try:
		homedir = shell.SHGetFolderPath(0, shellcon.CSIDL_APPDATA, 0, 0)
		homedir = os.path.join( homedir, 'PointsRaceMgr' )
		if not os.path.exists(homedir):
			os.mkdir( homedir )
	except:
		homedir = os.path.expanduser('~')
	return homedir
	
#---------------------------------------------------------------------------

from contextlib import contextmanager

@contextmanager
def tag( buf, name, attrs = {} ):
	if not isinstance(attrs, dict) and attrs:
		attrs = { 'class': attrs }
	buf.write(
		u'<{}>'.format(u' '.join( [name] + [u'{}="{}"'.format(attr, value) for attr, value in attrs.iteritems()] ) )
	)
	yield
	if name not in ('meta', 'img'):
		buf.write( u'</{}>\n'.format(name) )

#------------------------------------------------------------------------
try:
	dirName = os.path.dirname(os.path.abspath(__file__))
except:
	dirName = os.path.dirname(os.path.abspath(sys.argv[0]))

if os.path.basename(dirName) == 'library.zip':
	dirName = os.path.dirname(dirName)
imageFolder = os.path.join(dirName, 'images')
htmlFolder = os.path.join(dirName, 'html')

def getDirName():		return dirName
def getImageFolder():	return imageFolder
def getHtmlFolder():	return htmlFolder

def AlignHorizontalScroll( gFrom, gTo ):
	xFrom, yFrom = gFrom.GetViewStart()
	xTo,   yTo   = gTo.GetViewStart()
	gTo.Scroll( xFrom, yTo )

def AlignVerticalScroll( gFrom, gTo ):
	xFrom, yFrom = gFrom.GetViewStart()
	xTo,   yTo   = gTo.GetViewStart()
	gTo.Scroll( xTo, yFrom )

#------------------------------------------------------------------------
PlatformName = platform.system()
def writeLog( message ):
	try:
		dt = datetime.datetime.now()
		dt = dt.replace( microsecond = 0 )
		msg = u'{} ({}) {}{}'.format(
			dt.isoformat(),
			PlatformName,
			message,
			'\n' if not message or message[-1] != '\n' else ''
		)
		sys.stdout.write( removeDiacritic(msg) )
		sys.stdout.flush()
	except IOError:
		pass
		
def disable_stdout_buffering():
	fileno = sys.stdout.fileno()
	temp_fd = os.dup(fileno)
	sys.stdout.close()
	os.dup2(temp_fd, fileno)
	os.close(temp_fd)
	sys.stdout = os.fdopen(fileno, "w", 0)
		
def logCall( f ):
	def _getstr( x ):
		return u'{}'.format(x) if not isinstance(x, wx.Object) else u'<<{}>>'.format(x.__class__.__name__)
	
	def new_f( *args, **kwargs ):
		parameters = [_getstr(a) for a in args] + [ u'{}={}'.format( key, _getstr(value) ) for key, value in kwargs.iteritems() ]
		writeLog( 'call: {}({})'.format(f.__name__, removeDiacritic(u', '.join(parameters))) )
		return f( *args, **kwargs)
	return new_f
	
def logException( e, exc_info ):
	eType, eValue, eTraceback = exc_info
	ex = traceback.format_exception( eType, eValue, eTraceback )
	writeLog( '**** Begin Exception ****' )
	for d in ex:
		for line in d.split( '\n' ):
			writeLog( line )
	writeLog( '**** End Exception ****' )

def asInt( v ):
	return u'{}'.format(int(v))
def asFloat( v ):
	return u'{:.1f}'.format(float(v))

#------------------------------------------------------------------------

mainWin = None
def setMainWin( mw ):
	global mainWin
	mainWin = mw
	
def getMainWin():
	return mainWin

def refresh():
	if mainWin:
		mainWin.refresh()
		
def writeRace():
	if mainWin is not None:
		mainWin.writeRace()
	
def isMainWin():
	return mainWin is not None
	
if __name__ == '__main__':
	hd = getHomeDir()
	open( os.path.join(hd, 'Test.txt'), 'w' )
