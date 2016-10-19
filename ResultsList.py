import wx
import wx.grid as gridlib

import os
import sys
import cgi
import operator
import xlwt
import re
import datetime

import Utils
from Utils import tag
import Model

from ToPrintout import GrowTable

#--------------------------------------------------------------------------------
class ResultsList(wx.Panel):

	def __init__(self, parent):
		wx.Panel.__init__(self, parent)
		
		self.resultsGrid = parent.GetParent().scoreSheet.worksheet.gridBib
		self.summaryGrid = parent.GetParent().scoreSheet.results.gridResults
		
		self.fieldNames  = Model.RiderInfo.FieldNames
		self.headerNames = Model.RiderInfo.HeaderNames
		self.numericCols = set()
		
		self.grid = gridlib.Grid( self, style = wx.BORDER_SUNKEN )
		self.grid.DisableDragRowSize()
		self.grid.SetRowLabelSize( 0 )
		headers = self.getHeaders()
		self.grid.CreateGrid( 0, len(headers) )
		for col, h in enumerate(headers):
			self.grid.SetColLabelValue( col, h )
			attr = gridlib.GridCellAttr()
			attr.SetReadOnly()
			self.grid.SetColAttr( col, attr )

		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add(self.grid, 1, flag=wx.EXPAND|wx.ALL, border = 6)
		self.SetSizer(sizer)
		
	def getGrid( self ):
		return self.grid
		
	def getHeaders( self ):
		return (
			[self.resultsGrid.GetColLabelValue(0)] +
			list(self.headerNames[:-1]) +
			[self.resultsGrid.GetColLabelValue(c) for c in xrange(2,self.resultsGrid.GetNumberCols()-1)] +
			[self.summaryGrid.GetColLabelValue(c) for c in xrange(self.summaryGrid.GetNumberCols())]
		)
		
	def updateGrid( self ):
		race = Model.race
		headers = self.getHeaders()
		fieldNamesResults = self.fieldNames[:-1] # Ignore 'existing_points" as it comes from the scoreSheet.
		
		attrs = []
		for col, h in enumerate(headers):
			attr = gridlib.GridCellAttr()
			if col <= 1 or col > len(fieldNamesResults):
				attr.SetRenderer( gridlib.GridCellNumberRenderer() )
			attr.SetReadOnly()
			attrs.append( attr )

		content = []
		for row in xrange(self.resultsGrid.GetNumberRows()):
			bib = int( self.resultsGrid.GetCellValue(row, 1) )
			info = race.riderInfo.get( bib, None )
			if info is None:
				info = Model.RiderInfo( bib )
			values = (
				[self.resultsGrid.GetCellValue(row, 0)] +
				[unicode(getattr(info,f)).upper() if f == 'last_name' else unicode(getattr(info,f))
					for f in fieldNamesResults] +
				[self.resultsGrid.GetCellValue(row, c) for c in xrange(2,self.resultsGrid.GetNumberCols()-1)] +
				[self.summaryGrid.GetCellValue(row, c) for c in xrange(self.summaryGrid.GetNumberCols())]
			)
			content.append( values )
		
		empty = set(c for c in xrange(len(headers))
			if all((not values[c] or values[c] == '0.0') for values in content))
		
		def pluck( a ):
			return [v for c, v in enumerate(a) if c not in empty]
		headers = pluck( headers )
		attrs = pluck( attrs )
		content = [pluck(values) for values in content]

		def isIntOrBlank( v ):
			if not v:
				return True
			try:
				return float(v) == int(v.split('.')[0])
			except:
				return False

		def isFloatOrBlank( v ):
			if not v:
				return True
			try:
				f = float(v)
				return True
			except:
				return False
		
		self.numericCols = set()
		for c in xrange(len(headers)):
			if all(isIntOrBlank(values[c]) for values in content):
				self.numericCols.add( c )
				for values in content:
					if values[c].endswith(u'.0'):
						values[c] = values[c][:-2]
			elif all(isFloatOrBlank(values[c]) for values in content):
				self.numericCols.add( c )
				for values in content:
					if values[c] and not values[c].endswith(u'.0'):
						values[c] += u'.0'
		
		Utils.AdjustGridSize( self.grid, len(content), len(headers) )
		
		self.grid.BeginBatch()
		for c, (h,a) in enumerate(zip(headers, attrs)):
			self.grid.SetColLabelValue(c, h)
			self.grid.SetColAttr(c, a)
		for r, values in enumerate(content):
			for c, v in enumerate(values):
				self.grid.SetCellValue( r, c, v )
		self.grid.AutoSize()
		self.grid.EndBatch()
		
		self.Layout()
		
	def refresh( self ):
		self.updateGrid()
		
	def commit( self ):
		pass
		
	def toExcelSheet( self, ws ):
		self.refresh()
		
		race = Model.race
		
		labelStyle = xlwt.easyxf( "alignment: horizontal right;" )
		fieldStyle = xlwt.easyxf( "alignment: horizontal right;" )
		unitsStyle = xlwt.easyxf( "alignment: horizontal left;" )
	
		fnt = xlwt.Font()
		fnt.name = 'Arial'
		fnt.bold = True
		
		headerStyle = xlwt.XFStyle()
		headerStyle.font = fnt
		
		fnt = xlwt.Font()
		fnt.name = 'Arial'
		fnt.bold = True
		fnt.height = int(fnt.height * 1.5)
		
		titleStyle = xlwt.XFStyle()
		titleStyle.font = fnt

		rowCur = 0
		
		ws.write_merge( rowCur, rowCur, 0, 6, race.name, titleStyle )
		if race.communique:
			ws.write( rowCur, 7    , u'Communiqu\u00E9:', labelStyle )
			ws.write( rowCur, 7 + 1, race.communique, unitsStyle )
		ws.write( rowCur, 9, race.date.isoformat()[2:], unitsStyle )
		rowCur += 1
		ws.write( rowCur, 0, u'Category', labelStyle )
		ws.write( rowCur, 1, race.category, unitsStyle )
		
		rowCur += 2
		for c in xrange(self.grid.GetNumberCols()):
			ws.write( rowCur, c, self.grid.GetColLabelValue(c), headerStyle )
			
		def toInt( v ):
			try:
				return int(v)
			except:
				return v
			
		for r in xrange(self.grid.GetNumberRows()):
			rowCur += 1
			for c in xrange(self.grid.GetNumberCols()):
				v = self.grid.GetCellValue(r,c)
				if v != u'':
					ws.write( rowCur, c, toInt(self.grid.GetCellValue(r,c)) )

	def toPrintout( self, dc ):
		self.refresh()
		race = Model.race
		
		#---------------------------------------------------------------------------------------
		# Format on the page.
		(widthPix, heightPix) = dc.GetSizeTuple()
		
		# Get a reasonable border.
		borderPix = max(widthPix, heightPix) / 25
		
		widthFieldPix = widthPix - borderPix * 2
		heightFieldPix = heightPix - borderPix * 2
		
		xPix = borderPix
		yPix = borderPix

		# Race Information
		xLeft = xPix
		yTop = yPix
		
		gt = GrowTable(alignHorizontal=GrowTable.alignLeft, alignVertical=GrowTable.alignTop, cellBorder=False)
		titleAttr = GrowTable.bold | GrowTable.alignLeft
		rowCur = 0
		rowCur = gt.set( rowCur, 0, race.name, titleAttr )[0] + 1
		rowCur = gt.set( rowCur, 0, race.category, titleAttr )[0] + 1
		rowCur = gt.set( rowCur, 0, u'{} Laps, {} Sprints, {} km'.format(race.laps, race.getNumSprints(), race.getDistance()), titleAttr )[0] + 1
		rowCur = gt.set( rowCur, 0, race.date.strftime('%Y-%m-%d'), titleAttr )[0] + 1
		
		if race.communique:
			rowCur = gt.set( rowCur, 0, u'Communiqu\u00E9: {}'.format(race.communique), GrowTable.alignRight )[0] + 1
		rowCur = gt.set( rowCur, 0, u'Approved by:________', GrowTable.alignRight )[0] + 1
		
		# Draw the title
		titleHeight = heightFieldPix * 0.15
		
		image = wx.Image( os.path.join(Utils.getImageFolder(), 'Sprint1.png'), wx.BITMAP_TYPE_PNG )
		imageWidth, imageHeight = image.GetWidth(), image.GetHeight()
		imageScale = float(titleHeight) / float(imageHeight)
		newImageWidth, newImageHeight = int(imageWidth * imageScale), int(imageHeight * imageScale)
		image.Rescale( newImageWidth, newImageHeight, wx.IMAGE_QUALITY_HIGH )
		dc.DrawBitmap( wx.BitmapFromImage(image), xLeft, yTop )
		del image
		newImageWidth += titleHeight / 10
		
		gt.drawToFitDC( dc, xLeft + newImageWidth, yTop, widthFieldPix - newImageWidth, titleHeight )
		yTop += titleHeight * 1.20
		
		# Collect all the sprint and worksheet results information.
		gt = GrowTable(GrowTable.alignCenter, GrowTable.alignTop)
		gt.fromGrid( self.grid )
		
		gt.drawToFitDC( dc, xLeft, yTop, widthFieldPix, heightPix - borderPix - yTop )
		
		# Add a timestamp footer.
		fontSize = heightPix//85
		font = wx.FontFromPixelSize( wx.Size(0,fontSize), wx.FONTFAMILY_SWISS, wx.NORMAL, wx.FONTWEIGHT_NORMAL )
		dc.SetFont( font )
		text = u'Generated {}'.format( datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S') )
		footerTop = heightPix - borderPix + fontSize/2
		dc.DrawText( text, widthPix - borderPix - dc.GetTextExtent(text)[0], footerTop )
		
		# Add branding
		text = u'Powered by PointsRaceMgr'
		dc.DrawText( text, borderPix, footerTop )

	def toHtml( self, html ):
		with tag( html, 'table', {'class':'results'} ):
			with tag( html, 'thead' ):
				with tag( html, 'tr' ):
					for col in xrange(self.grid.GetNumberCols()):
						with tag( html, 'th' ):
							html.write( cgi.escape(self.grid.GetColLabelValue(col)).replace('\n', '<br\>') )
			with tag( html, 'tbody' ):
				for row in xrange(self.grid.GetNumberRows()):
					with tag( html, 'tr', {'class': 'odd'} if row & 1 else {} ):
						for col in xrange(self.grid.GetNumberCols()):
							with tag( html, 'td', {'class': 'numeric'} if col in self.numericCols else {} ):
								html.write( cgi.escape(self.grid.GetCellValue(row, col)) )

########################################################################

class ResultsListFrame(wx.Frame):
	def __init__(self):
		"""Constructor"""
		wx.Frame.__init__(self, None, title="ResultsList", size=(800,600) )
		panel = ResultsList(self)
		panel.refresh()
		self.Show()
 
#----------------------------------------------------------------------
if __name__ == "__main__":
	app = wx.App(False)
	frame = ResultsListFrame()
	app.MainLoop()
