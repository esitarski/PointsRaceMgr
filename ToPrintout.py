import wx
import os
import math
import datetime
import Utils
import Model

class GrowTable( object ):
	alignLeft = 1<<0
	alignCentre = alignCenter = 1<<1
	alignRight = 1<<2
	
	alignTop = 1<<3
	alignBottom = 1<<4
	
	def __init__( self, alignHorizontal=alignCentre, alignVertical=alignCentre, cellBorder=True ):
		self.table = []
		self.colWidths = []
		self.rowHeights = []
		self.vLines = []
		self.hLines = []
		self.alignHorizontal = alignHorizontal
		self.alignVertical = alignVertical
		self.cellBorder = cellBorder
		self.width = None
		self.height = None
		
	def set( self, row, col, value, highlight=False, align=alignRight ):
		self.table += [[] for i in xrange(max(0, row+1 - len(self.table)))]
		self.table[row] += [None for i in xrange(max(0, col+1 - len(self.table[row])))]
		self.table[row][col] = (value, highlight, align)
		return row, col
		
	def vLine( self, col, rowStart, rowEnd, thick = False ):
		self.vLines.append( (col, rowStart, rowEnd, thick) )
		
	def hLine( self, row, colStart, colEnd, thick = False ):
		self.hLines.append( (row, colStart, colEnd, thick) )
		
	def getNumberCols( self ):
		return max(len(r) for r in self.table)
		
	def getNumberRows( self ):
		return len(self.table)
		
	def getFonts( self, fontSize ):
		font = wx.FontFromPixelSize( wx.Size(0,fontSize), wx.FONTFAMILY_SWISS, wx.NORMAL, wx.FONTWEIGHT_NORMAL )
		fontBold = wx.FontFromPixelSize( wx.Size(0,fontSize), wx.FONTFAMILY_SWISS, wx.NORMAL, wx.FONTWEIGHT_BOLD )
		return font, fontBold
		
	def getCellBorder( self, fontSize ):
		return max(1, fontSize // 5) if self.cellBorder else 0
		
	def getSize( self, dc, fontSize ):
		font, fontBold = self.getFonts( fontSize )
		cellBorder = self.getCellBorder( fontSize )
		self.colWidths = [0] * self.getNumberCols()
		self.rowHeights = [0] * self.getNumberRows()
		height = 0
		for row, r in enumerate(self.table):
			for col, (value, highlight, align) in enumerate(r):
				vWidth, vHeight, lineHeight = dc.GetMultiLineTextExtent(value, fontBold if highlight else font)
				vWidth += cellBorder * 2
				vHeight += cellBorder * 2
				self.colWidths[col] = max(self.colWidths[col], vWidth)
				self.rowHeights[row] = max(self.rowHeights[row], vHeight)
		return sum( self.colWidths ), sum( self.rowHeights )
	
	def drawTextToFit( self, dc, text, x, y, width, height, align, font=None ):
		if font and font != dc.GetFont():
			dc.SetFont( font )
		fontheight = dc.GetFont().GetPixelSize()[1]
		cellBorder = self.getCellBorder( fontheight )
		tWidth, tHeight, lineHeight = dc.GetMultiLineTextExtent(text, dc.GetFont())
		lines = text.split( '\n' )
		xBorder, yBorder = (width - tWidth) / 2, height - cellBorder - lineHeight*len(lines)
		xLeft = x + cellBorder
		xRight = x + width - cellBorder
		yTop = y + yBorder
		for line in lines:
			if align == self.alignRight:
				dc.DrawText( line, xRight - dc.GetTextExtent(line)[0], yTop )
			elif align == self.alignLeft:
				dc.DrawText( line, xLeft, yTop )
			else:
				dc.DrawText( line, x + (width - dc.GetTextExtent(line)[0]) / 2, yTop )
				
			yTop += lineHeight
	
	def drawToFitDC( self, dc, x, y, width, height ):
		self.penThin = None
		self.penThick = None
		
		fontSizeLeft, fontSizeRight = 2, 512
		for i in xrange(20):
			fontSize = (fontSizeLeft + fontSizeRight) // 2
			tWidth, tHeight = self.getSize( dc, fontSize )
			if tWidth < width and tHeight < height:
				fontSizeLeft = fontSize
			else:
				fontSizeRight = fontSize
			if fontSizeLeft == fontSizeRight:
				break
		
		fontSize = fontSizeLeft
		tWidth, tHeight = self.getSize( dc, fontSize )
		self.width, self.height = tWidth, tHeight
		
		# Align the entire table in the space.
		if self.alignHorizontal == self.alignCentre:
			x += (width - tWidth) // 2
		elif self.alignHorizontal == self.alignRight:
			x += width - tWidth

		if self.alignVertical == self.alignCentre:
			y += (height - tHeight) // 2
		elif self.alignVertical == self.alignBottom:
			y += height - tHeight
			
		self.x = x
		self.y = y

		font, fontBold = self.getFonts( fontSize )
		yTop = y
		for row, r in enumerate(self.table):
			xLeft = x
			for col, (value, highlight, align) in enumerate(r):
				self.drawTextToFit( dc, value, xLeft, yTop, self.colWidths[col], self.rowHeights[row], align, fontBold if highlight else font )
				xLeft += self.colWidths[col]
			yTop += self.rowHeights[row]
			
		for col, rowStart, rowEnd, thick in self.vLines:
			xLine = x + sum(self.colWidths[:col])
			self.setPen( dc, thick )
			dc.DrawLine( xLine, y + sum(self.rowHeights[:rowStart]), xLine, y + sum(self.rowHeights[:rowEnd]) )
	
		for row, colStart, colEnd, thick in self.hLines:
			yLine = self.y + sum(self.rowHeights[:row])
			self.setPen( dc, thick )
			dc.DrawLine( x + sum(self.colWidths[:colStart]), yLine, x + sum(self.colWidths[:colEnd]), yLine )
			
	def setPen( self, dc, thick = False ):
		if not self.penThin:
			self.penThin = wx.Pen( wx.BLACK, 1, wx.SOLID )
			fontheight = dc.GetFont().GetPixelSize()[1]
			cellBorder = self.getCellBorder( fontheight )
			width = cellBorder / 2
			self.penThick = wx.Pen( wx.BLACK, width, wx.SOLID )
		newPen = self.penThick if thick else self.penThin
		if newPen != dc.GetPen():
			dc.SetPen( newPen )

def ToPrintout( dc ):
	race = Model.race
	mainWin = Utils.getMainWin()
	
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
	rowCur = 0
	rowCur = gt.set( rowCur, 0, race.name, highlight=True, align=GrowTable.alignLeft )[0] + 1
	rowCur = gt.set( rowCur, 0, race.category, True, align=GrowTable.alignLeft )[0] + 1
	rowCur = gt.set( rowCur, 0, u'{} Laps, {} Sprints, {} km'.format(race.laps, race.getNumSprints(), race.getDistance()), highlight=True, align=GrowTable.alignLeft )[0] + 1
	rowCur = gt.set( rowCur, 0, race.date.strftime('%Y-%m-%d'), highlight=True, align=GrowTable.alignLeft )[0] + 1
	
	if race.communique:
		rowCur = gt.set( rowCur, 0, u'Communiqu\u00E9: {}'.format(race.communique), align=GrowTable.alignRight )[0] + 1
	rowCur = gt.set( rowCur, 0, u'Approved by:________', align=GrowTable.alignRight )[0] + 1
	
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
	
	maxSprints = race.laps / race.sprintEvery
	
	gridPoints = mainWin.sprints.gridPoints
	gridSprint = mainWin.sprints.gridSprint
	
	colAdjust = {}
	colAdjust[gridPoints] = 1
	
	# First get the sprint results.
	rowCur = 0
	colCur = 0
	for grid in [gridPoints, gridSprint]:
		for col in xrange(maxSprints if grid == gridSprint else grid.GetNumberCols() - colAdjust.get(grid,0)):
			gt.set( rowCur, colCur, grid.GetColLabelValue(col), True )
			colCur += 1
	rowCur += 1
	gt.hLine( rowCur-1, 1, gt.getNumberCols(), False )
	gt.hLine( rowCur, 1, gt.getNumberCols(), True )
	
	# Find the maximum number of places for points.
	for rowMax in xrange(gridPoints.GetNumberRows()):
		if gridPoints.GetCellValue(rowMax, 2) == u'0':
			break
	
	# Add the values from the points and sprint tables.
	for row in xrange(rowMax):
		colCur = 0
		for grid in [gridPoints, gridSprint]:
			for col in xrange(maxSprints if grid == gridSprint else grid.GetNumberCols() - colAdjust.get(grid,0)):
				gt.set( rowCur, colCur, grid.GetCellValue(row,col) )
				colCur += 1
		rowCur += 1
		gt.hLine( rowCur, 1, colCur )
		
	for col in xrange( 1, colCur+1 ):
		gt.vLine( col, 0, rowCur )
	upperColMax = colCur
	
	# Collect the worksheet and results information
	gridBib = mainWin.worksheet.gridBib
	gridWorksheet = mainWin.worksheet.gridWorksheet
	gridResults = mainWin.results.gridResults
	
	colAdjust[gridBib] = 1
	
	rowWorksheet = rowCur
	
	colCur = 0
	for grid in [gridBib, gridWorksheet, gridResults]:
		for col in xrange(maxSprints if grid == gridWorksheet else grid.GetNumberCols() - colAdjust.get(grid,0)):
			gt.set( rowCur, colCur, grid.GetColLabelValue(col), highlight=True, align=GrowTable.alignCenter )
			colCur += 1
	rowCur += 1
	
	gt.hLine( rowCur-1, 0, colCur, True )
	gt.hLine( rowCur, 0, colCur, True )
	
	# Add the values from the bib, worksheet and results tables.
	for row in xrange(gridWorksheet.GetNumberRows()):
		colCur = 0
		for grid in [gridBib, gridWorksheet, gridResults]:
			if rowMax <= grid.GetNumberRows():
				for col in xrange(maxSprints if grid == gridWorksheet else grid.GetNumberCols() - colAdjust.get(grid,0)):
					gt.set( rowCur, colCur, grid.GetCellValue(row,col) )
					colCur += 1
		rowCur += 1
		gt.hLine( rowCur, 0, colCur )
		
	for col in xrange( 0, gt.getNumberCols()+1 ):
		gt.vLine( col, rowWorksheet, rowCur )
	
	gt.vLine( 3, 0, gt.getNumberRows(), True )
	gt.vLine( upperColMax, 0, gt.getNumberRows(), True )
	
	# Format the notes assuming a minimum readable size.
	notesHeight = 0
	gtNotes = None
	
	lines = [line for line in race.notes.split(u'\n') if line.strip()]
	if lines:
		gtNotes = GrowTable()
		maxLinesPerCol = 10
		numCols = int(math.ceil(len(lines) / float(maxLinesPerCol)))
		numRows = int(math.ceil(len(lines) / float(numCols)))
		rowCur, colCur = 0, 0
		for i, line in enumerate(lines):
			gtNotes.set( rowCur, colCur*2, u'{}.'.format(i+1), align=GrowTable.alignRight )
			gtNotes.set( rowCur, colCur*2+1, u'{}    '.format(line.strip()), align=GrowTable.alignLeft )
			rowCur += 1
			if rowCur == numRows:
				rowCur = 0
				colCur += 1
		lineHeight = heightPix // 65
		notesHeight = (lineHeight+1) * numRows
	
	gt.drawToFitDC( dc, xLeft, yTop, widthFieldPix, heightPix - borderPix - yTop - notesHeight )
	
	# Use any remaining space on the page for the notes.
	if gtNotes:
		notesTop = yTop + gt.height + lineHeight
		gtNotes.drawToFitDC( dc, xLeft, notesTop, widthFieldPix, heightPix - borderPix - notesTop )
	
	# Add a timestamp footer.
	fontSize = heightPix//85
	font = wx.FontFromPixelSize( wx.Size(0,fontSize), wx.FONTFAMILY_SWISS, wx.NORMAL, wx.FONTWEIGHT_NORMAL )
	dc.SetFont( font )
	text = u'Generated {}'.format( datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S') )
	dc.DrawText( text, widthPix - borderPix - dc.GetTextExtent(text)[0], heightPix - borderPix + fontSize/4 )
	
	# Add branding
	text = u'Powered by PointsRaceMgr'
	dc.DrawText( text, borderPix, heightPix - borderPix + fontSize/4 )

