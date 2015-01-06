import wx
import Utils
import Model

class GrowTable( object ):
	def __init__( self ):
		self.table = []
		self.colWidths = []
		self.rowHeights = []
		self.vLines = []
		self.hLines = []
		
	def set( self, row, col, value, highlight=False ):
		self.table += [[] for i in xrange(max(0, row+1 - len(self.table)))]
		self.table[row] += [None for i in xrange(max(0, col+1 - len(self.table[row])))]
		self.table[row][col] = (value, highlight)
		
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
		return fontSize / 5
		
	def getSize( self, dc, fontSize ):
		font, fontBold = self.getFonts( fontSize )
		cellBorder = self.getCellBorder( fontSize )
		self.colWidths = [0] * self.getNumberCols()
		self.rowHeights = [0] * self.getNumberRows()
		height = 0
		for row, r in enumerate(self.table):
			for col, (value, highlight) in enumerate(r):
				vWidth, vHeight, lineHeight = dc.GetMultiLineTextExtent(value, fontBold if highlight else font)
				vWidth += cellBorder * 2
				vHeight += cellBorder * 2
				self.colWidths[col] = max(self.colWidths[col], vWidth)
				self.rowHeights[row] = max(self.rowHeights[row], vHeight)
		return sum( self.colWidths ), sum( self.rowHeights )
	
	def drawTextToFit( self, dc, text, x, y, width, height, font=None ):
		if font:
			dc.SetFont( font )
		fontheight = dc.GetFont().GetPixelSize()[1]
		cellBorder = self.getCellBorder( fontheight )
		tWidth, tHeight, lineHeight = dc.GetMultiLineTextExtent(text, dc.GetFont())
		lines = text.split( '\n' )
		xBorder, yBorder = (width - tWidth) / 2, height - cellBorder - lineHeight*len(lines)
		xRight = x + width - cellBorder
		yTop = y + yBorder
		for line in lines:
			dc.DrawText( line, xRight - dc.GetTextExtent(line)[0], yTop )
			yTop += lineHeight
	
	def drawToFitDC( self, dc, x, y, width, height ):
		self.x = x
		self.y = y
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

		font, fontBold = self.getFonts( fontSize )
		yTop = y
		for row, r in enumerate(self.table):
			xLeft = x
			for col, (value, highlight) in enumerate(r):
				self.drawTextToFit( dc, value, xLeft, yTop, self.colWidths[col], self.rowHeights[row], fontBold if highlight else font )
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
		fontheight = dc.GetFont().GetPixelSize()[1]
		cellBorder = self.getCellBorder( fontheight )
		dc.SetPen( wx.Pen(wx.BLACK, cellBorder / 2 if thick else 1 ) )

def ToPrintout( dc ):
	race = Model.race
	mainWin = Utils.getMainWin()
	
	gt = GrowTable()
	
	# Collect all the sprint and worksheet results information.
	
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
	colAdjust[gridResults] = 2
	
	rowWorksheet = rowCur
	
	colCur = 0
	for grid in [gridBib, gridWorksheet, gridResults]:
		for col in xrange(maxSprints if grid == gridWorksheet else grid.GetNumberCols() - colAdjust.get(grid,0)):
			gt.set( rowCur, colCur, grid.GetColLabelValue(col), True )
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
	
	#---------------------------------------------------------------------------------------
	# Format on the page.
	(widthPix, heightPix) = dc.GetSizeTuple()
	
	# Get a reasonable border.
	borderPix = max(widthPix, heightPix) / 20
	
	widthFieldPix = widthPix - borderPix * 2
	heightFieldPix = heightPix - borderPix * 2
	
	xPix = borderPix
	yPix = borderPix

	# Race Name
	# Category
	# Date, Distance, Number of Laps
	xLeft = xPix
	yTop = xPix
	
	# Draw the main data.
	gt.drawToFitDC( dc, xLeft, yTop, widthFieldPix, heightPix - borderPix - yTop )