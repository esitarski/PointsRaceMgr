
import wx
import os
import xlwt
import Utils
import Model
import math
import cgi
import base64
from FitSheetWrapper import FitSheetWrapper
from contextlib import contextmanager

#---------------------------------------------------------------------------

@contextmanager
def tag( buf, name, attrs = {} ):
	if isinstance(attrs, basestring) and attrs:
		attrs = { 'class': attrs }
	buf.write( u'<{}>'.format( u' '.join(
			[name] + [u'{}="{}"'.format(attr, value) for attr, value in attrs.iteritems()]
		) ) )
	yield
	buf.write( u'</{}>\n'.format(name) )

brandText = u'Powered by PointsRaceMgr (sites.google.com/site/crossmgrsoftware)'

def ImageToPil( image ):
	"""Convert wx.Image to PIL Image."""
	w, h = image.GetSize()
	data = image.GetData()

	redImage = Image.new("L", (w, h))
	redImage.fromstring(data[0::3])
	greenImage = Image.new("L", (w, h))
	greenImage.fromstring(data[1::3])
	blueImage = Image.new("L", (w, h))
	blueImage.fromstring(data[2::3])

	if image.HasAlpha():
		alphaImage = Image.new("L", (w, h))
		alphaImage.fromstring(image.GetAlphaData())
		pil = Image.merge('RGBA', (redImage, greenImage, blueImage, alphaImage))
	else:
		pil = Image.merge('RGB', (redImage, greenImage, blueImage))
	return pil

def getHeaderFName():
	''' Get the header bitmap if specified and exists, or use a default.  '''
	try:
		graphicFName = Utils.getMainWin().getGraphicFName()
		with open(graphicFName, 'rb') as f:
			pass
		return graphicFName
	except:
		return os.path.join(Utils.getImageFolder(), 'PointsRaceMgr.png')

def getHeaderBitmap():
	''' Get the header bitmap if specified, or use a default.  '''
	if Utils.getMainWin():
		graphicFName = Utils.getMainWin().getGraphicFName()
		extension = os.path.splitext( graphicFName )[1].lower()
		bitmapType = {
			'.gif': wx.BITMAP_TYPE_GIF,
			'.png': wx.BITMAP_TYPE_PNG,
			'.jpg': wx.BITMAP_TYPE_JPEG,
			'.jpeg':wx.BITMAP_TYPE_JPEG }.get( extension, wx.BITMAP_TYPE_ANY )
		try:
			return wx.Bitmap( graphicFName, bitmapType )
		except Exception as e:
			pass
	
	return wx.Bitmap( os.path.join(Utils.getImageFolder(), 'SprintMgr.png'), wx.BITMAP_TYPE_PNG )

def writeHtmlHeader( buf, title ):
	with tag(buf, 'span', {'id': 'idRaceName'}):
		buf.write( unicode(cgi.escape(title).replace('\n', '<br/>\n')) )

class ExportGrid( object ):
	PDFLineFactor = 1.10

	def __init__( self, title, grid ):
		self.title = title
		self.grid = grid
		self.colnames = [grid.GetColLabelValue(c) for c in xrange(grid.GetNumberCols())]
		self.data = [ [grid.GetCellValue(r, c) for r in xrange(grid.GetNumberRows())] for c in xrange(len(self.colnames)) ]
		
		# Trim all empty rows.
		self.numRows = 0
		for col in self.data:
			for iRow, v in enumerate(col):
				if v.strip() and iRow >= self.numRows:
					self.numRows = iRow + 1
		
		for col in self.data:
			del col[self.numRows:]
		
		self.fontName = 'Helvetica'
		self.fontSize = 16
		
		self.leftJustifyCols = {}
		self.rightJustifyCols = {}
		
		for c, n in enumerate(self.colnames):
			if any(n.startswith(p) for p in ('Rank', 'Bib', 'Time', 'Sp', 'Finish', 'Points', '+', 'Existing',)):
				self.rightJustifyCols[c] = True
			else:
				self.leftJustifyCols[c] = True
	
	def _getFont( self, pixelSize = 28, bold = False ):
		return wx.FontFromPixelSize( (0,pixelSize), wx.FONTFAMILY_SWISS, wx.NORMAL,
									 wx.FONTWEIGHT_BOLD if bold else wx.FONTWEIGHT_NORMAL, False )
	
	def _getColSizeTuple( self, dc, font, col ):
		dc.SetFont( font )
		wSpace, hSpace = dc.GetTextExtent( '    ' )
		extents = [ dc.GetMultiLineTextExtent(self.colnames[col]) ]
		extents.extend( dc.GetMultiLineTextExtent(unicode(v), font) for v in self.data[col] )
		return max( e[0] for e in extents ), sum( e[1] for e in extents ) + hSpace/4
	
	def _getDataSizeTuple( self, dc, font ):
		dc.SetFont( font )
		wSpace, hSpace = dc.GetTextExtent( '    ' )
		
		wMax, hMax = 0, 0
		
		# Sum the width of each column.
		for col, c in enumerate(self.colnames):
			w, h = self._getColSizeTuple( dc, font, col )
			wMax += w + wSpace
			hMax = max( hMax, h )
			
		if wMax > 0:
			wMax -= wSpace
		
		return wMax, hMax
	
	def _drawMultiLineText( self, dc, text, x, y ):
		if not text:
			return
		wText, hText,  = dc.GetMultiLineTextExtent( text )
		lineHeightText = dc.GetTextExtent( 'PpYyJj' )[1]
		for line in text.split( '\n' ):
			dc.DrawText( line, x, y )
			y += lineHeightText

	def _getFontToFit( self, widthToFit, heightToFit, sizeFunc, isBold = False ):
		left = 1
		right = max(widthToFit, heightToFit)
		
		while right - left > 1:
			mid = (left + right) / 2.0
			font = self._getFont( mid, isBold )
			widthText, heightText = sizeFunc( font )
			if widthText <= widthToFit and heightText <= heightToFit:
				left = mid
			else:
				right = mid - 1
		
		return self._getFont( left, isBold )
	
	def drawToFitDC( self, dc ):
		# Get the dimensions of what we are printing on.
		(widthPix, heightPix) = dc.GetSizeTuple()
		
		# Get a reasonable border.
		borderPix = max(widthPix, heightPix) / 20
		
		widthFieldPix = widthPix - borderPix * 2
		heightFieldPix = heightPix - borderPix * 2
		
		xPix = borderPix
		yPix = borderPix
		
		graphicWidth = 0
		graphicHeight = heightPix * 0.15
		graphicBorder = 0
		qrWidth = 0
		
		# Draw the graphic.
		bitmap = getHeaderBitmap()
		bmWidth, bmHeight = bitmap.GetWidth(), bitmap.GetHeight()
		graphicHeight = int(heightPix * 0.12)
		graphicWidth = int(float(bmWidth) / float(bmHeight) * graphicHeight)
		graphicBorder = int(graphicWidth * 0.15)

		# Rescale the graphic to the correct size.
		# We cannot use a GraphicContext because it does not support a PrintDC.
		image = bitmap.ConvertToImage()
		image.Rescale( graphicWidth, graphicHeight, wx.IMAGE_QUALITY_HIGH )
		if dc.GetDepth() == 8:
			image = image.ConvertToGreyscale()
		bitmap = image.ConvertToBitmap( dc.GetDepth() )
		dc.DrawBitmap( bitmap, xPix, yPix )
		image, bitmap = None, None
		
		# Draw the title.
		def getTitleSize( font )
			dc.SetFont( font )
			return dc.GetMultiLineTextExtent(self.title)
		font = self._getFontToFit( widthFieldPix - graphicWidth - graphicBorder - qrWidth, graphicHeight, getTitleSize, True )
		dc.SetFont( font )
		self._drawMultiLineText( dc, self.title, xPix + graphicWidth + graphicBorder, yPix )
		yPix += graphicHeight + graphicBorder
		
		heightFieldPix = heightPix - yPix - borderPix
		
		# Draw the table.
		font = self._getFontToFit( widthFieldPix, heightFieldPix, lambda font: self._getDataSizeTuple(dc, font) )
		dc.SetFont( font )
		wSpace, hSpace  = dc.GetTextExtent( u'    ' )
		textHeight = hSpace
		
		# Get the max height per row.
		rowHeight = [0] * (self.numRows + 1)
		for r in xrange(self.numRows):
			rowHeight[r] = max( dc.GetMultiLineTextExtent(self.grid.GetCellValue(r, c))[1] for c in xrange(len(self.colnames)))
		
		yPixTop = yPix
		yPixMax = yPix
		for col, c in enumerate(self.colnames):
			colWidth = self._getColSizeTuple( dc, font, col )[0]
			yPix = yPixTop
			w, h = dc.GetMultiLineTextExtent( c )
			if col in self.leftJustifyCols:
				self._drawMultiLineText( dc, unicode(c), xPix, yPix )					# left justify
			else:
				self._drawMultiLineText( dc, unicode(c), xPix + colWidth - w, yPix )	# right justify
			yPix += h + hSpace/4
			if col == 0:
				yLine = yPix - hSpace/8
				dc.DrawLine( borderPix, yLine, widthPix - borderPix, yLine )
				for r in rowHeight:
					yLine += r
					dc.DrawLine( borderPix, yLine, widthPix - borderPix, yLine )
					
			for r, v in enumerate(self.data[col]):
				vStr = unicode(v)
				if vStr:
					w, h = dc.GetMultiLineTextExtent( vStr )
					if col in self.leftJustifyCols:
						self._drawMultiLineText( dc, vStr, xPix, yPix )					# left justify
					else:
						self._drawMultiLineText( dc, vStr, xPix + colWidth - w, yPix )	# right justify
				yPix += rowHeight[r]
			yPixMax = max(yPixMax, yPix)
			xPix += colWidth + wSpace
				
		# Put CrossMgr branding at the bottom of the page.
		font = self._getFont( borderPix // 5, False )
		dc.SetFont( font )
		w, h = dc.GetMultiLineTextExtent( brandText )
		self._drawMultiLineText( dc, brandText, borderPix, heightPix - borderPix + h )
	
	def toExcelSheet( self, sheet ):
		''' Write the contents of the grid to an xlwt excel sheet. '''
		titleStyle = xlwt.XFStyle()
		titleStyle.font.bold = True
		titleStyle.font.height += titleStyle.font.height / 2
		
		rowTop = 0
		if self.title:
			for line in self.title.split('\n'):
				sheet.write(rowTop, 0, line, titleStyle)
				rowTop += 1
			rowTop += 1
		
		sheetFit = FitSheetWrapper( sheet )
		
		# Write the colnames and data.
		headerStyleLeft = xlwt.XFStyle()
		headerStyleLeft.borders.bottom = xlwt.Borders.MEDIUM
		headerStyleLeft.font.bold = True
		headerStyleLeft.alignment.horz = xlwt.Alignment.HORZ_LEFT
		headerStyleLeft.alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT

		headerStyleRight = xlwt.XFStyle()
		headerStyleRight.borders.bottom = xlwt.Borders.MEDIUM
		headerStyleRight.font.bold = True
		headerStyleRight.alignment.horz = xlwt.Alignment.HORZ_RIGHT
		headerStyleRight.alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT

		styleLeft = xlwt.XFStyle()
		styleLeft.alignment.horz = xlwt.Alignment.HORZ_LEFT
		styleLeft.alignment.wrap = True
		styleLeft.alignment.vert = xlwt.Alignment.VERT_TOP
			
		styleRight = xlwt.XFStyle()
		styleRight.alignment.horz = xlwt.Alignment.HORZ_RIGHT
		styleRight.alignment.wrap = True
		styleRight.alignment.vert = xlwt.Alignment.VERT_TOP
			
		rowMax = 0
		for col, c in enumerate(self.colnames):
			sheetFit.write( rowTop, col, c, headerStyleLeft if col in self.leftJustifyCols else headerStyleRight, bold=True )
			style = styleLeft if col in self.leftJustifyCols else styleRight
			for row, v in enumerate(self.data[col]):
				rowCur = rowTop + 1 + row
				if rowCur > rowMax:
					rowMax = rowCur
				sheetFit.write( rowCur, col, v, style, bold=True )
				
		# Add branding at the bottom of the sheet.
		style = xlwt.XFStyle()
		style.alignment.horz = xlwt.Alignment.HORZ_LEFT
		sheet.write( rowMax + 2, 0, brandText, style )
		
	def toHtml( self, buf ):
		''' Write the contents to the buffer in HTML format. '''
		writeHtmlHeader( buf, self.title )
		
		with tag(buf, 'table', {'class': 'results'} ):
			with tag(buf, 'thead'):
				with tag(buf, 'tr'):
					for col in self.colnames:
						with tag(buf, 'th'):
							buf.write( cgi.escape(col).replace('\n', '<br/>\n') )
			with tag(buf, 'tbody'):
				for row in xrange(max(len(d) for d in self.data)):
					with tag(buf, 'tr', {'class':'odd'} if row & 1 else {}):
						for col in xrange(len(self.colnames)):
							with tag(buf, 'td', {'class':'rightAlign'} if col not in self.leftJustifyCols else {}):
								try:
									buf.write( cgi.escape(self.data[col][row]).replace('\n', '<br/>\n') )
								except IndexError:
									buf.write( u'&nbsp;' )									
		return buf
			
if __name__ == '__main__':
	pass
