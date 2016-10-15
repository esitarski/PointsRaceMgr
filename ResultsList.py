import wx
import wx.grid as gridlib

import os
import sys
import operator
import xlwt
import re

import Utils
import Model

#--------------------------------------------------------------------------------
class ResultsList(wx.Panel):

	def __init__(self, parent):
		wx.Panel.__init__(self, parent)
		
		self.resultsGrid = parent.GetParent().scoreSheet.worksheet.gridBib
		self.summaryGrid = parent.GetParent().scoreSheet.results.gridResults
		
		self.fieldNames  = Model.RiderInfo.FieldNames
		self.headerNames = Model.RiderInfo.HeaderNames
		
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
		fieldNamesResults = self.fieldNames[:-1]
		
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
				[unicode(getattr(info,f)) for f in fieldNamesResults] +
				[self.resultsGrid.GetCellValue(row, c) for c in xrange(2,self.resultsGrid.GetNumberCols()-1)] +
				[self.summaryGrid.GetCellValue(row, c) for c in xrange(self.summaryGrid.GetNumberCols())]
			)
			content.append( values )
		
		empty = set(c for c in xrange(len(headers))
			if all(not values[c] for values in content) or all(values[c]=='0.0' for values in content))
		
		def pluck( a ):
			return [v for c, v in enumerate(a) if c not in empty]
		headers = pluck( headers )
		attrs = pluck( attrs )
		content = [pluck(values) for values in content]
		
		for c in xrange(len(headers)):
			if all((not row[c] or row[c].endswith(u'.0')) for row in content):
				for row in content:
					row[c] = row[c][:-2]
		
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
