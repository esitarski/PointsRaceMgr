import wx
import wx.grid as gridlib

import os
import sys
import operator
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
			list(self.headerNames) +
			[self.resultsGrid.GetColLabelValue(c) for c in xrange(2,self.resultsGrid.GetNumberCols()-1)] +
			[self.summaryGrid.GetColLabelValue(c) for c in xrange(self.summaryGrid.GetNumberCols())]
		)
		
	def updateGrid( self ):
		race = Model.race
		headers = self.getHeaders()
		
		attrs = []
		for col, h in enumerate(headers):
			attr = gridlib.GridCellAttr()
			if col <= 1 or col > len(self.fieldNames):
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
				[unicode(getattr(info,f)) for f in self.fieldNames] +
				[self.resultsGrid.GetCellValue(row, c) for c in xrange(2,self.resultsGrid.GetNumberCols()-1)] +
				[self.summaryGrid.GetCellValue(row, c) for c in xrange(self.summaryGrid.GetNumberCols())]
			)
			content.append( values )
		
		empty = set(c for c in xrange(len(headers)) if all(not values[c] for values in content))
		
		def pluck( a ):
			return [v for c, v in enumerate(a) if c not in empty]
		headers = pluck( headers )
		attrs = pluck( attrs )
		content = [pluck(values) for values in content]
		
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
		
	def refresh( self ):
		self.updateGrid()
		
	def commit( self ):
		pass
			
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
