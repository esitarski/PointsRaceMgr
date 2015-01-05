import  wx
from  ToPrintout import ToPrintout

#----------------------------------------------------------------------

class PointsMgrPrintout(wx.Printout):
	def __init__(self):
		wx.Printout.__init__(self)

	def HasPage(self, page):
		return page == 1

	def GetPageInfo(self):
		return (1,1,1,1)

	def OnPrintPage(self, page):
		dc = self.GetDC()
		ToPrintout( dc )
		return True

