import  wx
from  ToPrintout import ToPrintout
import Utils

#----------------------------------------------------------------------

class PointsMgrPrintout(wx.Printout):
	def __init__(self):
		wx.Printout.__init__(self)

	def HasPage(self, page):
		return page in (1, 2)

	def GetPageInfo(self):
		return (1,2,1,2)

	def OnPrintPage(self, page):
		dc = self.GetDC()
		if page == 1:
			Utils.getMainWin().GetParent().GetParent().resultsList.toPrintout( dc )
		else:
			ToPrintout( dc )
		return True

