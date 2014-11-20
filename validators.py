# Created by Papagrigoriou Nikos, Thessaloniki, Greece
# nikos@papagrigoriou.gr - http://www.papagrigoriou.gr
# What is it: two classes for validating wxTextCtrl
# Your are free to use this program

from wxPython.wx import *
import string, sys

#if sys.platform == 'win32':
#    import win32api

class IntValidator(wxPyValidator):
    def __init__(self, minimum=0, maximum=None,
                       minstrict=TRUE, maxstrict=TRUE,
                       valid=wxColour(red=255,green=255,blue=255),
                       invalid=wxColour(red=216,green=191,blue=216)):
        wxPyValidator.__init__(self)
        
        self.minimum = minimum
        self.maximum = maximum
        if self.minimum >= self.maximum:
            self.maximum = None
        self.minstrict = minstrict
        self.maxstrict = maxstrict
        self.valid = valid
        self.invalid = invalid
        
        EVT_CHAR(self, self.OnChar)
        
    def Clone(self):
        return IntValidator(self.minimum, self.maximum,
                            self.minstrict, self.maxstrict,
                            self.valid, self.invalid)
    
    def TransferToWindow(self):
        return TRUE

    def TransferFromWindow(self):
        return TRUE
    
    def Validate(self, win):
        #print '*** IntValidator ***'
        tc = wxPyTypeCast(self.GetWindow(), "wxTextCtrl")
        val = tc.GetValue()
        #print val
        if tc.IsEnabled() and val == '':
            tc.SetBackgroundColour(self.invalid)
            tc.Refresh()
            return FALSE
        elif tc.IsEnabled() and val != '':
            for x in val:
                if x not in string.digits:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    return FALSE
            if not (self.minimum <= int(val) <= self.maximum):
                tc.SetBackgroundColour(self.invalid)
                tc.Refresh()
                return FALSE
        
        tc.SetBackgroundColour(self.valid)
        tc.Refresh()
        
        return TRUE

    def OnChar(self, event):
        key = event.KeyCode()
        tc = wxPyTypeCast(self.GetWindow(), "wxTextCtrl")
        val = tc.GetValue()
        inpoint = tc.GetInsertionPoint()

        if key == 8:
            if inpoint == 0:
                event.Skip()
                return
            elif inpoint == len(val):
                number = val[:-1]
                if number == '':
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return
                else:
                    number = int(number)
                
                if self.IsInRange(number) == 1:
                    tc.SetBackgroundColour(self.valid)
                    tc.Refresh()
                    event.Skip()
                    return
                elif self.IsInRange(number) == 0:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return
            else:
                val1 = val[:inpoint-1]
                val2 = val[inpoint:]
                number = int(val1 + val2)
                
                if self.IsInRange(number) == 1:
                    tc.SetBackgroundColour(self.valid)
                    tc.Refresh()
                    event.Skip()
                    return
                elif self.IsInRange(number) == 0:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return

            
        if key == WXK_DELETE:
            if inpoint == 0:
                number = val[1:]
                if number == '':
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return
                else:
                    number = int(number)
                
                if self.IsInRange(number) == 1:
                    tc.SetBackgroundColour(self.valid)
                    tc.Refresh()
                    event.Skip()
                    return
                elif self.IsInRange(number) == 0:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return
            elif inpoint == len(val):
                event.Skip()
                return
            else:
                val1 = val[:inpoint]
                val2 = val[inpoint+1:]
                number = int(val1 + val2)
                
                if self.IsInRange(number) == 1:
                    tc.SetBackgroundColour(self.valid)
                    tc.Refresh()
                    event.Skip()
                    return
                elif self.IsInRange(number) == 0:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return

        
        if key < WXK_SPACE or key > 255:
            event.Skip()
            return

        
        if chr(key) in string.digits:
            if self.IsAllSelected(): # first check if the whole text is selected
                number = int(chr(key))
            elif self.IsAnySelected(): # check if any text is selected
                selected = tc.GetSelection()
                val1 = val[:min(selected)]
                val2 = val[max(selected):]
                number = int(val1 + chr(key) + val2)
            elif inpoint == 0:
                number = int(chr(key) + val)
            elif inpoint == len(val):
                number = int(val + chr(key))
            else:
                val1 = val[:inpoint]
                val2 = val[inpoint:]
                number = int(val1 + chr(key) + val2)

            if self.IsInRange(number) == 1:
                tc.SetBackgroundColour(self.valid)
                tc.Refresh()
                event.Skip()
                return
            elif self.IsInRange(number) == 0:
                tc.SetBackgroundColour(self.invalid)
                tc.Refresh()
                event.Skip()
                return
        
        if not wxValidator_IsSilent():
            if sys.platform == 'win32':
                win32api.Beep(10, 10)
            else:
                wxBell()

        return
    
    def IsAnySelected(self):
        tc = wxPyTypeCast(self.GetWindow(), "wxTextCtrl")
        first, last = tc.GetSelection()
        if first == last:
            return 0
        else:
            return 1
    
    def IsAllSelected(self):
        tc = wxPyTypeCast(self.GetWindow(), "wxTextCtrl")
        selected = tc.GetSelection()
        if not self.IsAnySelected():
            return 0
        elif (max(selected) == len(tc.GetValue())) and \
             (min(selected) == 0 ):
            return 1
        else:
            return 0
    
    def IsInRange(self, num):
        if  self.minimum <= num <= self.maximum:
            return 1
        elif (num < self.minimum and not self.minstrict) or \
             (num > self.maximum and not self.maxstrict):
            return 0
        else:
            return -1
        
class FloatValidator(wxPyValidator):
    def __init__(self, minimum=None, maximum=None,
                       minstrict=TRUE, maxstrict=TRUE,
                       zero=TRUE,
                       valid=wxColour(red=255,green=255,blue=255),
                       invalid=wxColour(red=216,green=191,blue=216)):
        wxPyValidator.__init__(self)
        
        self.minimum = minimum
        self.maximum = maximum
        if self.minimum >= self.maximum:
            self.maximum = None
        self.minstrict = minstrict
        self.maxstrict = maxstrict
        self.zero = zero
        self.valid = valid
        self.invalid = invalid
        self.stringList = string.digits + '-.'
        
        EVT_CHAR(self, self.OnChar)
    
    def Clone(self):
        return FloatValidator(self.minimum, self.maximum,
                              self.minstrict, self.maxstrict,
                              self.zero,
                              self.valid, self.invalid)
    
    def TransferToWindow(self):
        return TRUE

    def TransferFromWindow(self):
        return TRUE
    
    def Validate(self, win):
        #print '*** Float Validator ***'
        tc = wxPyTypeCast(self.GetWindow(), "wxTextCtrl")
        val = tc.GetValue()
        #print val
        if tc.IsEnabled() and val == '':
            tc.SetBackgroundColour(self.invalid)
            tc.Refresh()
            return FALSE
        elif tc.IsEnabled() and val != '':
            for x in val:
                if x not in self.stringList:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    return FALSE
            if not (self.minimum <= float(val) <= self.maximum):
                tc.SetBackgroundColour(self.invalid)
                tc.Refresh()
                return FALSE

        tc.SetBackgroundColour(self.valid)
        tc.Refresh()
        
        return TRUE

    def OnChar(self, event):
        key = event.KeyCode()
        tc = wxPyTypeCast(self.GetWindow(), "wxTextCtrl")
        val = tc.GetValue()
        inpoint = tc.GetInsertionPoint()
        
        if key == 8:
            if inpoint == 0:
                event.Skip()
                return
            elif inpoint == len(val):
                number = val[:-1]
                if number in ['', '-', '.']:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return
                else:
                    number = float(number)
                
                if self.IsInRange(number) == 1:
                    tc.SetBackgroundColour(self.valid)
                    tc.Refresh()
                    event.Skip()
                    return
                elif self.IsInRange(number) == 0:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return
            else:
                val1 = val[:inpoint-1]
                val2 = val[inpoint:]
                number = float(val1 + val2)

                if self.IsInRange(number) == 1:
                    tc.SetBackgroundColour(self.valid)
                    tc.Refresh()
                    event.Skip()
                    return
                elif self.IsInRange(number) == 0:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return
                        
        if key == WXK_DELETE:
            if inpoint == 0:
                number = val[1:]
                if number in ['', '-', '.']:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return
                else:
                    number = float(number)
                
                if self.IsInRange(number) == 1:
                    tc.SetBackgroundColour(self.valid)
                    tc.Refresh()
                    event.Skip()
                    return
                elif self.IsInRange(number) == 0:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return
            elif inpoint == len(val):
                event.Skip()
                return
            else:
                val1 = val[:inpoint]
                val2 = val[inpoint+1:]
                number = float(val1 + val2)
                
                if self.IsInRange(number) == 1:
                    tc.SetBackgroundColour(self.valid)
                    tc.Refresh()
                    event.Skip()
                    return
                elif self.IsInRange(number) == 0:
                    tc.SetBackgroundColour(self.invalid)
                    tc.Refresh()
                    event.Skip()
                    return

        if key < WXK_SPACE or key > 255:
            event.Skip()
            return

        if chr(key) == '.':
            if inpoint == 0:
                event.Skip()
                return
            elif inpoint == len(val) and '.' not in val:
                event.Skip()
                return
            elif '.' not in val:
                event.Skip()
                return
                
        if chr(key) == '-':
            if inpoint == 0 and self.minimum < 0:
                event.Skip()
                return
        
        if chr(key) in string.digits:
            if inpoint == 0:
                number = float(chr(key) + val)
            elif inpoint == len(val):
                number = float(val + chr(key))
            else:
                val1 = val[:inpoint]
                val2 = val[inpoint:]
                number = float(val1 + chr(key) + val2)
            
            if self.IsInRange(number) == 1:
                tc.SetBackgroundColour(self.valid)
                tc.Refresh()
                event.Skip()
                return
            elif self.IsInRange(number) == 0:
                tc.SetBackgroundColour(self.invalid)
                tc.Refresh()
                event.Skip()
                return
        
        if not wxValidator_IsSilent():
            if sys.platform == 'win32':
                win32api.Beep(10, 10)
            else:
                wxBell()

        return
        
    def IsInRange(self, num):
        if num == 0.0 and not self.zero and self.minstrict:
            return -1
        elif num == 0.0 and not self.zero and not self.minstrict:
            return 0
        elif self.minimum <= num <= self.maximum:
            return 1
        elif (num < self.minimum and not self.minstrict) or \
             (num > self.maximum and not self.maxstrict):
            return 0


class MyDialog(wxDialog):
    def __init__(self, parent, title):
        wxDialog.__init__(self, parent, -1, title, wxPoint(100,100))
        
        iv1 = IntValidator(minimum=10, maximum=100,
                           minstrict=FALSE, maxstrict=TRUE)
        iv2 = IntValidator(maximum=100, minstrict=FALSE)
        
        wxTextCtrl(self, -1, pos=wxPoint(10,10), validator=iv1)
        wxStaticText(self, -1,
                     'Int: min=10, max=100, minstrict=FALSE, maxstrict=TRUE',
                     pos=wxPoint(100,10))
        
        wxTextCtrl(self, -1, pos=wxPoint(10,40), validator=iv2)
        wxStaticText(self, -1, 'Int: max=100, minstrict=FALSE',
                               pos=wxPoint(100,40))
        
        flt1 = FloatValidator(minimum=-213, maximum=200,
                              minstrict=FALSE, maxstrict=FALSE, zero=FALSE)
        flt2 = FloatValidator(minimum=100, maximum=200,
                              minstrict=FALSE, maxstrict=FALSE)
        
        wxTextCtrl(self, -1, pos=wxPoint(10,70), validator=flt1)
        wxStaticText(self, -1,
                     'Float: min=-213, max=200, minstrict=FALSE, '
                     'maxstrict=FALSE, zero=FALSE',
                     pos=wxPoint(100,70))
        
        wxTextCtrl(self, -1, pos=wxPoint(10,100), validator=flt2)
        wxStaticText(self, -1,
                     'Float: min=100, max=200, minstrict=FALSE, '
                     'maxstrict=FALSE',
                     pos=wxPoint(100,100))

        wxButton(self, wxID_OK, 'OK', pos=wxPoint(10,200))
        wxButton(self, wxID_CANCEL, 'Cancel', pos=wxPoint(100,200))
        
        self.Fit()
        
class TestFrame(wxFrame):
    def __init__(self, parent):
        wxFrame.__init__(self, parent, -1, "Testing Validators...",
                               pos=wxPoint(100,0),
                               size=wxSize(100,50))
        wxButton(self, 25, 'Start Dialog')
        EVT_BUTTON(self, 25, self.OnClick)
    
    def OnClick(self, event):
        dlg = MyDialog(self, "Testing Validators...")
        val = dlg.ShowModal()
        if val == wxID_OK:
            print 'Everything was OK.'
        dlg.Destroy


if __name__ == '__main__':
    app = wxPySimpleApp()
    frame = TestFrame(None)
    frame.Show(true)
    app.MainLoop()
