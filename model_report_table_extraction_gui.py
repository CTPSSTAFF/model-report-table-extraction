# Prototype Python code for GUI for tool to extract named tables from model-generated PRN report file
#
# NOTES: 
#   1. This script was written to run under Python 3.8.x
#   2. This script relies upon the wxPython librarybeing installed
#   3. To install wxPython under Python 3.8.x:
#       <Python_installation_folder>/python.exe -m pip install wxPython
#
# Original Author: Benjamin Krepp (for workscope-exhibit-tool project)
# Date: 1 August, 21 August 2018
# Adaptor: David Knudsen (for model-report-table-extraction project)
# Date: 23 February 2021
#

import wx, wx.html, sys
from model_report_table_extraction import main

# Code for the application's GUI begins here.
#
aboutText = """<p>Help text for this program is TBD.<br>
This program is running on version %(wxpy)s of <b>wxPython</b> and %(python)s of <b>Python</b>.
See <a href="http://wiki.wxpython.org">wxPython Wiki</a></p>""" 

class HtmlWindow(wx.html.HtmlWindow):
    def __init__(self, parent, id, size=(600,400)):
        wx.html.HtmlWindow.__init__(self,parent, id, size=size)
        if "gtk2" in wx.PlatformInfo:
            self.SetStandardFonts()
    # end_def __init__()

    def OnLinkClicked(self, link):
        wx.LaunchDefaultBrowser(link.GetHref())
    # end_def OnLinkClicked()
# end_class HtmlWindow

class AboutBox(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, None, -1, "About the Model Report Table Extraction Tool",
                           style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|wx.TAB_TRAVERSAL)
        hwin = HtmlWindow(self, -1, size=(400,200))
        vers = {}
        vers["python"] = sys.version.split()[0]
        vers["wxpy"] = wx.VERSION_STRING
        hwin.SetPage(aboutText % vers)
        btn = hwin.FindWindowById(wx.ID_OK)
        irep = hwin.GetInternalRepresentation()
        hwin.SetSize((irep.GetWidth()+25, irep.GetHeight()+10))
        self.SetClientSize(hwin.GetSize())
        self.CentreOnParent(wx.BOTH)
        self.SetFocus()
    # end_def __init__()
# end_class AboutBox

# This is the class for the main GUI itself.
class Frame(wx.Frame):
    # Name of input Excel file
    prnFileName = ''
    # Name of directory into which generated HTML files will be written
    outputDirName = ''
    def __init__(self, title):
        wx.Frame.__init__(self, None, title=title, pos=(250,250), size=(800,450),
                          style=wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        menuBar = wx.MenuBar()
        menu = wx.Menu()
        m_exit = menu.Append(wx.ID_EXIT, "E&xit\tAlt-X", "Close window and exit program.")
        self.Bind(wx.EVT_MENU, self.OnClose, m_exit)
        menuBar.Append(menu, "&File")
        menu = wx.Menu()
        m_about = menu.Append(wx.ID_ABOUT, "&About", "Information about this program")
        self.Bind(wx.EVT_MENU, self.OnAbout, m_about)
        menuBar.Append(menu, "&Help")
        self.SetMenuBar(menuBar)
        
        self.statusbar = self.CreateStatusBar()

        panel = wx.Panel(self)
        box = wx.BoxSizer(wx.VERTICAL)
        box.AddSpacer(20)
              
        m_select_file = wx.Button(panel, wx.ID_ANY, "Select model report PRN file")
        m_select_file.Bind(wx.EVT_BUTTON, self.OnSelectFile)
        box.Add(m_select_file, 0, wx.CENTER)
        box.AddSpacer(20)
        
        m_select_output_dir = wx.Button(panel, wx.ID_ANY, "Specify output directory")
        m_select_output_dir.Bind(wx.EVT_BUTTON, self.OnSelectOutputDir)
        box.Add(m_select_output_dir, 0, wx.CENTER)
        box.AddSpacer(20)       
        
        m_generate = wx.Button(panel, wx.ID_ANY, "Break out separate table files")
        m_generate.Bind(wx.EVT_BUTTON, self.OnGenerate)
        box.Add(m_generate, 0, wx.CENTER)
 
        # Placeholder for name of selected .prn file; it is populated in OnSelectFile(). 
        self.m_text = wx.StaticText(panel, -1, " ")
        self.m_text.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL))
        self.m_text.SetSize(self.m_text.GetBestSize())
        box.Add(self.m_text, 0, wx.ALL, 10)      
        
        # Placeholder for name of selected output directory; it is populated in OnSelectOutputDir().
        self.m_text_2 = wx.StaticText(panel, -1, " ")
        self.m_text_2.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL))
        self.m_text_2.SetSize(self.m_text.GetBestSize())
        box.Add(self.m_text_2, 0, wx.ALL, 10)             
        
        panel.SetSizer(box)
        panel.Layout()
    # end_def __init__()
        
    def OnClose(self, event):
        dlg = wx.MessageDialog(self, 
            "Do you really want to close this application?",
            "Confirm Exit", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
            self.Destroy()
    # end_def OnClose()

    def OnSelectFile(self, event):
        frame = wx.Frame(None, -1, 'win.py')
        frame.SetSize(0,0,200,50)
        openFileDialog = wx.FileDialog(frame, "Select model report PRN file", "", "", 
                                       "PRN files (*.prn)|*.prn", 
                                       wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        openFileDialog.ShowModal()
        self.prnFileName = openFileDialog.GetPath()
        self.m_text.SetLabel("Selected input PRN file:\n" + self.prnFileName)
        openFileDialog.Destroy()
        frame.Destroy()
    # end_def OnSelectFile()
    
    def OnSelectOutputDir(self, event):
        frame = wx.Frame(None, -1, 'win.py')
        frame.SetSize(0,0,200,50)
        dlg = wx.DirDialog(None, "Specify output directory", "",
                           wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        dlg.ShowModal()
        self.outputDirName = dlg.GetPath()
        self.m_text_2.SetLabel("Selected output directory:\n" + self.outputDirName)
        dlg.Destroy()
        frame.Destroy()
    # end_def OnSelectOutputDir()   
    
    def OnGenerate(self, event):
        dlg = wx.MessageDialog(self, 
            "Do you really want to run the table extraction tool?",
            "Confirm: OK/Cancel", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
            main(self.prnFileName, self.outputDirName, ('9.01','10.01','10.02','10.03','10.04','10.05','10.06'))
            message = "Separate table files generated. Process them using the Access template."
            caption = "Model Report Table Extraction Tool"
            dlg = wx.MessageDialog(None, message, caption, wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()
            # self.Destroy()
        else:
            pass
            # self.Destroy()
    # end_def OnGenerate()

    def OnAbout(self, event):
        dlg = AboutBox()
        dlg.ShowModal()
        dlg.Destroy() 
    # end_def OnAbout()
# end_class Frame

# The code for the GUI'd application itself begins here.
#
app = wx.App(redirect=True)   # Error messages go to popup window
top = Frame("Model Report Table Extraction Tool")
top.Show()
app.MainLoop()
