# -*- coding: cp1250 -*-
#Boa:Frame:Frame1

import wx
from wx.lib.anchors import LayoutAnchors
import firstscanner, os, sys, fnmatch
import ConfigParser
import time, random
sys.setrecursionlimit(3000)

ALL_OK = False

def create(parent):
    return Frame1(parent)

[wxID_FRAME1, wxID_FRAME1BUTTON1, wxID_FRAME1BUTTON2, wxID_FRAME1BUTTON8, 
 wxID_FRAME1BUTTONDATE, wxID_FRAME1BUTTONDB, wxID_FRAME1BUTTONDOCCODE, 
 wxID_FRAME1BUTTONDOCFOLDER, wxID_FRAME1BUTTONDOCNAME, 
 wxID_FRAME1BUTTONDOCPATH, wxID_FRAME1BUTTONFILENAME, 
 wxID_FRAME1BUTTONSAVECONFIG, wxID_FRAME1BUTTONVAR, wxID_FRAME1BUTTONVERSION, 
 wxID_FRAME1BUTTONVERSION1, wxID_FRAME1CHECKBOX1, wxID_FRAME1CHECKBOX2, 
 wxID_FRAME1CHOICEWORD, wxID_FRAME1ERRORTEXT1, wxID_FRAME1ERRORTEXT2, 
 wxID_FRAME1ERRORTEXT3, wxID_FRAME1ERRORTEXT4, wxID_FRAME1ERRORTEXT5, 
 wxID_FRAME1ERRORTEXT6, wxID_FRAME1PANEL1, wxID_FRAME1PANEL2, 
 wxID_FRAME1PANEL3, wxID_FRAME1STATICTEXT1, wxID_FRAME1STATICTEXT10, 
 wxID_FRAME1STATICTEXT11, wxID_FRAME1STATICTEXT12, wxID_FRAME1STATICTEXT2, 
 wxID_FRAME1STATICTEXT3, wxID_FRAME1STATICTEXT4, wxID_FRAME1STATICTEXT5, 
 wxID_FRAME1STATICTEXT6, wxID_FRAME1STATICTEXT7, wxID_FRAME1STATICTEXT8, 
 wxID_FRAME1STATICTEXT9, wxID_FRAME1STATICTEXTMASTER, wxID_FRAME1TEXTBODY, 
 wxID_FRAME1TEXTCTRL1, wxID_FRAME1TEXTCTRL2, wxID_FRAME1TEXTCTRL3, 
 wxID_FRAME1TEXTDATE, wxID_FRAME1TEXTDATEPREFIX, wxID_FRAME1TEXTFILENAME, 
 wxID_FRAME1TEXTMAILPATH, wxID_FRAME1TEXTSUBJECT, wxID_FRAME1TEXTVAR, 
 wxID_FRAME1TEXTVERSION, 
] = [wx.NewId() for _init_ctrls in range(51)]


# Define File Drop Target class
class FileDropTarget(wx.FileDropTarget):
    """ This object implements Drop Target functionality for Files """
    def __init__(self, obj, obj2):
        """ Initialize the Drop Target, passing in the Object Reference to
        indicate what should receive the dropped files """
        # Initialize the wxFileDropTarget Object
        wx.FileDropTarget.__init__(self)
        # Store the Object Reference for dropped files
        self.obj = obj
        self.obj2 = obj2

    def OnDropFiles(self, x, y, filenames):
        """ Implement File Drop """
        # For Demo purposes, this function appends a list of the files dropped at the end of the widget's text
        # Move Insertion Point to the end of the widget's text
        #self.obj.SetInsertionPointEnd()
        # append a list of the file names dropped
        #self.obj.WriteText("%d file(s) dropped at %d, %d:\n" % (len(filenames), x, y))
        #for file in filenames:
        if len(filenames) == 1:
            self.obj.Value = filenames[0]
        elif len(filenames) == 2:
            FILENAME, FIRST_PAGE = firstscanner.get_args(filenames)
            self.obj.Value = FIRST_PAGE
            self.obj2.Value = FILENAME
        #self.obj.WriteText('\n')
    #def OnDragOver(self, x, y, filenames): 
    #   self.obj2.SetBackgroundColour("Green") 
    #   self.obj2.Refresh()
    #   #self.obj.WriteText('hehe')
class FileDropTarget2(wx.FileDropTarget):
    """ obiekt upuszczania dla drugiego pola txt """
    def __init__(self, obj, obj2):
        """ Initialize the Drop Target, passing in the Object Reference to
        indicate what should receive the dropped files """
        # Initialize the wxFileDropTarget Object
        wx.FileDropTarget.__init__(self)
        # Store the Object Reference for dropped files
        self.obj = obj
        self.obj2 = obj2

    def OnDropFiles(self, x, y, filenames):
        """ Implement File Drop """
        # For Demo purposes, this function appends a list of the files dropped at the end of the widget's text
        # Move Insertion Point to the end of the widget's text
        #self.obj.SetInsertionPointEnd()
        # append a list of the file names dropped
        #self.obj.WriteText("%d file(s) dropped at %d, %d:\n" % (len(filenames), x, y))
        #for file in filenames:
        if len(filenames) == 1:
            self.obj.Value = filenames[0]
            F = filenames[0]
            if (os.path.isfile(F)) and ("." in F) and (F[F.rindex(".")+1:].lower() in ["doc","docx"]):
                #przeiagnieto plik worda, eksport tego pliku:
                self.obj.Value += " - eksport do pdf.........."
                self.obj.Enabled = False
                Result, F2 = firstscanner.saveword(F)
                if not Result:
                    self.obj.Value = ""
                    #self.obj.WriteText(u"Błąd. "+F2)
                    #time.sleep(3)
                    dlg = wx.MessageDialog(None,
                    F2,
                    "Uwaga", wx.OK)
                    result = dlg.ShowModal()
                    dlg.Destroy()
                    self.obj.Value = ""
                else:
                    self.obj.Value = F2
                self.obj.Enabled = True
        elif len(filenames) == 2:
            FILENAME, FIRST_PAGE = firstscanner.get_args(filenames)
            self.obj2.Value = FIRST_PAGE #tu zamienione miejscami wpisy
            self.obj.Value = FILENAME

class DirDropTarget(wx.FileDropTarget):
    """ obiekt upuszczania dla drugiego pola txt """
    def __init__(self, obj):
        """ Initialize the Drop Target, passing in the Object Reference to
        indicate what should receive the dropped files """
        # Initialize the wxFileDropTarget Object
        wx.FileDropTarget.__init__(self)
        # Store the Object Reference for dropped files
        self.obj = obj

    def OnDropFiles(self, x, y, filenames):
        """ Implement File Drop """
        # For Demo purposes, this function appends a list of the files dropped at the end of the widget's text
        # Move Insertion Point to the end of the widget's text
        #self.obj.SetInsertionPointEnd()
        # append a list of the file names dropped
        #self.obj.WriteText("%d file(s) dropped at %d, %d:\n" % (len(filenames), x, y))
        #for file in filenames:
        if len(filenames) == 1 and os.path.isdir(filenames[0]):
            self.obj.Value = filenames[0]

class TextDropTarget(wx.TextDropTarget):
   """ This object implements Drop Target functionality for Text """
   def __init__(self, obj):
      """ Initialize the Drop Target, passing in the Object Reference to
          indicate what should receive the dropped text """
      # Initialize the wx.TextDropTarget Object
      wx.TextDropTarget.__init__(self)
      # Store the Object Reference for dropped text
      self.obj = obj

   def OnDropText(self, x, y, data):
      """ Implement Text Drop """
      # When text is dropped, write it into the object specified
      self.obj.Value = data



class Frame1(wx.Frame):
    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRAME1, name='', parent=prnt,
              pos=wx.Point(507, 130), size=wx.Size(939, 773),
              style=wx.DEFAULT_FRAME_STYLE, title=u'FIRSTSCANNER by LDU')
        self.SetClientSize(wx.Size(931, 746))
        self.SetAutoLayout(False)

        self.panel1 = wx.Panel(id=wxID_FRAME1PANEL1, name='panel1', parent=self,
              pos=wx.Point(0, 0), size=wx.Size(931, 746),
              style=wx.TAB_TRAVERSAL)
        self.panel1.SetToolTipString(u'')

        self.button1 = wx.Button(id=wxID_FRAME1BUTTON1, label=u'START',
              name='button1', parent=self.panel1, pos=wx.Point(624, 288),
              size=wx.Size(272, 48), style=0)
        self.button1.SetConstraints(LayoutAnchors(self.button1, True, True,
              False, False))
        self.button1.SetToolTipString(u'start')
        self.button1.SetFont(wx.Font(16, wx.SWISS, wx.NORMAL, wx.NORMAL, False,
              u'Tahoma'))
        self.button1.SetDefault()
        self.button1.SetThemeEnabled(False)
        self.button1.SetAutoLayout(True)
        self.button1.Bind(wx.EVT_BUTTON, self.OnButton1Button,
              id=wxID_FRAME1BUTTON1)

        self.panel2 = wx.Panel(id=wxID_FRAME1PANEL2, name='panel2',
              parent=self.panel1, pos=wx.Point(8, 8), size=wx.Size(904, 88),
              style=wx.SUNKEN_BORDER | wx.TAB_TRAVERSAL)
        self.panel2.SetBackgroundColour(wx.Colour(192, 192, 192))
        self.panel2.SetToolTipString(u'Mo\u017cesz przeci\u0105gn\u0105\u0107 plik z eksploratora windows lub pulpitu')

        self.textCtrl1 = wx.TextCtrl(id=wxID_FRAME1TEXTCTRL1, name='textCtrl1',
              parent=self.panel2, pos=wx.Point(8, 16), size=wx.Size(880, 40),
              style=0, value=u'')
        self.textCtrl1.SetToolTipString(u'\u015acie\u017cka skanu pierwszej strony dokumentu')
        self.textCtrl1.Bind(wx.EVT_TEXT, self.OnTextCtrl1Text,
              id=wxID_FRAME1TEXTCTRL1)

        self.staticText3 = wx.StaticText(id=wxID_FRAME1STATICTEXT3,
              label=u'Plik (skan) pierwszej strony (przeci\u0105gnij plik na pole lub wklej link )',
              name='staticText3', parent=self.panel2, pos=wx.Point(8, 0),
              size=wx.Size(384, 16), style=0)
        self.staticText3.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, u'Tahoma'))
        self.staticText3.SetToolTipString(u'Mo\u017cesz przeci\u0105gn\u0105\u0107 plik z eksploratora windows lub pulpitu')

        self.panel3 = wx.Panel(id=wxID_FRAME1PANEL3, name='panel3',
              parent=self.panel1, pos=wx.Point(8, 104), size=wx.Size(904, 176),
              style=wx.SUNKEN_BORDER | wx.TAB_TRAVERSAL)
        self.panel3.SetBackgroundColour(wx.Colour(192, 192, 192))
        self.panel3.SetToolTipString(u'Mo\u017cesz przeci\u0105gn\u0105\u0107 plik z eksploratora windows lub pulpitu')

        self.textCtrl2 = wx.TextCtrl(id=wxID_FRAME1TEXTCTRL2, name='textCtrl2',
              parent=self.panel3, pos=wx.Point(8, 16), size=wx.Size(880, 48),
              style=0, value=u'')
        self.textCtrl2.SetToolTipString(u'\u015acie\u017cka dokumentu')
        self.textCtrl2.Bind(wx.EVT_TEXT, self.OnTextCtrl2Text,
              id=wxID_FRAME1TEXTCTRL2)

        self.staticText1 = wx.StaticText(id=wxID_FRAME1STATICTEXT1,
              label=u'Plik dokumentu (przeci\u0105gnij plik na pole lub wklej link )',
              name='staticText1', parent=self.panel3, pos=wx.Point(8, 0),
              size=wx.Size(313, 16), style=0)
        self.staticText1.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, u'Tahoma'))
        self.staticText1.SetToolTipString(u'Mo\u017cesz przeci\u0105gn\u0105\u0107 plik z eksploratora windows lub pulpitu')

        self.staticText2 = wx.StaticText(id=wxID_FRAME1STATICTEXT2,
              label=u'Lokalizacja dokumentu', name='staticText2',
              parent=self.panel3, pos=wx.Point(8, 64), size=wx.Size(129, 16),
              style=0)
        self.staticText2.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, u'Tahoma'))
        self.staticText2.SetToolTipString(u'Mo\u017cesz przeci\u0105gn\u0105\u0107 plik z eksploratora windows lub pulpitu')

        self.textCtrl3 = wx.TextCtrl(id=wxID_FRAME1TEXTCTRL3, name='textCtrl3',
              parent=self.panel3, pos=wx.Point(8, 104), size=wx.Size(880, 48),
              style=0, value=u'')
        self.textCtrl3.SetEditable(True)
        self.textCtrl3.Enable(False)
        self.textCtrl3.SetToolTipString(u'Katalog zapisu dokumentu')
        self.textCtrl3.Bind(wx.EVT_TEXT, self.OnTextCtrl3Text,
              id=wxID_FRAME1TEXTCTRL3)

        self.errortext1 = wx.StaticText(id=wxID_FRAME1ERRORTEXT1, label=u'',
              name=u'errortext1', parent=self.panel2, pos=wx.Point(8, 60),
              size=wx.Size(0, 13), style=0)
        self.errortext1.SetBackgroundColour(wx.Colour(255, 0, 0))

        self.errortext2 = wx.StaticText(id=wxID_FRAME1ERRORTEXT2, label=u'',
              name=u'errortext2', parent=self.panel3, pos=wx.Point(8, 68),
              size=wx.Size(0, 13), style=0)
        self.errortext2.SetBackgroundColour(wx.Colour(255, 0, 0))

        self.button2 = wx.Button(id=wxID_FRAME1BUTTON2,
              label=u'Eksportuj z MS Word', name='button2', parent=self.panel3,
              pos=wx.Point(728, 72), size=wx.Size(160, 23), style=0)
        self.button2.SetToolTipString(u'Kliknij aby u\u017cy\u0107 otwartego dokumentu MSWord')
        self.button2.Bind(wx.EVT_BUTTON, self.OnButton2Button,
              id=wxID_FRAME1BUTTON2)

        self.checkBox1 = wx.CheckBox(id=wxID_FRAME1CHECKBOX1,
              label=u'W katalogu dokumentu', name='checkBox1',
              parent=self.panel3, pos=wx.Point(8, 80), size=wx.Size(128, 21),
              style=0)
        self.checkBox1.SetValue(True)
        self.checkBox1.SetToolTipString(u'Mo\u017cesz przeci\u0105gn\u0105\u0107 plik z eksploratora windows lub pulpitu')
        self.checkBox1.Bind(wx.EVT_CHECKBOX, self.OnCheckBox1Checkbox,
              id=wxID_FRAME1CHECKBOX1)

        self.errortext3 = wx.StaticText(id=wxID_FRAME1ERRORTEXT3, label=u'',
              name=u'errortext3', parent=self.panel3, pos=wx.Point(8, 152),
              size=wx.Size(0, 13), style=0)
        self.errortext3.SetBackgroundColour(wx.Colour(255, 0, 0))

        self.staticText4 = wx.StaticText(id=wxID_FRAME1STATICTEXT4,
              label=u'Dane maila do wys\u0142ania:', name='staticText4',
              parent=self.panel1, pos=wx.Point(8, 312), size=wx.Size(116, 13),
              style=0)
        self.staticText4.SetToolTipString(u'')

        self.textSubject = wx.TextCtrl(id=wxID_FRAME1TEXTSUBJECT,
              name=u'textSubject', parent=self.panel1, pos=wx.Point(16, 344),
              size=wx.Size(880, 21), style=0, value=u'')
        self.textSubject.SetToolTipString(u'Temat maila')
        self.textSubject.Enable(False)
        self.textSubject.Bind(wx.EVT_SET_FOCUS, self.OnTextSubjectSetFocus)
        self.textSubject.Bind(wx.EVT_TEXT, self.OnTextSubjectText,
              id=wxID_FRAME1TEXTSUBJECT)

        self.textBody = wx.TextCtrl(id=wxID_FRAME1TEXTBODY, name=u'textBody',
              parent=self.panel1, pos=wx.Point(16, 392), size=wx.Size(880, 96),
              style=wx.TE_MULTILINE, value=u'')
        self.textBody.SetToolTipString(u'Tre\u015b\u0107 maila')
        self.textBody.Enable(False)
        self.textBody.Bind(wx.EVT_SET_FOCUS, self.OnTextBodySetFocus)
        self.textBody.Bind(wx.EVT_TEXT, self.OnTextBodyText,
              id=wxID_FRAME1TEXTBODY)

        self.staticText5 = wx.StaticText(id=wxID_FRAME1STATICTEXT5,
              label=u'Temat:', name='staticText5', parent=self.panel1,
              pos=wx.Point(16, 328), size=wx.Size(34, 13), style=0)
        self.staticText5.SetFont(wx.Font(8, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, u'Tahoma'))
        self.staticText5.SetToolTipString(u'')

        self.staticText6 = wx.StaticText(id=wxID_FRAME1STATICTEXT6,
              label=u'Tre\u015b\u0107:', name='staticText6', parent=self.panel1,
              pos=wx.Point(16, 376), size=wx.Size(30, 13), style=0)
        self.staticText6.SetToolTipString(u'')

        self.buttonFilename = wx.Button(id=wxID_FRAME1BUTTONFILENAME,
              label=u'<FILENAME>', name=u'buttonFilename', parent=self.panel1,
              pos=wx.Point(16, 602), size=wx.Size(75, 23), style=0)
        self.buttonFilename.SetToolTipString(u'Wstawi zmienn\u0105 nazwy pliku dokumentu')
        self.buttonFilename.Enable(False)
        self.buttonFilename.Bind(wx.EVT_BUTTON, self.OnButton3Button,
              id=wxID_FRAME1BUTTONFILENAME)

        self.buttonDocname = wx.Button(id=wxID_FRAME1BUTTONDOCNAME,
              label=u'<DOCNAME>', name=u'buttonDocname', parent=self.panel1,
              pos=wx.Point(104, 602), size=wx.Size(75, 23), style=0)
        self.buttonDocname.SetToolTipString(u'Wstawi zmienn\u0105 nazwy dokumentu (bez rozszerzenia)')
        self.buttonDocname.Enable(False)
        self.buttonDocname.Bind(wx.EVT_BUTTON, self.OnButtonDocnameButton,
              id=wxID_FRAME1BUTTONDOCNAME)

        self.buttonDocfolder = wx.Button(id=wxID_FRAME1BUTTONDOCFOLDER,
              label=u'<DOCFOLDER>', name=u'buttonDocfolder', parent=self.panel1,
              pos=wx.Point(192, 602), size=wx.Size(88, 23), style=0)
        self.buttonDocfolder.SetToolTipString(u'Wstawi zmienn\u0105 nazwy folderu gdzie b\u0119dzie zapisany dokument')
        self.buttonDocfolder.Enable(False)
        self.buttonDocfolder.Bind(wx.EVT_BUTTON, self.OnButtonDocfolderButton,
              id=wxID_FRAME1BUTTONDOCFOLDER)

        self.buttonDate = wx.Button(id=wxID_FRAME1BUTTONDATE, label=u'<DATE>',
              name=u'buttonDate', parent=self.panel1, pos=wx.Point(400, 602),
              size=wx.Size(75, 23), style=0)
        self.buttonDate.SetToolTipString(u'Wstawia zmienn\u0105 daty')
        self.buttonDate.Enable(False)
        self.buttonDate.Bind(wx.EVT_BUTTON, self.OnButtonDateButton,
              id=wxID_FRAME1BUTTONDATE)

        self.checkBox2 = wx.CheckBox(id=wxID_FRAME1CHECKBOX2, label=u'Edytuj',
              name='checkBox2', parent=self.panel1, pos=wx.Point(136, 320),
              size=wx.Size(48, 13), style=0)
        self.checkBox2.SetToolTipString(u'')
        self.checkBox2.SetValue(False)
        self.checkBox2.Bind(wx.EVT_CHECKBOX, self.OnCheckBox2Checkbox,
              id=wxID_FRAME1CHECKBOX2)

        self.buttonSaveConfig = wx.Button(id=wxID_FRAME1BUTTONSAVECONFIG,
              label=u'Zapisz zmiany w pliku config.txt',
              name=u'buttonSaveConfig', parent=self.panel1, pos=wx.Point(16,
              630), size=wx.Size(160, 34), style=0)
        self.buttonSaveConfig.Enable(False)
        self.buttonSaveConfig.SetToolTipString(u'Zapisanie tre\u015bci i innych parametr\xf3w maila (zostan\u0105 wczytane przy nast\u0119pnym starcie)')
        self.buttonSaveConfig.Bind(wx.EVT_BUTTON, self.OnButtonSaveConfigButton,
              id=wxID_FRAME1BUTTONSAVECONFIG)

        self.button8 = wx.Button(id=wxID_FRAME1BUTTON8,
              label=u'Zapisz mail jako msg', name='button8', parent=self.panel1,
              pos=wx.Point(632, 602), size=wx.Size(264, 56), style=0)
        self.button8.SetToolTipString(u'Zapisze maila w podanej \u015bcie\u017cce')
        self.button8.Enable(False)
        self.button8.Bind(wx.EVT_BUTTON, self.OnButton8Button,
              id=wxID_FRAME1BUTTON8)

        self.textFilename = wx.TextCtrl(id=wxID_FRAME1TEXTFILENAME,
              name=u'textFilename', parent=self.panel1, pos=wx.Point(16, 520),
              size=wx.Size(880, 21), style=0, value=u'')
        self.textFilename.SetToolTipString(u'')
        self.textFilename.Enable(False)
        self.textFilename.Bind(wx.EVT_SET_FOCUS, self.OnTextFilenameSetFocus)
        self.textFilename.Bind(wx.EVT_TEXT, self.OnTextFilenameText,
              id=wxID_FRAME1TEXTFILENAME)

        self.textMailpath = wx.TextCtrl(id=wxID_FRAME1TEXTMAILPATH,
              name=u'textMailpath', parent=self.panel1, pos=wx.Point(16, 572),
              size=wx.Size(880, 21), style=0, value=u'')
        self.textMailpath.SetToolTipString(u'')
        self.textMailpath.Enable(False)
        self.textMailpath.Bind(wx.EVT_SET_FOCUS, self.OnTextMailpathSetFocus)
        self.textMailpath.Bind(wx.EVT_TEXT, self.OnTextMailpathText,
              id=wxID_FRAME1TEXTMAILPATH)

        self.staticText7 = wx.StaticText(id=wxID_FRAME1STATICTEXT7,
              label=u'Nazwa pliku pod jak\u0105 zapisa\u0107 maila',
              name='staticText7', parent=self.panel1, pos=wx.Point(16, 496),
              size=wx.Size(165, 13), style=0)

        self.staticText8 = wx.StaticText(id=wxID_FRAME1STATICTEXT8,
              label=u'\u015acie\u017cka katalogu w kt\xf3rym zapisa\u0107 maila',
              name='staticText8', parent=self.panel1, pos=wx.Point(16, 552),
              size=wx.Size(191, 13), style=0)
        self.staticText8.SetToolTipString(u'')

        self.buttonDocpath = wx.Button(id=wxID_FRAME1BUTTONDOCPATH,
              label=u'<DOCPATH>', name=u'buttonDocpath', parent=self.panel1,
              pos=wx.Point(296, 602), size=wx.Size(75, 23), style=0)
        self.buttonDocpath.Enable(False)
        self.buttonDocpath.SetToolTipString(u'Wstawia zmienn\u0105 pe\u0142nej \u015bcie\u017cki do pliku')
        self.buttonDocpath.Bind(wx.EVT_BUTTON, self.OnButtonDocpathButton,
              id=wxID_FRAME1BUTTONDOCPATH)

        self.buttonVersion = wx.Button(id=wxID_FRAME1BUTTONVERSION,
              label=u'<VERSION>', name=u'buttonVersion', parent=self.panel1,
              pos=wx.Point(504, 602), size=wx.Size(80, 23), style=0)
        self.buttonVersion.Enable(False)
        self.buttonVersion.SetToolTipString(u'Wstawia zmienn\u0105 wersji')
        self.buttonVersion.Bind(wx.EVT_BUTTON, self.OnButtonVersionButton,
              id=wxID_FRAME1BUTTONVERSION)

        self.textVersion = wx.TextCtrl(id=wxID_FRAME1TEXTVERSION,
              name=u'textVersion', parent=self.panel1, pos=wx.Point(552, 288),
              size=wx.Size(64, 21), style=0, value=u'')
        self.textVersion.Enable(False)
        self.textVersion.SetToolTipString(u'Kliknij dwukrotnie aby wstawi\u0107 001')
        self.textVersion.Bind(wx.EVT_TEXT, self.OnTextVersionText,
              id=wxID_FRAME1TEXTVERSION)
        self.textVersion.Bind(wx.EVT_LEFT_DCLICK, self.OnTextVersionLeftDclick)
        self.textVersion.Bind(wx.EVT_MIDDLE_DCLICK,
              self.OnTextVersionMiddleDclick)

        self.staticText9 = wx.StaticText(id=wxID_FRAME1STATICTEXT9,
              label=u'Wersja dokumentu:', name='staticText9',
              parent=self.panel1, pos=wx.Point(456, 288), size=wx.Size(94, 13),
              style=0)
        self.staticText9.SetToolTipString(u'Kliknij dwukrotnie aby wstawi\u0107 001')
        self.staticText9.Bind(wx.EVT_LEFT_DCLICK, self.OnStaticText9LeftDclick)

        self.errortext4 = wx.StaticText(id=wxID_FRAME1ERRORTEXT4,
              label=u'45645645646', name=u'errortext4', parent=self.panel1,
              pos=wx.Point(472, 312), size=wx.Size(66, 13),
              style=wx.ALIGN_RIGHT)
        self.errortext4.SetBackgroundColour(wx.Colour(255, 0, 0))
        self.errortext4.SetToolTipString(u'')

        self.staticText10 = wx.StaticText(id=wxID_FRAME1STATICTEXT10,
              label=u'Data/batch:', name='staticText10', parent=self.panel1,
              pos=wx.Point(288, 288), size=wx.Size(58, 13), style=0)
        self.staticText10.SetToolTipString(u'Kliknij dwukrotnie aby wstawi\u0107 dzisiejsz\u0105 dat\u0119')
        self.staticText10.Bind(wx.EVT_LEFT_DCLICK,
              self.OnStaticText10LeftDclick)

        self.textDate = wx.TextCtrl(id=wxID_FRAME1TEXTDATE, name=u'textDate',
              parent=self.panel1, pos=wx.Point(352, 288), size=wx.Size(100, 21),
              style=0, value=u'')
        self.textDate.SetToolTipString(u'Kliknij dwukrotnie aby wstawi\u0107 dzisiejsz\u0105 dat\u0119')
        self.textDate.Enable(False)
        self.textDate.Bind(wx.EVT_TEXT, self.OnTextDateText,
              id=wxID_FRAME1TEXTDATE)
        self.textDate.Bind(wx.EVT_LEFT_DCLICK, self.OnTextDateLeftDclick)
        self.textDate.Bind(wx.EVT_MIDDLE_DCLICK, self.OnTextDateMiddleDclick)

        self.errortext5 = wx.StaticText(id=wxID_FRAME1ERRORTEXT5,
              label='staticText11', name=u'errortext5', parent=self.panel1,
              pos=wx.Point(352, 312), size=wx.Size(60, 13), style=0)
        self.errortext5.SetBackgroundColour(wx.Colour(255, 0, 0))

        self.buttonVersion1 = wx.Button(id=wxID_FRAME1BUTTONVERSION1,
              label=u'<VERSION-1>', name=u'buttonVersion1', parent=self.panel1,
              pos=wx.Point(504, 632), size=wx.Size(80, 23), style=0)
        self.buttonVersion1.Enable(False)
        self.buttonVersion1.SetToolTipString(u'Wstawia zmienn\u0105 wersji pomniejszon\u0105 o 1')
        self.buttonVersion1.Bind(wx.EVT_BUTTON, self.OnButtonVersion1Button,
              id=wxID_FRAME1BUTTONVERSION1)

        self.buttonDB = wx.Button(id=wxID_FRAME1BUTTONDB, label=u'<DATEPREFIX>',
              name=u'buttonDB', parent=self.panel1, pos=wx.Point(392, 632),
              size=wx.Size(88, 23), style=0)
        self.buttonDB.Enable(False)
        self.buttonDB.SetToolTipString(u'Wstawia zmienn\u0105 przedrostka daty')
        self.buttonDB.Bind(wx.EVT_BUTTON, self.OnButtonDBButton,
              id=wxID_FRAME1BUTTONDB)

        self.textDatePrefix = wx.TextCtrl(id=wxID_FRAME1TEXTDATEPREFIX,
              name=u'textDatePrefix', parent=self.panel1, pos=wx.Point(232,
              288), size=wx.Size(52, 21), style=0, value=u'')
        self.textDatePrefix.Bind(wx.EVT_TEXT, self.OnTextDatePrefixText,
              id=wxID_FRAME1TEXTDATEPREFIX)

        self.staticText11 = wx.StaticText(id=wxID_FRAME1STATICTEXT11,
              label=u'Prefix:', name='staticText11', parent=self.panel1,
              pos=wx.Point(192, 288), size=wx.Size(32, 13), style=0)

        self.textVar = wx.TextCtrl(id=wxID_FRAME1TEXTVAR, name=u'textVar',
              parent=self.panel1, pos=wx.Point(64, 288), size=wx.Size(124, 21),
              style=0, value=u'')
        self.textVar.Bind(wx.EVT_TEXT, self.OnTextVarText,
              id=wxID_FRAME1TEXTVAR)

        self.staticText12 = wx.StaticText(id=wxID_FRAME1STATICTEXT12,
              label=u'Zmienna:', name='staticText12', parent=self.panel1,
              pos=wx.Point(16, 288), size=wx.Size(44, 13), style=0)

        self.errortext6 = wx.StaticText(id=wxID_FRAME1ERRORTEXT6,
              label='staticText13', name=u'errortext6', parent=self.panel1,
              pos=wx.Point(200, 312), size=wx.Size(60, 13), style=0)
        self.errortext6.SetBackgroundColour(wx.Colour(255, 0, 0))

        self.buttonVar = wx.Button(id=wxID_FRAME1BUTTONVAR, label=u'<VAR>',
              name=u'buttonVar', parent=self.panel1, pos=wx.Point(296, 632),
              size=wx.Size(75, 23), style=0)
        self.buttonVar.SetToolTipString(u'Wstawia dowoln\u0105 zmienn\u0105 zdefiniowan\u0105 wy\u017cej')
        self.buttonVar.Enable(False)
        self.buttonVar.Bind(wx.EVT_BUTTON, self.OnButtonVarButton,
              id=wxID_FRAME1BUTTONVAR)

        self.buttonDoccode = wx.Button(id=wxID_FRAME1BUTTONDOCCODE,
              label=u'<DOCCODE>', name=u'buttonDoccode', parent=self.panel1,
              pos=wx.Point(200, 632), size=wx.Size(75, 23), style=0)
        self.buttonDoccode.SetToolTipString(u'Wstawia zmienn\u0105 kodu dokumentu')
        self.buttonDoccode.Enable(False)
        self.buttonDoccode.Bind(wx.EVT_BUTTON, self.OnButtonDoccodeButton,
              id=wxID_FRAME1BUTTONDOCCODE)

        self.choiceWord = wx.Choice(choices=[], id=wxID_FRAME1CHOICEWORD,
              name=u'choiceWord', parent=self.panel3, pos=wx.Point(144, 72),
              size=wx.Size(576, 19), style=0)
        self.choiceWord.SetLabel(u'')
        self.choiceWord.SetHelpText(u'')
        self.choiceWord.SetFont(wx.Font(7, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, u'Tahoma'))
        self.choiceWord.SetToolTipString(u'Wybierz otwarty dokument MSWord')

        self.staticTextMaster = wx.StaticText(id=wxID_FRAME1STATICTEXTMASTER,
              label=u'PAMI\u0118TAJ O WYPE\u0141NIENIU MASTER LISTY DOKUMENT\xd3W!',
              name=u'staticTextMaster', parent=self.panel1, pos=wx.Point(32,
              680), size=wx.Size(854, 39), style=wx.ALIGN_CENTRE)
        self.staticTextMaster.SetFont(wx.Font(24, wx.SWISS, wx.NORMAL,
              wx.NORMAL, False, u'Tahoma'))
        self.staticTextMaster.SetForegroundColour(wx.Colour(255, 0, 0))
        self.staticTextMaster.SetThemeEnabled(False)
        self.staticTextMaster.SetToolTipString(u'')
        self.staticTextMaster.SetConstraints(LayoutAnchors(self.staticTextMaster,
              True, True, False, False))

    def get_config(self):
        global ARGS
        CONF = firstscanner.config(ARGS[0])
        self.textSubject.Value = CONF.Subject()
        self.textFilename.Value = CONF.FileName()
        self.textBody.WriteText(CONF.Body())
        self.textMailpath.Value = CONF.MailPath()
        self._config = CONF
    def set_config(self):
        #zapisuje dane z porgramu do pliku conf.
        CONF = self._config
        #self.Refresh()
    def get_worddocs(self):
        #pobieera listę dokumentów worda i wstawia je na listę self.choiceWord
        self.choiceWord.AppendItems(firstscanner.getallworddocs())
        print self.choiceWord.Label
    def __init__(self, parent):
        self._init_ctrls(parent)
        self.get_config()
        dt1 = FileDropTarget(self.textCtrl1, self.textCtrl2)
        self.panel2.SetDropTarget(dt1)
        dt2 = FileDropTarget2(self.textCtrl2, self.textCtrl1)
        self.panel3.SetDropTarget(dt2)
        dt3 = DirDropTarget(self.textCtrl3)
        self.textCtrl3.SetDropTarget(dt3)
        self.focus = -1
        self._mailedititems = [self.buttonFilename, self.buttonDocname, self.buttonDocfolder, self.buttonDate, self.buttonSaveConfig, \
        self.buttonDocpath, self.buttonVersion, \
        self.textBody, self.textFilename, self.textMailpath, self.textSubject,\
        self.buttonDB, self.buttonVersion1, self.buttonDoccode, self.buttonVar]
        #wypełnienie Subject, Body i filename z pliku config.txt:
        self._object_by_code = [self.textSubject, self.textBody, self.textFilename, self.textMailpath]
        self.checkforversionneeded()
        self.get_worddocs()
        self._initial_ctrls = [self.button1, self.textCtrl1, self.textCtrl1, self.textCtrl2, self.textCtrl3, self.checkBox1, \
            self.textVersion, self.textDate, self.textVar, self.textDatePrefix, self.choiceWord]
        #umożliwienie migania textu.
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.blinkupdate, self.timer)
        self.Refresh()
    def blinkupdate(self, event):
        """"""
        #colors = ["blue", "green", "red", "yellow"]
        #self.staticTextMaster.SetForegroundColour(random.choice(colors))
        self.staticTextMaster.Shown = not self.staticTextMaster.Shown
        self.Refresh()
        
    def checkforversionneeded(self):
        #sprawdza, czy potrzebne jest wprowadzenie wersji maila
        global ALL_OK
        ALLTEXT = self.textBody.Value + self.textSubject.Value + self.textFilename.Value + self.textMailpath.Value
        result = True
        #Version:
        self.errortext4.Label = ""
        if "<VERSION>" in ALLTEXT:
            if not len(self.textVersion.Value) > 0:
                self.errortext4.Label = "Błąd: brak oznaczenia wersji"
                ALL_OK = False
                result = False
            if not self.textVersion.Enabled:
                self.textVersion.Enabled = True
        else:
            self.textVersion.Enabled = False
        #date:
        self.errortext5.Label = ""
        if "<DATE>" in ALLTEXT:
            if not len(self.textDate.Value) > 0:
                self.errortext5.Label = "Błąd: brak daty/batcha"
                ALL_OK = False
                result = False
            if not self.textDate.Enabled:
                self.textDate.Enabled = True
        else:
            self.textDate.Enabled = False
        #DatePrefix i Var:
        self.errortext6.Label = ""
        if "<DATEPREFIX>" in ALLTEXT:
            if not len(self.textDatePrefix.Value) > 0:
                self.errortext6.Label = "Błąd: brak danych"
                ALL_OK = False
                result = False
            else:
                self.errortext6.Label = ""
            if not self.textDatePrefix.Enabled:
                self.textDatePrefix.Enabled = True
        else:
            self.textDatePrefix.Enabled = False
            
        if "<VAR>" in ALLTEXT:
            if not len(self.textVar.Value) > 0:
                self.errortext6.Label = "Błąd: brak danych"
                ALL_OK = False
                result = False
            else:
                self.errortext6.Label = ""
            if not self.textVar.Enabled:
                self.textVar.Enabled = True
        else:
            self.textVar.Enabled = False
        self.Refresh()
        return result
    def checkinputdata(self):
        #sprawdza, czy wszystkie dane sa kompletne i odpowiednie:
        result = True
        result = result and firstscanner.isfileok(self.textCtrl1.Value, True)[0]
        result = result and firstscanner.isfileok(self.textCtrl2.Value, False)[0]
        if not self.checkBox1.Value:
            result = result and firstscanner.check_write(self.textCtrl3.Value)
        result = result and self.checkforversionneeded()
        return result
        
    def OnTextCtrl1Text(self, event):
        global ALL_OK
        F = self.textCtrl1.Value
        self.errortext1.Label = ""
        if F.strip() == "":
            event.Skip()
            ALL_OK = False
        else:
            OK, KOM = firstscanner.isfileok(F, True)
            self.errortext1.Label = KOM
            ALL_OK = OK
        event.Skip()

    def OnTextCtrl2Text(self, event):
        global ALL_OK
        F = self.textCtrl2.Value
        #self.textCtrl1
        self.errortext2.Label = ""
        ALL_OK = False
        if F.strip() == "":
            event.Skip()
        else:
            OK, KOM = firstscanner.isfileok(F)
            self.errortext2.Label = KOM
            if OK and self.checkBox1.Value:
                self.textCtrl3.Value = os.path.dirname(F)
            ALL_OK = OK
        event.Skip()
        
    def OnTextCtrl3Text(self, event):
        global ALL_OK
        F = self.textCtrl3.Value.strip()
        self.errortext3.Label = ""
        ALL_OK = False
        if F.strip() == "":
            event.Skip()
        else:
            if not os.path.isdir(F):
                self.errortext3.Label = "Błąd: katalog nie istnieje"
            elif not firstscanner.check_write(F):
                self.errortext3.Label = "Błąd: katalog tylko do odczytu"
            else:
                ALL_OK = True
        event.Skip()

    def OnButton2Button(self, event):
        W = firstscanner.getword()
        self.textCtrl2.Value = ""
        #self.textCtrl2.Value = W
        SEL = self.choiceWord.GetStringSelection()
        if SEL == "":
            SEL = None
        try:
            Result, P = firstscanner.saveword(SEL)
            if Result:
                self.textCtrl2.Value = P
            else:
                dlg = wx.MessageDialog(self,
                P,
                "Uwaga", wx.OK)
                result = dlg.ShowModal()
                dlg.Destroy()
        except:
            W = firstscanner.getword()
            self.errortext2.Label = u"Błąd podczas eksportu pliku "+W
        event.Skip()

    def OnCheckBox1Checkbox(self, event):
        if self.checkBox1.Value == True:
            self.textCtrl3.Enabled = False
            if os.path.isdir(os.path.dirname(self.textCtrl2.Value)):
                self.textCtrl3.Value = os.path.dirname(self.textCtrl2.Value)
        else:
            self.textCtrl3.Enabled = True
        event.Skip()

    def OnButton1Button(self, event):
        global ALL_OK
        #print ALL_OK
        if self.checkinputdata():
            for C in self._initial_ctrls:
                C.Enabled = False
            if not self.checkBox1.Value:
                PATH = self.textCtrl3.Value
                FULL_PATH = PATH + "\\" + os.path.basename(self.textCtrl2.Value)
            else:
                PATH = None
                FULL_PATH = self.textCtrl2.Value
            firstscanner.process_files(self.textCtrl2.Value, self.textCtrl1.Value, PATH)
            if "\\\\" in FULL_PATH:
                FULL_PATH = FULL_PATH.replace("\\\\","\\")
            self.M = firstscanner.Mail(firstscanner.resolve(self.textSubject.Value, FULL_PATH, self.textVersion.Value, self.textDate.Value, self.textDatePrefix.Value, self.textVar.Value),\
            firstscanner.resolve(self.textBody.Value, FULL_PATH, self.textVersion.Value, self.textDate.Value, self.textDatePrefix.Value, self.textVar.Value),\
            firstscanner.resolve(self.textFilename.Value, FULL_PATH, self.textVersion.Value, self.textDate.Value, self.textDatePrefix.Value, self.textVar.Value),\
            firstscanner.resolve(self.textMailpath.Value, FULL_PATH, self.textVersion.Value, self.textDate.Value, self.textDatePrefix.Value, self.textVar.Value))
            self.M.create()
            self.timer.Start(500)
            self.button8.Enabled = True
            #for C in self._mailedititems:
            #    C.Enabled = False
            #self.checkBox2.Enabled = False
            self.Refresh()
        else:
            dlg = wx.MessageDialog(self,
                "Niekompletne dane wejściowe!",
                "Uwaga", wx.OK)
            result = dlg.ShowModal()
            dlg.Destroy()
        event.Skip()
    def GetFocusedObject(self):
        return self._object_by_code[self.focus]

    def OnTextSubjectSetFocus(self, event):
        self.focus = 0
        event.Skip()

    def OnTextBodySetFocus(self, event):
        self.focus = 1
        event.Skip()
        
    def OnTextFilenameSetFocus(self, event):
        self.focus = 2
        event.Skip()

    def OnTextMailpathSetFocus(self, event):
        self.focus = 3
        event.Skip()


    def OnButton8Button(self, event):
        if firstscanner.check_write(self._config.MailPath()):
            result = self.M.save()
            if result == 0 and os.path.isfile(self.M.path()):
                #dlg = wx.MessageDialog(self,
                #    "Zapisano maila poprawnie",
                #    "Sukces!", wx.OK)
                #dlg.ShowModal()
                #dlg.Destroy()
                self.M.explore()
                self.Close()
            elif result == 7:
                dlg = wx.MessageDialog(self,
                "Plik o docelowej nazwie juz istnieje, sprawdź poprawność wykonanej operacji",
                "Uwaga", wx.OK)
                result = dlg.ShowModal()
                dlg.Destroy()
            elif result == 11:
                dlg = wx.MessageDialog(self,
                "Nie znaleziono maila w wysłanych, wyślij maila lub spróbuj ponownie za chwilę",
                "Uwaga", wx.OK)
                result = dlg.ShowModal()
                dlg.Destroy()
            else:
                dlg = wx.MessageDialog(self,
                "Błąd przy zapisie maila, sprawdź poprawność operacji",
                "Uwaga", wx.OK)
                result = dlg.ShowModal()
                dlg.Destroy()
        event.Skip()

    def OnButtonSaveConfigButton(self, event):
        if not self.textSubject.Value.strip() == "":
            self._config.set_Subject(self.textSubject.Value)
        if not self.textBody.Value.strip() == "":
            self._config.set_Body(self.textBody.Value)
        if not self.textFilename.Value.strip() == "":
            self._config.set_FileName(self.textFilename.Value)
        if os.path.isdir(self.textMailpath.Value.strip()):
            self._config.set_MailPath(self.textMailpath.Value)
        self._config.save()
        event.Skip()

    def OnCheckBox2Checkbox(self, event):
        if self.checkBox2.Value:
            #tresc i inne dane maila mozliwe do edycji
            for o in self._mailedititems:
                o.Enabled = True
        else:
            for o in self._mailedititems:
                o.Enabled = False
        event.Skip()

    def OnTextBodyText(self, event):
        self.checkforversionneeded()
        event.Skip()

    def OnTextVersionText(self, event):
        self.checkforversionneeded()
        event.Skip()


    def OnButton3Button(self, event):
        self.GetFocusedObject().WriteText("<FILENAME>")
        event.Skip()
    def OnButtonDocnameButton(self, event):
        self.GetFocusedObject().WriteText("<DOCNAME>")
        event.Skip()

    def OnButtonDocfolderButton(self, event):
        self.GetFocusedObject().WriteText("<DOCFOLDER>")
        event.Skip()

    def OnButtonDocpathButton(self, event):
        self.GetFocusedObject().WriteText("<DOCPATH>")
        event.Skip()

    def OnButtonDateButton(self, event):
        self.GetFocusedObject().WriteText("<DATE>")
        event.Skip()

    def OnButtonVersionButton(self, event):
        self.GetFocusedObject().WriteText("<VERSION>")
        event.Skip()
        
    def OnButtonVersion1Button(self, event):
        self.GetFocusedObject().WriteText("<VERSION-1>")
        event.Skip()

    def OnButtonDBButton(self, event):
        self.GetFocusedObject().WriteText("<DATEPREFIX>")
        event.Skip()

    def OnButtonVarButton(self, event):
        self.GetFocusedObject().WriteText("<VAR>")
        event.Skip()

    def OnButtonDoccodeButton(self, event):
        self.GetFocusedObject().WriteText("<DOCCODE>")
        event.Skip()

    def OnTextSubjectText(self, event):
        self.checkforversionneeded()
        event.Skip()

    def OnTextFilenameText(self, event):
        self.checkforversionneeded()
        event.Skip()
    
    def OnTextDatePrefixText(self, event):
        self.checkforversionneeded()
        event.Skip()

    def OnTextVarText(self, event):
        self.checkforversionneeded()
        event.Skip()

    def OnTextMailpathText(self, event):
        self.checkforversionneeded()
        event.Skip()

    def OnStaticText10LeftDclick(self, event):
        #wstawia dzisiejszą datę do okienka daty:
        DATESTR = time.strftime("%Y-%m-%d")
        self.textDate.WriteText(DATESTR)
        event.Skip()

    def OnTextDateText(self, event):
        self.checkforversionneeded()
        event.Skip()

    def OnTextDateLeftDclick(self, event):
        DATESTR = time.strftime("%Y-%m-%d")
        self.textDate.WriteText(DATESTR)
        self.textDatePrefix.WriteText("date of")
        event.Skip()

    def OnTextVersionLeftDclick(self, event):
        self.textVersion.Value = "001"
        event.Skip()

    def OnStaticText9LeftDclick(self, event):
        self.textVersion.WriteText("001")
        event.Skip()

    def OnTextDateMiddleDclick(self, event):
        if self.textVersion.Value == "666":
            self.textDate.Value = "Palikot na prezydenta!! ;)"
        event.Skip()

    def OnTextVersionMousewheel(self, event):
        self.textVersion.Value = firstscanner.plus1(self.textVersion.Value)
        event.Skip()

    def OnTextVersionMiddleDclick(self, event):
        self.textVersion.Value = firstscanner.plus1(self.textVersion.Value)
        event.Skip()


    
    
        
if __name__ == '__main__':
    ARGS = sys.argv
    app = wx.PySimpleApp()
    frame = create(None)
    frame.Show()
    if len(ARGS)>1:
        PDF_NAME, FIRST_PAGE = firstscanner.get_args(ARGS[1:], False)
    else:
        FIRST_PAGE = None
        PDF_NAME = None
    if not FIRST_PAGE == None:
        frame.textCtrl1.Value = FIRST_PAGE
    if not PDF_NAME == None:
        frame.textCtrl2.Value = PDF_NAME
    frame.Refresh()
    app.MainLoop()
