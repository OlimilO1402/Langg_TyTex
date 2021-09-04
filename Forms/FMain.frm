VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FMain 
   Caption         =   "TyTex"
   ClientHeight    =   7545
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11325
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   0
      Width           =   6735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   5580
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   360
      Width           =   6735
   End
   Begin MSComDlg.CommonDialog OFD 
      Left            =   2400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuTlb 
      Caption         =   "File"
      Begin VB.Menu mnuTlbOpenFile 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuTlbOpenReg 
         Caption         =   "Open Registrykey"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuTlbSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTlbOpenDir 
         Caption         =   "Goto Directory"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuTlbSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTlbSaveAs 
         Caption         =   "Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuTlbSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTlbExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuCode 
      Caption         =   "Code"
      Begin VB.Menu mnuCodeVB 
         Caption         =   "VB"
      End
      Begin VB.Menu mnuCodeJava 
         Caption         =   "Java"
      End
      Begin VB.Menu mnuCodeCpp 
         Caption         =   "C++"
      End
      Begin VB.Menu mnuCodeSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodeIndent 
         Caption         =   "Indentation"
      End
      Begin VB.Menu mnuCodeSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuCodeSearchNext 
         Caption         =   "Search Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuHelpstring 
      Caption         =   "HelpString"
      Begin VB.Menu mnuHelpstringLcid 
         Caption         =   "LCID"
         Begin VB.Menu mnuHelpstringLcidENUS 
            Caption         =   "LCID 0"
         End
         Begin VB.Menu mnuHelpstringLcidDEDE 
            Caption         =   "LCID 1"
         End
         Begin VB.Menu mnuHelpstringLcidSpecify 
            Caption         =   "Specify LCID X"
         End
      End
      Begin VB.Menu mnuHelpstringAuthor 
         Caption         =   "Author"
      End
   End
   Begin VB.Menu mnuExtras 
      Caption         =   "Extras"
      Begin VB.Menu mnuExtrasStatisticEvents 
         Caption         =   "Statistic Events"
      End
      Begin VB.Menu mnuExtrasStatisticMethod 
         Caption         =   "Statistic Methods"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " ? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TyITex, the tight typelibinfo extractor
Dim myTlb   As TypeLibrary    'the typelibrary file or from registry
Dim col     As New Collection 'for selecting from combobox
Dim myCoder As ICoder         'what do you want? VB, VB.NET, Java, Cpp, Delphi ... ?
Dim SearchForName As String
Dim SearchIndex   As Long

Private Sub Form_Load()
    
    ResetSearch
    
    MInd.IndentSize = 4
    MHelpCode.MyAuthorName = "OlimilO"
    Me.mnuCodeVB.Checked = True
    Set myCoder = New CoderVB
    'Combo1.AddItem "MBOFastGeoD":   col.Add "File | D:\VB60To_dotNETClasses\VB6_To_Delphi\FastGeoD\Tlb\MBOFastGeoD.tlb"
    Combo1.AddItem "VBA":           col.Add "File | C:\WINDOWS\system32\msvbvm60.dll"
    Combo1.AddItem "VBRUN":         col.Add "{EA544A21-C82D-11D1-A3E4-00A0C90AEA82} | 6 | 0 | 9"
    Combo1.AddItem "VB":            col.Add "File | C:\Program Files\Microsoft Visual Studio\VB98\VB6.OLB"
    Combo1.AddItem "VBIDE":         col.Add "File | C:\Program Files\Microsoft Visual Studio\VB98\VB6EXT.OLB" '
    Combo1.AddItem "stdole":        col.Add "File | C:\Windows\system32\stdole2.tlb"     ' OLE Automation
    Combo1.AddItem "TLI":           col.Add "File | C:\Windows\system32\tlbinf32.dll"    ' TypeLib Information
    Combo1.AddItem "ComCtl2":       col.Add "File | C:\Windows\system32\ComCt232.oca"    ' Microsoft Windows-Standardsteuerelemente-2 5.0
    Combo1.AddItem "ComCtl3":       col.Add "File | C:\Windows\system32\Comct332.oca"    ' Microsoft Windows Common Controls-3 6.0 (SP5)
    Combo1.AddItem "ComctlLib":     col.Add "File | C:\Windows\system32\comctl32.oca"    ' Microsoft Windows-Standardsteuerelemente 5.0 (SP2)
    Combo1.AddItem "MCI":           col.Add "File | C:\Windows\system32\mci32.oca"       ' Microsoft Multimedia-Steuerelement 6.0
    Combo1.AddItem "MSACAL":        col.Add "File | C:\Windows\system32\MSCAL.oca"       ' Microsoft Calendar Control 8.0
    Combo1.AddItem "MSComDlg":      col.Add "File | C:\Windows\system32\ComDlg32.oca"    ' Microsoft Standarddialog-Steuerelement 6.0
    Combo1.AddItem "MSCommLib":     col.Add "File | C:\Windows\system32\MSComm32.oca"    ' Microsoft Kommunikations-Steuerelement 6.0
    Combo1.AddItem "MSDBCtls":      col.Add "File | C:\Windows\system32\dblist32.oca"    ' Microsoft - Datengebundene Listensteuerelemente 6.0
    Combo1.AddItem "MSDBGrid":      col.Add "File | C:\Windows\system32\DBGRID32.oca"    ' Microsoft Datengebundenes Tabellensteuerelement 5.0 (SP3)
    Combo1.AddItem "MSFlexGridLib": col.Add "File | C:\Windows\system32\MSFlxGrd.oca"    ' Microsoft FlexTabelle-Steuerelement (FlexGrid) 6.0
    Combo1.AddItem "MSMask":        col.Add "File | C:\Windows\system32\msmask32.oca"    ' Microsoft Formatierte Bearbeitung-Steuerelement 6.0
    Combo1.AddItem "RichTextLib":   col.Add "File | C:\Windows\system32\richtx32.oca"    ' Microsoft RTF-Steuerelement 6.0
    Combo1.AddItem "TabDlg":        col.Add "File | C:\Windows\system32\TABCTL32.oca"    ' Microsoft Register-Steuerelement 6.0
    Combo1.AddItem "MSXML2":        col.Add "File | C:\Windows\system32\msxml6.dll"      ' Microsoft XML, v6.0
    Combo1.AddItem "Scripting":     col.Add "File | C:\Windows\system32\scrrun.dll"      ' Microsoft Scripting Runtime
    Combo1.AddItem "WScript":       col.Add "File | C:\Windows\system32\Wshom.ocx"       ' Windows Script Host Object Model
    Combo1.AddItem "Forms2":        col.Add "File | C:\Windows\System32\FM20.dll"        ' Microsoft Forms 2.0 Object Library
    Combo1.AddItem "Office":        col.Add "File | C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.dll" '
    Combo1.AddItem "Word":          col.Add "File | C:\Program Files\Microsoft Office\Office16\MSWORD.OLB" '
    Combo1.AddItem "Excel":         col.Add "File | C:\Program Files\Microsoft Office\OFFICE16\EXCEL.EXE"  ' Microsoft Excel 11.0 Object Library
    Combo1.AddItem "SHDocVwCtl":    col.Add "File | C:\WINDOWS\system32\ieframe.oca"                   ' Microsoft Internet Controls

'   Combo1.AddItem "rvbparsero":    col.Add "File | C:\Programme\Microsoft Visual Studio\Common\Tools\VS-Ent98\vmodeler\rvbparsero.dll" 'Rose 4.0/Visual Basic Parser
'

'   Combo1.AddItem "TabDlg":        col.Add "File |
   ' Combo1.AddItem
   Combo1.ListIndex = 0
   UpdateForm
   Call SetMnuHelpstringChecked(MHelpCode.MyLCID)
End Sub

Sub ResetSearch()
    SearchIndex = -1
    SearchForName = ""
End Sub

Private Sub Form_Resize()
   Dim l As Single, t As Single, W As Single, H As Single
   Dim brdr As Single: 'brdr = 8 * IIf(Me.ScaleMode = vbTwips, Screen.TwipsPerPixelX, 1)
   l = List1.Left: t = List1.Top
   W = List1.Width
   H = Me.ScaleHeight - t - brdr
   If W > 0 And H > 0 Then List1.Move l, t, W, H
   l = l + W + brdr
   W = Me.ScaleWidth - l - brdr
   If W > 0 And H > 0 Then
      Text1.Move l, t, W, H
      Text2.Width = W
   End If
End Sub

Private Sub Combo1_Click()
   Call OpenTlB(Combo1.ListIndex, True)
End Sub

Private Sub List1_Click()
   mnuExtrasStatisticEvents.Checked = False
   mnuExtrasStatisticMethod.Checked = False
   Call UpdateCode
End Sub

Private Sub mnuCodeSearch_Click()
    Dim s As String: s = InputBox("What do you want to find, give a name:", "Search for a name", SearchForName)
    If StrPtr(s) = 0 Then Exit Sub 'Cancel
    SearchForName = s
    
    'If Len(s) Then
    '    SearchForName = s
    SearchIndex = -1
    SearchIndex = SearchNext(SearchForName, SearchIndex)
        
    mnuCodeSearchNext.Enabled = True
        
'        Dim si As String
'        Dim i As Long
'        Dim n As Long: n = List1.ListCount
'        For i = 0 To n - 1
'            si = List1.List(i)
'            If InStr(1, si, SearchForName, vbTextCompare) > 0 Then
'                'List1.Text = si
'                List1.ListIndex = i
'            End If
'        Next
'    End If
End Sub
Function SearchNext(aName As String, ByVal startIndex As Long) As Long
    'return the found Index
    Dim si As String
    Dim i As Long
    Dim n As Long: n = List1.ListCount
    For i = startIndex + 1 To n - 1
        si = List1.List(i)
        If InStr(1, si, aName, vbTextCompare) > 0 Then
            'List1.Text = si
            List1.ListIndex = i
            SearchNext = i
            Exit Function
        End If
    Next
End Function

Private Sub mnuCodeSearchNext_Click()
    SearchIndex = SearchNext(SearchForName, SearchIndex)
End Sub

Private Sub mnuHelpInfo_Click()
    FAbout.Show vbModal, Me
End Sub

Private Sub mnuTlbOpenFile_Click()
   'On Error GoTo catch
   If Not myTlb Is Nothing Then
      OFD.InitDir = GetDir(myTlb.FileName)
   Else
      OFD.InitDir = "C:\Windows\System32\"
   End If
   Dim flt As String
   flt = flt & "VB-Komponenten (*.ocx, *.oca)|*.ocx;*.oca|"
   flt = flt & "ActiveXdlls (*.dll)|*.dll|"
   flt = flt & "Typelibraries (*.tlb, *.olb)|*.tlb;*olb|"
   flt = flt & "(ocx, dll, tlb) |*.ocx;*.oca;*.dll;*.tlb;*.olb|"
   flt = flt & "Alle Dateien (*.*)|*.*"
   OFD.Filter = flt
   OFD.FilterIndex = 1
   'OFD.CancelError = True
   'OFD.FileName = "D:\VB60To_dotNETClasses\VB6_To_Delphi\FastGeoD\Tlb\MBOFastGeoD.tlb"
   'OFD.FileName = "C:\Windows\System32\
   'OFD.FileName = ""
   'OFD.FileName = "msvbvm60.dll"
   'OFD.FileName = "tlbinf32.dll"
   'OFD.FileName = "stdole2.tlb"
   'OFD.FileName = "C:\Programme\Microsoft Visual Studio\VB98\VB6.OLB"
   OFD.ShowOpen
   If Len(OFD.FileName) > 0 Then
      Dim f As String: f = OFD.FileName
      Dim s As String: s = "File | " & f
      Combo1.AddItem myTlb.TypeLibInfo.Name:  col.Add s
      Combo1.ListIndex = Combo1.ListCount - 1
      Call OpenFile(f)
   End If
End Sub
Private Sub mnuTlbOpenReg_Click()
   Dim StrVal As String
   StrVal = InputBox("Geben sie ein: guid maj min lcid")
   If StrVal <> vbNullString Then
      Dim sa() As String: sa = Split(StrVal, " ")
      Call OpenReg(sa)
      Dim s As String
      s = Join(sa, " | ") ' sa(1) & " | " & sa(2) & " | " & sa(3)
      Combo1.AddItem myTlb.TypeLibInfo.Name:  col.Add s
   End If
End Sub
Private Sub mnuTlbOpenDir_Click()
   If Not myTlb Is Nothing Then
      Shell "Explorer " & GetDir(myTlb.FileName), vbNormalFocus
   End If
End Sub
Private Sub mnuTlbSaveAs_Click()
   OFD.ShowSave
   If OFD.FileName <> "" Then
      Dim f As String: f = OFD.FileName
      Dim d As String
      Dim i As Integer
      For i = 0 To Combo1.ListCount - 1
         Call OpenTlB(i, False)
         d = GetDir(f) & "\VB"
         If Not IsDir(d) Then Call MkDir(d)
         Call myTlb.SaveFiles(d, New CoderVB)
         d = GetDir(f) & "\Java"
         If Not IsDir(d) Then Call MkDir(d)
         Call myTlb.SaveFiles(d, New CoderJava)
      Next
   End If
End Sub
Private Sub mnuTlbExit_Click()
   Unload Me
End Sub

Private Sub mnuCodeVB_Click()
   Call SelectCoder(New CoderVB)
End Sub
Private Sub mnuCodeJava_Click()
   Call SelectCoder(New CoderJava)
End Sub
Private Sub mnuCodeCpp_Click()
   Call SelectCoder(New CoderCpp)
End Sub

Private Sub SelectCoder(aCoder As ICoder)
   Set myCoder = aCoder
   Me.mnuCodeJava.Checked = False
   Me.mnuCodeVB.Checked = False
   Me.mnuCodeCpp.Checked = False
   Select Case True
   Case TypeOf aCoder Is CoderVB:    Me.mnuCodeVB.Checked = True
   Case TypeOf aCoder Is CoderJava:  Me.mnuCodeJava.Checked = True
   Case TypeOf aCoder Is CoderCpp:   Me.mnuCodeCpp.Checked = True
   End Select
   UpdateCode
End Sub
Private Sub mnuCodeIndent_Click()
   'Dim s As String: s = InputBox("Indentation?", "Indentation Size:", CStr(MInd.IndentSize))
   'If Len(s) And IsNumeric(s) Then MInd.MyIndentSize = CLng(s)
   FIndent.Show vbModal, Me
   UpdateCode
End Sub

Private Sub mnuHelpstringLcidENUS_Click()
   MHelpCode.MyLCID = 0
   Call SetMnuHelpstringChecked(MHelpCode.MyLCID)
   UpdateCode
End Sub
Private Sub mnuHelpstringLcidDEDE_Click()
   MHelpCode.MyLCID = 1
   Call SetMnuHelpstringChecked(MHelpCode.MyLCID)
   UpdateCode
End Sub
Private Sub mnuHelpstringLcidSpecify_Click()
   Dim s As String: s = InputBox("LCID:", "LCID?", "&H" & Hex$(MHelpCode.MyLCID))
   If s <> vbNullString Then
      If IsNumeric(s) Then
         MHelpCode.MyLCID = CLng(s)
         Call SetMnuHelpstringChecked(MHelpCode.MyLCID)
         UpdateCode
      End If
   End If
End Sub
Private Sub mnuHelpstringAuthor_Click()
   If Len(MHelpCode.MyAuthorName) = 0 Then
      MHelpCode.MyAuthorName = "OlimilO"
   End If
   MHelpCode.MyAuthorName = InputBox("Authors name:", "Authors name?", MHelpCode.MyAuthorName)
   UpdateCode
End Sub
Private Sub SetMnuHelpstringChecked(ByVal LCID As Long)
   mnuHelpstringLcidENUS.Checked = IIf(LCID = 0, True, False)
   mnuHelpstringLcidDEDE.Checked = IIf(LCID = 1, True, False)
   mnuHelpstringLcidSpecify.Checked = IIf(((LCID <> 0) And (LCID <> 1)), True, False)
   Dim sa() As String: sa = Split(mnuHelpstringLcidSpecify.Caption, " ")
   sa(2) = "&&H" & Hex$(LCID)
   mnuHelpstringLcidSpecify.Caption = Join(sa, " ")
End Sub
Private Sub mnuExtrasStatisticEvents_Click()
   Text1.Text = myTlb.getEventsStatistic
   mnuExtrasStatisticEvents.Checked = True
   mnuExtrasStatisticMethod.Checked = False
End Sub
Private Sub mnuExtrasStatisticMethod_Click()
   Text1.Text = myTlb.getMethodsStatistic
   mnuExtrasStatisticEvents.Checked = False
   mnuExtrasStatisticMethod.Checked = True
End Sub


Private Sub OpenTlB(i As Integer, ByVal bUpdateView As Boolean)
   Dim s As String: s = col.item(i + 1)
   Dim sa() As String: sa = Split(s, " | ")
   Dim aNewTlb As New TypeLibrary
   If sa(0) = "File" Then
      'Call aNewTlb.LoadFile(aFilename)
      Call aNewTlb.LoadFile(sa(1))
   Else
      Call aNewTlb.LoadReg(sa(0), sa(1), sa(2), sa(3))
   End If
   If Not aNewTlb.TypeLibInfo Is Nothing Then
      Set myTlb = aNewTlb
      If bUpdateView Then
         Call UpdateForm
         Call UpdateView
      End If
   End If
End Sub
Private Sub OpenFile(aFilename As String)
   Dim aNewTlb As New TypeLibrary
   Call aNewTlb.LoadFile(aFilename)
   Set NewTlB = aNewTlb
End Sub
Private Sub OpenReg(sa() As String)
   Dim aNewTlb As New TypeLibrary
   Call aNewTlb.LoadReg(sa(0), sa(1), sa(2), sa(3))
   Set NewTlB = aNewTlb
End Sub
Private Property Set NewTlB(aNewTlb As TypeLibrary)
   If Not aNewTlb.TypeLibInfo Is Nothing Then
      Set myTlb = aNewTlb
      Call UpdateForm
      Call UpdateView
   End If
End Property

Private Sub UpdateForm()
   Me.Caption = "TyITEx - [" & myTlb.TypeLibInfo.Name & ": " & myTlb.FileName & "]"
End Sub
Private Sub UpdateView()
   Call myTlb.TypesToListBox(List1)
   Text2.Text = myTlb.ToString 'myTlb.TypeLibInfo.Name & " " & myTlb.TypeLibInfo.Guid
   UpdateCode
End Sub
Private Sub UpdateCode()
   If Me.mnuExtrasStatisticEvents.Checked = True Then
      Text1.Text = myTlb.getEventsStatistic
   ElseIf Me.mnuExtrasStatisticMethod.Checked = True Then
      Text1.Text = myTlb.getEventsStatistic
   Else
      If List1.ListIndex < 0 Then
         List1.ListIndex = 0
      End If
      Dim item As String: item = List1.List(List1.ListIndex)
      If Len(item) > 0 Then
         Call myTlb.ItemToTextBox(myCoder, item, Text1)
         'Call Scin1.addText(myTlb.ItemToString(myCoder, item)) '???
         'Scin1.Redraw = True
      End If
   End If
End Sub

