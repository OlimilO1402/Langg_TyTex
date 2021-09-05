VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "TyTex"
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10215
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Panel1 
      BorderStyle     =   0  'Kein
      Height          =   6015
      Left            =   0
      ScaleHeight     =   6015
      ScaleWidth      =   9975
      TabIndex        =   2
      Top             =   360
      Width           =   9975
      Begin VB.TextBox TBCode 
         BeginProperty Font 
            Name            =   "Consolas"
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
         TabIndex        =   4
         Top             =   0
         Width           =   6735
      End
      Begin VB.ListBox LBTypes 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5580
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   6735
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   3015
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
'TyTex, the tight typelibinfo extractor
Dim myTlb   As TypeLibrary    'the typelibrary file or from registry
Dim col     As New Collection 'for selecting from combobox
Dim myCoder As ICoder         'what do you want? VB, VB.NET, Java, Cpp, Delphi ... ?
Dim SearchForName As String
Dim SearchIndex   As Long
Dim WithEvents Splitter1 As Splitter
Attribute Splitter1.VB_VarHelpID = -1

Private Sub Form_Load()
    
    ResetSearch
    
    MInd.IndentSize = 4
    MHelpCode.MyAuthorName = "YourNameHere"
    Me.mnuCodeVB.Checked = True
    Set myCoder = New CoderVB
    'Combo1.AddItem "MBOFastGeoD":   col.Add "File | D:\VB60To_dotNETClasses\VB6_To_Delphi\FastGeoD\Tlb\MBOFastGeoD.tlb"
    Combo1.AddItem "VBA":           col.Add "File | C:\WINDOWS\system32\msvbvm60.dll"
    Combo1.AddItem "VBRUN":         col.Add "{EA544A21-C82D-11D1-A3E4-00A0C90AEA82} | 6 | 0 | 9"
    
    'Combo1.AddItem "VB":            col.Add "File | C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.OLB"
    'Combo1.AddItem "VBIDE":         col.Add "File | C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6EXT.OLB" '
    
    Combo1.AddItem "VB":            col.Add "File | C:\Program Files (x86)\VB98\VB98\VB6.OLB"
    Combo1.AddItem "VBIDE":         col.Add "File | C:\Program Files (x86)\VB98\VB98\VB6EXT.OLB" '
    
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
'   Combo1.AddItem "TabDlg":        col.Add "File |
   
   Combo1.ListIndex = 0
   UpdateForm
   Call SetMnuHelpstringChecked(MHelpCode.MyLCID)
   Set Splitter1 = New Splitter
   Splitter1.New_ False, Me, Me.Panel1, "Splitter1", Me.LBTypes, Me.TBCode
   Splitter1.LeftTopPos = Me.LBTypes.Width
   Splitter1.BorderStyle = bsXPStyl

End Sub

Sub ResetSearch()
    SearchIndex = -1
    SearchForName = ""
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    T = Panel1.Top
    W = Me.ScaleWidth
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then
        Panel1.Move L, T, W, H
        Text2.Width = TBCode.Width
    End If
End Sub

Public Property Get OFD() As OpenFileDialog
    Set OFD = New OpenFileDialog
    OFD.Filter = GetFilter
    OFD.FilterIndex = 1
    OFD.InitialDirectory = App.Path
End Property

Public Property Get SFD() As SaveFileDialog
    Set SFD = New SaveFileDialog
    SFD.Filter = GetFilter
    SFD.FilterIndex = 1
    SFD.InitialDirectory = App.Path
End Property

Private Function GetFilter() As String
    Dim flt As String
    flt = "VB-Components [*.ocx, *.oca] |*.ocx;*.oca|" & _
          "ActiveXDlls [*.dll]|*.dll|" & _
          "Typelibraries [*.tlb, *.olb]|*.tlb;*olb|" & _
          "[ocx, dll, tlb]|*.ocx;*.oca;*.dll;*.tlb;*.olb|" & _
          "All files [*.*]|*.*"
    GetFilter = flt
End Function

Private Sub Combo1_Click()
   Call OpenTlB(Combo1.ListIndex, True)
End Sub

Private Sub LBTypes_Click()
   mnuExtrasStatisticEvents.Checked = False
   mnuExtrasStatisticMethod.Checked = False
   Call UpdateCode
End Sub

Private Sub mnuCodeSearch_Click()
    Dim s As String: s = InputBox("What do you want to find, give a name:", "Search for a name", SearchForName)
    If StrPtr(s) = 0 Then Exit Sub 'Cancel
    SearchForName = s
    SearchIndex = -1
    SearchIndex = SearchNext(SearchForName, SearchIndex)
    mnuCodeSearchNext.Enabled = True
End Sub

Function SearchNext(aName As String, ByVal startIndex As Long) As Long
    'return the found Index
    Dim si As String
    Dim i As Long
    Dim n As Long: n = LBTypes.ListCount
    For i = startIndex + 1 To n - 1
        si = LBTypes.List(i)
        If InStr(1, si, aName, vbTextCompare) > 0 Then
            'LBTypes.Text = si
            LBTypes.ListIndex = i
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
    Dim aOFD As OpenFileDialog: Set aOFD = Me.OFD
    If Not myTlb Is Nothing Then 'no not IIf !!!
        aOFD.InitialDirectory = GetDir(myTlb.FileName)
    Else
        aOFD.InitialDirectory = "C:\Windows\System32\"
    End If
    If aOFD.ShowDialog = vbCancel Then Exit Sub
    Dim f As String: f = aOFD.FileName
    If Len(f) > 0 Then
        Dim s As String: s = "File | " & f
        If OpenFile(f) Then
            Combo1.AddItem myTlb.TypeLibInfo.Name:  col.Add s
            Combo1.ListIndex = Combo1.ListCount - 1
        End If
    End If
End Sub

Private Sub mnuTlbOpenReg_Click()
    Dim s As String: s = InputBox("Please give: guid maj min lcid")
    If StrPtr(s) = 0 Then Exit Sub
    Dim sa() As String: sa = Split(s, " ")
    If OpenReg(sa) Then
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
   Dim aSFD As SaveFileDialog: Set aSFD = Me.SFD
   'OFD.ShowSave
   If aSFD.ShowDialog = vbCancel Then Exit Sub
   If aSFD.FileName <> "" Then
      Dim f As String: f = aSFD.FileName
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
    TBCode.Text = myTlb.getEventsStatistic
    mnuExtrasStatisticEvents.Checked = True
    mnuExtrasStatisticMethod.Checked = False
End Sub

Private Sub mnuExtrasStatisticMethod_Click()
    TBCode.Text = myTlb.getMethodsStatistic
    mnuExtrasStatisticEvents.Checked = False
    mnuExtrasStatisticMethod.Checked = True
End Sub

Private Sub OpenTlB(i As Integer, ByVal bUpdateView As Boolean)
    Dim s As String: s = col.item(i + 1)
    Dim sa() As String: sa = Split(s, " | ")
    Dim aNewTlb As New TypeLibrary
    If sa(0) = "File" Then
        aNewTlb.LoadFile sa(1)
    Else
        aNewTlb.LoadReg sa(0), sa(1), sa(2), sa(3)
    End If
    If Not aNewTlb.TypeLibInfo Is Nothing Then
        Set myTlb = aNewTlb
        If bUpdateView Then
            UpdateForm
            UpdateView
        End If
    End If
End Sub

Private Function OpenFile(aFilename As String) As Boolean
    Dim aNewTlb As New TypeLibrary
    OpenFile = aNewTlb.LoadFile(aFilename)
    If OpenFile Then Set NewTlB = aNewTlb
End Function

Private Function OpenReg(sa() As String) As Boolean
    Dim aNewTlb As New TypeLibrary
    OpenReg = aNewTlb.LoadReg(sa(0), sa(1), sa(2), sa(3))
    If OpenReg Then Set NewTlB = aNewTlb
End Function

Private Property Set NewTlB(aNewTlb As TypeLibrary)
    If Not aNewTlb.TypeLibInfo Is Nothing Then
        Set myTlb = aNewTlb
        Call UpdateForm
        Call UpdateView
    End If
End Property

Private Sub UpdateForm()
    Me.Caption = App.EXEName & " - [" & myTlb.TypeLibInfo.Name & ": " & myTlb.FileName & "]"
End Sub

Private Sub UpdateView()
    myTlb.TypesToListBox LBTypes
    Text2.Text = myTlb.ToString
    UpdateCode
End Sub

Private Sub UpdateCode()
    If Me.mnuExtrasStatisticEvents.Checked = True Then
        TBCode.Text = myTlb.getEventsStatistic
    ElseIf Me.mnuExtrasStatisticMethod.Checked = True Then
        TBCode.Text = myTlb.getEventsStatistic
    Else
        If LBTypes.ListIndex < 0 Then
            LBTypes.ListIndex = 0
        End If
        Dim item As String: item = LBTypes.List(LBTypes.ListIndex)
        If Len(item) > 0 Then
            myTlb.ItemToTextBox myCoder, item, TBCode
        End If
    End If
End Sub

Private Sub Splitter1_OnMove(Sender As Splitter)
    Combo1.Width = Sender.LeftTopPos
    Text2.Left = Sender.LeftTopPos + Sender.Width
    Text2.Width = TBCode.Width
End Sub
