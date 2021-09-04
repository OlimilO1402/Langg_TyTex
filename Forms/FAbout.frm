VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Info zu meiner Anwendung"
   ClientHeight    =   2310
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5775
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1594.403
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   5423.023
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "FAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2880
      TabIndex        =   0
      Top             =   1800
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Systeminfo..."
      Height          =   345
      Left            =   4320
      TabIndex        =   2
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Innen ausgefüllt
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1107.799
      Y2              =   1107.799
   End
   Begin VB.Label lblDescription 
      Caption         =   "Beschreibung"
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   1050
      TabIndex        =   3
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Name der Anwendung"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   1845
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1118.153
      Y2              =   1118.153
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   345
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   1965
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Registrierungsschlüssel-Sicherheitsoptionen...
'Const READ_CONTROL       As Long = &H20000
'Const KEY_QUERY_VALUE    As Long = &H1
'Const KEY_SET_VALUE      As Long = &H2
'Const KEY_CREATE_SUB_KEY As Long = &H4
'Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
'Const KEY_NOTIFY         As Long = &H10
'Const KEY_CREATE_LINK    As Long = &H20
'Const KEY_ALL_ACCESS     As Long = KEY_QUERY_VALUE + KEY_SET_VALUE + _
'                                   KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
'                                   KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
'
'' Registrierungsschlüssel-Stammtypen...
'Const HKEY_LOCAL_MACHINE = &H80000002
'Const ERROR_SUCCESS = 0
'Const REG_SZ = 1                         ' Null-terminierte Unicode-Zeichenfolge
'Const REG_DWORD = 4                      ' 32-Bit-Zahl
'
'Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
'Const gREGVALSYSINFOLOC = "MSINFO"
'Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
'Const gREGVALSYSINFO = "PATH"
'
'Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    Me.Caption = "Info zu " & App.EXEName
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.EXEName
    lblDescription.Caption = App.FileDescription
End Sub

Private Sub cmdSysInfo_Click()
    StartSysInfo
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Public Sub StartSysInfo()

Try: On Error GoTo Catch
    Dim SysInfoPath As String: SysInfoPath = "MSINFO32.EXE"
    
    Shell SysInfoPath, vbNormalFocus

'    Exit Sub
'    Dim rc As Long
'
'
'    ' Versuchen, den Systeminfo-Programmpfad/-namen aus der Registrierung abzurufen...
'    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
'    ' Versuchen, nur den Systeminfo-Programmpfad aus der Registrierung abzurufen...
'    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
'        ' Überprüfen, ob bekannte 32-Dateiversion vorhanden ist
'        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
'            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
'            Call Shell(SysInfoPath, vbNormalFocus)
'            Exit Sub
'        ' Fehler - Datei wurde nicht gefunden...
'        Else
'            SysInfoPath = "MSINFO32.EXE"
'            Call Shell(SysInfoPath, vbNormalFocus)
'            Exit Sub
'            'GoTo SysInfoErr
'        End If
'    ' Fehler - Registrierungseintrag wurde nicht gefunden...
'    Else
'        'GoTo Catch
'        SysInfoPath = "MSINFO32.EXE"
'        Call Shell(SysInfoPath, vbNormalFocus)
'        Exit Sub
'
'    End If
'
'    'SysInfoPath = "MSINFO32.EXE"
'
'
    Exit Sub
Catch:
    MsgBox "Systeminformationen sind momentan nicht verfügbar", vbOKOnly
End Sub

'Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'    Dim i As Long                                           ' Schleifenzähler
'    Dim rc As Long                                          ' Rückgabe-Code
'    Dim hKey As Long                                        ' Zugriffsnummer für einen offenen Registrierungsschlüssel
'    Dim hDepth As Long                                      '
'    Dim KeyValType As Long                                  ' Datentyp eines Registrierungsschlüssels
'    Dim tmpVal As String                                    ' Temporärer Speicher eines Registrierungsschlüsselwertes
'    Dim KeyValSize As Long                                  ' Größe der Registrierungsschlüsselvariablen
'    '------------------------------------------------------------
'    ' Registrierungsschlüssel unter KeyRoot {HKEY_LOCAL_MACHINE...} öffnen
'    '------------------------------------------------------------
'    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Registrierungsschlüssel öffnen
'
'    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln...
'
'    tmpVal = String$(1024, 0)                             ' Platz für Variable reservieren
'    KeyValSize = 1024                                       ' Größe der Variable markieren
'
'    '------------------------------------------------------------
'    ' Registrierungsschlüsselwert abrufen...
'    '------------------------------------------------------------
'    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)                            ' Schlüsselwert abrufen/erstellen
'
'    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln
'
'    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 fügt null-terminierte Zeichenfolge hinzu...
'        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null gefunden, aus Zeichenfolge extrahieren
'    Else                                                    ' Keine null-terminierte Zeichenfolge für WinNT...
'        tmpVal = Left(tmpVal, KeyValSize)                   ' Null nicht gefunden, nur Zeichenfolge extrahieren
'    End If
'    '------------------------------------------------------------
'    ' Schlüsselwerttyp für Konvertierung bestimmen...
'    '------------------------------------------------------------
'    Select Case KeyValType                                  ' Datentypen durchsuchen...
'    Case REG_SZ                                             ' Zeichenfolge für Registrierungsschlüsseldatentyp
'        KeyVal = tmpVal                                     ' Zeichenfolgenwert kopieren
'    Case REG_DWORD                                          ' Registrierungsschlüsseldatentyp DWORD
'        For i = Len(tmpVal) To 1 Step -1                    ' Jedes Bit konvertieren
'            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Wert Zeichen für Zeichen erstellen
'        Next
'        KeyVal = Format$("&h" + KeyVal)                     ' DWORD in Zeichenfolge konvertieren
'    End Select
'
'    GetKeyValue = True                                      ' Erfolgreiche Ausführung zurückgeben
'    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
'    Exit Function                                           ' Beenden
'
'GetKeyError:      ' Bereinigen, nachdem ein Fehler aufgetreten ist...
'    KeyVal = ""                                             ' Rückgabewert auf leere Zeichenfolge setzen
'    GetKeyValue = False                                     ' Fehlgeschlagene Ausführung zurückgeben
'    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
'End Function
