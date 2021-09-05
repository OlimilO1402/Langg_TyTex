VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private myTLIApp As New TLI.TLIApplication
Private myTli    As TLI.TypeLibInfo
Private myFile   As String

Public Sub NewC(aFilename As String)
   Call LoadFile(aFilename)
End Sub
Public Function LoadFile(aFilename As String) As Boolean
Try: On Error GoTo Catch
   Dim newTLI As TypeLibInfo
   myFile = aFilename
   Set newTLI = myTLIApp.TypeLibInfoFromFile(aFilename)
   Set myTli = newTLI
   LoadFile = True
   Exit Function
Catch:
   MsgBox Err.Description
End Function
Public Function LoadReg(ByVal aGUID As String, ByVal maj As Integer, ByVal Min As Integer, ByVal LCID As Long) As Boolean
Try: On Error GoTo Catch
   Dim newTLI As TypeLibInfo
   'myFile = aGUID & " " & CStr(maj) & " " & CStr(min) & " " & CStr(LCID)
   Set newTLI = myTLIApp.TypeLibInfoFromRegistry(aGUID, maj, Min, LCID)
   myFile = newTLI.ContainingFile
   Set myTli = newTLI
   LoadReg = True
   Exit Function
Catch:
   MsgBox Err.Description
End Function
Public Property Get FileName() As String
   FileName = myFile
End Property
'Public Property Get Name() As String
'   myTli.ShowLibName = True
'   Name = myTli.Name & " " & myTli.Guid
'   'Name = myTli.Guid
'End Property
Public Function ToString() As String
   ToString = myTli.Name & " " & myTli.Guid & " " & myTli.MajorVersion & " " & myTli.MinorVersion & " " & myTli.LCID
End Function
Public Property Get TypeLibInfo() As TypeLibInfo
   Set TypeLibInfo = myTli
End Property

Public Sub TypesToListBox(aLB As ListBox)
    aLB.Clear
    EnumsToListBox aLB
    UDTypesToListBox aLB
    ModulesToListBox aLB
    ClassesToListBox aLB
    InterfacesToListBox aLB
    AliasesToListBox aLB
    UnionsToListBox aLB
End Sub
Private Sub EnumsToListBox(aLB As ListBox)
    Dim ci As TLI.ConstantInfo
    For Each ci In myTli.Constants
        Call aLB.AddItem("Enum " & ci.Name)
    Next
End Sub
Private Sub UDTypesToListBox(aLB As ListBox)
    Dim ri As TLI.RecordInfo
    For Each ri In myTli.Records
        Call aLB.AddItem("Type " & ri.Name)
    Next
End Sub
Private Sub ModulesToListBox(aLB As ListBox)
    Dim di As DeclarationInfo
    For Each di In myTli.Declarations
        Call aLB.AddItem("Module " & di.Name)
    Next
End Sub
Private Sub ClassesToListBox(aLB As ListBox)
    Dim ci As TLI.CoClassInfo
    For Each ci In myTli.CoClasses
        Call aLB.AddItem("Class " & ci.Name)
    Next
End Sub
Private Sub InterfacesToListBox(aLB As ListBox)
    Dim ii As TLI.InterfaceInfo
    For Each ii In myTli.Interfaces
        Call aLB.AddItem("Interface " & ii.Name)
    Next
End Sub
Private Sub AliasesToListBox(aLB As ListBox)
    Dim ai As TLI.IntrinsicAliasInfo
    For Each ai In myTli.IntrinsicAliases
        Call aLB.AddItem("Alias " & ai.Name)
    Next
End Sub
Private Sub UnionsToListBox(aLB As ListBox)
    Dim ui As TLI.UnionInfo
    For Each ui In myTli.Unions
        Call aLB.AddItem("Union " & ui.Name)
    Next
End Sub

Public Sub ItemToTextBox(aCoder As ICoder, aItem As String, aTB As TextBox)
   Dim s() As String: s = Split(aItem, " ")
   aTB.Text = ""
   Select Case s(0)
   Case "Enum":       aTB.Text = aCoder.EnumToCode(GetEnum(s(1)))
   Case "Type":       aTB.Text = aCoder.TypeToCode(GetType(s(1)))
   Case "Module":     aTB.Text = aCoder.ModuleToCode(GetModul(s(1)))
   Case "Interface":  aTB.Text = aCoder.InterfaceToCode(GetInterface(s(1)))
   Case "Class":      aTB.Text = aCoder.ClassToCode(GetClass(s(1)))
   Case "Alias":      aTB.Text = aCoder.AliasToCode(GetAlias(s(1)))
   Case "Union":      aTB.Text = aCoder.UnionToCode(GetUnion(s(1)))
   End Select
End Sub
Public Function ItemToString(aCoder As ICoder, aItem As String) As String
   Dim s() As String: s = Split(aItem, " ")
   Select Case s(0)
   Case "Enum":       ItemToString = aCoder.EnumToCode(GetEnum(s(1)))
   Case "Type":       ItemToString = aCoder.TypeToCode(GetType(s(1)))
   Case "Module":     ItemToString = aCoder.ModuleToCode(GetModul(s(1)))
   Case "Interface":  ItemToString = aCoder.InterfaceToCode(GetInterface(s(1)))
   Case "Class":      ItemToString = aCoder.ClassToCode(GetClass(s(1)))
   Case "Alias":      ItemToString = aCoder.AliasToCode(GetAlias(s(1)))
   Case "Union":      ItemToString = aCoder.UnionToCode(GetUnion(s(1)))
   End Select
End Function

Private Function GetEnum(aEnumName As String) As ConstantInfo
   For Each GetEnum In myTli.Constants
      If GetEnum.Name = aEnumName Then Exit Function
   Next
End Function
Private Function GetType(aTypeName As String) As RecordInfo
   For Each GetType In myTli.Records
      If GetType.Name = aTypeName Then Exit Function
   Next
End Function
Private Function GetModul(aModulName As String) As DeclarationInfo
   For Each GetModul In myTli.Declarations
      If GetModul.Name = aModulName Then Exit Function
   Next
End Function
Private Function GetInterface(aInterfaceName As String) As InterfaceInfo
   For Each GetInterface In myTli.Interfaces
      If GetInterface.Name = aInterfaceName Then Exit Function
   Next
End Function
Private Function GetClass(aClassName As String) As CoClassInfo
   For Each GetClass In myTli.CoClasses
      If GetClass.Name = aClassName Then Exit Function
   Next
End Function
Private Function GetAlias(aAliasName As String) As IntrinsicAliasInfo
   For Each GetAlias In myTli.IntrinsicAliases
      If GetAlias.Name = aAliasName Then Exit Function
   Next
End Function
'Unions ' not supported by VB and Java
Private Function GetUnion(aUnionName As String) As UnionInfo
   For Each GetUnion In myTli.Unions
      If GetUnion.Name = aUnionName Then Exit Function
   Next
End Function

Public Sub SaveFiles(ByVal ToDir As String, aCoder As ICoder)
   '1. alles in eine Datei speichern?
   '2. alles in mehrere Dateien speichern
   '2.1 Alle Enums in eine Datei speichern
   '    Alle UD-Types in eine Datei speichern
   '    jede Klasse und Interface in eine eigene Datei speichern
   '    jeweils in ein Unterverzeichnis das so heißt wie die tlb
   
   'save all enums
   ToDir = ToDir & "\" & myTli.Name & "\"
   If Dir(ToDir) = "" Then Call MkDir(ToDir)
   
   Call SaveEnums(ToDir & "Enums.bas", aCoder)
   Call SaveTypes(ToDir & "Types.bas", aCoder)
   Call SaveModules(ToDir, aCoder)
   Call SaveClasses(ToDir, aCoder)
   Call SaveInterfaces(ToDir, aCoder)
   
End Sub
Private Function SaveEnums(FNm As String, aCoder As ICoder)
   Call MFStream.OOpen(FNm)
   Dim ci As TLI.ConstantInfo
   For Each ci In myTli.Constants
      Call MFStream.Append(aCoder.EnumToCode(ci))
   Next
   Call MFStream.CClose
End Function
Private Function SaveTypes(FNm As String, aCoder As ICoder)
   Call MFStream.OOpen(FNm)
   Dim ri As TLI.RecordInfo
   For Each ri In myTli.Records
      Call MFStream.Append(aCoder.TypeToCode(ri))
   Next
   Call MFStream.CClose
End Function
Private Function SaveModules(ToDir As String, aCoder As ICoder)
   Dim mi As TLI.DeclarationInfo
   For Each mi In myTli.Declarations
      Call MFStream.SaveString(ToDir & mi.Name & ".bas", aCoder.ModuleToCode(mi))
   Next
End Function
Private Function SaveClasses(ToDir As String, aCoder As ICoder)
   Dim ci As TLI.CoClassInfo
   For Each ci In myTli.CoClasses
      Call MFStream.SaveString(ToDir & ci.Name & ".cls", aCoder.ClassToCode(ci))
   Next
End Function
Private Function SaveInterfaces(ToDir As String, aCoder As ICoder)
   Dim ii As TLI.InterfaceInfo
   For Each ii In myTli.Interfaces
      Call MFStream.SaveString(ToDir & ii.Name & ".cls", aCoder.InterfaceToCode(ii))
   Next
End Function

Public Function getEventsStatistic() As String
'die DefaultEventInterfaces aller CoClasses durchwandern
'jedes Event in eine Collection packen
'und zugehörige Controls finden
   Dim keys As New Collection
   Dim ci As CoClassInfo
   Dim ii As InterfaceInfo
   Dim mi As MemberInfo
   On Error Resume Next
   'alle vorhandenen verschiedenen Events sammeln
   For Each ci In myTli.CoClasses
      Set ii = ci.DefaultEventInterface
      If Not ii Is Nothing Then
         For Each mi In ii.Members
            If Not ContainsKey(keys, mi.Name) Then
               keys.Add mi.Name, mi.Name
            End If
         Next
      End If
   Next
   'Call MUtil.SortCol(keys)
   Dim sKeys() As String
   Call CopyColToArray(sKeys, keys)
   'Call Sort(sKeys)
   'Debug.Print Join(sKeys, " ")
   Dim s As String
   Dim i As Integer
   'für jedes Event die Klasse dazunehmen die das Event hat
   ReDim sVals(0 To UBound(sKeys)) As String
   Dim L As Integer: L = GetMaxStringLen(sKeys)
   Dim sHL As String: sHL = Space(L + 2)
   For Each ci In myTli.CoClasses
      sHL = sHL & ci.Name & "; "
   Next
   For i = 0 To UBound(sKeys)
      s = PadRight(sKeys(i), L) & ": "
      For Each ci In myTli.CoClasses
         'If ContainsMember(ci.DefaultEventInterface, sKeys(i)) Then
         '   s = s & ci.Name & "; "
         'End If
         If ContainsMember(ci.DefaultEventInterface, sKeys(i)) Then
            s = s & PadRight("  *", Len(ci.Name) + 2)
         Else
            s = s & PadRight("   ", Len(ci.Name) + 2)
         End If
      Next
      sVals(i) = s
   Next
   getEventsStatistic = sHL & vbCrLf & Join(sVals, vbCrLf)

End Function
Public Function getMethodsStatistic() As String
'die DefaultInterfaces aller CoClasses durchwandern
'jede Methode in eine Collection packen
'und zugehörige Controls finden
   Dim keys As New Collection
   Dim ci As CoClassInfo
   Dim ii As InterfaceInfo
   Dim mi As MemberInfo
   On Error Resume Next
   'alle vorhandenen verschiedenen Methoden sammeln
   For Each ci In myTli.CoClasses
      Set ii = ci.DefaultInterface
      If Not ii Is Nothing Then
         For Each mi In ii.Members
            If Not ContainsKey(keys, mi.Name) Then
               keys.Add mi.Name, mi.Name
            End If
         Next
      End If
   Next
   
   'Call MUtil.SortCol(keys)
   Dim sKeys() As String
   Call CopyColToArray(sKeys, keys)
   'Call Sort(sKeys)
   'Debug.Print Join(sKeys, " ")
   Dim s As String
   Dim i As Integer
   'für jede Methode die Klasse dazunehmen die diese Methode hat
   ReDim sVals(0 To UBound(sKeys)) As String
   Dim L As Integer: L = GetMaxStringLen(sKeys)
   Dim sHL As String: sHL = Space(L + 2)
   For Each ci In myTli.CoClasses
      sHL = sHL & ci.Name & "; "
   Next
   For i = 0 To UBound(sKeys)
      s = PadRight(sKeys(i), L) & ": "
      For Each ci In myTli.CoClasses
         'If ContainsMember(ci.DefaultEventInterface, sKeys(i)) Then
         '   s = s & ci.Name & "; "
         'End If
         If ContainsMember(ci.DefaultInterface, sKeys(i)) Then
            s = s & PadRight("  *", Len(ci.Name) + 2)
         Else
            s = s & PadRight("   ", Len(ci.Name) + 2)
         End If
      Next
      sVals(i) = s
   Next
   getMethodsStatistic = sHL & vbCrLf & Join(sVals, vbCrLf)

End Function

'hmm aber so macht das eigentlich keinen Sinn,
'man müßte zuerst nach der Anzahl der Methoden/Events und erst dann nach dem Name alphabetisch sortieren.
