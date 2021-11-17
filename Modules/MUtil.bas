Attribute VB_Name = "MUtil"
Option Explicit '
'Public Declare Function DllGetDocumentation Lib "activeds" (ByVal lcid As Long, strBuffer As String) As Long 'nop
' ########## '   Enum    ' ########## '
Private Const uFlags As Integer = 31
Private FlagList(0 To uFlags) As Variant 'Decimal 'Currency
' ########## '  Vector   ' ########## '
Private m_TLIArrayBoundsChecked As Boolean
Private m_FixArrayBounds As Boolean
' ########## ' Interface ' ########## '
Private IUnkIDisp(0 To 6) As String
Private m_IUnkIDispFilled As Boolean

' ########## '  Vector   ' ########## '
'due to Errors in TLI
Private Sub CheckTLIArrayBounds()
    Debug.Assert Not m_TLIArrayBoundsChecked
    On Error Resume Next
    Dim Bounds() As Long
    Call TLI.TypeLibInfoFromFile("stdole2.tlb").TypeInfos.NamedItem("GUID").Members(4).ReturnType.ArrayBounds(Bounds)
    If Bounds(1, 2) = 8 Then m_FixArrayBounds = True
    m_TLIArrayBoundsChecked = True
End Sub

Public Function GetArrayBounds(vti As VarTypeInfo, ByRef outBounds() As Long) As Long
    If Not m_TLIArrayBoundsChecked Then CheckTLIArrayBounds
    GetArrayBounds = vti.ArrayBounds(outBounds())
    If m_FixArrayBounds Then
        Dim i As Integer
        For i = 1 To GetArrayBounds
            outBounds(i, 2) = outBounds(i, 2) + 2 * outBounds(i, 1) - 1
        Next
    End If
End Function

' ########## '   Enum    ' ########## '
'used by EnumToString
'returns true if a enum acts like a flaglist
Public Function IsFlags(aEnum As ConstantInfo)
    On Error Resume Next
    If FlagList(uFlags) = 0 Then CreateFlagList
    Dim mi As MemberInfo
    'IsFlags = False 'eh klar
    For Each mi In aEnum.Members
        If Not FlagListContains(mi.Value) Then Exit Function
    Next
    IsFlags = True
End Function
Private Sub CreateFlagList()
    Dim i As Integer
    For i = 0 To uFlags
        FlagList(i) = CDec(CDec(2) ^ CDec(i))
    Next
End Sub
Private Function FlagListContains(Value) As Boolean
    Dim i As Integer
    For i = 0 To uFlags
        If FlagList(i) = CDec(Value) Then
            FlagListContains = True
            Exit Function
        End If
    Next
End Function

' Utils
Public Function GetMaxNameLenMembers(ms As Members) As Integer
    Dim mi As MemberInfo
    Dim mx As Integer
    For Each mi In ms
        mx = Max(Len(mi.Name), mx)
    Next
    GetMaxNameLenMembers = mx
End Function
Public Function GetMaxNameLenParams(ps As Parameters) As Integer
    Dim pi As ParameterInfo
    Dim mx As Integer
    For Each pi In ps
        mx = Max(Len(pi.Name), mx)
    Next
    GetMaxNameLenParams = mx
End Function
Public Function GetMaxStringLen(sArr() As String) As Integer
    Dim i As Integer
    For i = 0 To UBound(sArr)
        GetMaxStringLen = Max(Len(sArr(i)), GetMaxStringLen)
    Next
End Function

Public Function Max(Val1, Val2)
    If Val1 > Val2 Then Max = Val1 Else Max = Val2
End Function

' ########## ' Interface ' ########## '
Public Function IsRealInterface(aInterface As InterfaceInfo, aClass As CoClassInfo) As Boolean
    If aClass Is Nothing Then Exit Function
    If aInterface Is Nothing Then Exit Function
    If Not aClass.DefaultEventInterface Is Nothing Then
        If (aInterface.Name = aClass.DefaultEventInterface.Name) Then Exit Function
    End If
    If Not aClass.DefaultInterface Is Nothing Then
        If (aInterface.Name = aClass.DefaultInterface.Name) Then Exit Function
    End If
    IsRealInterface = True
End Function


'blind out annoying members from IUnknown and IDispatch
''from IUnknown
'QueryInterface   ' Function
'AddRef           ' Function
'Release          ' Function

'from IDispatch
'GetTypeInfoCount ' Function
'GetTypeInfo      ' Function
'GetIDsOfNames    ' Function
'Invoke           ' Function
Private Sub FillIUnkIDispArr()
    Dim i As Integer
    IUnkIDisp(i) = "QueryInterface": i = i + 1
    IUnkIDisp(i) = "AddRef": i = i + 1
    IUnkIDisp(i) = "Release": i = i + 1
    IUnkIDisp(i) = "GetTypeInfoCount": i = i + 1
    IUnkIDisp(i) = "GetTypeInfo": i = i + 1
    IUnkIDisp(i) = "GetIDsOfNames": i = i + 1
    IUnkIDisp(i) = "Invoke": i = i + 1
    m_IUnkIDispFilled = True
End Sub
Public Function IsIUnkIDispProc(aProcName As String) As Integer
    'returns the index of the function
    'so it is possible here to determine if the procedure is from IUnknown or IDispatch
    Dim i As Integer
    If Not m_IUnkIDispFilled Then FillIUnkIDispArr
    For i = 0 To 6
        If StrComp(IUnkIDisp(i), aProcName, vbTextCompare) = 0 Then
            IsIUnkIDispProc = i + 1
            Exit Function
        End If
    Next
End Function
Public Function IsSub(vt As TliVarType) As Boolean
    IsSub = ((vt And VT_HRESULT) = VT_HRESULT) Or ((vt And VT_VOID) = VT_VOID)
End Function

' ########## ' Interface ' ########## '
Public Function CountRealInterfaces(aClass As CoClassInfo) As Integer
    If aClass Is Nothing Then Exit Function
    Dim ii As InterfaceInfo
    Dim c As Integer
    For Each ii In aClass.Interfaces
       If IsRealInterface(ii, aClass) Then
            c = c + 1
       End If
    '      If Not aClass.DefaultEventInterface Is Nothing Then
    '         If Not (ii.Name = aClass.DefaultEventInterface.Name) Then
    '            c = c + 1
    '         End If
    '      Else
    '         If Not aClass.DefaultInterface Is Nothing Then
    '            If Not (ii.Name = aClass.DefaultInterface.Name) Then
    '               c = c + 1
    '            End If
    '         Else
    '            c = c + 1
    '         End If
    '      End If
    Next
    CountRealInterfaces = c
    'Debug.Print C
End Function
Public Function ContainsMember(aInterface As InterfaceInfo, MemberName As String) As Boolean
    Dim mi As MemberInfo
    For Each mi In aInterface.Members
        If mi.Name = MemberName Then
            ContainsMember = True
            Exit Function
        End If
    Next
End Function

'PathFileName
Public Function GetDir(ByVal aPath As String) As String
    On Error Resume Next
    GetDir = aPath
    If IsDir(GetDir) Then Exit Function
    Dim pos As Long
    
    pos = InStrRev(GetDir, "\")
    If pos > 0 Then GetDir = Left$(GetDir, pos - 1)
    If IsDir(GetDir) Then Exit Function
    
    pos = InStrRev(GetDir, "\")
    If pos > 0 Then GetDir = Left$(GetDir, pos - 1)
    If IsDir(GetDir) Then Exit Function
    GetDir = ""
End Function
Public Function IsDir(p As String) As Boolean
    On Error Resume Next
    IsDir = (GetAttr(p) = vbDirectory)
    On Error GoTo 0
End Function

'Collection
Public Function ContainsKey(aCol As Collection, aKey As String) As Boolean
    On Error Resume Next
    If IsEmpty(aCol.item(aKey)) Then: 'DoNothing
    ContainsKey = (Err.Number = 0)
    On Error GoTo 0
End Function

'Strings, Arrays, Collections
Public Sub SortCol(aCol As Collection)
    Dim sArr() As String
    Call CopyColToArray(sArr, aCol)
    Call Sort(sArr)
    Call CopyArrayToCol(aCol, sArr)
End Sub
Public Sub CopyColToArray(aDstArrOut() As String, aSrcColIn As Collection)
    ReDim aDstArrOut(0 To aSrcColIn.Count - 1)
    Dim i As Integer
    Dim v
    For Each v In aSrcColIn
        aDstArrOut(i) = CStr(v): i = i + 1
    Next
End Sub
Public Sub CopyArrayToCol(aDstColOut As Collection, aSrcArrIn() As String)
    Set aDstColOut = New Collection
    Dim i As Integer
    For i = LBound(aSrcArrIn) To UBound(aSrcArrIn)
        Call aDstColOut.Add(aSrcArrIn(i))
    Next
End Sub

Public Sub Sort(sArr() As String)
    Call QuickSort(sArr, LBound(sArr), UBound(sArr))
End Sub
Private Function Compare(sArr() As String, ByVal i1 As Long, ByVal i2 As Long) As Long
    Compare = StrComp(sArr(i1), sArr(i2), vbTextCompare) 'vbBinaryCompare)
End Function
Private Sub Swap(sArr() As String, ByVal i1 As Long, ByVal i2 As Long)
    Dim aTemp As String: aTemp = sArr(i1)
    sArr(i1) = sArr(i2): sArr(i2) = aTemp
End Sub
' QuickSort
' Dieser QuickSort-Algorithmus ist unabhängig von den zu sortierenden
' Daten, da der Vergleich von Daten (Compare) und ein Vertauschen der
' Daten (Swap) aus dem Algorithmus in andere Routinen ausgelagert ist.
Private Sub QuickSort(sArr() As String, ByVal i1 As Long, ByVal i2 As Long)
    Dim T As Long
    If i2 > i1 Then
        T = divide(sArr, i1, i2)
        Call QuickSort(sArr, i1, T - 1)
        Call QuickSort(sArr, T + 1, i2)
    End If
End Sub
Private Function divide(sArr() As String, ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (Compare(sArr, i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (Compare(sArr, p, j) < 0))
        If i < j Then Call Swap(sArr, i, j)
    Loop While (i < j)
    Call Swap(sArr, i, p)
    divide = i
End Function

