VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoderVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2009_09_01 OM Zeilen: 515
Implements ICoder

Private Sub Class_Initialize()
   Call MHelpCode.NewHelpFlavour("' ", "' ", "'")
End Sub

' ########## '   Enum    ' ########## '
Private Function ICoder_EnumToCode(aEnum As TLI.ConstantInfo) As String
   If aEnum Is Nothing Then Exit Function
   Dim s  As String
   Dim mi As MemberInfo
   
   Dim maxlen As Integer: maxlen = GetMaxNameLenMembers(aEnum.Members)
   Dim bIsFlagEnum As Boolean: bIsFlagEnum = IsFlags(aEnum)
   s = s & "'Namespace " & aEnum.Parent.Name & vbCrLf
   s = s & MHelpCode.AttributesToCode(aEnum)
   s = s & MHelpCode.HelpMultiLine(aEnum.HelpString(MyLCID))
   s = s & MInd.Peek & "Public Enum " & aEnum.Name & vbCrLf
   MInd.Push
   For Each mi In aEnum.Members
      s = s & MHelpCode.MemberHelpToCode(mi)
      s = s & MInd.Peek & PadRight(mi.Name, maxlen) & EnumConstToCode(mi, bIsFlagEnum) & vbCrLf
   Next
   s = s & MInd.Pop & "End Enum" & vbCrLf
   ICoder_EnumToCode = s
End Function
Private Function EnumConstToCode(mi As MemberInfo, ByVal bIsFlagEnum As Boolean) As String
   If mi Is Nothing Then Exit Function
   Dim s As String, v As Integer
   Dim b() As Byte, i As Integer
   Dim hx As String: hx = "&H"
   If Not IsMissing(mi.Value) Then
      If IsNumeric(mi.Value) Then
         s = s & " = " & IIf((bIsFlagEnum Or (mi.Value < 0)), hx & Hex$(mi.Value), CStr(mi.Value))
      Else
         If Len(mi.Value) > 0 Then
            s = s & " = "
            b = mi.Value
            s = s & "AscW(ChrW(" & hx
            For i = UBound(b) To 0 Step -1
               s = s & Hex(b(i))
            Next
            s = s & "))"
         End If
      End If
   End If
   EnumConstToCode = s
End Function


' ########## '   Type    ' ########## '
Private Function ICoder_TypeToCode(aType As TLI.RecordInfo) As String
   If aType Is Nothing Then Exit Function
   Dim s As String
   Dim mi As MemberInfo
   Dim sh As String
   Dim maxlen As Integer: maxlen = GetMaxNameLenMembers(aType.Members)
   s = s & MHelpCode.AttributesToCode(aType)
   s = s & MInd.Peek & "Public Type " & aType.Name & vbCrLf
   MInd.Push
   For Each mi In aType.Members
      s = s & MHelpCode.MemberHelpToCode(mi)
      s = s & MInd.Peek & PadRight(mi.Name, maxlen) & ReturnTypeToCode(mi.ReturnType) & vbCrLf
   Next
   s = s & MInd.Pop & "End Type" & vbCrLf
   ICoder_TypeToCode = s
End Function
Private Function ReturnTypeToCode(vti As VarTypeInfo, Optional ByVal bIsFunction As Boolean) As String
   If vti Is Nothing Then Exit Function
   Dim s As String
   Dim vt As TliVarType: vt = vti.VarType
   'Debug.Print VarTypeToString(vt)
   If vt And VT_ARRAY Then
      If Not bIsFunction Then
         s = s & "()" 'no boundaries here
         vt = vt Xor VT_ARRAY
      End If
   End If
   If vt And VT_VECTOR Then
      s = s & VectorBoundsToCode(vti)
      vt = vt Xor VT_VECTOR
   End If
   If Not IsSub(vt) Then
      s = s & " As "
      Dim ti As TypeInfo: Set ti = vti.TypeInfo
      If Not ti Is Nothing Then
         s = s & ti.Name
      Else
         s = s & VarTypeToCode(vt)
      End If
      If bIsFunction Then
         If vt And VT_ARRAY Then
            s = s & "()" 'no boundaries here
         End If
      End If
   End If
   ReturnTypeToCode = s
End Function
Private Function VectorBoundsToCode(vti As VarTypeInfo) As String
   If vti Is Nothing Then Exit Function
   Dim s As String
   s = s & "("
   Dim nBounds As Long
   Dim Bounds() As Long
   Dim i As Long, lb As Long, ub As Long
   nBounds = GetArrayBounds(vti, Bounds) 'vti.ArrayBounds(Bounds())
   For i = 1 To nBounds
      lb = Bounds(i, 1)
      ub = Bounds(i, 2)
      s = s & CStr(lb) & " To " & CStr(ub)
      If i < nBounds Then s = s & ", "
   Next
   s = s & ")"
   VectorBoundsToCode = s
End Function
Private Function VarTypeToCode(ByVal vt As TliVarType) As String
    VarTypeToCode = MEVbVarType.EVbVarType_ToStr(vt)
'   Dim s As String
'   Select Case vt
'   Case VT_EMPTY:       s = "" ' Variant" '"Empty" 'no not Variant!
'   Case VT_NULL:        s = "Null"
'   Case VT_I2:          s = "Integer"
'   Case VT_I4:          s = "Long"
'   Case VT_R4:          s = "Single"
'   Case VT_R8:          s = "Double"
'   Case VT_CY:          s = "Currency" '???
'   Case VT_DATE:        s = "Date"
'   Case VT_BSTR:        s = "String"
'   Case VT_DISPATCH:    s = "IDispatch"
'   Case VT_ERROR:       s = "Error"
'   Case VT_BOOL:        s = "Boolean"
'   Case VT_VARIANT:     s = "Variant"
'   Case VT_UNKNOWN:     s = "IUnknown" '???
'   Case VT_DECIMAL:     s = "Decimal"
'   Case VT_I1:          s = "Byte"
'   Case VT_UI1:         s = "Byte" '"Unsigned Byte"
'   Case VT_UI2:         s = "Integer" '"Unsigned Integer"
'   Case VT_UI4:         s = "Long" '"Unsigned Long"
'   Case VT_I8:          s = "Currency" '"LongLong"
'   Case VT_UI8:         s = "Currency" '"UnsignedLongLong"
'   Case VT_INT:         s = "Long"
'   Case VT_UINT:        s = "Long" '"Unsigned Long"
'   Case VT_VOID:        s = "Long" '"Void"
'   Case VT_HRESULT:     s = "Long" '"HResult"
'   Case VT_PTR:         s = "Long" '"*"
'   Case VT_SAFEARRAY:   s = "SafeArray"
'   Case VT_CARRAY:      s = "()" '"[]"
'   Case VT_USERDEFINED: s = "UDType"
'   Case VT_LPSTR:       s = "String" '"LPSTR"
'   Case VT_LPWSTR:      s = "String" '"LPWSTR"
'   Case VT_RECORD:      s = "UDType"
'   Case VT_FILETIME:    s = "FileTime"
'   Case VT_BLOB:        s = "Blob"
'   Case VT_STREAM:      s = "Stream"
'   Case VT_STORAGE:     s = "Storage"
'   Case VT_STREAMED_OBJECT: s = "Streamed Object"
'   Case VT_STORED_OBJECT:   s = "Stored Object"
'   Case VT_BLOB_OBJECT:     s = "Blob Object"
'   Case VT_CF:          s = "CF" '???
'   Case VT_CLSID:       s = "CLSID"
'   Case VT_VECTOR:      s = "Vector" 'fixed array!
'   Case VT_BYREF:       s = "ByRef"
'   Case VT_RESERVED:    s = "Reserved"
'   Case Else:           s = "&H" & Hex$(vt)
'   End Select
'
''   If vt And VT_EMPTY Then s = s & "Empty"
''   If vt And VT_NULL Then s = s & "Null"
''   If vt And VT_I2 Then s = s & "Integer"
''   If vt And VT_I4 Then s = s & "Long"
''   If vt And VT_R4 Then s = s & "Single"
''   If vt And VT_R8 Then s = s & "Double"
''   If vt And VT_CY Then s = s & "Currency"
''   If vt And VT_DATE Then s = s & "Date"
''   If vt And VT_BSTR Then s = s & "String"
''   If vt And VT_DISPATCH Then s = s & "IDispatch"
''   If vt And VT_ERROR Then s = s & "Error"
''   If vt And VT_BOOL Then s = s & "Boolean"
''   If vt And VT_VARIANT Then s = s & "Variant"
''   If vt And VT_UNKNOWN Then s = s & "IUnknown"
''   If vt And VT_DECIMAL Then s = s & "Decimal"
''   If vt And VT_I1 Then s = s & "Byte"
''   If vt And VT_UI1 Then s = s & "Unsigned Byte"
''   If vt And VT_UI2 Then s = s & "Unsigned Integer"
''   If vt And VT_UI4 Then s = s & "Unsigned Long"
''   If vt And VT_I8 Then s = s & "LongLong"
''   If vt And VT_UI8 Then s = s & "Unsigned LongLong"
''   If vt And VT_INT Then s = s & "Long"
''   If vt And VT_UINT Then s = s & "Unsigned Long"
''   If vt And VT_VOID Then s = s & "Void"
''   If vt And VT_HRESULT Then s = s & "HResult"
''   If vt And VT_PTR Then s = s & "*"
''   If vt And VT_SAFEARRAY Then s = s & "SafeArray"
''   If vt And VT_CARRAY Then s = s & "[]"
''   If vt And VT_USERDEFINED Then s = s & "UDType"
''   If vt And VT_LPSTR Then s = s & "String"
''   If vt And VT_LPWSTR Then s = s & "String"
''   If vt And VT_RECORD Then s = s & "UDType"
''   If vt And VT_FILETIME Then s = s & "FileTime"
''   If vt And VT_BLOB Then s = s & "Blob"
''   If vt And VT_STREAM Then s = s & "Stream"
''   If vt And VT_STORAGE Then s = s & "Storage"
''   If vt And VT_STREAMED_OBJECT Then s = s & "Streamed Object"
''   If vt And VT_STORED_OBJECT Then s = s & "Stored Object"
''   If vt And VT_BLOB_OBJECT Then s = s & "Blob Object"
''   If vt And VT_CF Then s = s & "CF"
''   If vt And VT_CLSID Then s = s & "CLSID"
''   If vt And VT_VECTOR Then s = s & "Vector"
''   If vt And VT_BYREF Then s = s & "ByRef"
''   If vt And VT_RESERVED Then s = s & "Reserved"
'
'   VarTypeToCode = s
End Function




' ########## ' Procedure ' ########## '
Private Function ProcToCode(mi As MemberInfo, lastProcName As String, Optional bIsEvent As Boolean) As String
   On Error Resume Next
   Dim s As String
   Dim sp1 As String, sp2 As String, params As String
   Dim hs As String
   If (lastProcName <> mi.Name) Then
      s = s & vbCrLf 'newline at the beginning
      s = s & MHelpCode.MethodHelpToCode(mi)
   End If
   s = s & ModifierToCode(mi)
   If bIsEvent Then
      sp1 = "Event "
   Else
      Select Case mi.InvokeKind
      Case INVOKE_CONST:          sp1 = "Const "
      Case INVOKE_EVENTFUNC:      sp1 = "Event "
      Case INVOKE_FUNC:
         'wie erkennt man Sub?
         If Not mi.ReturnType Is Nothing Then
            sp1 = IIf(IsSub(mi.ReturnType.VarType), "Sub ", "Function ")
         Else
            sp1 = "Sub "
         End If
      Case INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF
         sp1 = "Property "
         Select Case mi.InvokeKind
         Case INVOKE_PROPERTYGET:    sp2 = "Get "
         Case INVOKE_PROPERTYPUT:    sp2 = "Let "
         Case INVOKE_PROPERTYPUTREF: sp2 = "Set "
         End Select
      Case INVOKE_UNKNOWN:        sp1 = "Sub "
      End Select
   End If
   s = s & sp1 & sp2 & mi.Name & "("
   params = ParamsToCode(mi)
   If (mi.InvokeKind = INVOKE_PROPERTYPUT) Or (mi.InvokeKind = INVOKE_PROPERTYPUTREF) Then
      s = s & "RHS" & ReturnTypeToCode(mi.ReturnType)
      If Len(params) > 0 Then s = s & ", "
   End If
   s = s & params
   s = s & ")"
   Select Case mi.InvokeKind
   Case INVOKE_FUNC, INVOKE_PROPERTYGET
      's = s & " As "
      If Not mi.ReturnType Is Nothing Then
         s = s & ReturnTypeToCode(mi.ReturnType, True)
      End If
   End Select
   If Not bIsEvent Then
      s = s & vbCrLf
      s = s & MInd.Peek & "End " & Trim(sp1)
   End If
   ProcToCode = s
   On Error GoTo 0
End Function
Private Function ModifierToCode(mi As MemberInfo) As String
   If mi Is Nothing Then Exit Function
   Dim s As String
   Dim a() As String
   Dim n As Integer
   n = mi.AttributeStrings(a)
   If mi.AttributeMask <> 0 Then
      'Debug.Print Join(a, " ") & " " & mi.Name
   End If
   If mi.AttributeMask And FuncFlags.FUNCFLAG_FHIDDEN Then
      s = s & "Private "
      's = s & Join(a, " ") & " "
   Else
      s = s & IIf(Left(mi.Name, 1) = "_", "Private ", "Public ")
   End If
   ModifierToCode = s
End Function
Private Function ParamsToCode(mi As MemberInfo) As String
   Dim s As String
   Dim svt As String
   Dim pi As ParameterInfo
   Dim i As Integer, cnt As Integer
   cnt = mi.Parameters.Count
   For i = 1 To cnt
      Set pi = mi.Parameters(i)
      s = s & ParamFlagsToCode(pi.flags) & pi.Name & ReturnTypeToCode(pi.VarTypeInfo)
      If pi.flags And PARAMFLAG_FOPT Then
         If pi.flags And PARAMFLAG_FHASDEFAULT Then
            s = s & " = " & pi.DefaultValue
         End If
      End If
      If i < cnt Then
         s = s & ", "
      End If
   Next
   ParamsToCode = s
End Function
Private Function ParamFlagsToCode(aFlag As ParamFlags) As String
   Dim s As String
   If aFlag And PARAMFLAG_FHASCUSTDATA Then s = s & ""
   If aFlag And PARAMFLAG_FOPT Then s = s & "Optional "
   If aFlag And PARAMFLAG_FLCID Then s = s & "LCID "
   If aFlag And PARAMFLAG_FOUT Then
      s = s & "ByRef "
   ElseIf aFlag And PARAMFLAG_FIN Then
      s = s & "ByVal "
   End If
   If aFlag And PARAMFLAG_FRETVAL Then s = s & "ByRef "
   'If aFlag And PARAMFLAG_FHASDEFAULT Then s = s & " = "
   If aFlag And PARAMFLAG_NONE Then s = s & ""
   ParamFlagsToCode = s
End Function



' ########## '  Module   ' ########## '
Private Function ICoder_ModuleToCode(aModule As TLI.DeclarationInfo) As String
   If aModule Is Nothing Then Exit Function
   Dim mi As MemberInfo
   Dim s As String
   Dim mh As String
   Dim lastminame As String
   s = s & MHelpCode.AttributesToCode(aModule)
   s = s & MHelpCode.HelpSingleLine(aModule.HelpString(MyLCID))
   's = s & "Public Modul " & aModul.Name & vbCrLf
   For Each mi In aModule.Members
      s = s & ProcToCode(mi, lastminame) & vbCrLf
      lastminame = mi.Name
   Next
   's = s & vbCrLf & "End Modul" & " ' " & aModul.Name
   ICoder_ModuleToCode = s
End Function


' ########## ' Interface ' ########## '
Private Function ICoder_InterfaceToCode(aInterface As TLI.InterfaceInfo) As String
   ICoder_InterfaceToCode = InterfaceToCode(aInterface)
End Function
Private Function InterfaceToCode(aInterface As InterfaceInfo, Optional ByVal bIsDefClass As Boolean, _
                                                              Optional ByVal bIsEvent As Boolean) As String
   If aInterface Is Nothing Then Exit Function
   Dim s As String
   Dim mi As MemberInfo
   Dim lastminame As String
   
   s = s & MHelpCode.AttributesToCode(aInterface)
   If Not bIsDefClass Then
      s = s & "Public Interface " & aInterface.Name & vbCrLf
   End If
   Dim ii As InterfaceInfo
   For Each ii In aInterface.ImpliedInterfaces
   'ainterface.
      If (StrComp(ii.Name, "IDispatch", vbTextCompare) = False) And (StrComp(ii.Name, "IUnknown", vbTextCompare) = False) Then
      '   'blind out IDispatch- and IUnknown-methods
      '   s = s & InterfaceToCode(ii)
          s = s & "Implements " & ii.Name & vbCrLf
      End If
   Next
   For Each mi In aInterface.Members
      'Debug.Print mi.Name
      If (IsIUnkIDispProc(mi.Name) = False) Or _
           ((aInterface.Name = "IDispatch") Or _
            (aInterface.Name = "IUnknown")) Then
         s = s & ProcToCode(mi, lastminame, bIsEvent) & vbCrLf
      End If
      lastminame = mi.Name
   Next
   
   If Not bIsDefClass Then
      s = s & vbCrLf & "End Interface" & " ' " & aInterface.Name & vbCrLf
   End If
   InterfaceToCode = s
End Function


' ########## '   Class   ' ########## '
'A class exists of
'  private friend or protected memberfields (variables) not possible to see
'  public memberfields (in tlb as property get/let/set)
'  one public defaultinterface
'  optional one eventinterface
'  optional one or more additional implemented interfaces
Private Function ICoder_ClassToCode(aClass As TLI.CoClassInfo) As String
   If aClass Is Nothing Then Exit Function
   Dim s As String
   Dim mi As MemberInfo
   
   s = s & MHelpCode.AttributesToCode(aClass) ' & vbCrLf
   s = s & "Public Class " & aClass.Name & vbCrLf
   s = s & "' " & aClass.Guid & vbCrLf
   s = s & ImplementsToCode(aClass) & vbCrLf
   If Not aClass.DefaultEventInterface Is Nothing Then
      If Not (aClass.DefaultEventInterface.Name = "IDispatch") Then
         s = s & InterfaceToCode(aClass.DefaultEventInterface, True, True)
      End If
   End If
   s = s & InterfaceToCode(aClass.DefaultInterface, True)
   s = s & vbCrLf & "End Class" & " ' " & aClass.Name
   ICoder_ClassToCode = s
End Function
Private Function ImplementsToCode(aClass As TLI.CoClassInfo) As String
   If aClass Is Nothing Then Exit Function
   Dim s As String
   Dim c As Integer: c = CountRealInterfaces(aClass)
   If c = 0 Then Exit Function
   MInd.Push
   Dim ii As InterfaceInfo
   For Each ii In aClass.Interfaces
      If IsRealInterface(ii, aClass) Then
         If c > 0 Then
            s = vbCrLf & s
         End If
         s = s & MInd.Peek & "Implements " & ii.Name & vbCrLf
      End If
   Next
   MInd.Pop
   ImplementsToCode = s
End Function

'Private Function VBVarType_ToStr(vt As VbVarType) As String
'    Dim s As String
'    Select Case vt
'    Case vbArray:           s = "Array"
'    Case vbBoolean:         s = "Boolean"
'    Case vbByte:            s = "Byte"
'    Case vbCurrency:        s = "Currency"
'    Case vbDataObject:      s = "DataObject"
'    Case vbDate:            s = "Date"
'    Case vbDecimal:         s = "Decimal"
'    Case vbDouble:          s = "Double"
'    Case vbEmpty:           s = "Empty"
'    Case vbError:           s = "Error"
'    Case vbInteger:         s = "Integer"
'    Case vbLong:            s = "Long"
'    Case vbNull:            s = "Null"
'    Case vbObject:          s = "Object"
'    Case vbSingle:          s = "Single"
'    Case vbString:          s = "String"
'    Case vbUserDefinedType: s = "UserDefinedType"
'    Case vbVariant:         s = "Variant"
'    Case Else:              s = CStr(vt)
'    End Select
'    VBVarType_ToStr = s
'End Function
' ########## '   Alias   ' ########## '
Private Function ICoder_AliasToCode(aAlias As TLI.IntrinsicAliasInfo) As String
    Dim s As String
    Dim mi As MemberInfo
    s = s & MHelpCode.AttributesToCode(aAlias)
    'das ja Blödsinn, ein Alias hat doch keine Member
'    If TypeOf aAlias Is MemberInfo Then
'        s = s & MHelpCode.MemberHelpToCode(aAlias, True)
'        MInd.Push
'        With aAlias
'           s = s & MInd.Peek & "Public Alias " & .Name & ReturnTypeToCode(.ResolvedType)
'        End With
'        MInd.Pop
'    End If

    s = s & "Alias " & aAlias.Name & " As " & VarTypeToCode(aAlias.ResolvedType) '& " " & aAlias.TypeKindString
    ICoder_AliasToCode = s
End Function

' ########## '   Union   ' ########## '
Private Function ICoder_UnionToCode(aUnion As TLI.UnionInfo) As String
   Dim s As String
   Dim mi As MemberInfo
   With aUnion
      s = s & MInd.Push & "Public Union " & .Name '& " As " & .TypeKindString
      MInd.Push
      For Each mi In .Members
         s = s & MInd.Peek & mi.Name & " As " & .TypeKindString
      Next
      s = s & MInd.Pop & "End Union"
   End With
   ICoder_UnionToCode = s
End Function
