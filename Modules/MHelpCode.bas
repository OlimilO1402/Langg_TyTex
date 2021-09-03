Attribute VB_Name = "MHelpCode"
Option Explicit
' ########## '   Help    ' ########## '
Public MyLCID       As Long '= 0
'Private Const LCID_en_US As Long = &H409 ' = 1033  'en-US 'noop
'Private Const LCID_en_EN As Long = &H809 ' = 2057  'en-EN 'noop
'Private Const LCID_de_DE As Long = &H407 ' = 1031  'de-DE 'noop
Public MyAuthorName As String
' the following will be selected by the Coder itself
' please make a decision
' which flavour do you like?  ' VB  or Java-Style
Private myHlpBeg As String    ' "'" or "/**"
Private myHlpInt As String    ' "'" or " * "
Private myHlpEnd As String    ' "'" or " */"

'Also uses Indentation Ind

'Create the help-Begin, intermediate and end-strings
Public Sub NewHelpFlavour(hlpBegin As String, hlpIntermed As String, hlpEnd As String)
   myHlpBeg = hlpBegin
   myHlpInt = hlpIntermed
   myHlpEnd = hlpEnd
End Sub


' ########## '   Help    ' ########## '
Public Function AttributesToCode(ti As TypeInfo) As String
   Dim s As String, a() As String
   Dim n As Integer: n = ti.AttributeStrings(a)
   If n > 0 Then s = s & HelpSingleLine(Join(a, " ")) & vbCrLf
   AttributesToCode = s
End Function

Public Function MemberHelpToCode(mi As MemberInfo, Optional ByVal bIsMultiLine As Boolean = False) As String
   If bIsMultiLine Then
      MemberHelpToCode = HelpMultiLine(mi.HelpString(MyLCID))
   Else
      MemberHelpToCode = HelpSingleLine(mi.HelpString(MyLCID))
   End If
End Function
Public Function MethodHelpToCode(mi As MemberInfo) As String
   Dim s As String
   Dim hs As String: hs = mi.HelpString(MyLCID)
   Dim ph() As String: Call ParamsGetHelpStrings(mi.Parameters, ph)
   Dim psh As String
   If Len(hs) > 0 Then
      If Len(myHlpBeg) > 0 Then
         s = s & MInd.Peek & myHlpBeg & vbCrLf
      End If
      s = s & MInd.Peek & myHlpInt & hs & vbCrLf
      's = s & MInd.Peek & myHlpInt & "@author " & MyAuthorName & vbCrLf
      's = s & MInd.Peek & myHlpInt & "@version " & "1.0" & vbCrLf
      
      's = s & MInd.Peek & myHlpInt & "@since  " & HelpDate & vbCrLf
      
      psh = Join(ph, vbCrLf & MInd.Peek & myHlpInt & "@param ")
      If Len(psh) > 0 Then
         s = s & MInd.Peek & myHlpInt & "@param " & psh & vbCrLf
      End If
      
      If Len(myHlpEnd) > 0 Then
         s = s & MInd.Peek & myHlpEnd & vbCrLf
      End If
   End If
   MethodHelpToCode = s
End Function
Public Function HelpSingleLine(hs As String) As String
   Dim s As String
   If Len(hs) > 0 Then
      s = s & MInd.Peek & myHlpBeg & hs & IIf(Left(myHlpEnd, 1) <> " ", " ", "") & myHlpEnd & vbCrLf
   End If
   HelpSingleLine = s
End Function
Public Function HelpMultiLine(hs As String) As String
   Dim s As String
   If Len(hs) > 0 Then
      If Len(myHlpBeg) > 0 Then
         s = s & MInd.Peek & myHlpBeg & vbCrLf
      End If
      s = s & MInd.Peek & myHlpInt & hs & vbCrLf
      s = s & MInd.Peek & myHlpInt & "@author " & MyAuthorName & vbCrLf
      s = s & MInd.Peek & myHlpInt & "@since  " & HelpDate & vbCrLf
      If Len(myHlpEnd) > 0 Then
         s = s & MInd.Peek & myHlpEnd & vbCrLf
      End If
   End If
   HelpMultiLine = s
End Function
Private Function HelpDate() As String
    Dim d As Date: d = Now
    HelpDate = CStr(Year(d)) & "-" & PadLeft(Month(d), 2, "0") & "-" & PadLeft(Day(d), 2, "0")
End Function
Private Sub ParamsGetHelpStrings(ps As Parameters, outPH() As String)
   Dim s As String
   Dim pi As ParameterInfo
   Dim hasHelp As Boolean
   Dim hs As String
   Dim maxNameLen As Integer
   maxNameLen = GetMaxNameLenParams(ps)
   If ps.Count > 0 Then
      ReDim outPH(0 To ps.Count - 1)
      Dim i As Integer
      For i = 1 To ps.Count '- 1
         Set pi = ps.item(i)
         hs = ParamGetHelpString(pi)
         If Len(hs) > 0 Then hasHelp = True
         
         outPH(i - 1) = PadRight(pi.Name, maxNameLen) & " : " & hs
      Next
   End If
'   For Each pi In ps
'      hs = ParamGetHelpString(pi)
'      If Len(hs) > 0 Then hasHelp = True
'      s = s & MInd.Peek & " * @" & PadRight(pi.Name, maxNameLen) & " : " & hs & vbCrLf
'   Next
'   If hasHelp Then
'      ParamsGetHelpString = s
'   End If
End Sub
Private Function ParamGetHelpString(pi As ParameterInfo) As String
   Dim s As String
   If Not pi.VarTypeInfo Is Nothing Then
      If Not pi.VarTypeInfo.TypeInfo Is Nothing Then
         s = s & pi.VarTypeInfo.TypeInfo.HelpString(MyLCID)
      End If
   End If
   ParamGetHelpString = s
End Function



