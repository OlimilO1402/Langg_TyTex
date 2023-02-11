Attribute VB_Name = "MInd"
Option Explicit
Public IsTab        As Boolean
Public IndentSize   As Integer
Private myIndStack  As Integer
Private myIndent    As String
'hmm what if someone likes tabs more than spaces?

Public Function Push() As String
    myIndStack = myIndStack + 1
    Push = Create
End Function
Public Function Pop() As String
    myIndStack = myIndStack - 1
    Pop = Create
End Function
Private Function Create() As String
    'Debug.Print CStr(myIndStack) & " " & CStr(MyIndentSize)
    'IIf nop
    If (myIndStack > 0) Then
        If IsTab Then
            myIndent = String$(myIndStack, Chr$(9))
        Else
           If (IndentSize > 0) Then
                myIndent = Space$(myIndStack * IndentSize)
           End If
        End If
    Else
        myIndStack = 0
        myIndent = vbNullString
    End If
    Create = myIndent
End Function
Public Function Peek() As String
    Peek = myIndent
End Function

Public Sub Clear()
    myIndStack = 0
End Sub


