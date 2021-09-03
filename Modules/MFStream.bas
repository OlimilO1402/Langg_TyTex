Attribute VB_Name = "MFStream"
Option Explicit
Dim myFNr As Integer
Dim myFNm As String

Public Sub OOpen(FNam As String)
   myFNr = FreeFile
   Open FNam For Append As myFNr
End Sub

Public Sub Append(aStrVal As String)
   Print #myFNr, aStrVal
End Sub

Public Sub CClose()
   Close myFNr
End Sub

Public Sub SaveString(FNm As String, aStrVal As String)
   Dim FNr As Integer: FNr = FreeFile
   Call MakeFileName(FNm)
   Open FNm For Output As FNr
   Print #FNr, aStrVal
   Close FNr
End Sub
Private Sub MakeFileName(FNm As String)
   Dim i As Integer
   For i = 0 To 32
      FNm = Replace(FNm, Chr(i), "_")
   Next
End Sub


