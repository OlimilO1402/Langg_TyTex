VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type Stat
   Name  As String
   Count As Integer
   Index As Integer
End Type
Private mbSort As Boolean
Private myClasses  As Collection
Private myProcs(0) As Stat
Private myClass(0) As Stat
Private myBolMat(0, 0) As Boolean

Public Sub NewC(Classes As Collection, bSort As Boolean)
   Set myClasses = Classes
   mbSort = bSort
   'under construction
End Sub

Public Function getEventsStatistic() As String
   'under construction

End Function

Public Function getMethodsStatistic() As String
   'under construction

End Function
