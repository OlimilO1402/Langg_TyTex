VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2009_09_01 OM Zeilen: 23

Public Function EnumToCode(aEnum As ConstantInfo) As String
End Function

Public Function TypeToCode(aType As RecordInfo) As String
End Function

Public Function ModuleToCode(aModule As DeclarationInfo) As String
End Function

Public Function InterfaceToCode(aInterface As InterfaceInfo) As String
End Function

Public Function ClassToCode(aClass As CoClassInfo) As String
End Function

Public Function AliasToCode(aAlias As IntrinsicAliasInfo) As String
End Function

Public Function UnionToCode(aUnion As UnionInfo) As String
End Function
