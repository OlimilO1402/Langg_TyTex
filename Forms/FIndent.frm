VERSION 5.00
Begin VB.Form FIndent 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Indentation"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2895
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbIndSize 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   600
      Width           =   735
   End
   Begin VB.OptionButton OptSpace 
      Caption         =   "Space"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton OptTabs 
      Caption         =   "Tabs"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Size:"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "FIndent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim i
   For i = 0 To 10
      CmbIndSize.AddItem i
   Next
   OptTabs.value = MInd.IsTab: OptSpace.value = Not MInd.IsTab
   CmbIndSize.ListIndex = MInd.IndentSize
End Sub

Private Sub BtnOK_Click()
   MInd.IsTab = OptTabs.value
   MInd.IndentSize = CmbIndSize.ListIndex
   Unload Me
End Sub
Private Sub BtnCancel_Click()
   Unload Me
End Sub


