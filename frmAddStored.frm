VERSION 5.00
Begin VB.Form frmAddStored 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Stored Procedure"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtCode 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmAddStored"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Call addStored(txtCode.Text)
Call Unload(Me)
End Sub
