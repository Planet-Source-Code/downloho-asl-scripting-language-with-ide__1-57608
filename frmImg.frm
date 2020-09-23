VERSION 5.00
Begin VB.Form frmImg 
   BackColor       =   &H80000005&
   Caption         =   "Image"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmImg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   WindowState     =   2  'Maximized
   Begin VB.Image imgMain 
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
If gblCanClose = False Then Cancel = -1: Me.Hide
End Sub
