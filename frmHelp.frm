VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   6885
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
   Begin VB.FileListBox flbHelp 
      Height          =   5355
      Left            =   0
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub flbHelp_Click()
Dim s As String

Open App.Path & "\res\" & flbHelp.List(flbHelp.ListIndex) For Input As #1
 s = Input(LOF(1), #1)
Close #1

txtHelp.Text = s
End Sub

Private Sub Form_Load()
flbHelp.Path = App.Path & "\res"
flbHelp.Refresh
End Sub

Private Sub Form_Resize()
flbHelp.Height = Me.Height - 300
txtHelp.Height = Me.Height - 400
txtHelp.Width = Me.Width - flbHelp.Width - 100
End Sub
