VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProc 
   Caption         =   "Stored Code"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   Icon            =   "frmProc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   8790
   Begin VB.CommandButton cmdRem 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtbTemp 
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      _Version        =   393217
      TextRTF         =   $"frmProc.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbMain 
      Height          =   3375
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5953
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmProc.frx":008C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lstProc 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2295
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Is"
         Object.Width           =   1905
      EndProperty
   End
End
Attribute VB_Name = "frmProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
frmAddStored.Show vbModal
End Sub

Private Sub cmdRem_Click()
On Error GoTo 1

Dim s As String, i As Integer
 s = lstProc.SelectedItem.Text
 
 For i = 0 To gblStoredCnt
  If gblStored(0, i) = s Then
   gblStored(0, i) = ""
   gblStored(1, i) = ""
   Exit For
  End If
 Next i
 
 lstProc.ListItems.Remove lstProc.SelectedItem.Index
1
End Sub

Private Sub Form_Load()
Dim i As Integer, l As ListItem
For i = 0 To gblStoredCnt
 If gblStored(1, i) <> "" Then
  Set l = lstProc.ListItems.Add(, , gblStored(0, i))
  If LCase(Left(Trim(gblStored(1, i)), 5)) = "!type" Then
   l.SubItems(1) = "Type"
  Else
   l.SubItems(1) = "Procedure"
  End If
 End If
Next i
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
rtbMain.Width = Me.Width - rtbMain.Left - 100
rtbMain.Height = Me.Height - cmdAdd.Height

lstProc.Height = Me.Height - (cmdAdd.Height * 2) - 50
cmdAdd.Top = lstProc.Height
cmdRem.Top = lstProc.Height

End Sub

Private Sub lstProc_Click()
'1300
'1080
Dim l As Integer

For l = 0 To gblStoredCnt
 If gblStored(0, l) = lstProc.SelectedItem.Text Then rtbMain.Text = "": Call PrintText(gblStored(1, l), rtbMain, rtbTemp): Exit For
Next l
End Sub

