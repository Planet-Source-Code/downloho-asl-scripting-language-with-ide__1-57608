VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu Editor"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbWin 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdRem 
      Caption         =   "Rem"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdPos 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdPos 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdPos 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdPos 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.ListBox lstMenu 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Window:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Text:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Name:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   915
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuSub 
         Caption         =   "Sub Menu"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public b_Sender As String

Private Function FindEdit(ByVal str As String) As Form
Dim frm As Form
 For Each frm In Forms()
  If frm.Name = "frmEdit" And InStr(frm.Tag, Chr(0)) <> 0 Then
   If Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = str Then Exit For
  End If
 Next frm
 Set FindEdit = frm
End Function

Private Sub cmdAdd_Click()
lstMenu.AddItem txtName.Text & " - " & txtText.Text & " ^" & cmbWin.Text
End Sub

Public Sub ShowMenu(ByVal Win As String)
For i = 0 To mnuSub.Count - 1
 If InStr(mnuSub(i).Tag, Chr(1)) Then
  If Mid(mnuSub(i).Tag, InStr(mnuSub(i).Tag, Chr(1)) + 1) = Win Then mnuSub(i).Visible = True Else mnuSub(i).Visible = False
 End If
Next i

Call PopupMenu(mnuMain)
End Sub

Public Sub cmdOk_Click()
Dim i As Integer, s As String, a As String, b As String, c As String, j As Integer

For i = 0 To mnuSub.Count - 1
 If i <> 0 Then mnuSub(i).Visible = False
 mnuSub(i).Tag = ""
Next i

For i = 0 To lstMenu.ListCount - 1
 If Left(lstMenu.List(i), 6) = ". . . " Then
  a = Left(Mid(lstMenu.List(i), 7), InStrRev(Mid(lstMenu.List(i), 7), " - ") - 1)
  b = Mid(Mid(lstMenu.List(i), 7), InStrRev(Mid(lstMenu.List(i), 7), " - ") + 3)
  c = Mid(b, InStrRev(b, "^") + 1)
  b = Left(b, InStrRev(b, "^") - 1)
  s = s & "newsubmenu(" & c & ",""" & a & """,""" & b & """)" & vbCrLf


  For j = 0 To mnuSub.Count - 1
   If mnuSub(j).Tag = "" Then GoTo 1
  Next j
   j = mnuSub.Count
   Call Load(mnuSub(j))
1
   
  
  'MsgBox mnuSub(j).Caption & " - " & j
  mnuSub(j).Visible = True
  mnuSub(j).Tag = a & Chr(1) & c
  mnuSub(j).Caption = b
 Else
  a = Left(lstMenu.List(i), InStrRev(lstMenu.List(i), " - ") - 1)
  b = Mid(lstMenu.List(i), InStrRev(lstMenu.List(i), " - ") + 3)
  c = Mid(b, InStrRev(b, "^") + 1)
  b = Left(b, InStrRev(b, "^") - 1)
  s = s & "newmenu(" & c & ",""" & a & """,""" & b & """)" & vbCrLf
 End If
Next i
gblMenu = s
If Me.Visible = True Then Me.Hide
End Sub

Private Sub cmdPos_Click(Index As Integer)
On Error GoTo 1
Dim s As String, i As Integer
i = lstMenu.ListIndex
If i = -1 Then GoTo 1
s = lstMenu.List(i)
Select Case Index
 Case 0
  If Left(s, 6) = ". . . " Then lstMenu.List(i) = Mid(s, 7)
 Case 1
  If Left(s, 6) <> ". . . " Then lstMenu.List(i) = ". . . " & s
 Case 2
  s = lstMenu.List(i - 1)
  If s <> "" Then
   lstMenu.List(i - 1) = lstMenu.List(i)
   lstMenu.List(i) = s
   lstMenu.ListIndex = i - 1
  End If
 Case 3
  s = lstMenu.List(i + 1)
  If s <> "" Then
   lstMenu.List(i + 1) = lstMenu.List(i)
   lstMenu.List(i) = s
   lstMenu.ListIndex = i + 1
  End If
End Select
1
End Sub

Private Sub cmdRem_Click()

If lstMenu.ListIndex <> -1 Then Call lstMenu.RemoveItem(lstMenu.ListIndex)
End Sub

Private Sub Form_Activate()
Dim frm As Form, i As Integer, j As Integer
Call cmbWin.Clear
For Each frm In Forms
 If frm.Tag <> "" And frm.Name = "frmWin" Then
  cmbWin.AddItem Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1)
  If Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = b_Sender Then j = i
  i = i + 1
 End If
Next
cmbWin.ListIndex = j
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If gblCanClose = False Then Cancel = -1: Me.Hide
End Sub

Private Sub lstMenu_Click()
Dim i As Integer
Dim a As String, b As String, c As String
i = lstMenu.ListIndex
 If Left(lstMenu.List(i), 6) = ". . . " Then
  a = Left(Mid(lstMenu.List(i), 7), InStr(Mid(lstMenu.List(i), 7), " - ") - 1)
  b = Mid(Mid(lstMenu.List(i), 7), InStr(Mid(lstMenu.List(i), 7), " - ") + 3)
  c = Mid(b, InStrRev(b, "^") + 1)
  b = Left(b, InStrRev(b, "^") - 2)
 Else
  a = Left(lstMenu.List(i), InStrRev(lstMenu.List(i), " - ") - 1)
  b = Mid(lstMenu.List(i), InStrRev(lstMenu.List(i), " - ") + 3)
  c = Mid(b, InStrRev(b, "^") + 1)
  b = Left(b, InStrRev(b, "^") - 2)
 End If
txtName = a
txtText = b
For i = 0 To cmbWin.ListCount - 1
 If cmbWin.List(i) = c Then cmbWin.ListIndex = i: Exit For
Next i
End Sub

Private Sub mnuSub_Click(Index As Integer)
If mnuSub(Index).Tag = "" Then Exit Sub
Dim a As String, b As String
a = Left(mnuSub(Index).Tag, InStr(mnuSub(Index).Tag, Chr(1)) - 1)
b = Mid(mnuSub(Index).Tag, InStr(mnuSub(Index).Tag, Chr(1)) + 1)

Dim frm As Form, i As Integer
'Dim frmW As Form
Set frm = FindEdit(b)
'Set frmW = FormIndex(b)

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & a & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & a & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub
