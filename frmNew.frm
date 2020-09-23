VERSION 5.00
Begin VB.Form frmNew 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   129
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   2  'CenterScreen
   Begin VisualAce.winConnect wc 
      Index           =   0
      Left            =   3720
      Top             =   0
      _ExtentX        =   1535
      _ExtentY        =   661
   End
   Begin VB.DriveListBox drvNew 
      Height          =   315
      Index           =   0
      Left            =   4080
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.DirListBox dirNew 
      Height          =   315
      Index           =   0
      Left            =   3840
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.FileListBox flbNew 
      Height          =   480
      Index           =   0
      Left            =   3840
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   1560
   End
   Begin VB.Timer tmrNew 
      Enabled         =   0   'False
      Index           =   0
      Left            =   2640
      Top             =   1320
   End
   Begin VB.OptionButton optNew 
      Caption         =   "Option"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkNew 
      Caption         =   "Check"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbNew 
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Text            =   "Combo"
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstNew 
      Height          =   450
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox memNew 
      Height          =   615
      Index           =   0
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmNew.frx":0CCA
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Button"
      Height          =   375
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblNew 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgNew 
      Height          =   375
      Index           =   0
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblANew 
      Caption         =   "Label"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Menu mnuNew 
      Caption         =   "Menu"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu subNew 
         Caption         =   "Sub Menu"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Res As New Collection

Public Sub LoadInactiveCtrl(ByVal oName As String, ByVal sName As String, l As Integer, t As Integer, w As Integer, h As Integer)
Dim i As Integer
i = wc.Count
Call Load(wc(i))
wc(i).Tag = sName & Chr(1) & l & Chr(2) & t & Chr(2) & w & Chr(2) & h
If modLan.FileExist(App.Path & "\" & oName & ".xco") = True Then
 If modLan.FileExist(TEMP_PATH & "\" & oName & ".exe") = False Then
  CopyFile App.Path & "\" & oName & ".xco", TEMP_PATH & "\" & oName & ".exe", 0
  m_Res.Add TEMP_PATH & "\" & oName & ".exe"
 End If
ElseIf modLan.FileExist(SYS_PATH & "\" & oName & ".xco") = True Then
 If modLan.FileExist(TEMP_PATH & "\" & oName & ".exe") = False Then
  CopyFile SYS_PATH & "\" & oName & ".xco", TEMP_PATH & "\" & oName & ".exe", 0
  m_Res.Add TEMP_PATH & "\" & oName & ".exe"
 End If
Else
 MsgBox "Error: Can't find InActiveControl", vbCritical, "Error #100"
 Exit Sub
End If
'MsgBox TEMP_PATH & "\" & oName & ".exe"
Call wc(i).Run(TEMP_PATH & oName & ".exe", "0")
Me.SetFocus
End Sub

Public Sub LoadNLL(ByVal oName As String)
End Sub

Private Sub chkNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & chkNew(Index%).Tag, Me.Tag & "_" & chkNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & chkNew(Index%).Tag & "_Click")
End Sub

Private Sub chkNew_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Call modLan.setString("$button", Button, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$shift", Shift, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$x", Int(X / Screen.TwipsPerPixelX), Me.Tag & "_" & chkNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$y", Int(Y / Screen.TwipsPerPixelY), Me.Tag & "_" & chkNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$_", Me.Tag & "." & chkNew(Index%).Tag, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseDown")

Call modLan.Execute(modLan.sString, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseDown")
End Sub

Private Sub chkNew_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call modLan.setString("$button", Button, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseMove")
Call modLan.setString("$shift", Shift, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseMove")
Call modLan.setString("$x", Int(X / Screen.TwipsPerPixelX), Me.Tag & "_" & chkNew(Index%).Tag & "_MouseMove")
Call modLan.setString("$y", Int(Y / Screen.TwipsPerPixelY), Me.Tag & "_" & chkNew(Index%).Tag & "_MouseMove")
Call modLan.setString("$_", Me.Tag & "." & chkNew(Index%).Tag, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseMove")

Call modLan.Execute(modLan.sString, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseMove")
End Sub

Private Sub chkNew_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call modLan.setString("$button", Button, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$shift", Shift, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$x", Int(X / Screen.TwipsPerPixelX), Me.Tag & "_" & chkNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$y", Int(Y / Screen.TwipsPerPixelY), Me.Tag & "_" & chkNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$_", Me.Tag & "." & chkNew(Index%).Tag, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseUp")

Call modLan.Execute(modLan.sString, Me.Tag & "_" & chkNew(Index%).Tag & "_MouseUp")
End Sub

Private Sub cmbNew_Change(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & cmbNew(Index%).Tag, Me.Tag & "_" & cmbNew(Index%).Tag & "_Change")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & cmbNew(Index%).Tag & "_Change")
End Sub

Private Sub cmbNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & cmbNew(Index%).Tag, Me.Tag & "_" & cmbNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & cmbNew(Index%).Tag & "_Click")
End Sub

Private Sub cmdNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & cmdNew(Index%).Tag, Me.Tag & "_" & cmdNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & cmdNew(Index%).Tag & "_Click")
End Sub

Private Sub dirNew_Change(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & dirNew(Index%).Tag, Me.Tag & "_" & dirNew(Index%).Tag & "_Change")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & dirNew(Index%).Tag & "_Change")
End Sub

Private Sub dirNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & dirNew(Index%).Tag, Me.Tag & "_" & dirNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & dirNew(Index%).Tag & "_Click")
End Sub

Private Sub drvNew_Change(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & drvNew(Index%).Tag, Me.Tag & "_" & drvNew(Index%).Tag & "_Change")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & drvNew(Index%).Tag & "_Change")
End Sub

Private Sub flbNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & flbNew(Index%).Tag, Me.Tag & "_" & flbNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & flbNew(Index%).Tag & "_Click")
End Sub

Private Sub flbNew_DblClick(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & flbNew(Index%).Tag, Me.Tag & "_" & flbNew(Index%).Tag & "_DblClick")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & flbNew(Index%).Tag & "_DblClick")
End Sub

Private Sub Form_Click()
Call modLan.setString("$_", Me.Tag, Me.Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_Click")
End Sub

Private Sub Form_Load()
tmrLoad.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim i As Integer
For i = 0 To wc.Count - 1
 wc(i).Send "~"
Next i

For i = 1 To m_Res.Count
 Call Kill(m_Res(i))
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
If modLan.gblEnd = False Then Cancel = -1: Me.Hide
End Sub

Private Sub imgNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & imgNew(Index%).Tag, Me.Tag & "_" & imgNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & imgNew(Index%).Tag & "_Click")
End Sub

Private Sub imgNew_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call modLan.setString("$_", Me.Tag & "." & imgNew(Index%).Tag, Me.Tag & "_" & imgNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$Button", CStr(Button), Me.Tag & "_" & imgNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$x", CStr(Int(X / Screen.TwipsPerPixelX)), Me.Tag & "_" & imgNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$y", CStr(Int(Y / Screen.TwipsPerPixelY)), Me.Tag & "_" & imgNew(Index%).Tag & "_MouseDown")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & imgNew(Index%).Tag & "_MouseDown")
End Sub

Private Sub imgNew_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call modLan.setString("$_", Me.Tag & "." & imgNew(Index%).Tag, Me.Tag & "_" & imgNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$Button", CStr(Button), Me.Tag & "_" & imgNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$x", CStr(Int(X / Screen.TwipsPerPixelX)), Me.Tag & "_" & imgNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$y", CStr(Int(Y / Screen.TwipsPerPixelY)), Me.Tag & "_" & imgNew(Index%).Tag & "_MouseUp")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & imgNew(Index%).Tag & "_MouseUp")
End Sub

Private Sub lblNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & lblNew(Index%).Tag, Me.Tag & "_" & lblNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & lblNew(Index%).Tag & "_Click")
End Sub

Private Sub lblNew_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call modLan.setString("$_", Me.Tag & "." & lblNew(Index%).Tag, Me.Tag & "_" & lblNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$Button", CStr(Button), Me.Tag & "_" & lblNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$x", CStr(Int(X / Screen.TwipsPerPixelX)), Me.Tag & "_" & lblNew(Index%).Tag & "_MouseDown")
Call modLan.setString("$y", CStr(Int(Y / Screen.TwipsPerPixelY)), Me.Tag & "_" & lblNew(Index%).Tag & "_MouseDown")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & lblNew(Index%).Tag & "_MouseDown")
End Sub

Private Sub lblNew_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call modLan.setString("$_", Me.Tag & "." & lblNew(Index%).Tag, Me.Tag & "_" & lblNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$Button", CStr(Button), Me.Tag & "_" & lblNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$x", CStr(Int(X / Screen.TwipsPerPixelX)), Me.Tag & "_" & lblNew(Index%).Tag & "_MouseUp")
Call modLan.setString("$y", CStr(Int(Y / Screen.TwipsPerPixelY)), Me.Tag & "_" & lblNew(Index%).Tag & "_MouseUp")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & lblNew(Index%).Tag & "_MouseUp")
End Sub

Private Sub lstNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & lstNew(Index%).Tag, Me.Tag & "_" & lstNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & lstNew(Index%).Tag & "_Click")
End Sub

Private Sub lstNew_DblClick(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & lstNew(Index%).Tag, Me.Tag & "_" & lstNew(Index%).Tag & "_DblClick")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & lstNew(Index%).Tag & "_DblClick")
End Sub

Private Sub memNew_Change(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & memNew(Index%).Tag, Me.Tag & "_" & memNew(Index%).Tag & "_Change")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & memNew(Index%).Tag & "_Change")
End Sub

Private Sub memNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & memNew(Index%).Tag, Me.Tag & "_" & memNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & memNew(Index%).Tag & "_Click")
End Sub

Private Sub mnuNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & mnuNew(Index%).Tag, Me.Tag & "_" & mnuNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & mnuNew(Index%).Tag & "_Click")
End Sub

Private Sub optNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & optNew(Index%).Tag, Me.Tag & "_" & optNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & optNew(Index%).Tag & "_Click")
End Sub

Private Sub subNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & subNew(Index%).Tag, Me.Tag & "_" & subNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & subNew(Index%).Tag & "_Click")
End Sub

Private Sub tmrLoad_Timer()
If Me.Tag <> "" Then
Call modLan.setString("$_", Me.Tag, Me.Tag & "_Init")
Call modLan.Execute$(modLan.sString, Me.Tag & "_init")
tmrLoad.Enabled = False
End If
End Sub

Private Sub tmrNew_Timer(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & tmrNew(Index%).Tag, Me.Tag & "_" & tmrNew(Index%).Tag & "_Timer")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & tmrNew(Index%).Tag & "_Timer")
End Sub

Private Sub txtNew_Change(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & txtNew(Index%).Tag, Me.Tag & "_" & txtNew(Index%).Tag & "_Change")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & txtNew(Index%).Tag & "_Change")
End Sub

Private Sub txtNew_Click(Index As Integer)
Call modLan.setString("$_", Me.Tag & "." & txtNew(Index%).Tag, Me.Tag & "_" & txtNew(Index%).Tag & "_Click")
Call modLan.Execute(modLan.sString, Me.Tag & "_" & txtNew(Index%).Tag & "_Click")
End Sub

Private Sub wc_Got(Index As Integer, ByVal Msg As String)
Dim arr() As String
arr() = Split(Mid(Msg, 2), Chr(2))
'MsgBox Msg
Select Case Left(Msg, 1)
 Case "@"
  wc(Index).mHwnd = arr(0)
  wc(Index).Send "!" & Me.hwnd & Chr(2) & Mid(wc(Index).Tag, InStr(wc(Index).Tag, Chr(1)) + 1)
  wc(Index).Tag = Left(wc(Index).Tag, InStr(wc(Index).Tag, Chr(1)) - 1)
 Case "$" 'function
  arr(0) = Me.Tag & "_" & wc(Index).Tag & "_" & arr(0)
  For i = 1 To UBound(arr())
   Call modLan.doCode(arr(i), arr(0))
  Next i
  Call modLan.Execute(modLan.sString, arr(0))
End Select
End Sub
