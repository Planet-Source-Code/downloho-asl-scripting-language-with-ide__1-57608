VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEdit 
   Caption         =   "Edit ( )"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6570
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5716
            MinWidth        =   1985
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1984
            MinWidth        =   1984
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1984
            MinWidth        =   1984
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1984
            MinWidth        =   1984
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1984
            MinWidth        =   1984
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbx 
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"frmEdit.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstCode 
      Appearance      =   0  'Flat
      Height          =   1230
      ItemData        =   "frmEdit.frx":04C2
      Left            =   3240
      List            =   "frmEdit.frx":058C
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cmbProc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmEdit.frx":07A2
      Left            =   3120
      List            =   "frmEdit.frx":07A9
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cmbObj 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmEdit.frx":07BB
      Left            =   0
      List            =   "frmEdit.frx":07C2
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin RichTextLib.RichTextBox rtbEdit 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9128
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmEdit.frx":07D1
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
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2

Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETLINE = &HC4
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_GETLINECOUNT = &HBA

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public IsAttached As Boolean

Private Function Rgb2Html(ByVal l As Long) As String
Rgb2Html$ = "#" & IIf(Len(Hex(GetRGB(l&).Red)) = 1, "0", "") & Hex(GetRGB(l&).Red) & _
            IIf(Len(Hex(GetRGB(l&).Green)) = 1, "0", "") & Hex(GetRGB(l&).Green) & _
            IIf(Len(Hex(GetRGB(l&).Blue)) = 1, "0", "") & Hex(GetRGB(l&).Blue)
End Function

Private Sub Form_Load()
cmbProc.ListIndex = 0
cmbObj.ListIndex = 0
's$ = "!proc Window1_Button1_Click()" & vbCrLf & "msgbox(""Button Clicked"",0,""Button"")" & vbCrLf & "end!" & vbCrLf
'MsgBox s$
'Call modColor.PrintText(s$, rtbEdit, rtbx)
'rtbEdit.Text = "!proc Window1_Button1_Click()" & vbCrLf & "msgbox(""Button Clicked"",0,""Button"")" & vbCrLf & "end!" & vbCrLf
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Or mdiMain.WindowState = vbMinimized Then Exit Sub
If Me.Width <= 1000 Then Me.Width = 1000: Exit Sub
If Me.Height <= 1000 Then Me.Height = 1000: Exit Sub

rtbEdit.Width = Me.ScaleWidth
rtbEdit.Top = 2 'cmbObj.Height + 2
rtbEdit.Height = Me.ScaleHeight - rtbEdit.Top - stBar.Height

cmbObj.Width = (Me.ScaleWidth / 2) - 2
cmbProc.Width = cmbObj.Width
cmbProc.Left = cmbObj.Left + cmbObj.Width + 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
If gblCanClose = False Then Cancel = -1: Me.Hide
End Sub

Private Sub lstCode_Click()
rtbEdit.SelText = lstCode.Text
lstCode.Visible = False
End Sub

Private Sub rtbEdit_Change()
On Error Resume Next
stBar.Panels(2).Text = "Lines: " & SendMessage(rtbEdit.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
Exit Sub
'needs to be fixed!
Dim l As Long

l& = InStr(l& + 1, LCase$(rtbEdit.Text), "!proc")

If InStr(Tag, Chr(0)) = 0 Then
cmbProc.Clear
cmbProc.AddItem "(Procedures)"
'cmbProc.ListIndex = 0
Do Until l& = 0
 cmbProc.AddItem Mid$(rtbEdit.Text, InStr(l& + 1, rtbEdit.Text, " ") + 1, InStr(l& + 1, rtbEdit.Text, "(") - InStr(l& + 1, rtbEdit.Text, " ") - 1)

  If Mid$(Mid$(rtbEdit.Text, InStrRev(rtbEdit.Text, "!proc", rtbEdit.SelStart - 1)), InStr(l& + 1, Mid$(rtbEdit.Text, InStrRev(rtbEdit.Text, "!proc", rtbEdit.SelStart - 1)), " ") + 1, InStr(l& + 1, Mid$(rtbEdit.Text, InStrRev(rtbEdit.Text, "!proc", rtbEdit.SelStart - 1)), "(") - InStr(l& + 1, Mid$(rtbEdit.Text, InStrRev(rtbEdit.Text, "!proc", rtbEdit.SelStart - 1)), " ") - 1) = Mid$(rtbEdit.Text, InStr(l& + 1, rtbEdit.Text, " ") + 1, InStr(l& + 1, rtbEdit.Text, "(") - InStr(l& + 1, rtbEdit.Text, " ") - 1) Then cmbProc.ListIndex = cmbProc.ListCount - 1
  l& = InStr(l& + 1, LCase$(rtbEdit.Text), "!proc")
 DoEvents
Loop

End If
End Sub

Private Sub rtbEdit_Click()
lstCode.Visible = False
End Sub

Private Sub rtbEdit_GotFocus()
For i = 1 To mdiMain.tvwFiles.Nodes.Count
 If InStr(Tag, Chr(0)) <> 0 Then
  If mdiMain.tvwFiles.Nodes(i).Text = Left(Tag, InStr(Tag, Chr(0)) - 1) Then mdiMain.tvwFiles.Nodes(i).Selected = True: Exit For
 Else
  If mdiMain.tvwFiles.Nodes(i).Text = Left(GetFileName(Tag), Len(GetFileName(Tag)) - 4) Then mdiMain.tvwFiles.Nodes(i).Selected = True: Exit For
 End If
Next i
End Sub

Private Sub rtbEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'Exit Sub
Dim pt As POINTAPI
Static s As String
Call GetCaretPos(pt)

If lstCode.Visible = True Then
 s = s & Chr(KeyCode)
 Dim l As Long
 l = SendMessageByString(lstCode.hwnd, LB_FINDSTRING, ByVal 0&, s)
 If l <> -1 Then lstCode.ListIndex = l - 1
 Call lstCode.Move(rtbEdit.Left + (rtbEdit.Font.Size / 2) + pt.X, rtbEdit.Top + (rtbEdit.Font.Size * 2) + pt.Y)
ElseIf Shift = 2 And KeyCode = 32 Then
 s = ""
 lstCode.Visible = True
 Call lstCode.Move(rtbEdit.Left + (rtbEdit.Font.Size / 2) + pt.X, rtbEdit.Top + (rtbEdit.Font.Size * 2) + pt.Y)
 
End If

End Sub

Private Sub rtbEdit_KeyUp(KeyCode As Integer, Shift As Integer)
Dim pt As POINTAPI
Call GetCaretPos(pt)
'Caption = pt.X & " - " & pt.Y
Call lstCode.Move(rtbEdit.Left + (rtbEdit.Font.Size / 2) + pt.X, rtbEdit.Top + (rtbEdit.Font.Size * 2) + pt.Y)

If rtbEdit.SelLength <> 0 Then Exit Sub

Dim l&

 l& = rtbEdit.SelStart: Y& = l&: c& = rtbEdit.SelColor
 If l& <> 0 Then
  m& = InStrRev(rtbEdit.Text, Chr(32), l&)
  n& = InStrRev(rtbEdit.Text, Chr(10), l&)
  o& = InStrRev(rtbEdit.Text, Chr(13), l&)

  If m& > n& Then a& = m& Else a& = n& + 1
  If a& < o& Then a& = o& + 1
  If a& = 0 Then a& = 1
 Else
  a& = 1
  l& = 1
 End If

 p& = InStr(l& + 1, rtbEdit.Text, Chr(32))
 r& = InStr(l& + 1, rtbEdit.Text, Chr(13))
 q& = InStr(l& + 1, rtbEdit.Text, Chr(10))
 s& = InStr(l& + 1, rtbEdit.Text, "(")
 
 'If p& < r& And r& <> 0 Then b& = p& Else b& = r& - 1
 If b& < q& And q& <> 0 Then b& = q& - 1
 If b& <= 0 Then b& = Len(rtbEdit.Text) + 1
 If p& < b& And p& <> 0 Then b& = p&
 If s& < b& And s& <> 0 Then b& = s&

With rtbEdit
 Select Case LCase$(Trim(Mid$(rtbEdit.Text, a&, b& - a&)))
  Case "var", "!proc", "end!", "!type", "with", "include", "set"
   .SelStart = a& - 1
   .SelLength = b& - a&
   .SelBold = True
   .SelStart = Y&
   .SelBold = False
  Case "//"
   .SelStart = a - 1
   .SelLength = b - a
   .SelColor = RGB(0, 196, 0)
   .SelStart = Y
   .SelColor = vbWindowText
  Case "$", "@", "%", "&", "#", "^"
   .SelStart = .SelStart - 1
   .SelLength = 1
   .SelColor = RGB(200, 55, 0)
   .SelStart = .SelStart + 1
   .SelColor = vbBlack
  Case "(", ")"
   .SelStart = .SelStart - 1
   .SelLength = 1
   .SelColor = RGB(0, 55, 255)
   .SelStart = .SelStart + 1
   .SelColor = vbBlack
  Case Else
   If Shift = 1 Then
    Select Case KeyCode
     Case 52, 50, 53, 55
      .SelStart = .SelStart - 1
      .SelLength = 1
      .SelColor = vbRed
      .SelStart = .SelStart + 1
      .SelColor = vbBlack
     Case 57, 48
      .SelStart = .SelStart - 1
      .SelLength = 1
      .SelColor = vbBlue
      .SelStart = .SelStart + 1
      .SelColor = vbBlack
    End Select
   ElseIf Shift = 2 Then
   'MsgBox Shift & " - " & KeyCode
    Select Case KeyCode
     Case 38
      If l <> 1 Then a = InStrRev(.Text, "!proc ", l - 1)
      If a = 0 Then
       .SelStart = 1
      Else
       .SelStart = a
      End If
     Case 40
      a = InStr(l + 1, .Text, "!proc ")
      If a = 0 Then
       .SelStart = Len(.Text)
      Else
       .SelStart = a
      End If
    End Select
   Else
    .SelBold = False
    .SelColor = vbBlack
   End If
 End Select
End With
Call rtbEdit_MouseUp(0, 0, 0, 0)
End Sub

Private Sub rtbEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
stBar.Panels(3).Text = "Line: " & (rtbEdit.GetLineFromChar(rtbEdit.SelStart) + 1)

If InStrRev(LCase(rtbEdit.Text), "!proc ", rtbEdit.SelStart + 1) <> 0 Then stBar.Panels(5).Text = "Proc: " & (rtbEdit.GetLineFromChar(rtbEdit.SelStart) - rtbEdit.GetLineFromChar(InStrRev(LCase(rtbEdit.Text), "!proc ", rtbEdit.SelStart + 1)) + 1)


Dim l&, b As Long, s As String
If rtbEdit.SelStart <> 0 Then
 l& = InStrRev(rtbEdit.Text, Chr(13), rtbEdit.SelStart)
 If l& < InStrRev(rtbEdit.Text, Chr(13), rtbEdit.SelStart) And l& <> 0 Then l& = InStrRev(rtbEdit.Text, Chr(13), rtbEdit.SelStart)
 'Caption = l&
 If l& = 0 Then l& = -1
 stBar.Panels(4).Text = "Off: " & rtbEdit.SelStart - l&
'l& = InStrRev(rtbEdit.Text, "!proc", rtbEdit.SelStart)
 'b = InStr(l& + 1, rtbEdit.Text, "_")
 
 'If b = 0 Then b& = InStrRev(rtbEdit.Text, "_", rtbEdit.SelStart)
 'If l <> 0 And b <> 0 Then s = Mid(rtbEdit.Text, l + 6, b - l - 6)
Else
 'l& = InStr(rtbEdit.Text, "!proc")
 'b = InStr(l&, rtbEdit.Text, "_")
 
 'If b = 0 Then b& = InStrRev(rtbEdit.Text, "_", rtbEdit.SelStart)
 'If l <> 0 And b <> 0 Then s = Mid(rtbEdit.Text, l + 6, b - l - 6)
 stBar.Panels(4).Text = "Off: 1"
 stBar.Panels(5).Text = "Proc: 1"
End If

End Sub
