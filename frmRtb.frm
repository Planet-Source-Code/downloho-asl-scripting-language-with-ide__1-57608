VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6060
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "dddd"
            TextSave        =   "dddd"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "dddddd"
            TextSave        =   "dddddd"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmRtb.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_GETLINECOUNT = &HBA

Private Sub Command1_Click()
MsgBox rtb.GetLineFromChar(rtb.SelStart)
End Sub

Private Sub rtb_Change()
stBar.Panels(1).Text = "Lines: " & SendMessage(rtb.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
End Sub

Private Sub rtb_KeyUp(KeyCode As Integer, Shift As Integer)
If rtb.SelLength <> 0 Then Exit Sub

Dim l&

 l& = rtb.SelStart: y& = l&: c& = rtb.SelColor
 If l& <> 0 Then
  m& = InStrRev(rtb.Text, Chr(32), l&)
  n& = InStrRev(rtb.Text, Chr(10), l&)
  o& = InStrRev(rtb.Text, Chr(13), l&)

  If m& > n& Then a& = m& Else a& = n& + 1
  If a& < o& Then a& = o& + 1
  If a& = 0 Then a& = 1
 Else
  a& = 1
  l& = 1
 End If

 p& = InStr(l& + 1, rtb.Text, Chr(32))
 r& = InStr(l& + 1, rtb.Text, Chr(13))
 q& = InStr(l& + 1, rtb.Text, Chr(10))
 s& = InStr(l& + 1, rtb.Text, "(")
 
 'If p& < r& And r& <> 0 Then b& = p& Else b& = r& - 1
 If b& < q& And q& <> 0 Then b& = q& - 1
 If b& <= 0 Then b& = Len(rtb.Text) + 1
 If p& < b& And p& <> 0 Then b& = p&
 If s& < b& And s& <> 0 Then b& = s&

 Caption = LCase$(Trim(Mid$(rtb.Text, a&, b& - a&)))

With rtb
 Select Case LCase$(Trim(Mid$(rtb.Text, a&, b& - a&)))
  Case "var", "!proc", "end!"
   .SelStart = a& - 1
   .SelLength = b& - a&
   'If LCase$(Trim(Mid$(rtb.Text, a&, b& - a&))) = "var" Then
   rtb.SelColor = RGB(23, 23, 200)
   'Else .SelBold = True
   .SelStart = y&
   .SelBold = False
  Case "$", "(", ")"
   .SelStart = rtb.SelStart - 1
   .SelLength = 1
   .SelColor = RGB(200, 23, 23)
   .SelStart = rtb.SelStart + 1
   .SelColor = vbBlack
  Case Else
   If Shift = 1 Then
    Select Case KeyCode
     Case 52, 50
      .SelStart = rtb.SelStart - 1
      .SelLength = 1
      .SelColor = RGB(200, 23, 23)
      .SelStart = rtb.SelStart + 1
      .SelColor = vbBlack
     Case 57, 48
      .SelStart = rtb.SelStart - 1
      .SelLength = 1
      .SelColor = RGB(100, 23, 100)
      .SelStart = rtb.SelStart + 1
      .SelColor = vbBlack
    End Select
   Else
    .SelBold = False
    .SelColor = vbBlack
   End If
 End Select
End With

Call rtb_MouseUp(0, 0, 0, 0)
End Sub

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
stBar.Panels(2).Text = "Line: " & (rtb.GetLineFromChar(rtb.SelStart) + 1)
If rtb.SelStart <> 0 Then
 Dim l&
 l& = InStrRev(rtb.Text, Chr(13), rtb.SelStart)
 If l& < InStrRev(rtb.Text, Chr(13), rtb.SelStart) And l& <> 0 Then l& = InStrRev(rtb.Text, Chr(13), rtb.SelStart)
 Caption = l&
 If l& = 0 Then l& = -1
 stBar.Panels(3).Text = "Chr: " & rtb.SelStart - l&
Else
 stBar.Panels(3).Text = "Chr: 1"
End If
End Sub
