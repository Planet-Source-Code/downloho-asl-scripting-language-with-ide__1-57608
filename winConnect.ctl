VERSION 5.00
Begin VB.UserControl winConnect 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   870
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   405
   ScaleWidth      =   870
   ToolboxBitmap   =   "winConnect.ctx":0000
   Begin VB.TextBox txtRe 
      Height          =   285
      Left            =   1665
      TabIndex        =   0
      Text            =   "?"
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image img 
      Height          =   375
      Left            =   0
      Picture         =   "winConnect.ctx":0312
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "winConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1
Private Const WM_SETTEXT = &HC

Public mHwnd As Long

Event Got(ByVal Msg As String)
Event WindowGone(ByVal hwnd As Long)

Public Function IsWinGone(ByVal hwnd As Long) As Boolean
IsWinGone = CBool(IsWindow(hwnd&))
End Function

Private Sub SetText(ByVal Window As Long, ByVal Text As String)
'send a message to mHwnd to change the text
 If IsWinGone(Window&) = False Then RaiseEvent WindowGone(Window&): Exit Sub
 Call SendMessageByString(Window&, WM_SETTEXT, 0&, Text$)
End Sub

Public Sub Run(ByVal sFile As String, Optional ByVal mCommand As String)
  mCommand$ = "&" & txtRe.hwnd ' & Chr(2) & mCommand$
  Dim ret&
  ret& = ShellExecute(0, vbNullString, sFile$, IIf(mCommand$ = "", vbNullString, mCommand$), "c:\", SW_SHOWNORMAL)
End Sub

Public Sub Send(ByVal Msg As String)
 If mHwnd& <> 0 Then Call SetText(mHwnd&, Msg$)
End Sub

Public Sub SetCommand(ByVal Msg As String)
On Error GoTo 1
If Left$(Msg$, 1) = "&" Then
 mHwnd& = CLng(Mid$(Msg$, 2))
 Call Send("@" & txtRe.hwnd)
End If
Exit Sub
1
mHwnd = 0
End Sub

Private Sub txtRe_Change()
 If txtRe.Text = "?" Then Exit Sub
  RaiseEvent Got(txtRe.Text)
 txtRe.Text = "?"
End Sub

Private Sub UserControl_Resize()
Width = img.Width
Height = img.Height
End Sub
