VERSION 5.00
Begin VB.UserControl xBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   ForeColor       =   &H80000008&
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   166
   Begin VB.Image imgDown 
      Height          =   240
      Left            =   0
      Picture         =   "xBox.ctx":0000
      Top             =   0
      Width           =   270
   End
End
Attribute VB_Name = "xBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum XBOXStyle
 TextBox = 0
 ListBox = 1
 Memo = 2
 StaticLabel = 3
 ComboBox = 4
End Enum

Event Click()
Event DblClick()
Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event KeyPress(ByVal KeyAscii As Integer)
Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim mEnabled As Boolean, mStyle As XBOXStyle, mVisible As Boolean
Dim mLMouseDown As Boolean, mRMouseDown As Boolean, mLWidth As Integer, mLHeight As Integer
Dim mLastX As Integer, mLastY As Integer, mText As String, DefaultCl(1) As Long

Public Property Get DefaultBCl() As Long
DefaultBCl = DefaultCl(0)
End Property

Public Property Get DefaultFCl() As Long
DefaultFCl = DefaultCl(1)
End Property

Public Property Let pBackColor(b As Long)
BackColor = b
Call PrintTxt
End Property

Public Property Get pBackColor() As Long
pBackColor = BackColor
End Property

Public Property Let pLimitHeight(b As Integer)
mLHeight = b
End Property

Public Property Get pLimitHeight() As Integer
pLimitHeight = mLHeight
End Property

Public Property Let pLimitWidth(b As Long)
mLWidth = b
End Property

Public Property Get pLimitWidth() As Long
pLimitWidth = mLWidth
End Property

Public Property Let pForeColor(b As Long)
ForeColor = b
Call PrintTxt
End Property

Public Property Get pForeColor() As Long
pForeColor = ForeColor
End Property

Public Property Let pEnabled(b As Boolean)
mEnabled = b
End Property

Public Property Get pEnabled() As Boolean
pEnabled = mEnabled
End Property

Public Property Get pLastX() As Integer
pLastX = mLastX
End Property

Public Property Get pLastY() As Integer
pLastY = mLastY
End Property

Public Property Get pLMouseDown() As Boolean
pLMouseDown = mLMouseDown
End Property

Public Property Get pRMouseDown() As Boolean
pRMouseDown = mRMouseDown
End Property

Public Property Let pStyle(b As XBOXStyle)
If b = ListBox Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False, True, False, False, True)
ElseIf b = Memo Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False, True, True, False, True)
ElseIf b = TextBox Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False, False, False, False, True)
ElseIf b = StaticLabel Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False)
 UserControl.BackColor = vbButtonFace
ElseIf b = CombBox Then
 Call ChangeWin(hwnd, False, False, False, False, False, False, False, False, False, True)
End If

mStyle = b
End Property

Public Property Get pStyle() As XBOXStyle
pStyle = mStyle
End Property

Public Property Let pText(b As String)
mText = b
Call PrintTxt
End Property

Public Property Get pText() As String
pText = mText
End Property

Public Property Let pVisible(b As Boolean)
mVisible = b
End Property

Public Property Get pVisible() As Boolean
pVisible = mVisible
End Property

Private Sub PrintTxt(Optional ByVal OffSet As Integer = 0)
Dim s$
s$ = mText
Cls
CurrentX = 0 '(UserControl.ScaleWidth / 2) - (UserControl.TextWidth(s$) / 2) - 2 + OffSet%
CurrentY = 0 '(UserControl.ScaleHeight / 2) - (UserControl.TextHeight(s$) / 2) + OffSet%
Print s$
Refresh
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
Call ChangeWin(hwnd, False, False, False, False, False, False, False, False, False, True)
mEnabled = True
Call UserControl.PropertyChanged("pEnabled")
mVisible = True
Call UserControl.PropertyChanged("pVisible")
DefaultCl(0) = BackColor
DefaultCl(1) = ForeColor
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 mLMouseDown = True
 mLastX = X: mLastY = Y
ElseIf Button = 2 Then
 mRMouseDown = True
 mLastX = X: mLastY = Y
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
If Button = 1 Then
 mLMouseDown = False
 mLastX = 0: mLastY = 0
ElseIf Button = 2 Then
 mRMouseDown = False
 mLastX = 0: mLastY = 0
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
mEnabled = PropBag.ReadProperty("pEnabled", True)
mStyle = PropBag.ReadProperty("pStyle", TextBox)

If mStyle = ListBox Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False, True, False, False, True)
ElseIf mStyle = Memo Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False, True, True, False, True)
ElseIf mStyle = TextBox Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False, False, False, False, True)
ElseIf mStyle = StaticLabel Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False)
 UserControl.BackColor = vbButtonFace
ElseIf mStyle = CombBox Then
 Call ChangeWin(hwnd, False, False, False, False, False, False, False, False, False, True)
End If

mLastX = PropBag.ReadProperty("pLastX", 0)
mLastY = PropBag.ReadProperty("pLastY", 0)
mLWidth = PropBag.ReadProperty("pLimitWidth", 0)
mLHeight = PropBag.ReadProperty("pLimitHeight", 0)
mLMouseDown = PropBag.ReadProperty("pLMouseDown", False)
mRMouseDown = PropBag.ReadProperty("pRMouseDown", False)
mText = PropBag.ReadProperty("pText", UserControl.Name)
Call PrintTxt
End Sub

Private Sub UserControl_Resize()
If mLWidth <> 0 And UserControl.Width <> (mLWidth * Screen.TwipsPerPixelX) Then UserControl.Width = (mLWidth * Screen.TwipsPerPixelX)
If mLHeight <> 0 And UserControl.Height <> (mLHeight * Screen.TwipsPerPixelY) Then UserControl.Height = (mLHeight * Screen.TwipsPerPixelY)
Call PrintTxt
imgDown.Top = 0
imgDown.Left = UserControl.ScaleWidth - imgDown.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("pEnabled", mEnabled, True)
Call PropBag.WriteProperty("pStyle", mStyle, TextBox)

If mStyle = ListBox Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False, True, False, False, True)
ElseIf mStyle = Memo Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False, True, True, False, True)
ElseIf mStyle = TextBox Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False, False, False, False, True)
ElseIf mStyle = StaticLabel Then
 imgDown.Visible = False
 Call ChangeWin(hwnd, False, False, False, False, False, False)
 UserControl.BackColor = vbButtonFace
ElseIf mStyle = CombBox Then
 Call ChangeWin(hwnd, False, False, False, False, False, False, False, False, False, True)
End If

Call PropBag.WriteProperty("pLastX", mLastX, 0)
Call PropBag.WriteProperty("pLastY", mLastY, 0)
Call PropBag.WriteProperty("pLimitWidth", mLWidth, 0)
Call PropBag.WriteProperty("pLimitHeight", mLHeight, 0)
Call PropBag.WriteProperty("pRMouseDown", mRMouseDown, False)
Call PropBag.WriteProperty("pLMouseDown", mLMouseDown, False)
Call PropBag.WriteProperty("pText", mText, UserControl.Name)
End Sub
