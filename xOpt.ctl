VERSION 5.00
Begin VB.UserControl xOpt 
   AutoRedraw      =   -1  'True
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   Begin VB.Image img 
      Height          =   165
      Index           =   3
      Left            =   1560
      Picture         =   "xOpt.ctx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image img 
      Height          =   180
      Index           =   2
      Left            =   1200
      Picture         =   "xOpt.ctx":01CE
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image img 
      Height          =   165
      Index           =   1
      Left            =   480
      Picture         =   "xOpt.ctx":03C0
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image img 
      Height          =   180
      Index           =   0
      Left            =   120
      Picture         =   "xOpt.ctx":058E
      Top             =   120
      Width           =   180
   End
End
Attribute VB_Name = "xOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum XSELStyle
 CheckBox = 0
 Radio = 1
End Enum

Event Click()
Event DblClick()
Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event KeyPress(ByVal KeyAscii As Integer)
Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim mEnabled As Boolean, mStyle As XSELStyle, mVisible As Boolean
Dim mLMouseDown As Boolean, mRMouseDown As Boolean, mValue As Boolean
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

Public Property Let pStyle(b As XSELStyle)
If b = CheckBox Then
 img(0).Visible = True
 img(1).Visible = False
Else
 img(0).Visible = False
 img(1).Visible = True
End If
mStyle = b
End Property

Public Property Get pStyle() As XSELStyle
pStyle = mStyle
End Property

Public Property Let pText(b As String)
mText = b
Call PrintTxt
End Property

Public Property Get pText() As String
pText = mText
End Property

Public Property Get pValue() As Boolean
pValue = mValue
End Property

Public Property Let pValue(b As Boolean)
mValue = b
If mStyle = CheckBox Then
 If mValue = False Then
  img(0).Visible = True
  img(1).Visible = False
  img(2).Visible = False
  img(3).Visible = False
 Else
  img(2).Visible = True
  img(1).Visible = False
  img(0).Visible = False
  img(3).Visible = False
 End If
Else
 If mValue = False Then
  img(0).Visible = False
  img(1).Visible = True
  img(2).Visible = False
  img(3).Visible = False
 Else
  img(0).Visible = False
  img(3).Visible = True
  img(2).Visible = False
  img(1).Visible = False
 End If
End If
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
CurrentX = img(0).Left + img(0).Width + 5  '(UserControl.ScaleWidth / 2) - (UserControl.TextWidth(s$) / 2) - 2 + OffSet%
CurrentY = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(s$) / 2) + OffSet%
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
mValue = PropBag.ReadProperty("pValue", False)
mStyle = PropBag.ReadProperty("pStyle", TextBox)

If mStyle = CheckBox Then
 If mValue = False Then
  img(0).Visible = True
  img(1).Visible = False
  img(2).Visible = False
  img(3).Visible = False
 Else
  img(2).Visible = True
  img(1).Visible = False
  img(0).Visible = False
  img(3).Visible = False
 End If
Else
 If mValue = False Then
  img(0).Visible = False
  img(1).Visible = True
  img(2).Visible = False
  img(3).Visible = False
 Else
  img(0).Visible = False
  img(3).Visible = True
  img(2).Visible = False
  img(1).Visible = False
 End If
End If

mLastX = PropBag.ReadProperty("pLastX", 0)
mLastY = PropBag.ReadProperty("pLastY", 0)
mLMouseDown = PropBag.ReadProperty("pLMouseDown", False)
mRMouseDown = PropBag.ReadProperty("pRMouseDown", False)
mText = PropBag.ReadProperty("pText", UserControl.Name)

Call PrintTxt
End Sub

Private Sub UserControl_Resize()
img(0).Left = 2
img(1).Left = 2
img(0).Top = ((Height / Screen.TwipsPerPixelY) / 2) - (img(0).Height / 2)
img(1).Top = ((Height / Screen.TwipsPerPixelY) / 2) - (img(1).Height / 2)

img(2).Left = 2
img(2).Top = img(0).Top

img(3).Left = 2
img(3).Top = img(0).Top

Call PrintTxt
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("pEnabled", mEnabled, True)
Call PropBag.WriteProperty("pValue", mValue, False)
Call PropBag.WriteProperty("pStyle", mStyle, TextBox)

If mStyle = CheckBox Then
 If mValue = False Then
  img(0).Visible = True
  img(1).Visible = False
  img(2).Visible = False
  img(3).Visible = False
 Else
  img(2).Visible = True
  img(1).Visible = False
  img(0).Visible = False
  img(3).Visible = False
 End If
Else
 If mValue = False Then
  img(0).Visible = False
  img(1).Visible = True
  img(2).Visible = False
  img(3).Visible = False
 Else
  img(0).Visible = False
  img(3).Visible = True
  img(2).Visible = False
  img(1).Visible = False
 End If
End If

Call PropBag.WriteProperty("pLastX", mLastX, 0)
Call PropBag.WriteProperty("pLastY", mLastY, 0)
Call PropBag.WriteProperty("pRMouseDown", mRMouseDown, False)
Call PropBag.WriteProperty("pLMouseDown", mLMouseDown, False)
Call PropBag.WriteProperty("pText", mText, UserControl.Name)
End Sub

