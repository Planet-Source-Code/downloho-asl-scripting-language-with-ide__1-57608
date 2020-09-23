VERSION 5.00
Begin VB.UserControl xDialog 
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   ScaleHeight     =   62
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   93
   ToolboxBitmap   =   "xDialog.ctx":0000
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Width           =   510
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   2
      Left            =   960
      Picture         =   "xDialog.ctx":0312
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "xDialog.ctx":0F54
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "xDialog.ctx":1B96
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "xDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Click()
Event DblClick()
Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event KeyPress(ByVal KeyAscii As Integer)
Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Public Enum UDE_DIALOG
 dTimer = 0
 dMenu = 1
 dControl = 2
End Enum

Dim mEnabled As Boolean, mLMouseDown As Boolean, mRMouseDown As Boolean, mStyle As UDE_DIALOG
Dim mLastX As Integer, mLastY As Integer, mVisible As Boolean, mInterval As Long

Public Property Let pEnabled(b As Boolean)
mEnabled = b
End Property

Public Property Get pEnabled() As Boolean
pEnabled = mEnabled
End Property

Public Property Let pInterval(b As Long)
mInterval = b
End Property

Public Property Get pInterval() As Long
pInterval = mInterval
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

Public Property Let pVisible(b As Boolean)
mVisible = b
End Property

Public Property Get pVisible() As Boolean
pVisible = mVisible
End Property

Public Property Let pStyle(b As UDE_DIALOG)
mStyle = b
Set pic.Picture = img(b).Picture
End Property

Public Property Get pStyle() As UDE_DIALOG
pStyle = mStyle
End Property

Private Sub pic_Click()
RaiseEvent Click
End Sub

Private Sub pic_DblClick()
RaiseEvent DblClick
End Sub

Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub pic_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub pic_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 mLMouseDown = True
 mLastX = X: mLastY = Y
ElseIf Button = 2 Then
 mRMouseDown = True
 mLastX = X: mLastY = Y
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
If Button = 1 Then
 mLMouseDown = False
 mLastX = 0: mLastY = 0
ElseIf Button = 2 Then
 mRMouseDown = False
 mLastX = 0: mLastY = 0
End If
End Sub

Private Sub UserControl_Initialize()
mStyle = dTimer
Call UserControl.PropertyChanged("pStyle")
mEnabled = True
Call UserControl.PropertyChanged("pEnabled")
mVisible = True
Call UserControl.PropertyChanged("pVisible")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
mEnabled = PropBag.ReadProperty("pEnabled", True)
mVisible = PropBag.ReadProperty("pVisible", True)
mLastX = PropBag.ReadProperty("pLastX", 0)
mLastY = PropBag.ReadProperty("pLastY", 0)
mLMouseDown = PropBag.ReadProperty("pLMouseDown", False)
mRMouseDown = PropBag.ReadProperty("pRMouseDown", False)
mStyle = PropBag.ReadProperty("pStyle", dTimer)
Set pic.Picture = img(mStyle).Picture
End Sub

Private Sub UserControl_Resize()
Width = pic.Width * Screen.TwipsPerPixelX
Height = pic.Height * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("pEnabled", mEnabled, True)
Call PropBag.WriteProperty("pVisible", mVisible, True)
Call PropBag.WriteProperty("pLastX", mLastX, 0)
Call PropBag.WriteProperty("pLastY", mLastY, 0)
Call PropBag.WriteProperty("pRMouseDown", mRMouseDown, False)
Call PropBag.WriteProperty("pLMouseDown", mLMouseDown, False)
Call PropBag.WriteProperty("pStyle", mStyle, dTimer)
End Sub

