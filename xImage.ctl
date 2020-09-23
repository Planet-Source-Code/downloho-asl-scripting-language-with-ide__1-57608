VERSION 5.00
Begin VB.UserControl xImage 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Line xLine 
      BorderStyle     =   3  'Dot
      Index           =   3
      X1              =   48
      X2              =   48
      Y1              =   144
      Y2              =   32
   End
   Begin VB.Line xLine 
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   48
      X2              =   168
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Line xLine 
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   168
      X2              =   168
      Y1              =   32
      Y2              =   144
   End
   Begin VB.Line xLine 
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   0
      X2              =   120
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "xImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Click()
Event SetPicture(ByVal pName As String)
Event DblClick()
Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event KeyPress(ByVal KeyAscii As Integer)
Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim mVisible As Boolean, mPicture As String
Dim mLMouseDown As Boolean, mRMouseDown As Boolean
Dim mLastX As Integer, mLastY As Integer

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

Public Property Let pPicture(b As String)
mPicture = b
RaiseEvent SetPicture(b)
End Property

Public Property Get pPicture() As String
pPicture = mPicture
End Property

Public Property Set zPicture(b As Picture)
Set imgMain.Picture = b
End Property

Private Sub imgMain_Click()
RaiseEvent Click
End Sub

Private Sub imgMain_DblClick()
RaiseEvent DblClick
End Sub

Private Sub imgMain_Initialize()
mVisible = True
Call UserControl.PropertyChanged("pVisible")
End Sub

Private Sub imgMain_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub imgMain_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub imgMain_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub imgMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = X / Screen.TwipsPerPixelX
Y = Y / Screen.TwipsPerPixelY
If Button = 1 Then
 mLMouseDown = True
 mLastX = X: mLastY = Y
 
ElseIf Button = 2 Then
 mRMouseDown = True
 mLastX = X: mLastY = Y
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = X / Screen.TwipsPerPixelX
Y = Y / Screen.TwipsPerPixelY
'Debug.Print X, Y, "G"
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = X / Screen.TwipsPerPixelX
Y = Y / Screen.TwipsPerPixelY
RaiseEvent MouseUp(Button, Shift, X, Y)
If Button = 1 Then
 mLMouseDown = False
 mLastX = 0: mLastY = 0
ElseIf Button = 2 Then
 mRMouseDown = False
 mLastX = 0: mLastY = 0
End If
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
mVisible = True
Call UserControl.PropertyChanged("pVisible")
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
'Debug.Print X, Y, "D"
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
mVisible = PropBag.ReadProperty("pVisible", True)
mLastX = PropBag.ReadProperty("pLastX", 0)
mLastY = PropBag.ReadProperty("pLastY", 0)
mLMouseDown = PropBag.ReadProperty("pLMouseDown", False)
mRMouseDown = PropBag.ReadProperty("pRMouseDown", False)
mPicture = PropBag.ReadProperty("pPicture", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("pVisible", mVisible, True)
Call PropBag.WriteProperty("pLastX", mLastX, 0)
Call PropBag.WriteProperty("pLastY", mLastY, 0)
Call PropBag.WriteProperty("pRMouseDown", mRMouseDown, False)
Call PropBag.WriteProperty("pLMouseDown", mLMouseDown, False)
Call PropBag.WriteProperty("pPicture", mPicture, "")
End Sub

Private Sub UserControl_Resize()
Dim w As Integer, h As Integer
w = (UserControl.Width / Screen.TwipsPerPixelX)
h = (UserControl.Height / Screen.TwipsPerPixelY)

xLine(0).x1 = 0
xLine(0).y1 = 0

xLine(0).x2 = w - 1
xLine(0).y2 = 0

xLine(1).x1 = w - 1
xLine(1).y1 = 0

xLine(1).x2 = w - 1
xLine(1).y2 = h - 1

xLine(2).x1 = 0
xLine(2).y1 = h - 1

xLine(2).x2 = w
xLine(2).y2 = h - 1

xLine(3).x1 = 0
xLine(3).y1 = 0

xLine(3).x2 = 0
xLine(3).y2 = h - 1

imgMain.Width = w
imgMain.Height = h
End Sub
