VERSION 5.00
Begin VB.UserControl xButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ToolboxBitmap   =   "xButton.ctx":0000
   Begin VB.Line tLine 
      BorderColor     =   &H80000010&
      Index           =   7
      X1              =   56
      X2              =   168
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line tLine 
      BorderColor     =   &H80000016&
      BorderWidth     =   2
      Index           =   6
      X1              =   56
      X2              =   168
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line tLine 
      BorderColor     =   &H80000016&
      BorderWidth     =   2
      Index           =   5
      X1              =   168
      X2              =   168
      Y1              =   104
      Y2              =   24
   End
   Begin VB.Line tLine 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   152
      X2              =   152
      Y1              =   96
      Y2              =   24
   End
   Begin VB.Line tLine 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Index           =   3
      X1              =   56
      X2              =   128
      Y1              =   32
      Y2              =   32
   End
   Begin VB.Line tLine 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      Index           =   2
      X1              =   56
      X2              =   136
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Line tLine 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Index           =   1
      X1              =   41
      X2              =   41
      Y1              =   104
      Y2              =   24
   End
   Begin VB.Line tLine 
      BorderColor     =   &H80000015&
      BorderWidth     =   2
      Index           =   0
      X1              =   56
      X2              =   56
      Y1              =   112
      Y2              =   24
   End
End
Attribute VB_Name = "xButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
 X As Long
 Y As Long
End Type

Event Click()
Event DblClick()
Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event KeyPress(ByVal KeyAscii As Integer)
Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim mEnabled As Boolean, mLMouseDown As Boolean, mRMouseDown As Boolean
Dim mLastX As Integer, mLastY As Integer, mText As String
Dim mVisible As Boolean, DefaultCl(1) As Long

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
End Property

Public Property Get pForeColor() As Long
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

Private Sub MakeButton(ByVal Pressed As Boolean)
Dim i%
For i% = 0 To tLine.Count - 1
 tLine(i%).Visible = True
Next i%
If Pressed = False Then
    With tLine(5)
     .x1 = 0
     .y1 = 0
     .x2 = 0
     .y2 = Height / Screen.TwipsPerPixelY
    End With
    With tLine(1)
     .x1 = 1
     .y1 = 0
     .x2 = 1
     .y2 = Height / Screen.TwipsPerPixelY
    End With
    With tLine(6)
     .x1 = 0
     .y1 = 0
     .x2 = Width / Screen.TwipsPerPixelX
     .y2 = 0
    End With
    With tLine(3)
     .x1 = 0
     .y1 = 1
     .x2 = Width / Screen.TwipsPerPixelX
     .y2 = 1
    End With
    With tLine(0)
     .x1 = Width / Screen.TwipsPerPixelX - 1
     .y1 = 0
     .x2 = Width / Screen.TwipsPerPixelX - 1
     .y2 = Height / Screen.TwipsPerPixelY
    End With
    With tLine(4)
     .x1 = (Width - 30) / Screen.TwipsPerPixelX
     .y1 = 0
     .x2 = (Width - 30) / Screen.TwipsPerPixelX
     .y2 = Height / Screen.TwipsPerPixelY
    End With
    With tLine(2)
     .x1 = 0
     .y1 = Height / Screen.TwipsPerPixelY - 1
     .x2 = Width / Screen.TwipsPerPixelX
     .y2 = Height / Screen.TwipsPerPixelY - 1
    End With
    With tLine(7)
     .x1 = 0
     .y1 = (Height - 30) / Screen.TwipsPerPixelY
     .x2 = Width / Screen.TwipsPerPixelX
     .y2 = (Height - 30) / Screen.TwipsPerPixelY
    End With
Cls
Call PrintTxt
Else
    With tLine(5)
     .x1 = Width / Screen.TwipsPerPixelX - 1
     .y1 = 0
     .x2 = Width / Screen.TwipsPerPixelX - 1
     .y2 = Height / Screen.TwipsPerPixelY
    End With
    With tLine(1)
     .x1 = Width / Screen.TwipsPerPixelX - 1
     .y1 = 0
     .x2 = Width / Screen.TwipsPerPixelX - 1
     .y2 = Height / Screen.TwipsPerPixelY
    End With
    With tLine(6)
     .x1 = 0
     .y1 = Height / Screen.TwipsPerPixelY - 1
     .x2 = Width / Screen.TwipsPerPixelX
     .y2 = Height / Screen.TwipsPerPixelY - 1
     End With
    With tLine(3)
     .x1 = 0
     .y1 = Height / Screen.TwipsPerPixelY - 1
     .x2 = Width / Screen.TwipsPerPixelX
     .y2 = Height / Screen.TwipsPerPixelY - 1
    End With '#
    With tLine(0)
     .x1 = 0
     .y1 = 0
     .x2 = 0
     .y2 = Height / Screen.TwipsPerPixelY
    End With
    With tLine(4)
     .x1 = 1
     .y1 = 0
     .x2 = 1
     .y2 = Height / Screen.TwipsPerPixelY
    End With
    With tLine(2)
     .x1 = 0
     .y1 = 0
     .x2 = Width / Screen.TwipsPerPixelX
     .y2 = 0
    End With
    With tLine(7)
     .x1 = 0
     .y1 = 1
     .x2 = Width / Screen.TwipsPerPixelX
     .y2 = 1
    End With
Cls
Call PrintTxt(2)
End If
End Sub

Private Sub PrintTxt(Optional ByVal OffSet As Integer = 0)
Dim s$
Cls
s$ = mText
CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(s$) / 2) - 2 + OffSet%
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
Call PrintTxt
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
mVisible = PropBag.ReadProperty("pVisible", True)
mLastX = PropBag.ReadProperty("pLastX", 0)
mLastY = PropBag.ReadProperty("pLastY", 0)
mLMouseDown = PropBag.ReadProperty("pLMouseDown", False)
mRMouseDown = PropBag.ReadProperty("pRMouseDown", False)
mText = PropBag.ReadProperty("pText", UserControl.Name)
Call PrintTxt
End Sub

Private Sub UserControl_Resize()
Call MakeButton(False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("pEnabled", mEnabled, True)
Call PropBag.WriteProperty("pVisible", mVisible, True)
Call PropBag.WriteProperty("pLastX", mLastX, 0)
Call PropBag.WriteProperty("pLastY", mLastY, 0)
Call PropBag.WriteProperty("pRMouseDown", mRMouseDown, False)
Call PropBag.WriteProperty("pLMouseDown", mLMouseDown, False)
Call PropBag.WriteProperty("pText", mText, UserControl.Name)
End Sub
