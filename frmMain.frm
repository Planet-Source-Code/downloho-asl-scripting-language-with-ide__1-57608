VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Window Editor"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWin 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   120
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.PictureBox picTool 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         FillColor       =   &H80000008&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   47
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox picLine 
         BackColor       =   &H00808080&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   30
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.PictureBox picLine 
         BackColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   30
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.PictureBox picLine 
         BackColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   30
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.PictureBox picLine 
         BackColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   30
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.PictureBox picBox 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   120
         Index           =   0
         Left            =   120
         MousePointer    =   8  'Size NW SE
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         Begin VB.PictureBox PicIn 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   90
            Index           =   0
            Left            =   15
            ScaleHeight     =   6
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   6
            TabIndex        =   16
            Top             =   15
            Width           =   90
         End
      End
      Begin VB.PictureBox picBox 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   120
         Index           =   1
         Left            =   240
         MousePointer    =   7  'Size N S
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         Begin VB.PictureBox PicIn 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   90
            Index           =   1
            Left            =   15
            ScaleHeight     =   6
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   6
            TabIndex        =   14
            Top             =   15
            Width           =   90
         End
      End
      Begin VB.PictureBox picBox 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   120
         Index           =   2
         Left            =   0
         MousePointer    =   6  'Size NE SW
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         Begin VB.PictureBox PicIn 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   90
            Index           =   2
            Left            =   15
            ScaleHeight     =   6
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   6
            TabIndex        =   12
            Top             =   15
            Width           =   90
         End
      End
      Begin VB.PictureBox picBox 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   120
         Index           =   3
         Left            =   0
         MousePointer    =   6  'Size NE SW
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         Begin VB.PictureBox PicIn 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   90
            Index           =   3
            Left            =   15
            ScaleHeight     =   6
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   6
            TabIndex        =   10
            Top             =   15
            Width           =   90
         End
      End
      Begin VB.PictureBox picBox 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   120
         Index           =   4
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         Begin VB.PictureBox PicIn 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   90
            Index           =   4
            Left            =   15
            ScaleHeight     =   6
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   6
            TabIndex        =   8
            Top             =   15
            Width           =   90
         End
      End
      Begin VB.PictureBox picBox 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   120
         Index           =   5
         Left            =   0
         MousePointer    =   8  'Size NW SE
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         Begin VB.PictureBox PicIn 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   90
            Index           =   5
            Left            =   15
            ScaleHeight     =   6
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   6
            TabIndex        =   6
            Top             =   15
            Width           =   90
         End
      End
      Begin VB.PictureBox picBox 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   120
         Index           =   6
         Left            =   0
         MousePointer    =   9  'Size W E
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         Begin VB.PictureBox PicIn 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   90
            Index           =   6
            Left            =   15
            ScaleHeight     =   6
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   6
            TabIndex        =   4
            Top             =   15
            Width           =   90
         End
      End
      Begin VB.PictureBox picBox 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   120
         Index           =   7
         Left            =   0
         MousePointer    =   9  'Size W E
         ScaleHeight     =   8
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   8
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   120
         Begin VB.PictureBox PicIn 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   90
            Index           =   7
            Left            =   15
            ScaleHeight     =   6
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   6
            TabIndex        =   2
            Top             =   15
            Width           =   90
         End
      End
      Begin VisualAce.xButton xButton1 
         Height          =   495
         Left            =   720
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         pText           =   "xButton1"
      End
      Begin VisualAce.xBox xBox1 
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         pStyle          =   1
         pText           =   "xBox1"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim selObj As Object, mLastX As Integer, mLastY As Integer

Public Sub BoxHide(ByVal b As Boolean, ByVal c As Boolean)
Dim i%

 For i% = 0 To 7
  picBox(i%).Visible = b
 Next i%

 For i% = 0 To 3
  picLine(i%).Visible = c
 Next i%
End Sub

Public Sub DrawFocus(ByRef obj As Object, ByVal b As Boolean)
Dim i%
Set selObj = obj
With obj
If b = True Then
 Call BoxHide(False, True)
 picLine(0).Left = .Left
 picLine(0).Top = .Top
 picLine(0).Height = .Height

 picLine(1).Left = .Left
 picLine(1).Top = .Top
 picLine(1).Width = .Width
 picLine(1).Height = 2
 
 picLine(2).Left = .Left + .Width - 2
 picLine(2).Top = .Top
 picLine(2).Height = .Height

 picLine(3).Left = .Left
 picLine(3).Top = .Top + .Height - 2
 picLine(3).Width = .Width
 picLine(3).Height = 2
Else
 Call BoxHide(True, False)
 picBox(0).Left = .Left - 8
 picBox(0).Top = .Top - 8

 picBox(1).Left = .Left + (.Width / 2) - 4
 picBox(1).Top = .Top - 8

 picBox(2).Left = .Left + .Width
 picBox(2).Top = .Top - 8

 picBox(3).Left = .Left - 8
 picBox(3).Top = .Top + .Height

 picBox(4).Left = .Left + (.Width / 2) - 4
 picBox(4).Top = .Top + .Height

 picBox(5).Left = .Left + .Width
 picBox(5).Top = .Top + .Height

 picBox(6).Left = .Left - 8
 picBox(6).Top = .Top + (.Height / 2) - 4

 picBox(7).Left = .Left + .Width
 picBox(7).Top = .Top + (.Height / 2) - 4
End If
End With
End Sub

Private Sub MoveObj(ByRef obj As Object, ByVal OffX As Integer, OffY As Integer)
If obj.pLMouseDown = True Then
 Dim pt As POINTAPI
 Call GetCursorPos(pt)

 With obj
  picLine(0).Left = (pt.X - (Me.Left / Screen.TwipsPerPixelX)) - .pLastX - picWin.Left - OffX
  picLine(1).Left = (pt.X - (Me.Left / Screen.TwipsPerPixelX)) - .pLastX - picWin.Left - OffX
  picLine(2).Left = ((pt.X - (Me.Left / Screen.TwipsPerPixelX)) - .pLastX) + .Width - picWin.Left - (OffX + 2)
  picLine(3).Left = (pt.X - (Me.Left / Screen.TwipsPerPixelX)) - .pLastX - picWin.Left - OffX

  picLine(0).Top = (pt.Y - ((Me.Top / Screen.TwipsPerPixelY) + 23) - (picWin.Top + OffY)) - .pLastY
  picLine(1).Top = (pt.Y - ((Me.Top / Screen.TwipsPerPixelY) + 23) - (picWin.Top + OffY)) - .pLastY
  picLine(2).Top = (pt.Y - ((Me.Top / Screen.TwipsPerPixelY) + 23) - (picWin.Top + OffY)) - .pLastY
  picLine(3).Top = ((pt.Y - ((Me.Top / Screen.TwipsPerPixelY) + 23) - (picWin.Top + (OffY + 2))) - .pLastY) + .Height

 End With
 Dim s$

   s$ = picLine(0).Left & "," & picLine(0).Top
   picTool.Visible = True
   picTool.Top = picLine(3).Top + 10
   picTool.Left = picLine(3).Left + picLine(3).Width - 10
   picTool.Cls
   picTool.CurrentX = (picTool.Width / 2) - (picTool.TextWidth(s$) / 2) - 2
   picTool.CurrentY = (picTool.Height / 2) - (picTool.TextHeight(s$) / 2)
   picTool.Print s$
   picTool.Refresh

End If
End Sub

Private Sub Form_Load()
Dim i%

 For i% = 0 To 7
  picBox(i%).Visible = False
 Next i%

Dim hMenu As Long

 Call ChangeWin(picWin.hwnd, , , , False)

 hMenu& = GetSystemMenu(picWin.hwnd, False)
 Call RemoveMenu(hMenu&, GetMenuItemCount(hMenu&) - 1, MF_BYPOSITION)
 Call RemoveMenu(hMenu, GetMenuItemCount(hMenu&) - 1, MF_BYPOSITION)

 Call SetText(picWin.hwnd, "Window1")
End Sub

Private Sub Form_Paint()
'picWin.Top = 6
'picWin.Left = 6
picWin.Left = 10
picWin.Top = 10
End Sub

Private Sub PicIn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 mLastX = X
 mLastY = Y
 Call DrawFocus(selObj, True)
End Sub

Private Sub PicIn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Dim pt As POINTAPI, iL%, iT%, s$
 Call GetCursorPos(pt)

 iL% = (pt.X - (Me.Left / Screen.TwipsPerPixelX)) - mLastX - picWin.Left - 10
 iT% = pt.Y - ((Me.Top / Screen.TwipsPerPixelY) + (23 * 2)) - picWin.Top - mLastY

  If iL% - selObj.Left <= 16 Then iL% = 16 + selObj.Left
  If iT% - selObj.Top <= 16 Then iT% = 16 + selObj.Top

Select Case Index%
 Case 55
   picLine(0).Left = iL%
   picLine(0).Height = iT% - picLine(2).Top

   picLine(2).Height = picLine(2).Height

   picLine(1).Top = iT%
   picLine(1).Width = picLine(3).Left - iL%

   picLine(3).Width = picLine(3).Width
 Case 4
   picLine(3).Top = iT%

   picLine(2).Height = iT% - selObj.Top
   picLine(0).Height = picLine(2).Height
   s$ = picLine(0).Height & "," & picLine(1).Width
 Case 5
   picLine(2).Left = iL%
   picLine(2).Height = iT% - picLine(2).Top

   picLine(0).Height = picLine(2).Height

   picLine(3).Top = iT%
   picLine(3).Width = iL% - picLine(3).Left

   picLine(1).Width = picLine(3).Width
   s$ = picLine(0).Height & "," & picLine(1).Width
 Case 7
   picLine(2).Left = iL%

   picLine(3).Width = iL% - picLine(3).Left

   picLine(1).Width = picLine(3).Width
   s$ = picLine(0).Height & "," & picLine(1).Width
End Select
   
   picTool.Visible = True
   picTool.Top = picLine(3).Top + 10
   picTool.Left = picLine(3).Left + picLine(3).Width - 10
   picTool.Cls
   picTool.CurrentX = (picTool.Width / 2) - (picTool.TextWidth(s$) / 2) - 2
   picTool.CurrentY = (picTool.Height / 2) - (picTool.TextHeight(s$) / 2)
   picTool.Print s$
   picTool.Refresh

picWin.Refresh
End Sub

Private Sub PicIn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call BoxHide(True, True)
selObj.Height = picLine(0).Height
selObj.Width = picLine(3).Width
Call DrawFocus(selObj, False)
End Sub

Private Sub picWin_Resize()
If CBool(IsZoomed(picWin.hwnd)) = True Then
 picWin.Width = Me.ScaleWidth - 20
 picWin.Height = Me.ScaleHeight - 20
End If
Call DrawDots(picWin)
End Sub

Private Sub xBox1_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(xBox1, KeyCode, Me)
End Sub

Private Sub xBox1_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(xBox1, KeyCode, Me)
End Sub

Private Sub xBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 Call DrawFocus(xBox1, True)
 Set selObj = xBox1
End If
End Sub

Private Sub xBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(xBox1, 10, 25)
End Sub

Private Sub xBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If xBox1.pLMouseDown = True Then Call xBox1.Move(picLine(0).Left, picLine(0).Top): picTool.Visible = False
 Call DrawFocus(xBox1, False)
End If
End Sub

Private Sub xButton1_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(xButton1, KeyCode, Me)
End Sub

Private Sub xButton1_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(xButton1, KeyCode, Me)
End Sub

Private Sub xButton1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 Call DrawFocus(xButton1, True)
 Set selObj = xButton1
End If
End Sub

Private Sub xButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(xButton1, 8, 23)
End Sub

Private Sub xButton1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If xButton1.pLMouseDown = True Then Call xButton1.Move(picLine(0).Left, picLine(0).Top): picTool.Visible = False
 Call DrawFocus(xButton1, False)
End If
End Sub
