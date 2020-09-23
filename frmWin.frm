VERSION 5.00
Begin VB.Form frmWin 
   BackColor       =   &H80000005&
   Caption         =   "Window Editor"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   Icon            =   "frmWin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   406
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTool 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1800
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   29
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picWin 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   120
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Tag             =   "Window1"
      Top             =   120
      Width           =   4575
      Begin VB.PictureBox picTool 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         FillColor       =   &H80000008&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   47
         TabIndex        =   28
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
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
      Begin VB.PictureBox picLine 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   30
         Index           =   3
         Left            =   120
         ScaleHeight     =   30
         ScaleWidth      =   495
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picLine 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   30
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.PictureBox picLine 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   30
         Index           =   1
         Left            =   120
         ScaleHeight     =   30
         ScaleWidth      =   495
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picLine 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
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
      Begin VisualAce.xButton cmdNew 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         pText           =   "Button"
      End
      Begin VisualAce.xBox memNew 
         Height          =   495
         Index           =   0
         Left            =   720
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         pStyle          =   2
         pText           =   "Memo"
      End
      Begin VisualAce.xBox txtNew 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         pText           =   "TextBox"
      End
      Begin VisualAce.xBox lstNew 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         pStyle          =   1
         pText           =   "ListBox"
      End
      Begin VisualAce.xOpt chkNew 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         pText           =   "CheckBox"
      End
      Begin VisualAce.xOpt optNew 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         pStyle          =   1
         pText           =   "Option"
      End
      Begin VisualAce.xBox lblNew 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   30
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         pStyle          =   3
         pText           =   "Label"
      End
      Begin VisualAce.xBox cmbNew 
         Height          =   300
         Index           =   0
         Left            =   1320
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         pStyle          =   4
         pLimitHeight    =   20
         pText           =   "ComboBox"
      End
      Begin VisualAce.xImage imgNew 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
      End
      Begin VisualAce.xDialog mnuNew 
         Height          =   510
         Index           =   0
         Left            =   2400
         TabIndex        =   33
         Top             =   1800
         Visible         =   0   'False
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   900
         pStyle          =   1
      End
      Begin VisualAce.xDialog tmrNew 
         Height          =   510
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   900
      End
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public xX%, xY%
Dim mLastX As Integer, mLastY As Integer, mIsCreate As Boolean

Private Function FindEdit() As Form
Dim frm As Form
 For Each frm In Forms()
  If frm.Name = "frmEdit" And InStr(frm.Tag, Chr(0)) <> 0 Then
   If Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = Left$(Tag, InStr(Tag, Chr(0)) - 1) Then Exit For
  End If
 Next frm
 Set FindEdit = frm
End Function

Public Sub onLoad()
Dim i%

 For i% = 0 To 7
  picBox(i%).Visible = False
 Next i%

Dim hMenu As Long

 Call ChangeWin(picWin.hwnd, , , , False)
 Call FlashWindow(picWin.hwnd, 1)

 hMenu& = GetSystemMenu(picWin.hwnd, False)
 Call RemoveMenu(hMenu&, GetMenuItemCount(hMenu&) - 1, MF_BYPOSITION)
 Call RemoveMenu(hMenu, GetMenuItemCount(hMenu&) - 1, MF_BYPOSITION)
'MsgBox Me.Tag
 Call SetText(picWin.hwnd, Left$(Tag, InStr(Tag, Chr(0)) - 1))
 picWin.Tag = Left$(Tag, InStr(Tag, Chr(0)) - 1)
End Sub

Public Sub BoxHide(ByVal b As Boolean, ByVal c As Boolean)
Dim i%

' Call mdiMain.PropList.Clear

 For i% = 0 To 7
  picBox(i%).Visible = b
 Next i%

 For i% = 0 To 3
  picLine(i%).Visible = c
 Next i%

picTool(0).Visible = False
picTool(1).Visible = False
End Sub

Public Sub DrawFocus(ByRef obj As Object, ByVal b As Boolean)
Dim i%
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
picTool(0).Visible = False
picTool(1).Visible = False
End Sub

Private Sub MoveObj(ByRef obj As Object, ByVal OffX As Integer, OffY As Integer)

If obj.pLMouseDown = True Then
 Dim pt As POINTAPI, i%, lC(3) As Long
 
 Call GetCursorPos(pt)
 With obj
  lC(0) = (pt.X - (mdiMain.Left / Screen.TwipsPerPixelX + (mdiMain.picLeft.Width / Screen.TwipsPerPixelX) + 6) - (Me.Left / Screen.TwipsPerPixelX)) - .pLastX - picWin.Left - OffX
  lC(1) = ((pt.X - (mdiMain.Left / Screen.TwipsPerPixelX + (mdiMain.picLeft.Width / Screen.TwipsPerPixelX) + 6) - (Me.Left / Screen.TwipsPerPixelX)) - .pLastX) + .Width - picWin.Left - (OffX + 2)
  lC(2) = (pt.Y - (((mdiMain.Top / Screen.TwipsPerPixelY) + (mdiMain.tbMain.Height / Screen.TwipsPerPixelY) - 2) - (Me.Top / Screen.TwipsPerPixelY) + 23) - (picWin.Top + OffY)) - .pLastY
  lC(3) = ((pt.Y - (((mdiMain.Top / Screen.TwipsPerPixelY) + (mdiMain.tbMain.Height / Screen.TwipsPerPixelY) - 2) - (Me.Top / Screen.TwipsPerPixelY) + 23) - (picWin.Top + (OffY + 2))) - .pLastY) + .Height

  picLine(0).Left = lC(0)
  picLine(1).Left = lC(0)
  picLine(2).Left = lC(1)
  picLine(3).Left = lC(0)

  picLine(0).Top = lC(2)
  picLine(1).Top = lC(2)
  picLine(2).Top = lC(2)
  picLine(3).Top = lC(3)

 End With
 Dim s$

   s$ = picLine(0).Left & "," & picLine(0).Top
   i% = 0
    picTool(i%).Visible = True
    picTool(i%).Top = picLine(3).Top + picWin.Top + 30
    picTool(i%).Left = picLine(3).Left + picLine(3).Width + picWin.Left + 10
    picTool(i%).Cls
    picTool(i%).CurrentX = (picTool(i%).Width / 2) - (picTool(i%).TextWidth(s$) / 2) - 2
    picTool(i%).CurrentY = (picTool(i%).Height / 2) - (picTool(i%).TextHeight(s$) / 2)
    picTool(i%).Print s$
    picTool(i%).Refresh

   i% = 1
    picTool(i%).Visible = True
    picTool(i%).Top = picLine(3).Top + picWin.Top + 63
    picTool(i%).Left = picLine(3).Left + picLine(3).Width + picWin.Left + 24
    picTool(i%).Cls
    picTool(i%).CurrentX = (picTool(i%).Width / 2) - (picTool(i%).TextWidth(s$) / 2) - 2
    picTool(i%).CurrentY = (picTool(i%).Height / 2) - (picTool(i%).TextHeight(s$) / 2)
    picTool(i%).Print s$
    picTool(i%).Refresh


End If
End Sub

Public Sub ResetWin()
Dim butX As Button
 If picWin.MousePointer = ccCross Then picWin.MousePointer = ccDefault
 objNew = eNULL

  For Each butX In mdiMain.tbTool.Buttons()
   butX.Value = tbrUnpressed
  Next butX
  mdiMain.tbTool.Refresh
  mdiMain.tbTool.Buttons(1).Value = tbrPressed
  picWin.Cls
End Sub

Public Sub SetProp(ByRef obj As Object, Optional bHeight As Boolean = True, Optional bNULL As Boolean = True, Optional bVis As Boolean = True, Optional bWidth As Boolean = True, Optional bEnabled As Boolean = True, Optional bLeft As Boolean = True, Optional bTop As Boolean = True, Optional bInterval As Boolean = False, Optional bText As Boolean = False, Optional bValue As Boolean = False, Optional bPicture As Boolean = False)
mSetProp = True
Call mdiMain.PropList.Clear
With obj

 Call mdiMain.PropList.Add("Name", .Tag)
 If bEnabled = True Then Call mdiMain.PropList.Add("Enabled", .pEnabled)
 If bInterval = True Then Call mdiMain.PropList.Add("Interval", .pInterval)
 If bLeft = True Then Call mdiMain.PropList.Add("Left", .Left)
 If bHeight = True Then Call mdiMain.PropList.Add("Height", .Height)
 If bPicture = True Then Call mdiMain.PropList.Add("Picture", .pPicture)
 If bText = True Then
  If obj.Name <> "picWin" Then
   Call mdiMain.PropList.Add("Text", .pText)
  Else
   Call mdiMain.PropList.Add("Text", GetCaption$(picWin.hwnd))
  End If
 End If

 If bTop = True Then Call mdiMain.PropList.Add("Top", .Top)
 If bValue = True Then Call mdiMain.PropList.Add("Value", .pValue)
 If bVis = True Then Call mdiMain.PropList.Add("Visible", .pVisible)
 If bWidth = True Then Call mdiMain.PropList.Add("Width", .Width)
 If obj.Name <> "picWin" And obj.Name <> "tmrNew" And obj.Name <> "imgNew" And obj.Name <> "mnuNew" Then Call mdiMain.PropList.Add("BackColor", IIf(Left$(.pBackColor, 1) = "-", "<default>", Rgb2Html(.pBackColor)))
 If obj.Name <> "picWin" And obj.Name <> "tmrNew" And obj.Name <> "imgNew" And obj.Name <> "mnuNew" Then Call mdiMain.PropList.Add("ForeColor", IIf(Left$(.pForeColor, 1) = "-", "<default>", Rgb2Html(.pForeColor)))
End With
Call mdiMain.PropList.Update
mSetProp = False
End Sub

Private Sub sqaurebox(obj As Object, X, Y)
 obj.Cls
 obj.DrawWidth = 2 ': obj.DrawStyle = 2
  obj.Line (mLastX, mLastY)-(CInt(X), CInt(Y)), &H808080, B
 obj.DrawWidth = 1
 'obj.DrawStyle = 0
End Sub

Private Sub chkNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = chkNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub chkNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & chkNew(Index%).Tag & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & chkNew(Index%).Tag & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
 Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub chkNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(chkNew(Index%), KeyCode, Me)
End Sub

Private Sub chkNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(chkNew(Index%), KeyCode, Me)
End Sub

Private Sub chkNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then

 Call DrawFocus(chkNew(Index%), True)
End If
End Sub

Private Sub chkNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(chkNew(Index%), 8, 23)
End Sub

Private Sub chkNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If chkNew(Index%).pLMouseDown = True Then Call chkNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = chkNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(chkNew(Index%), False)
 Call SetProp(chkNew(Index%), , , , , , , , , True, True)
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub

Private Sub cmbNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = cmbNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub cmbNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & cmbNew(Index%).Tag & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & cmbNew(Index%).Tag & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub cmbNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(cmbNew(Index%), KeyCode, Me)
End Sub

Private Sub cmbNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(cmbNew(Index%), KeyCode, Me)
End Sub

Private Sub cmbNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then

 Call DrawFocus(cmbNew(Index%), True)
End If
End Sub

Private Sub cmbNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(cmbNew(Index%), 10, 25)
End Sub

Private Sub cmbNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If cmbNew(Index%).pLMouseDown = True Then Call cmbNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = cmbNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(cmbNew(Index%), False)
 Call SetProp(cmbNew(Index%), False, , , , , , , , True)
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub

Private Sub cmdNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = cmdNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub cmdNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & cmdNew(Index%).Tag & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & cmdNew(Index%).Tag & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub cmdNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(cmdNew(Index%), KeyCode, Me)
End Sub

Private Sub cmdNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(cmdNew(Index%), KeyCode, Me)
End Sub

Private Sub cmdNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then

 Call DrawFocus(cmdNew(Index%), True)
End If
End Sub

Private Sub cmdNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(cmdNew(Index%), 8, 23)
End Sub

Private Sub cmdNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If cmdNew(Index%).pLMouseDown = True Then Call cmdNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = cmdNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(cmdNew(Index%), False)
 Call SetProp(cmdNew(Index%), , , , , , , , , True)
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub

Private Sub Form_Click()
Set gblSelWinObj = picWin
Call BoxHide(False, False)
End Sub

Private Sub Form_GotFocus()
Set gblSelWinObj = picWin
mdiMain.tbTool.Enabled = True
End Sub

Private Sub Form_LostFocus()
mdiMain.tbTool.Enabled = False
End Sub

Private Sub Form_Paint()
picWin.Left = 10
picWin.Top = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set gblSelWinObj = gblNotObj
Set gblSelObj = gblNotObj
If gblCanClose = False Then Cancel = -1: Me.Hide
mdiMain.cmbCon.Clear
mdiMain.PropList.Clear
mdiMain.PropList.Update
End Sub

Private Sub imgNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = imgNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub imgNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & imgNew(Index%).Tag & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & imgNew(Index%).Tag & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub imgNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(imgNew(Index%), KeyCode, Me)
End Sub

Private Sub imgNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(imgNew(Index%), KeyCode, Me)
End Sub

Private Sub imgNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then

 Call DrawFocus(imgNew(Index%), True)
End If
End Sub

Private Sub imgNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(imgNew(Index%), 8, 23)
End Sub

Private Sub imgNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If imgNew(Index%).pLMouseDown = True Then Call imgNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = imgNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(imgNew(Index%), False)
 Call SetProp(imgNew(Index%), , , , , False, , , , , , True)
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub

Private Sub imgNew_SetPicture(Index As Integer, ByVal pName As String)
Dim frm As Form
For Each frm In Forms()
 If frm.Name = "frmImg" And frm.Tag <> "" Then
  If LCase(Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1)) = LCase(pName) Then Set imgNew(Index).zPicture = frm.imgMain.Picture: Exit Sub
 End If
Next frm
End Sub

Private Sub lblNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = lblNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub lblNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & lblNew(Index%).Tag & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & lblNew(Index%).Tag & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub lblNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(lblNew(Index%), KeyCode, Me)
End Sub

Private Sub lblNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(lblNew(Index%), KeyCode, Me)
End Sub

Private Sub lblNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then

 Call DrawFocus(lblNew(Index%), True)
End If
End Sub

Private Sub lblNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(lblNew(Index%), 8, 23)
End Sub

Private Sub lblNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If lblNew(Index%).pLMouseDown = True Then Call lblNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = lblNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(lblNew(Index%), False)
 Call SetProp(lblNew(Index%), , , , , False, , , , True)
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub

Private Sub lstNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = lstNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub lstNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & lstNew(Index%).Tag & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & lstNew(Index%).Tag & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub lstNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(lstNew(Index%), KeyCode, Me)
End Sub

Private Sub lstNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(lstNew(Index%), KeyCode, Me)
End Sub

Private Sub lstNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 Call DrawFocus(lstNew(Index%), True)
End If
End Sub

Private Sub lstNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(lstNew(Index%), 10, 25)
End Sub

Private Sub lstNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If lstNew(Index%).pLMouseDown = True Then Call lstNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = lstNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(lstNew(Index%), False)
 Call SetProp(lstNew(Index%))
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub

Private Sub memNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = memNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub memNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & memNew(Index%).Tag & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & memNew(Index%).Tag & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub memNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(memNew(Index%), KeyCode, Me)
End Sub

Private Sub memNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(memNew(Index%), KeyCode, Me)
End Sub

Private Sub memNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 Call DrawFocus(memNew(Index%), True)
End If
End Sub

Private Sub memNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(memNew(Index%), 10, 25)
End Sub

Private Sub memNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If memNew(Index%).pLMouseDown = True Then Call memNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = memNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(memNew(Index%), False)
 Call SetProp(memNew(Index%), , , , , , , , , True)
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub

Private Sub mnuNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = mnuNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub mnuNew_DblClick(Index As Integer)
frmMenu.b_Sender = Left(Tag, InStr(Tag, Chr(0)) - 1)
frmMenu.Show vbModal
frmMenu.b_Sender = ""
End Sub

Private Sub mnuNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(mnuNew(Index%), KeyCode, Me)
End Sub

Private Sub mnuNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(mnuNew(Index%), KeyCode, Me)
End Sub

Private Sub mnuNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 Call DrawFocus(mnuNew(Index%), True)
End If
End Sub

Private Sub mnuNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(mnuNew(Index%), 9, 24)
End Sub

Private Sub mnuNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If mnuNew(Index%).pLMouseDown = True Then Call mnuNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = mnuNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(mnuNew(Index%), False)
 Call SetProp(mnuNew(Index%), False, False, False, False, , , , True)
 picTool(0).Visible = False
 picTool(1).Visible = False
ElseIf Button = 2 Then
 Call frmMenu.ShowMenu(picWin.Tag)
End If
End Sub

Private Sub optNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = optNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub optNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & optNew(Index%).Tag & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & optNew(Index%).Tag & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub optNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(optNew(Index%), KeyCode, Me)
End Sub

Private Sub optNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(optNew(Index%), KeyCode, Me)
End Sub

Private Sub optNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then

 Call DrawFocus(optNew(Index%), True)
End If
End Sub

Private Sub optNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(optNew(Index%), 8, 23)
End Sub

Private Sub optNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If optNew(Index%).pLMouseDown = True Then Call optNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = optNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(optNew(Index%), False)
 Call SetProp(optNew(Index%), , , , , , , , , True, True)
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub


Private Sub picBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicIn_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub picBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicIn_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub picBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PicIn_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub PicIn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 mLastX = X
 mLastY = Y
 Call DrawFocus(gblSelObj, True)
End Sub

Private Sub PicIn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Dim pt As POINTAPI, iL%, iT%, s$
 Call GetCursorPos(pt)

 iL% = (pt.X - (mdiMain.Left / Screen.TwipsPerPixelX + (mdiMain.picLeft.Width / Screen.TwipsPerPixelX) + 5) - (Me.Left / Screen.TwipsPerPixelX)) - mLastX - picWin.Left - 10
 iT% = pt.Y - ((mdiMain.Top / Screen.TwipsPerPixelY) + (20 * 2) + (mdiMain.tbMain.Height / Screen.TwipsPerPixelY)) - ((Me.Top / Screen.TwipsPerPixelY) + (25.5 * 2)) - picWin.Top - mLastY

Select Case Index%
 Case 0
  iL% = iL% + 8
  If iL% >= (gblSelObj.Left + gblSelObj.Width) - 16 Then iL% = gblSelObj.Left + gblSelObj.Width - 16
  iT% = iT% + 8
  If iT% >= (gblSelObj.Top + gblSelObj.Height) - 16 Then iT% = gblSelObj.Top + gblSelObj.Height - 16
 Case 1, 2
  If iL% - gblSelObj.Left <= 16 Then iL% = 16 + gblSelObj.Left
  iT% = iT% + 8
  If iT% >= (gblSelObj.Top + gblSelObj.Height) - 16 Then iT% = gblSelObj.Top + gblSelObj.Height - 16
 Case 3, 6
  iL% = iL% + 8
  If iL% >= (gblSelObj.Left + gblSelObj.Width) - 16 Then iL% = gblSelObj.Left + gblSelObj.Width - 16
  If iT% - gblSelObj.Top <= 16 Then iT% = 16 + gblSelObj.Top
 Case 4, 5, 7
  If iL% - gblSelObj.Left <= 16 Then iL% = 16 + gblSelObj.Left
  If iT% - gblSelObj.Top <= 16 Then iT% = 16 + gblSelObj.Top
End Select

Select Case Index%
 Case 0
 
   picLine(2).Top = iT%
   picLine(2).Height = gblSelObj.Top + gblSelObj.Height - iT%
   
   picLine(0).Left = iL%
   picLine(0).Top = iT%
   picLine(0).Height = gblSelObj.Top + gblSelObj.Height - iT%
   
   picLine(3).Left = iL%
   picLine(3).Width = (gblSelObj.Left + gblSelObj.Width) - iL%

   picLine(1).Left = iL%
   picLine(1).Top = iT%
   picLine(1).Width = (gblSelObj.Left + gblSelObj.Width) - iL%
   
   s$ = picLine(3).Width & "," & picLine(0).Height
 Case 1
   picLine(0).Top = iT%
   picLine(1).Top = iT%
   picLine(2).Top = iT%

   picLine(2).Height = gblSelObj.Top + gblSelObj.Height - iT%
   picLine(0).Height = picLine(2).Height

   s$ = picLine(3).Width & "," & picLine(2).Height
 Case 2
   picLine(0).Top = iT%
   picLine(1).Top = iT%
   picLine(2).Top = iT%

   picLine(2).Height = gblSelObj.Top + gblSelObj.Height - iT%
   picLine(0).Height = picLine(2).Height
   
   picLine(3).Width = iL% - picLine(3).Left
   picLine(1).Width = picLine(3).Width

   picLine(2).Left = iL%

   s$ = picLine(3).Width & "," & picLine(2).Height
 Case 3
   picLine(0).Left = iL%

   picLine(3).Left = iL%
   picLine(3).Width = (gblSelObj.Left + gblSelObj.Width) - iL%

   picLine(1).Left = iL%
   picLine(1).Width = picLine(3).Width

   picLine(3).Top = iT%

   picLine(2).Height = iT% - gblSelObj.Top
   picLine(0).Height = picLine(2).Height

   s$ = picLine(3).Width & "," & picLine(2).Height
 Case 4
   picLine(3).Top = iT%

   picLine(2).Height = iT% - gblSelObj.Top
   picLine(0).Height = picLine(2).Height
   s$ = picLine(1).Width & "," & picLine(2).Height
 Case 5
   picLine(2).Left = iL%
   picLine(2).Height = iT% - picLine(2).Top

   picLine(0).Height = picLine(2).Height

   picLine(3).Top = iT%
   picLine(3).Width = iL% - picLine(3).Left

   picLine(1).Width = picLine(3).Width
   s$ = picLine(3).Width & "," & picLine(0).Height
 Case 6
   picLine(0).Left = iL%

   picLine(3).Left = iL%
   picLine(3).Width = (gblSelObj.Left + gblSelObj.Width) - iL%

   picLine(1).Left = iL%
   picLine(1).Width = picLine(3).Width

   s$ = picLine(3).Width & "," & picLine(0).Height
 Case 7
   picLine(2).Left = iL%

   picLine(3).Width = iL% - picLine(3).Left

   picLine(1).Width = picLine(3).Width
   s$ = picLine(3).Width & "," & picLine(2).Height
End Select
   
   For iL% = 0 To 1
    picTool(iL%).Visible = True
    picTool(iL%).Top = picLine(3).Top + picWin.Top + IIf(iL% = 0, 30, 63)
    picTool(iL%).Left = picLine(3).Left + picLine(3).Width + picWin.Left + IIf(iL% = 0, 10, 24)
    picTool(iL%).Cls
    picTool(iL%).CurrentX = (picTool(iL%).Width / 2) - (picTool(iL%).TextWidth(s$) / 2) - 2
    picTool(iL%).CurrentY = (picTool(iL%).Height / 2) - (picTool(iL%).TextHeight(s$) / 2)
    picTool(iL%).Print s$
    picTool(iL%).Refresh
   Next iL%

picWin.Refresh
End Sub

Private Sub PicIn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call BoxHide(True, True)
Call gblSelObj.Move(picLine(0).Left, picLine(0).Top, picLine(3).Width, picLine(0).Height)
Call DrawFocus(gblSelObj, False)
 Dim b As Boolean
Select Case gblSelObj.Name
 Case "cmdNew", "txtNew", "memNew", "cmbNew", "chkNew", "optNew"
  b = True
 Case Else
  b = False
End Select

 Dim jk As Boolean, g As Boolean, h As Boolean
 g = True
 If gblSelObj.Name = "imgNew" Then g = False: h = True
 If gblSelObj.Name = "chkNew" Or gblSelObj.Name = "optNew" Then jk = True
Call SetProp(gblSelObj, , , , , g, , , , b, jk, h)
picTool(0).Visible = False
picTool(1).Visible = False
End Sub

Private Sub picWin_Click()
mdiMain.tbTool.Enabled = True
Dim con As Control, i%

If mIsCreate = False Then
 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
  End If
 Next con
 
 mSetProp = True
  mdiMain.cmbCon.ListIndex = 0
 mSetProp = False

For i = 1 To mdiMain.tvwFiles.Nodes.Count
 If mdiMain.tvwFiles.Nodes(i).Text = Left(Tag, InStr(Tag, Chr(0)) - 1) Then mdiMain.tvwFiles.Nodes(i).Selected = True: Exit For
Next i
mdiMain.tbProp.Buttons(1).Enabled = True

 Set gblSelWinObj = picWin
 Set gblSelObj = picWin
 Call SetProp(picWin, , , False, , False, False, False, , True)
Else
 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = gblSelObj.Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
 mIsCreate = False
End If
End Sub

Private Sub picWin_DblClick()
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc " & picWin.Tag & "_Init()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc " & picWin.Tag & "_Init()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub picWin_GotFocus()
Set gblSelWinObj = picWin
End Sub

Private Sub picWin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picWin.MousePointer = ccCross Then
 mLastX = X: mLastY = Y
  Call sqaurebox(picWin, X, Y)
 Exit Sub
End If
End Sub

Private Sub picWin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picWin.MousePointer = ccCross Then
  If Button = 1 Then Call sqaurebox(picWin, X, Y)
 Exit Sub
End If
End Sub

Private Sub picWin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Call ResetWin: Exit Sub

  Set gblSelWinObj = picWin
 Select Case objNew
  Case eButton
   mIsCreate = True
   Call MakeButton(Me, IIf(X <= mLastX, X, mLastX), IIf(Y <= mLastY, Y, mLastY), IIf(X < mLastX, mLastX - X, X - mLastX), IIf(Y < mLastY, mLastY - Y, Y - mLastY))
  Case eTextBox, eListBox, eMemo, eLabel, eComboBox
   mIsCreate = True
   Call MakeBox(Me, objNew, IIf(X <= mLastX, X, mLastX), IIf(Y <= mLastY, Y, mLastY), IIf(X < mLastX, mLastX - X, X - mLastX), IIf(Y < mLastY, mLastY - Y, Y - mLastY))
  Case eTimer, eMenu
   mIsCreate = True
   Call MakeDial(Me, objNew, IIf(X <= mLastX, X, mLastX), IIf(Y <= mLastY, Y, mLastY), IIf(X < mLastX, mLastX - X, X - mLastX), IIf(Y < mLastY, mLastY - Y, Y - mLastY))
  Case eCheckBox, eOption
   mIsCreate = True
   Call MakeOpt(Me, objNew, IIf(X <= mLastX, X, mLastX), IIf(Y <= mLastY, Y, mLastY), IIf(X < mLastX, mLastX - X, X - mLastX), IIf(Y < mLastY, mLastY - Y, Y - mLastY))
  Case eImage
   mIsCreate = True
   Call MakeImg(Me, objNew, IIf(X <= mLastX, X, mLastX), IIf(Y <= mLastY, Y, mLastY), IIf(X < mLastX, mLastX - X, X - mLastX), IIf(Y < mLastY, mLastY - Y, Y - mLastY))
  Case Else
   Call BoxHide(False, False)
 End Select

Call ResetWin
End Sub

Private Sub picWin_Resize()

 If picWin.Left <> 10 Then picWin.Left = 10
 If picWin.Top <> 10 Then picWin.Top = 10

If CBool(IsZoomed(picWin.hwnd)) = True Then
 picWin.Width = Me.ScaleWidth - 20
 picWin.Height = Me.ScaleHeight - 20
End If
Call DrawDots(picWin)
End Sub

Private Sub tmrNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = tmrNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub tmrNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & tmrNew(Index%).Tag & "_Timer()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & tmrNew(Index%).Tag & "_Timer()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub tmrNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(tmrNew(Index%), KeyCode, Me)
End Sub

Private Sub tmrNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(tmrNew(Index%), KeyCode, Me)
End Sub

Private Sub tmrNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 Call DrawFocus(tmrNew(Index%), True)
End If
End Sub

Private Sub tmrNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(tmrNew(Index%), 9, 24)
End Sub

Private Sub tmrNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If tmrNew(Index%).pLMouseDown = True Then Call tmrNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = tmrNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(tmrNew(Index%), False)
 Call SetProp(tmrNew(Index%), False, False, False, False, , , , True)
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub

Private Sub txtNew_Click(Index As Integer)
Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = txtNew(Index%).Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con
End Sub

Private Sub txtNew_DblClick(Index As Integer)
Dim frm As Form, i As Integer
Set frm = FindEdit

i = InStr("DD" & frm.rtbEdit.Text, "!proc ^" & txtNew(Index%).Tag & "_Click()" & vbCrLf)
If i = 0 Then
 Call modColor.PrintText(vbCrLf & "<font face=""Courier New"">!proc ^" & txtNew(Index%).Tag & "_Click()" & vbCrLf & vbCrLf & "end!", frm.rtbEdit, frm.rtbx)
 frm.rtbEdit.SelStart = Len(frm.rtbEdit.Text) - 7
Else
 frm.rtbEdit.SelStart = InStr(i + 1, frm.rtbEdit.Text, vbCrLf) + 1
End If
Call ShowWindow(frm.hwnd, 3)
End Sub

Private Sub txtNew_KeyDown(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyDown(txtNew(Index%), KeyCode, Me)
End Sub

Private Sub txtNew_KeyUp(Index As Integer, ByVal KeyCode As Integer, ByVal Shift As Integer)
Call OnKeyUp(txtNew(Index%), KeyCode, Me)
End Sub

Private Sub txtNew_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 Call DrawFocus(txtNew(Index%), True)
End If
End Sub

Private Sub txtNew_MouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call MoveObj(txtNew(Index%), 10, 25)
End Sub

Private Sub txtNew_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 1 Then
 If txtNew(Index%).pLMouseDown = True Then Call txtNew(Index%).Move(picLine(0).Left, picLine(0).Top)

  Set gblSelObj = txtNew(Index%)
  Set gblSelWinObj = picWin

 Call DrawFocus(txtNew(Index%), False)
 Call SetProp(txtNew(Index%), , , , , , , , , True)
 picTool(0).Visible = False
 picTool(1).Visible = False
End If
End Sub
