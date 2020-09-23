VERSION 5.00
Begin VB.UserControl PropList 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2580
   ForeColor       =   &H80000016&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   172
   Begin VB.VScrollBar vs 
      Height          =   3255
      LargeChange     =   32
      Left            =   2400
      Max             =   0
      SmallChange     =   16
      TabIndex        =   4
      Top             =   240
      Width           =   135
   End
   Begin VisualAce.xButton colHead 
      Height          =   240
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   423
      pText           =   "Value"
   End
   Begin VisualAce.xButton colHead 
      Height          =   240
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   423
      pText           =   "Property"
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000016&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   2
      Top             =   240
      Width           =   2415
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   1230
         TabIndex        =   3
         Top             =   15
         Visible         =   0   'False
         Width           =   1305
      End
   End
End
Attribute VB_Name = "PropList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type UDT_ITEMS
 Item As New Collection
 Text As New Collection
End Type

Event PropChange(ByVal Prop As String, ByVal Value As String)

Dim Items As UDT_ITEMS, CurrI As Integer, upProp As Boolean

Public Sub Add(ByVal bItem As String, ByVal bDefault As String)
Items.Item.Add bItem$
Items.Text.Add bDefault$
End Sub

Public Sub Clear()
Dim temp As New Collection, tempX As New Collection
Dim con As TextBox
 'For Each con In txt()
 ' con.Visible = False
 'Next con
 'pic.Cls
 'pic.Refresh
 Set Items.Item = temp
 Set Items.Text = tempX
End Sub

Public Sub Update()
upProp = True
Dim i%, l&
pic.Cls
For i% = 1 To Items.Item.Count
 
 With txt(i% - 1)
  .Visible = True
  .Text = Items.Text(i%)
 End With

 If i% - 1 = CurrI% Then
  pic.Line (0, (CurrI%) * 16 + 1)-((ScaleWidth / 2) - 1, ((CurrI% + 1) * 16) - 1), vbHighlight, BF

  l& = pic.ForeColor
  pic.ForeColor = vbHighlightText
  pic.CurrentX = 4
  pic.CurrentY = (i%) * 16 - 15
  pic.Print Items.Item(i%)
  pic.ForeColor = l&
 Else
  l& = pic.ForeColor
  pic.ForeColor = vbWindowText
  pic.CurrentX = 4
  pic.CurrentY = i% * 16 - 15
  pic.Print Items.Item(i%)
  pic.ForeColor = l&
 End If
Next i%
 For i% = Items.Item.Count To txt.Count - 1
  txt(i%).Visible = False
 Next i%

pic.Refresh
'If txt.Count - 1 > CurrI% Then txt(CurrI%).SetFocus
upProp = False
End Sub

Private Sub pic_Resize()
If pic.Height < ((Items.Item.Count + 1) * 16) Then pic.Height = ((Items.Item.Count + 1) * 16)
Call Update
pic.Refresh
End Sub

Private Sub txt_DblClick(Index As Integer)
If LCase(txt(Index).Text) = "true" Then
 txt(Index).Text = "False"
ElseIf LCase(txt(Index).Text) = "false" Then
 txt(Index).Text = "True"
End If

End Sub

Private Sub txt_GotFocus(Index As Integer)
pic.Cls
CurrI% = Index%
Call Update
pic.Line (0, (Index%) * 16 + 1)-((ScaleWidth / 2) - 1, ((Index% + 1) * 16) - 1), vbHighlight, BF

 l& = pic.ForeColor
 pic.ForeColor = vbHighlightText
 pic.CurrentX = 4
 pic.CurrentY = (Index% + 1) * 16 - 15
 pic.Print Items.Item(Index% + 1)
 pic.ForeColor = l&
 txt(Index%).BackColor = vbWindowBackground
pic.Refresh
End Sub

Private Sub txt_LostFocus(Index As Integer)
txt(Index%).BackColor = vbWindowBackground
If upProp = False Then RaiseEvent PropChange(Items.Item(Index% + 1), txt(Index%).Text)
End Sub

Private Sub UserControl_Initialize()
Call ChangeWin(hwnd, False, False, False, False, False, False, False, False, False, True)
Dim i%
For i% = 1 To 16
 Call Load(txt(i%))
 With txt(i%)
  .Left = colHead(1).Left + 2
  .Width = colHead(1).Width - 2
  .Height = 15
  .Top = (i% * 16) + 1
 End With
Next i%
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Int(Y / 16) >= txt.Count Then Refresh: Exit Sub
If txt(Int(Y / 16)).Visible = False Then Exit Sub
txt(Int(Y / 16)).SelStart = 0
txt(Int(Y / 16)).SelLength = Len(txt(Int(Y / 16)).Text)
txt(Int(Y / 16)).SetFocus
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
colHead(0).Width = (ScaleWidth / 2) - 1
colHead(1).Left = (ScaleWidth / 2) - 1
colHead(1).Width = ScaleWidth / 2

Dim con As TextBox
 For Each con In txt()
  con.Left = colHead(1).Left + 2
  con.Height = colHead(0).Height - 2
 Next con

With pic
.Height = Height / Screen.TwipsPerPixelY ' + 120
.Width = Width / Screen.TwipsPerPixelX
.Cls
Set .Picture = LoadPicture()
pic.Line ((ScaleWidth / 2) - 1, 0)-((ScaleWidth / 2) - 1, ScaleHeight)

Dim i%

 For i% = 1 To ScaleHeight / 16
  pic.Line (0, i% * 16)-(ScaleWidth, i% * 16)
 Next i%
.Refresh
Set .Picture = .Image

vs.Left = UserControl.ScaleWidth - vs.Width
vs.Height = UserControl.ScaleHeight - colHead(0).Height

If UserControl.ScaleHeight < ((Items.Item.Count + 1) * 16) Then vs.Max = ((Items.Item.Count + 1) * 16) - UserControl.ScaleHeight Else vs.Max = 0

Call Update
.Refresh
End With
End Sub

Private Sub vs_Change()
pic.Top = Int("-" & vs.Value) + colHead(0).Height

With pic

.Cls
Set .Picture = LoadPicture()
pic.Line ((ScaleWidth / 2) - 1, 0)-((ScaleWidth / 2) - 1, pic.ScaleHeight)

Dim i%

 For i% = 1 To pic.ScaleHeight / 16
  pic.Line (0, i% * 16)-(ScaleWidth, i% * 16)
 Next i%
.Refresh
Set .Picture = .Image

Call Update
.Refresh
End With
End Sub
