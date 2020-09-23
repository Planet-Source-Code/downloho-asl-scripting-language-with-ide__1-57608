VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCodeBrowser 
   Caption         =   "Code Browser"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCodeBrowser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   9135
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBack 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   605
      TabIndex        =   3
      Top             =   5880
      Width           =   9135
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   10
         Left            =   7560
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   9
         Left            =   6840
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   8
         Left            =   6120
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   7
         Left            =   5400
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   6
         Left            =   4680
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblDes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   5
         Left            =   3960
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   4
         Left            =   3240
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   3
         Left            =   2520
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "code1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   1080
         MouseIcon       =   "frmCodeBrowser.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "See Also:"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   870
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type: function"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1245
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeBrowser.frx":110C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeBrowser.frx":1DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeBrowser.frx":2AC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwCode 
      Height          =   5655
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   9975
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtCode 
      Height          =   5655
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.ListBox lstType 
      Height          =   5580
      ItemData        =   "frmCodeBrowser.frx":37A0
      Left            =   120
      List            =   "frmCodeBrowser.frx":37A7
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCodeBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrCode(110, 5) As String

Private Function RetIndex(ByVal sType As String)
Select Case sType
 Case "String", "Array"
  RetIndex = 1
 Case "System"
  RetIndex = 3
 Case Else
  RetIndex = 2
End Select
End Function

Private Function IsInList(ByVal s As String) As Boolean
Dim i As Integer

 For i = 0 To lstType.ListCount - 1
  If lstType.List(i) = s Then IsInList = True: Exit Function
 Next i
End Function

Private Sub Form_Load()
Dim arr() As String, arrX() As String
Dim s As String, v As Variant
Dim i As Integer, l As Long, k As Integer

For i = 2 To lbl.Count - 1
 Set lbl(i).MouseIcon = lbl(1).MouseIcon
 lbl(i).MousePointer = 99
Next i
i = 0
 Open App.Path & "\res\code.brw" For Input As #1
  s = Input(LOF(1), #1)
 Close #1

arr() = Split(s, ";" & vbCrLf)

For l = 0 To UBound(arr())
s = Replace(arr(l), "\sc", ";")

 If i = 0 Then
 k = Int(l / 5)
 
  arrCode(k, i) = Left(s, InStr(s, ":") - 1)
  arrCode(k, i + 1) = Mid(s, InStr(s, ":") + 1)
  lvwCode.ListItems.Add , "n" & k, arrCode(k, i + 1), , RetIndex(arrCode(k, i))
  If IsInList(arrCode(k, i)) = False Then lstType.AddItem arrCode(k, i)
   'd = d & "<a name=""" & arrCode(k, i + 1) & """></a><b>" & arrCode(k, i) & "</b>: <font color=""blue"">" & arrCode(k, i + 1) & "</font><br>"
 ElseIf i = 1 Then
  arrCode(k, i + 1) = arr(l)
   'd = d & s & "<br>" & vbCrLf
 ElseIf i = 2 Then
  arrCode(k, i + 1) = s
   'd = d & s & "<br>" & vbCrLf
 ElseIf i = 3 Then
  arrCode(k, i + 1) = s
   'd = d & "<ul>" & Replace(s, vbCrLf, "<br>") & "</ul>" & vbCrLf
 Else
  arrCode(k, i + 1) = s
   'Dim arf() As String
   'arf = Split(s, ",")
   'd = d & "See Also: "
   'For Each v In arf()
    'd = d & "<a href=""#" & v & """>" & v & "</a>, "
   'Next v
    'd = d & " <a href=""#"">back to top</a><br>" & vbCrLf
  i = -1
 End If
' If i = -1 Then d = d & "<br>" & vbCrLf
i = i + 1
Next l
'Clipboard.Clear
'Clipboard.SetText d
'lstType.ListIndex = 0
End Sub

Private Sub Form_Resize()
If mdiMain.WindowState = vbMinimized Then Exit Sub
lstType.Height = Height - picBack.Height - 500
lvwCode.Height = lstType.Height
'lvwCode.Width = Width - lvwCode.Left
lvwCode.ColumnHeaders(1).Width = lvwCode.Width - 300
txtCode.Left = lvwCode.Left + lvwCode.Width + 100
txtCode.Width = Width - txtCode.Left - 200
txtCode.Height = lstType.Height

End Sub

Private Sub lbl_Click(Index As Integer)
If Index = 0 Then Exit Sub
Dim i As Integer, j As Integer
Dim arr() As String

For i = 0 To UBound(arrCode())
 If arrCode(i, 0) = "" Then Exit For

 If arrCode(i, 1) = lbl(Index).Caption Then
 l = i
 
txtCode.Text = arrCode(l, 4)
lblCode.Caption = arrCode(l, 0) & ": " & arrCode(l, 2)
lblDes.Caption = arrCode(l, 3)

For j = 1 To lbl.Count - 1
 lbl(j).Visible = False
Next j

arr() = Split(arrCode(l, 5), ",")
For j = 0 To UBound(arr())
 lbl(j + 1).Caption = arr(j)
 If j > 0 Then lbl(j + 1).Left = lbl(j).Left + lbl(j).Width + 4
 lbl(j + 1).Visible = True
Next j
Exit For
 End If
Next i
End Sub

Private Sub lstType_Click()
Dim i As Integer
lvwCode.ListItems.Clear
For i = 0 To UBound(arrCode())
 If arrCode(i, 0) = "" Then Exit For
 If lstType.Text = "All" Then
  lvwCode.ListItems.Add , "n" & i, arrCode(i, 1), , RetIndex(arrCode(i, 0))
 Else
  If arrCode(i, 0) = lstType.Text Then
   lvwCode.ListItems.Add , "n" & i, arrCode(i, 1), , RetIndex(arrCode(i, 0))
  End If
 End If
Next i
End Sub

Private Sub lvwCode_Click()
Dim arr() As String, i As Integer
Dim l As Integer
l = CInt(Mid(lvwCode.SelectedItem.Key, 2))

txtCode.Text = arrCode(l, 4)
lblCode.Caption = arrCode(l, 0) & ": " & arrCode(l, 2)
lblDes.Caption = arrCode(l, 3)

For i = 1 To lbl.Count - 1
 lbl(i).Visible = False
Next i

arr() = Split(arrCode(l, 5), ",")
For i = 0 To UBound(arr())
 lbl(i + 1).Caption = arr(i)
 If i > 0 Then lbl(i + 1).Left = lbl(i).Left + lbl(i).Width + 8
 lbl(i + 1).Visible = True
Next i
End Sub

Private Sub lvwCode_DblClick()
Exit Sub
Dim arr() As String, i As Integer
Dim l As Integer
l = CInt(Mid(lvwCode.SelectedItem.Key, 2))

txtCode.Text = arrCode(l, 4)
lblCode.Caption = arrCode(l, 0) & ": " & arrCode(l, 2)
lblDes.Caption = arrCode(l, 3)

For i = 1 To lbl.Count - 1
 lbl(i).Visible = False
Next i

arr() = Split(arrCode(l, 5), ",")
For i = 0 To UBound(arr())
 lbl(i + 1).Caption = arr(i)
 If i > 0 Then lbl(i + 1).Left = lbl(i).Left + lbl(i).Width + 4
 lbl(i + 1).Visible = True
Next i

If Left(txtCode.Text, 1) = "!" Then
 modLan.sString = txtCode.Text
Else
 modLan.sString = "!proc test()" & vbCrLf & txtCode.Text & vbCrLf & "end!"
End If
modLan.clrStrings

Call Execute(modLan.sString, "Test")
End Sub
