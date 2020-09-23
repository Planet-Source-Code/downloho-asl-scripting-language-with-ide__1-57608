VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl rDlg 
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   ScaleHeight     =   3525
   ScaleWidth      =   5160
   Begin VB.FileListBox File 
      Height          =   1650
      Left            =   5760
      Pattern         =   "*.vpr"
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.DirListBox Dir 
      Height          =   1665
      Left            =   5760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   1680
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rDlg.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rDlg.ctx":039C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rDlg.ctx":07F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rDlg.ctx":0C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rDlg.ctx":1098
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbFile 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   2760
      Width           =   2775
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3836
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imglst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdDesk 
      Height          =   315
      Left            =   4560
      Picture         =   "rDlg.ctx":13B4
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Height          =   315
      Left            =   4080
      Picture         =   "rDlg.ctx":173E
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files of &type:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File &name:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Look &in:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   190
      Width           =   570
   End
End
Attribute VB_Name = "rDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim iIndex%, kFile As String

Event FileName(ByVal File As String)

Private Function StripPath(ByVal sTxt As String) As String
Dim i%, s$
s$ = sTxt$
If s$ = "" Then Exit Function
i% = InStrRev(s$, "\")
StripPath$ = Mid$(s$, i% + 1)
End Function

Private Function StripFile(ByVal sTxt As String) As String
Dim i%, s$
s$ = sTxt$
If s$ = "" Then Exit Function
i% = InStrRev(s$, "\")
StripFile$ = Left$(s$, i% - 1)
End Function

Private Sub cmdDesk_Click()
On Error Resume Next
Dir.Path = "C:\Windows\Desktop\"
End Sub

Private Sub cmdOpen_Click()
If txtFileName.Text = "" Or FileExist(Dir.Path & IIf(Right$(Dir.Path, 1) = "\", "", "\") & txtFileName.Text) = False Then Exit Sub
RaiseEvent FileName(Dir.Path & IIf(Right$(Dir.Path, 1) = "\", "", "\") & txtFileName.Text)
txtFileName.Text = ""
End Sub

Private Sub cmdUp_Click()
'MsgBox StripFile$(Dir.Path)
Dir.Path = IIf(Right$(StripFile$(Dir.Path), 1) <> "\", StripFile$(Dir.Path) & "\", StripFile$(Dir.Path))
End Sub

Private Sub Dir_Change()
Dim i%, s$, arr$(), v As Variant
lstFiles.ListItems.Clear
For i% = 0 To Dir.ListCount - 1
 arr$() = Split(StripPath$(Dir.List(i%)), " ")
 s$ = ""
  For Each v In arr$()
   s$ = s$ & UCase$(Left$(StripPath$(v), 1)) & LCase$(Mid$(StripPath$(v), 2)) & " "
  Next v
  s$ = Left$(s$, Len(s) - 1)
  s$ = IIf(Len(s) <= 1, StripPath$(Dir.List(i%)), StripPath$(s$))
 lstFiles.ListItems.Add , "!" & Dir.List(i%), s$, , 2
Next i%
File.Path = Dir.Path
End Sub

Private Sub Drive_Change()
On Error GoTo 1
'Dir.Path = Drive.List(Drive.ListIndex)
'Dir.Path = Mid(Drive.List(Drive.ListIndex), 1, 2) & "\"
Dir.Path = Drive.Drive
Dir.Refresh
iIndex% = Drive.ListIndex
Exit Sub
1
If Err.Number = 68 Then Call MsgBox(UCase$(Drive.List(Drive.ListIndex)) & "\ is not accessible." & vbCrLf & vbCrLf & "The device is not ready", vbCritical, "Open"): Drive.ListIndex = iIndex%
End Sub

Private Sub File_PathChange()
Dim i%
For i% = 0 To File.ListCount - 1
 lstFiles.ListItems.Add , "@" & File.List(i%), File.List(i%), , IIf(LCase$(Right$(File.List(i%), 4)) = ".prj", 5, 4)
Next i%
End Sub

Private Sub lstFiles_Click()
On Error GoTo 1
Dim s$
s$ = lstFiles.SelectedItem.Key

 Select Case Left$(s$, 1)
  Case "@"
   txtFileName.Text = lstFiles.SelectedItem.Text
 End Select
1
End Sub

Private Sub lstFiles_DblClick()
Dim s$
s$ = lstFiles.SelectedItem.Key

 Select Case Left$(s$, 1)
  Case "!"
   Dir.Path = Mid$(s$, 2)
  Case "@"
   'call open
   txtFileName.Text = lstFiles.SelectedItem.Text
   Call cmdOpen_Click
 End Select
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0: Call cmdOpen_Click
End Sub

Private Sub UserControl_Initialize()
cmbFile.AddItem "Visual Ace Project (*.vpr)"
cmbFile.ListIndex = 0
Dir.Path = Drive.Drive
iIndex% = 1: kFile$ = "!cancel!"
End Sub
