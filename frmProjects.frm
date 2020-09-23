VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjects 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Projects"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox flbPrj 
      Height          =   285
      Left            =   120
      Pattern         =   "*.vpr"
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLogo 
      Height          =   1255
      Left            =   160
      Picture         =   "frmProjects.frx":0000
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   3
      Top             =   120
      Width           =   5310
   End
   Begin MSComctlLib.ImageList imlPrj 
      Left            =   4800
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16776960
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjects.frx":29C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjects.frx":2E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjects.frx":3270
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VisualAce.rDlg rDlg 
      Height          =   3615
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6376
   End
   Begin MSComctlLib.ListView lvwPrj 
      Height          =   3615
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6376
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "imlPrj"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7223
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Existing"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()

End Sub

Private Sub File1_Click()

End Sub

Private Sub Form_Load()
lvwPrj.ListItems.Add , , "Empty Project", 2
lvwPrj.ListItems.Add , , "Typical Project", 1
flbPrj.Path = App.Path & "\Templates\"
flbPrj.Refresh

Dim i As Integer
For i = 0 To flbPrj.ListCount - 1
 lvwPrj.ListItems.Add , , Left(GetFileName(flbPrj.List(i)), Len(GetFileName(flbPrj.List(i))) - 4), 3
Next i
End Sub

Private Sub lvwPrj_DblClick()
Me.Hide
Call ClosePrj
Select Case lvwPrj.SelectedItem.Text
 Case "Empty Project"
  PROJECT_FILE = "Project1"
  STARTUP_OBJ = ""
  ICON_FILE = ""
  EXEC_FILE = "Project1"
  COMPILER_FILE = "Ace"
  COMPILER_METH = "default"
  RUN_BUILD = 0
 Case "Typical Project"
  PROJECT_FILE = "Project1"
  STARTUP_OBJ = "Window1"
  ICON_FILE = ""
  EXEC_FILE = "Project1"
  COMPILER_FILE = "Ace"
  COMPILER_METH = "default"
  RUN_BUILD = 0
  Call NewWindow
 Case Else
  Me.Hide
  Call OpenProject(App.Path & "\Templates\" & lvwPrj.SelectedItem.Text & ".vpr", True)
End Select
End Sub

Private Sub picLogo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 1 And Button = 2 Then
If InputBox("") <> "dump.code" Then Exit Sub
Dim s As String, sDat As String, a As String
a = InputBox("")
If a = "" Then Exit Sub

Open App.Path & "\res\ace.dat" For Binary Access Read As #1
 sDat = Input(LOF(1), #1)
Close #1

Open a For Input As #1
 s = Input(LOF(1), #1)
Close #1

s = mdiMain.Compile(s)

Open GetFilePath(a) & "test.exe" For Binary Access Write As #1
 Put #1, , sDat & "!CD" & s & "!/CD"
Close #1
End If
End Sub

Private Sub rDlg_FileName(ByVal File As String)
 Call OpenProject(File)
 Call Me.Hide
End Sub

Private Sub tabMain_Click()
Select Case tabMain.SelectedItem.Index
 Case 1
  lvwPrj.Visible = True
  rDlg.Visible = False
 Case 2
  lvwPrj.Visible = False
  rDlg.Visible = True
End Select
End Sub
