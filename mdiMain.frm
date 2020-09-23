VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Visual Ace"
   ClientHeight    =   8025
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10260
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD 
      Left            =   3000
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   7
   End
   Begin VB.PictureBox picDebug 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1845
      Left            =   0
      ScaleHeight     =   1845
      ScaleWidth      =   10260
      TabIndex        =   12
      Top             =   6180
      Visible         =   0   'False
      Width           =   10260
      Begin VB.CommandButton cmdMenu 
         Caption         =   "."
         Height          =   135
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   135
      End
      Begin VB.TextBox txtDebug 
         Height          =   1695
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   120
         Width           =   10215
      End
   End
   Begin MSComctlLib.ImageList ilMain 
      Left            =   2040
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":143A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":188E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2136
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2292
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":23EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":254A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":299E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5820
      Left            =   0
      ScaleHeight     =   388
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   2
      Top             =   360
      Width           =   1035
      Begin VB.CommandButton cmdClose 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   11
         Top             =   0
         Width           =   195
      End
      Begin MSComctlLib.Toolbar tbTool 
         Height          =   2340
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   4128
         ButtonWidth     =   714
         ButtonHeight    =   688
         Style           =   1
         ImageList       =   "ilTool"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pointer"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Button"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "TextBox"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Memo"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "ListBox"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "ComboBox"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "CheckBox"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Option"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Timer"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Image"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Label"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Menu"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000002&
         Caption         =   " ToolBox"
         ForeColor       =   &H80000009&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.PictureBox picRight 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5820
      Left            =   7185
      ScaleHeight     =   388
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   1
      Top             =   360
      Width           =   3075
      Begin VB.CommandButton cmdClose 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   9
         Top             =   0
         Width           =   195
      End
      Begin VB.ComboBox cmbCon 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3000
         Width           =   2895
      End
      Begin MSComctlLib.Toolbar tbProp 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "ilMain"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "View Code"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "View Object"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VisualAce.PropList PropList 
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   6165
      End
      Begin MSComctlLib.TreeView tvwFiles 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4048
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ilFile"
         Appearance      =   1
      End
      Begin VB.Label lbl 
         BackColor       =   &H80000002&
         Caption         =   " Poperties"
         ForeColor       =   &H80000009&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   2520
      End
   End
   Begin MSComctlLib.ImageList ilTool 
      Left            =   2040
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":32F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":37FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3CFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4202
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4706
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":510E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5612
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":601A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":651E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilFile 
      Left            =   2040
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   12
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":72CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":77D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":83A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   100
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open Project"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Project"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Image"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "ai"
                  Text            =   "Add"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Script"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "as"
                  Text            =   "Add"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "ns"
                  Text            =   "New"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add Window"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "aw"
                  Text            =   "Add"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "nw"
                  Text            =   "New"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Run"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Pause"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Code Browser"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stored Procedures and Types"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNPrj 
         Caption         =   "New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOPrj 
         Caption         =   "Open Project"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSPrj 
         Caption         =   "Save Project"
      End
      Begin VB.Menu mnuFileCPrj 
         Caption         =   "Close Project"
      End
      Begin VB.Menu mnuLine44 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAdd 
         Caption         =   "Add File"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileRem 
         Caption         =   "Remove File"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuLine22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSFile 
         Caption         =   "Save File"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSFile 
         Caption         =   "Save File As..."
         Index           =   1
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNScr 
         Caption         =   "New Script"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileNWin 
         Caption         =   "New Window"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMExe 
         Caption         =   "Make Executable"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuLine231z 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileX 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPrj 
      Caption         =   "Project"
      Begin VB.Menu mnuPrjProp 
         Caption         =   "Properties"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuLine123 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrjRun 
         Caption         =   "Run"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewDebug 
         Caption         =   "Debug"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuViewProp 
         Caption         =   "Properties"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuViewTool 
         Caption         =   "ToolBox"
         Checked         =   -1  'True
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu mnuLine738 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuHide 
      Caption         =   "mnuHide"
      Visible         =   0   'False
      Begin VB.Menu mnuDebugClr 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim onLoad As Boolean

Private Function comCode(Optional ByVal isAct As Boolean, Optional ByVal JustCode As Boolean)
  Dim con As Control, frm As Form
  Dim a$, b$, c$, cPar As Control
  Dim d As Boolean, e As Boolean, f As Boolean, g As Boolean
  modLan.gblEnd = False
  For Each frm In Forms()
    If frm.Name = "frmWin" Or frm.Name = "frmEdit" Then frm.Enabled = isAct
   If frm.Name = "frmWin" And JustCode = False Then

   For Each con In frm.Controls()
    If con.Tag <> "" Then
     Select Case con.Name
      Case "picWin"
      'newwindow(winname,caption string,left int,top int,width int,height int)
       Set cPar = con
       a$ = a$ & "newwindow(""" & con.Tag & """,""" & modScript.GetCaption(con.hwnd) & """,0,0," & con.Width & "," & con.Height & ")" & vbCrLf
      Case "cmbNew"
      'newbutton(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newcombo(""" & cPar.Tag & """,""" & con.Tag & """," & con.Left & "," & con.Top & "," & con.Width & ")" & vbCrLf
       b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Text = """ & con.pText & """" & vbCrLf
       d = False: e = True: f = True: g = True
      Case "cmdNew"
      'newbutton(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newbutton(""" & cPar.Tag & """,""" & con.Tag & """," & con.Left & "," & con.Top & "," & con.Width & "," & con.Height & ")" & vbCrLf
       d = True: e = False: f = True: g = True
      Case "lblNew"
      'newbutton(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newlabel(""" & cPar.Tag & """,""" & con.Tag & """," & con.Left & "," & con.Top & "," & con.Width & "," & con.Height & ")" & vbCrLf
       d = True: e = True: f = True: g = False
      Case "imgNew"
      'newimage(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newimage(""" & cPar.Tag & """,""" & con.Tag & """," & con.Left & "," & con.Top & "," & con.Width & "," & con.Height & ")" & vbCrLf
       b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".picture = """ & con.pPicture & """" & vbCrLf
       d = False: e = False: f = False: g = False
      Case "chkNew"
      'newcheck(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newcheck(""" & cPar.Tag & """,""" & con.Tag & """," & con.Left & "," & con.Top & "," & con.Width & "," & con.Height & ")" & vbCrLf
       b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Value = """ & con.pValue & """" & vbCrLf
       d = True: e = True: f = True: g = True
      Case "lstNew"
      'newlist(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newlist(""" & cPar.Tag & """,""" & con.Tag & """," & con.Left & "," & con.Top & "," & con.Width & "," & con.Height & ")" & vbCrLf
       d = False: e = True: f = True: g = True
      Case "memNew"
      'newmemo(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newmemo(""" & cPar.Tag & """,""" & con.Tag & """," & con.Left & "," & con.Top & "," & con.Width & "," & con.Height & ")" & vbCrLf
       b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Text = """ & con.pText & """" & vbCrLf
       d = False: e = True: f = True: g = True
      Case "optNew"
      'newoption(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newoption(""" & cPar.Tag & """,""" & con.Tag & """," & con.Left & "," & con.Top & "," & con.Width & "," & con.Height & ")" & vbCrLf
       b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Value = """ & con.pValue & """" & vbCrLf
       d = True: e = True: f = True: g = True
      Case "tmrNew"
      'newtimer(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newtimer(""" & cPar.Tag & """,""" & con.Tag & """)" & vbCrLf
       b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Interval = """ & con.pInterval & """" & vbCrLf
       d = False: e = False: f = True: g = False
      Case "txtNew"
      'newtext(winname,control name,left int,top int,width int,height int)
       b$ = b$ & "newtext(""" & cPar.Tag & """,""" & con.Tag & """," & con.Left & "," & con.Top & "," & con.Width & "," & con.Height & ")" & vbCrLf
       b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Text = """ & con.pText & """" & vbCrLf
       d = False: e = True: f = True: g = True
      Case Else
       d = False: e = False: f = False: g = False
     End Select
       If d = True Then b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Caption = """ & con.pText & """" & vbCrLf
       If e = True Then If Left$(con.pBackColor, 1) <> "-" Then b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Backcolor = """ & Rgb2Html$(con.pBackColor) & """" & vbCrLf
       If f = True Then b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Enabled = """ & con.pEnabled & """" & vbCrLf
       If g = True Then If Left$(con.pForeColor, 1) <> "-" Then b$ = b$ & "%" & cPar.Tag & "." & con.Tag & ".Forecolor = """ & Rgb2Html$(con.pForeColor) & """" & vbCrLf
    End If
   Next con
   d = False: e = False: f = False: g = False
   ElseIf frm.Name = "frmEdit" And frm.Tag <> "" Then
    c = c & vbCrLf & frm.rtbEdit.Text
   End If
  Next frm
Dim s$
s$ = ""
'If JustCode = False Then
   Dim arr() As String, v As Variant, j As String
 For Each frm In Forms()
  If frm.Name = "frmEdit" Then
   If InStr(frm.Tag, Chr(0)) = 0 Then
   
   j = ""
   arr() = Split(frm.rtbEdit.Text, vbCrLf)
    For Each v In arr()
     If Left(v, 7) = "include" Then
      j = j & retCode(Mid(v, 8)) & vbCrLf
     Else
      j = j & v & vbCrLf
     End If
    Next v
    s$ = s$ & j
   Else

   j = ""
   arr() = Split(frm.rtbEdit.Text, vbCrLf)
    For Each v In arr()
     If Left(v, 7) = "include" Then
      j = j & retCode(Mid(v, 9)) & vbCrLf
     Else
      j = j & v & vbCrLf
     End If
    Next v

    s$ = s$ & Replace$(j, "%Me.", "%" & Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) & ".", , , vbTextCompare)
    s$ = Replace$(s, "#Me.", "#" & Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) & ".", , , vbTextCompare)
    s$ = Replace$(s, "!proc ^", "!proc " & Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) & "_", , , vbTextCompare)
   End If
  End If
 Next frm
 
'MsgBox s$'v
If JustCode = False Then
  c$ = "!proc WinMain()" & vbCrLf & a$ & b$ & gblMenu & IIf(BEFORE_SHOWWIN = 1, WINMAIN_CODE & vbCrLf, "") & "showwin(""" & STARTUP_OBJ & """)" & vbCrLf & IIf(BEFORE_SHOWWIN = 0, WINMAIN_CODE & vbCrLf, "") & "end!" & vbCrLf & s$
Else
  c$ = s$
End If
  comCode = c
End Function

Private Sub cmbCon_Click()
If mSetProp = False Then
 Dim b As Boolean
 Dim i%, con As Control

 i% = cmbCon.ListIndex
 For Each con In gblSelWinObj.Parent.Controls()
  If con.Tag = Left$(cmbCon.List(i%), InStr(cmbCon.List(i%), ":") - 2) Then
 
   Select Case con.Name
    Case "cmdNew", "txtNew", "memNew", "cmbNew", "chkNew", "optNew"
     b = True
    Case Else
     b = False
   End Select
 
  Dim jk As Boolean
  If con.Name = "chkNew" Or con.Name = "optNew" Then jk = True
    If con.Name <> gblSelWinObj.Name Then Call gblSelWinObj.Parent.DrawFocus(con, False) Else Call gblSelWinObj.Parent.BoxHide(False, False)
    If con.Name = "tmrNew" Then
     Call gblSelWinObj.Parent.SetProp(con, False, False, False, False, , , , True)
    Else
     Dim g As Boolean, h As Boolean
     g = True
     If con.Name = "imgNew" Then g = False: h = True
     If con.Name = gblSelWinObj.Name Then Call gblSelWinObj.Parent.SetProp(con, , , False, , False, False, False, False, False) Else Call gblSelWinObj.Parent.SetProp(con, , , , , g, , , , b, jk, h)
    End If
    Set gblSelObj = con
    Exit For

  End If
 Next con
End If
End Sub

Private Sub cmdClose_Click(Index As Integer)
Select Case Index%
 Case 0
  Call mnuViewProp_Click
 Case 1
 Call mnuViewTool_Click
End Select
End Sub

Private Sub cmdMenu_Click()
Call PopupMenu(mnuHide)
End Sub

Private Sub MDIForm_Load()
'Call ChangeWin(picRight.hwnd, False, False, False, False, False, False, False, False, False, False, True)
'Call ChangeWin(picLeft.hwnd, False, True, False, False, False, False, False, False, False, False, True)
'Call SetText(picLeft.hwnd, "ToolBox")
'Call SetText(picRight.hwnd, "Properties")
'Call FlashWindow(picLeft.hwnd, 1)
'Call FlashWindow(picRight.hwnd, 1)

Call modReg.assFile(App.Path & "\VisualAce.exe", App.Path & "\res\project.ico", "vpr", "Visual Ace Project")
Call modReg.assFile(App.Path & "\VisualAce.exe", App.Path & "\res\window.ico", "vaw", "Visual Ace Window")
Call modReg.assFile(App.Path & "\VisualAce.exe", App.Path & "\res\script.ico", "vas", "Visual Ace Script")

Dim nod As Node

 Set nod = tvwFiles.Nodes.Add(, , "img", "Images", 1)
 Set nod = tvwFiles.Nodes.Add(, , "scr", "Scripts", 1)
'  Set nod = tvwFiles.Nodes.Add("scr", tvwChild, "script", "Window1", 4)
 Set nod = tvwFiles.Nodes.Add(, , "win", "Windows", 1)
'  Set nod = tvwFiles.Nodes.Add("win", tvwChild, "form", "Window1", 5)

  tbTool.Buttons(1).Value = tbrPressed
  tbTool.Refresh
  
STARTUP_OBJ = "Window1"
modLan.gErrorsEnd = True
modLan.gErrorsOn = True
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If PROJECT_FILE <> "" Then
If Me.WindowState = vbMinimized Then Me.WindowState = 0
 frmSave.Show vbModal
 If frmSave.Ok = False Then Cancel = -1: Exit Sub
End If

Call SaveStored

gblCanClose = True
modLan.gblEnd = True
End
End Sub

Private Sub mnuDebugClr_Click()
txtDebug.Text = ""
End Sub

Private Sub mnuFileAdd_Click()
On Error GoTo 1
Dim sF$

 With mdiMain.CD
  .Filter = "VA Window File (*.vaw)|*.vaw|VA Script File (*.vas)|*.vas|Image File (*.bmp, *.gif, *.jpg)|*.bmp;*.gif;*.jpg|"
  .CancelError = True
  .FileName = ""
  .ShowOpen
  sF$ = .FileName
 End With

 Select Case LCase$(Right$(sF$, 4))
  Case ".vaw"
   Call OpenWindow(sF$)
  Case ".vas"
   Call OpenScript(sF$)
  Case Else
   Call OpenImage(sF)
 End Select

1
End Sub

Private Sub mnuFileCPrj_Click()
If PROJECT_FILE <> "" Then
 frmSave.Show vbModal
 If frmSave.Ok = False Then Exit Sub
End If
Call ClosePrj
End Sub

Public Function Compile(ByVal sCode As String) As String
Dim l As Long, s As String, i As Integer
 sCode = StrReverse(sCode)
 For l = 1 To Len(sCode)
  If i = 0 Then
   s = s & Chr(Asc(Mid(sCode, l, 1)) - 8)
   i = 1
  ElseIf i = 1 Then
   s = s & Chr(Asc(Mid(sCode, l, 1)) + 5)
   i = 2
  Else
   s = s & Chr(Asc(Mid(sCode, l, 1)) - 4)
   i = 0
  End If
 Next l
 Compile = s
End Function

Sub MakeExe(ByVal sFile As String, Optional ByVal Run As Boolean)
Dim sDat As String

sFile = sFile
EXEC_FILE = sFile

Dim frm As Form

Open App.Path & "\res\" & COMPILER_FILE & ".dat" For Binary Access Read As #1
 sDat = Input(LOF(1), #1)
Close #1
picDebug.Visible = True
 txtDebug.Text = txtDebug.Text & vbCrLf & ">> Fetching Code"
 DoEvents
Dim c As String, sImg As String
c = comCode(True)

Dim o$
o$ = SynChk$(c$)
 Select Case Left$(o$, 1)
  Case "q"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing qoute." & vbCrLf & Mid$(o$, 2)
   Exit Sub
  Case "p"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing paryntheses." & vbCrLf & Mid$(o$, 2)
   Exit Sub
  Case "e"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing procedure 'end!' tag." & vbCrLf & Mid$(o$, 2)
   Exit Sub
  Case "l"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing end Loop command." & vbCrLf & Mid$(o$, 2)
   Exit Sub
  Case "b"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing bracket." & vbCrLf & Mid$(o$, 2)
   Exit Sub
 End Select
'C$
'sFile = "testace.exe"
txtDebug.Text = txtDebug.Text & vbCrLf & ">> Compilling Code"
DoEvents
txtDebug.Text = txtDebug.Text & vbCrLf & ">> Using " & COMPILER_FILE & " with " & COMPILER_METH
DoEvents
'add encryption methods here
If COMPILER_METH <> "default" Then
 
 modLan.clrStrings
 Open App.Path & "\res\" & COMPILER_METH & ".vas" For Input As #1
  modLan.sString = Input(LOF(1), #1)
 Close #1
 
 Call modLan.setString("$Code", c, "Compiler")
 c = Execute(modLan.sString, "Compiler")
Else
 c = Compile(c)
End If

'Open "C:\windows\desktop\dump.code" For Output As #1
' Print #1, c
'Close #1

txtDebug.Text = txtDebug.Text & vbCrLf & ">> Processing Images"
DoEvents
For Each frm In Forms()
 If frm.Name = "frmImg" And frm.Tag <> "" Then
  If ICON_FILE <> "" And Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = ICON_FILE Then _
     txtDebug.Text = txtDebug.Text & vbCrLf & ">> Adding Icons": _
     sDat = ReplaceIcon(Mid(frm.Tag, InStr(frm.Tag, Chr(0)) + 1), sDat)
DoEvents
  Open Mid(frm.Tag, InStr(frm.Tag, Chr(0)) + 1) For Binary Access Read As #1
   sImg = sImg & GetFileName(Mid(frm.Tag, InStr(frm.Tag, Chr(0)) + 1)) & Space(50 - Len(GetFileName(Mid(frm.Tag, InStr(frm.Tag, Chr(0)) + 1)))) & Input(LOF(1), #1) & "&file&"
  Close #1
  'Debug.Print Mid(frm.Tag, InStr(frm.Tag, Chr(0)) + 1)
 End If
Next frm
If sImg <> "" Then sImg = Left(sImg, Len(sImg) - 6)
txtDebug.Text = txtDebug.Text & vbCrLf & ">> Saving File"
DoEvents

Open sFile For Binary Access Write As #1
 Put #1, , sDat & "!CD" & c & "!/CD" & sImg
Close #1


If Run = True Then
 txtDebug.Text = txtDebug.Text & vbCrLf & ">> Build Complete."
 DoEvents
 Call Shell(sFile)
Else
 txtDebug.Text = txtDebug.Text & vbCrLf & ">> Made: " & sFile
 DoEvents
End If

End Sub

Private Sub mnuFileMExe_Click()
On Error GoTo 1
CD.FileName = EXEC_FILE
CD.Filter = "Executable (*.exe)|*.exe"
CD.ShowSave
If CD.FileName = "" Then Exit Sub

Call MakeExe(CD.FileName)
1
End Sub

Private Sub mnuFileNPrj_Click()
frmProjects.Show vbModal
End Sub

Private Sub mnuFileNScr_Click()
Call NewScript
End Sub

Private Sub mnuFileNWin_Click()
Call NewWindow
End Sub

Private Sub mnuFileOPrj_Click()
On Error GoTo 1
CD.Filter = "Visual Ace Project (*.vpr)|*.vpr"
CD.CancelError = True
CD.ShowOpen

Call OpenProject(CD.FileName, IIf(CD.Flags = 1025, True, False))
1
End Sub

Private Sub mnuFileRem_Click()
On Error Resume Next
Dim nod As Node
Dim frm As Form, frmX As Form

Set nod = tvwFiles.SelectedItem
   
   For Each frm In Forms()
    If frm.Name = "frmWin" And InStr(frm.Tag, Chr(0)) <> 0 Then
     If Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = nod.Text Then Exit For
    ElseIf frm.Name = "frmEdit" And InStr(frm.Tag, Chr(0)) <> 0 Then
     'If Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = nod.Text Then Exit For
     Set frmX = frm
    End If
   Next frm

  If MsgBox("Are you sure you want to remove " & nod.Text & "?", vbCritical + vbYesNo, "Remove File") = vbYes Then
   gblCanClose = True
   Call Unload(frm)
   Call Unload(frmX)
   gblCanClose = False
   Call tvwFiles.Nodes.Remove(nod.Index)
  End If
'1
End Sub

Private Sub mnuFileSFile_Click(Index As Integer)
'On Error GoTo 1
Dim nod As Node
Dim frm As Form

Set nod = tvwFiles.SelectedItem

Select Case nod.Parent.Key
 Case "win"
   For Each frm In Forms()
    If frm.Name = "frmWin" And InStr(frm.Tag, Chr(0)) <> 0 Then
     If Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = nod.Text Then Exit For
    End If
   Next frm
  Call SaveWindow(frm, CBool(Index%))
 Case "scr"
   For Each frm In Forms()
    If frm.Name = "frmEdit" And InStr(frm.Tag, Chr(0)) = 0 Then
     If GetFileName$(frm.Tag) = nod.Text & ".vas" Then Exit For
    End If
   Next frm
  Call SaveScript(frm, CBool(Index%))
End Select
1
End Sub

Private Sub mnuFileSPrj_Click()
Call SavePrj
End Sub

Private Sub mnuFileX_Click()
Call Unload(Me)
End Sub

Private Sub mnuHelpAbout_Click()
Dim s As String
s = s & "All In One Chat program >> Aio Chat" & vbCrLf
s = s & "Aio Scripting Language >> Asl Scripts" & vbCrLf
s = s & "Asl Code Editor >> Visual Ace" & vbCrLf
MsgBox s
End Sub

Private Sub mnuHelpTopics_Click()
Call ShowWindow(frmHelp.hwnd, 3)
End Sub

Private Sub mnuPrjProp_Click()
Call frmProp.Show(vbModal)
End Sub

Private Sub mnuPrjRun_Click()
Dim but As MSComctlLib.Button

 Set but = tbMain.Buttons(9)

 Call tbMain_ButtonClick(but)
End Sub

Private Sub mnuViewDebug_Click()
If mnuViewDebug.Checked = True Then mnuViewDebug.Checked = False Else mnuViewDebug.Checked = True
picDebug.Visible = mnuViewDebug.Checked
End Sub

Private Sub mnuViewProp_Click()
If mnuViewProp.Checked = True Then mnuViewProp.Checked = False Else mnuViewProp.Checked = True
picRight.Visible = mnuViewProp.Checked
End Sub

Private Sub mnuViewTool_Click()
If mnuViewTool.Checked = True Then mnuViewTool.Checked = False Else mnuViewTool.Checked = True
picLeft.Visible = mnuViewTool.Checked
End Sub

Private Sub picDebug_Resize()
txtDebug.Width = picDebug.Width
End Sub

Private Sub picLeft_Resize()
On Error GoTo 1
If Me.WindowState = vbMinimized Then Exit Sub
 tbTool.Width = picLeft.Width / Screen.TwipsPerPixelX - 7
 tbTool.Height = picLeft.Height / Screen.TwipsPerPixelY - 10
 lbl(1).Left = 2
 lbl(1).Width = picLeft.ScaleWidth - 20
 'cmdClose(1).Left = picLeft.ScaleWidth - cmdClose(1).Width - 6
 'cmdClose(1).Height = 13
1
End Sub

Private Sub picRight_Resize()
On Error GoTo 1
If Me.WindowState = vbMinimized Then Exit Sub
Dim iW%
iW% = (picRight.Width / Screen.TwipsPerPixelX)

 tbProp.Width = iW% - tbProp.Left - 10
 tvwFiles.Width = iW% - tvwFiles.Left - 10
 cmbCon.Width = iW% - cmbCon.Left - 10

 PropList.Top = cmbCon.Top + cmbCon.Height + 8
 PropList.Width = iW% - PropList.Left - 10
 PropList.Height = (picRight.Height / Screen.TwipsPerPixelY) - PropList.Top
 lbl(0).Width = picRight.ScaleWidth - 20

1
End Sub

Private Sub PropList_PropChange(ByVal Prop As String, ByVal Value As String)
'On Error GoTo 1
Dim b As Boolean

With gblSelObj
 Select Case Prop
  Case "Name"
   If .Name = "lstNew" Then .pText = Value$
   .Tag = Value$
  Case "Enabled"
   If LCase$(Value$) <> "false" And LCase$(Value$) <> "true" Then Exit Sub
   .pEnabled = CBool(Value$)
  Case "Interval"
   If IsNumeric(Value$) = False Then Exit Sub
   .pInterval = CLng(Value$)
  Case "Picture"
   .pPicture = Value
  Case "Left"
   If IsNumeric(Value$) = False Then Exit Sub
   .Left = CInt(Value$)
  Case "Height"
   If IsNumeric(Value$) = False Then Exit Sub
   .Height = CInt(Value$)
  Case "Text"
   If gblSelObj.Name = "picWin" Then
    Call SetText(gblSelObj.hwnd, Value$)
   Else
    .pText = Value$
   End If
  Case "Top"
   If IsNumeric(Value$) = False Then Exit Sub
   .Top = CInt(Value$)
  Case "Value"
   If LCase$(Value$) <> "false" And LCase$(Value$) <> "true" Then Exit Sub
   .pValue = CBool(Value$)
  Case "Visible"
   If LCase$(Value$) <> "false" And LCase$(Value$) <> "true" Then Exit Sub
   .pVisible = CBool(Value$)
  Case "Width"
   If IsNumeric(Value$) = False Then Exit Sub
   .Width = CInt(Value$)
  Case "BackColor"
   If .Name <> "picWin" Then If LCase$(Value$) <> "<default>" Then .pBackColor = HTML2RGB(Value$) Else .pBackColor = .DefaultBCl
  Case "ForeColor"
   If .Name <> "picWin" Then If LCase$(Value$) <> "<default>" Then .pForeColor = HTML2RGB(Value$) Else .pForeColor = .DefaultFCl
 End Select
End With

Dim con As Control, i%

 mdiMain.cmbCon.Clear
 For Each con In gblSelWinObj.Parent.Controls()
  If con.Tag <> "" Then
   mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
   If con.Tag = gblSelObj.Tag Then mSetProp = True: mdiMain.cmbCon.ListIndex = i%: mSetProp = False
   i% = i% + 1
  End If
 Next con

  Select Case gblSelObj.Name
   Case "cmdNew", "txtNew", "memNew", "cmbNew", "chkNew", "optNew"
    b = True
   Case Else
    b = False
  End Select

 Dim jk As Boolean
 If gblSelObj.Name = "chkNew" Or gblSelObj.Name = "optNew" Then jk = True

 If gblSelObj.Name <> gblSelWinObj.Name Then Call gblSelWinObj.Parent.DrawFocus(gblSelObj, False)
 If gblSelObj.Name = "tmrNew" Then
  Call gblSelWinObj.Parent.SetProp(gblSelObj, False, False, False, False, , , , True)
 Else 'ByRef obj As Object, Optional bHeight As Boolean = True, Optional bNULL As Boolean = True, Optional bVis As Boolean = True, Optional bWidth As Boolean = True, Optional bEnabled As Boolean = True, Optional bLeft As Boolean = True, Optional bTop As Boolean = True, Optional bInterval As Boolean = False, Optional bText As Boolean = False
   Dim g As Boolean, h As Boolean
   g = True
   If gblSelObj.Name = "imgNew" Then g = False: h = True
  If gblSelObj.Name = gblSelWinObj.Name Then Call gblSelWinObj.Parent.SetProp(gblSelObj, , , False, , False, False, False, False, True) Else Call gblSelWinObj.Parent.SetProp(gblSelObj, , , , , g, , , , b, jk, h)
 End If
1
End Sub

Private Sub tbMainClick(ByVal Index As Integer)
On Error GoTo 1
Dim sF$
Select Case Index%
 Case 5  'add image
 With CD
  .Filter = "Image File (*.bmp, *.gif, *.jpg, *.ico)|*.bmp;*.gif;*.jpg;*.ico|"
  .CancelError = True
  .FileName = ""
  .ShowOpen
  sF$ = .FileName
 End With

   Call OpenImage(sF$)

 Case 6  'add script
 With CD
  .Filter = "VA Script File (*.vas)|*.vas|"
  .CancelError = True
  .FileName = ""
  .ShowOpen
  sF$ = .FileName
 End With

   Call OpenScript(sF$)

 Case 7  'add window
 With CD
  .Filter = "VA Window File (*.vaw)|*.vaw|"
  .CancelError = True
  .FileName = ""
  .ShowOpen
  sF$ = .FileName
 End With

   Call OpenWindow(sF$)
End Select
1
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo 1
Dim frm As Form
Dim sF$
'MsgBox Button.Index
Select Case Button.Index
 Case 5, 6, 7: Call tbMainClick(Button.Index)
 Case 9
 
 If RUN_BUILD = 1 Then
  Call MakeExe(App.Path & "\build.exe", True)
  Exit Sub
 End If
 
 Dim c As String
 If tbMain.Buttons(11).Enabled = True Then
  tbMain.Buttons(10).Enabled = True
  tbMain.Buttons(9).Enabled = False
  c = comCode(, True)
  modLan.sString = c

  Exit Sub
 End If
 
 picDebug.Visible = True
 txtDebug.Text = txtDebug.Text & vbCrLf & ">> Compiling"
 

c = comCode()
modLan.sString = c
'Debug.Print C
'Exit Sub
Dim o$
o$ = SynChk$(c$)
 Select Case Left$(o$, 1)
  Case "q"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing qoute." & vbCrLf & Mid$(o$, 2)
     For Each frm In Forms()
      If frm.Name = "frmWin" Or frm.Name = "frmEdit" Then frm.Enabled = True
     Next frm
   Exit Sub
  Case "p"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing paryntheses." & vbCrLf & Mid$(o$, 2)
     For Each frm In Forms()
      If frm.Name = "frmWin" Or frm.Name = "frmEdit" Then frm.Enabled = True
     Next frm
   Exit Sub
  Case "e"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing procedure 'end!' tag." & vbCrLf & Mid$(o$, 2)
     For Each frm In Forms()
      If frm.Name = "frmWin" Or frm.Name = "frmEdit" Then frm.Enabled = True
     Next frm
   Exit Sub
  Case "l"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing end Loop command." & vbCrLf & Mid$(o$, 2)
     For Each frm In Forms()
      If frm.Name = "frmWin" Or frm.Name = "frmEdit" Then frm.Enabled = True
     Next frm
   Exit Sub
  Case "b"
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> Syntax Error: Missing bracket." & vbCrLf & Mid$(o$, 2)
     For Each frm In Forms()
      If frm.Name = "frmWin" Or frm.Name = "frmEdit" Then frm.Enabled = True
     Next frm
   Exit Sub
 End Select

   Call modLan.clrStrings
   txtDebug.Text = txtDebug.Text & vbCrLf & ">> " & Execute(c$, "WinMain")

   tbMain.Buttons(9).Enabled = False
   tbMain.Buttons(10).Enabled = True
   tbMain.Buttons(11).Enabled = True
 Case 10 'pause
   tbMain.Buttons(9).Enabled = True
   tbMain.Buttons(10).Enabled = False
   tbMain.Buttons(11).Enabled = True
   gblEnd = True

  For Each frm In Forms()
   If frm.Name = "frmEdit" Then frm.Enabled = True
  Next frm
 Case 11 'stop
  modLan.gblEnd = True
   For Each frm In Forms()
    If frm.Name = "frmNew" Then Call Unload(frm)
   Next frm
  modLan.gblEnd = True
  For Each frm In Forms()
   If frm.Name = "frmWin" Or frm.Name = "frmEdit" Then frm.Enabled = True
  Next frm
   tbMain.Buttons(9).Enabled = True
   tbMain.Buttons(10).Enabled = False
   tbMain.Buttons(11).Enabled = False
   picDebug.Visible = False
 Case 13
  Call ShowWindow(frmCodeBrowser.hwnd, 3)
 Case 14
  Call ShowWindow(frmProc.hwnd, 3)
 Case Else
  modLan.sString = "!proc Main()" & vbCrLf & "msgbox(""else"",0,""Caption"")" & vbCrLf & "end!"
  Call Execute(modLan.sString, "Main")
End Select
1
End Sub

Private Sub tbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo 1
Dim but As MSComctlLib.Button
Select Case ButtonMenu.Tag
 Case "ai": Call tbMainClick(5)
 Case "as": Call tbMainClick(6)
 Case "aw": Call tbMainClick(7)
 Case "ns": Call NewScript
 Case "nw": Call NewWindow
End Select
1
End Sub

Private Sub tbProp_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo 1
Dim frm As Form

Select Case Button.Index
 Case 1
  Dim frmX As Form
 For Each frmX In Forms()
  If frmX.Name = "frmWin" And InStr(frmX.Tag, Chr(0)) <> 0 Then
   If Left$(frmX.Tag, InStr(frmX.Tag, Chr(0)) - 1) = tvwFiles.SelectedItem.Text Then Exit For
  End If
 Next frmX
  
  For Each frm In Forms()
  If frm.Name = "frmEdit" And InStr(frm.Tag, Chr(0)) <> 0 Then
   If Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = tvwFiles.SelectedItem.Text Then
   Dim con As Control
    frm.cmbObj.Clear
    frm.cmbObj.AddItem "(Objects)"
    frm.cmbObj.ListIndex = 0
    For Each con In frmX.Controls()
     If con.Tag <> "" Then frm.cmbObj.AddItem con.Tag
    Next con
    Call ShowWindow(frm.hwnd, 3) 'Show window API. for some reason
   Exit For
   End If
  End If
 Next frm
 Case 2
  For Each frm In Forms()
  If frm.Name = "frmWin" And InStr(frm.Tag, Chr(0)) <> 0 And Left$(frm.Tag, Len(tvwFiles.SelectedItem.Text)) = tvwFiles.SelectedItem.Text Then
    Call ShowWindow(frm.hwnd, 3) 'Show window API. for some reason
   Exit For
  End If
 Next frm
End Select
1
End Sub

Private Sub tbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
objNew = Button.Index
On Error GoTo err1

 If Button.Index <> 1 And gblSelWinObj <> 0 Then gblSelWinObj.MousePointer = ccCross Else gblSelWinObj.MousePointer = vbDefault
  Button.Value = tbrPressed
  tbTool.Refresh
Exit Sub
err1:
  tbTool.Buttons(1).Value = tbrPressed
  tbTool.Refresh
End Sub

Private Sub tvwFiles_Collapse(ByVal Node As MSComctlLib.Node)
Node.Image = 1
End Sub

Private Sub tvwFiles_Expand(ByVal Node As MSComctlLib.Node)
Node.Image = 2
End Sub

Private Sub tvwFiles_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next

If Node.Key = "scr" Or Node.Key = "img" Then tbProp.Buttons(1).Enabled = False Else tbProp.Buttons(1).Enabled = True

Dim frm As Form
If Node.Parent.Key = "scr" Then
 For Each frm In Forms()
  If frm.Name = "frmEdit" And GetFileName(frm.Tag) = Node.Text & ".vas" Then
    Call ShowWindow(frm.hwnd, 3) 'Show window API. for some reason
   Exit For
  End If
 Next frm
 tbProp.Buttons(1).Enabled = False
ElseIf Node.Parent.Key = "img" Then

 For Each frm In Forms()
  If frm.Name = "frmImg" And InStr(frm.Tag, Chr(0)) <> 0 Then
   If Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = Node.Text Then
    Call ShowWindow(frm.hwnd, 3) 'Show window API. for some reason
    Exit For
   End If
  End If
 Next frm
 tbProp.Buttons(1).Enabled = False
ElseIf Node.Parent.Key = "win" Then
 For Each frm In Forms()
  If frm.Name = "frmWin" And GetFileName(frm.Tag) = Node.Text & ".vaw" Then
    Call ShowWindow(frm.hwnd, 3) 'Show window API. for some reason
   Exit For
  End If
 Next frm
 tbProp.Buttons(1).Enabled = True
End If
1
End Sub

Private Sub txtDebug_Change()
txtDebug.SelStart = Len(txtDebug.Text)
End Sub
