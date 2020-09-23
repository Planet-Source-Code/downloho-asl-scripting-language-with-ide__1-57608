VERSION 5.00
Begin VB.Form frmProp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdCan 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkBuild 
         Caption         =   "Build and Run"
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbStart 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmbIco 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "StartUp Object: "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Icon: "
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   4455
      Begin VB.CheckBox chkErrorEnd 
         Caption         =   "End Process on Error"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkError 
         Caption         =   "Use Error Control"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4455
      Begin VB.CheckBox chkBeforeShow 
         Caption         =   "Put before showwin()"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   2660
         Width           =   2295
      End
      Begin VB.TextBox txtMain 
         Height          =   1455
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ComboBox cmbMeth 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox cmbCom 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.FileListBox flbCom 
         Height          =   285
         Left            =   1320
         Pattern         =   "*.dat"
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WinMain Code: "
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Method: "
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compiler: "
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkError_Click()
chkErrorEnd.Enabled = CBool(chkError.Value)
End Sub

Private Sub cmdCan_Click()
Call Me.Hide
End Sub

Private Sub cmdOk_Click()
STARTUP_OBJ = cmbStart.Text
ICON_FILE = cmbIco.Text
COMPILER_FILE = cmbCom.Text
COMPILER_METH = cmbMeth.Text
WINMAIN_CODE = txtMain.Text
BEFORE_SHOWWIN = chkBeforeShow.Enabled
RUN_BUILD = chkBuild.Value
modLan.gErrorsOn = CBool(chkError.Value)
If chkErrorEnd.Enabled = True Then modLan.gErrorsEnd = CBool(chkErrorEnd.Value)
Call Me.Hide
End Sub

Private Sub Form_Activate()
Dim frm As Form, i As Integer, j As Integer
Call cmbStart.Clear

j = 0: i = -1
For Each frm In Forms
 If frm.Tag <> "" And frm.Name = "frmWin" Then
  If Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = STARTUP_OBJ Then i = j
  cmbStart.AddItem Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1)
  j = j + 1
 End If
Next
If i <> -1 Then cmbStart.ListIndex = i

j = 0: i = -1
cmbIco.Clear
For Each frm In Forms
 If frm.Tag <> "" And frm.Name = "frmImg" Then
  If LCase(Right(frm.Tag, 4)) = ".ico" Then
  If Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = ICON_FILE Then i = j
   cmbIco.AddItem Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1)
  End If
 End If
Next
If i <> -1 Then cmbIco.ListIndex = i

For j = 0 To cmbCom.ListCount - 1
 If cmbCom.List(j) = COMPILER_FILE Then cmbCom.ListIndex = j: Exit For
Next

For j = 0 To cmbMeth.ListCount - 1
 If cmbMeth.List(j) = COMPILER_METH Then cmbMeth.ListIndex = j: Exit For
Next

txtMain.Text = WINMAIN_CODE
chkBuild.Value = RUN_BUILD
End Sub

Private Sub Form_Load()
flbCom.Pattern = "*.dat"
flbCom.Path = App.Path & "\res\"
Dim i As Integer, j As Integer

For i = 0 To flbCom.ListCount - 1
 If Left(flbCom.List(i), Len(flbCom.List(i)) - 4) = "ace" Then j = i
 cmbCom.AddItem Left(flbCom.List(i), Len(flbCom.List(i)) - 4)
Next i

cmbCom.ListIndex = j

flbCom.Pattern = "*.vas"
flbCom.Refresh

cmbMeth.AddItem "default"
For i = 0 To flbCom.ListCount - 1
 cmbMeth.AddItem Left(flbCom.List(i), Len(flbCom.List(i)) - 4)
Next i

cmbMeth.ListIndex = 0
End Sub
