VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save Files"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Folder"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Save changes to the following files?"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2550
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ok As Boolean

Private Sub cmdCan_Click()
Ok = False
Me.Hide
End Sub

Private Function GetFormFromHwnd(ByVal lhWnd As Long) As Form
Dim frm As Form
For Each frm In Forms
 If frm.hwnd = lhWnd Then Set GetFormFromHwnd = frm: Exit Function
Next frm
End Function

Private Sub cmdNo_Click()
Ok = True
Me.Hide
End Sub

Private Sub cmdYes_Click()
Dim l As ListItem, frm As Form

For Each l In lvwFiles.ListItems
 If l.Key = "D" & 1 Then
  Call SavePrjFile(l.Text, IIf(l.SubItems(1) <> "\", False, True))
 Else
  If l.Selected = True Then
   Set frm = GetFormFromHwnd(Mid(l.Key, 2))
   If frm.Name = "frmWin" Then Call SaveWindow(frm)
   If frm.Name = "frmEdit" Then Call SaveScript(frm)
  End If
 End If
Next l
Ok = True
Me.Hide

End Sub

Private Sub AddFile(ByVal File As String, ByVal lhWnd As Long)
Dim l As ListItem
Set l = lvwFiles.ListItems.Add(, "D" & lhWnd, GetFileName(File))
    l.SubItems(1) = GetFilePath(File)
    l.Selected = True
End Sub

Private Sub Form_Activate()
Dim frm As Form, sF As String
Call lvwFiles.ListItems.Clear

Call AddFile(IIf(InStr(PROJECT_FILE, "\") <> 0, PROJECT_FILE, "\" & PROJECT_FILE & ".vpr"), 1)

For Each frm In Forms()
 If InStr(frm.Tag, Chr(0)) <> 0 Then
  sF = Mid(frm.Tag, InStr(frm.Tag, Chr(0)) + 1)
  sN = Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1)
 Else
  sF = "": sN = ""
 End If
 
 If frm.Name = "frmWin" And sF <> "" Then
  If sF <> "\" Then
   Call AddFile(sF, frm.hwnd)
  Else
   Call AddFile("\" & sN & ".vaw", frm.hwnd)
  End If
 End If
 
 If frm.Name = "frmEdit" Then
  If frm.IsAttached = False Then
   If sF <> "\" Then
    Call AddFile(frm.Tag, frm.hwnd)
   Else
    Call AddFile("\" & sN & ".vas", frm.hwnd)
   End If
  End If
 End If
Next frm

End Sub

