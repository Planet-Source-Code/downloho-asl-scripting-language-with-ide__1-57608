Attribute VB_Name = "modPrj"
Option Explicit

Global STARTUP_OBJ As String, PROJECT_FILE As String
Global ICON_FILE As String, EXEC_FILE As String
Global COMPILER_FILE As String, COMPILER_METH As String
Global WINMAIN_CODE As String, RUN_BUILD As Integer, BEFORE_SHOWWIN As Integer

Public Sub ClosePrj()
'ask to save
Dim frm As Form
For Each frm In Forms()
 If frm.Name = "frmWin" Or frm.Name = "frmEdit" Or frm.Name = "frmImg" Then
  gblCanClose = True
  Call Unload(frm)
  gblCanClose = False
 End If
Next frm

  PROJECT_FILE = ""
  STARTUP_OBJ = ""
  ICON_FILE = ""
  EXEC_FILE = ""

With mdiMain
 .tvwFiles.Nodes.Clear
 .cmbCon.Clear
Dim nod As Node
 Set nod = .tvwFiles.Nodes.Add(, , "img", "Images", 1)
 Set nod = .tvwFiles.Nodes.Add(, , "scr", "Scripts", 1)
 Set nod = .tvwFiles.Nodes.Add(, , "win", "Windows", 1)

End With
End Sub

Public Function GetFileName(ByVal File As String) As String
Dim i As Integer
If File = "" Then Exit Function
i = InStrRev(File, "\")
GetFileName = Mid(File, i + 1)
End Function

Public Function GetFilePath(ByVal File As String) As String
Dim i As Integer
If File = "" Then Exit Function
i = InStrRev(File, "\")
GetFilePath = Left(File, i)
End Function

Public Sub NewPrj()

End Sub

Public Function NewScript(Optional Parent As String) As Form
Dim frmE As New frmEdit
Dim frm As Form, i%, s$

With frmE
If Parent$ = "" Then
 For Each frm In Forms()
  If frm.Name = "frmEdit" And InStr(frm.Tag, Chr(0)) = 0 Then i% = i% + 1
 Next frm
 i% = i% + 1
  s$ = "Script" & i%
  .Tag = "\" & s & ".vas"
  .IsAttached = False
 Dim nod As Node
 Set nod = mdiMain.tvwFiles.Nodes.Add("scr", tvwChild, , s$, 4)
     nod.Parent.Expanded = True
Else
 s$ = Parent$
  .Tag = Parent$ & Chr(0)
  .IsAttached = True
End If
  .Caption = s$ & " - Script"
  
  Call ShowWindow(.hwnd, 3)
 End With

Set NewScript = frm
End Function

Public Sub NewWindow()
Dim frm As Form, i%
Dim frmW As New frmWin
 
 For Each frm In Forms()
  If frm.Name = "frmWin" Then i% = i% + 1
 Next frm
i% = i% + 1
With frmW
 .Tag = "Window" & i% & Chr(0) & "\" & "Window" & i% & ".vaw"
 .Caption = "Window" & i% & " - Window"
 Call NewScript("Window" & i%)
 Call .onLoad
 Call ShowWindow(.hwnd, 3)
End With
 Dim nod As Node
 Set nod = mdiMain.tvwFiles.Nodes.Add("win", tvwChild, , "Window" & i%, 5)
     nod.Parent.Expanded = True
End Sub

Public Sub OpenImage(ByVal File As String)
Dim s As String, frmI As New frmImg

  With frmI
  .Tag = Left$(GetFileName$(File$), Len(GetFileName$(File$)) - 4) & Chr(0) & File
  Set .imgMain.Picture = LoadPicture(File)
  .Caption = Left$(GetFileName$(File$), Len(GetFileName$(File$)) - 4) & " - Image [" & .imgMain.Width & " x " & .imgMain.Height & "]"
  Call ShowWindow(.hwnd, 3)
 End With
 
 Dim nod As Node
 Set nod = mdiMain.tvwFiles.Nodes.Add("img", tvwChild, , Left$(GetFileName$(File$), Len(GetFileName$(File$)) - 4), 3)
     nod.Parent.Expanded = True
End Sub

Public Sub OpenProject(ByVal File As String, Optional ByVal ReadOnly As Boolean)
Call ClosePrj
Dim s As String, arrX() As String
Dim arr() As String, i As Integer
Open File$ For Input As #1
 s$ = Input(LOF(1), #1)
Close #1
s = StripEndCrLf(s)

arr() = Split(s, vbCrLf)
 If arr(i) <> "Visual Ace Project File" Then MsgBox "Not a valid Visual Ace Project file", vbCritical, "Visual Ace Error: #401": Exit Sub
 For i = 1 To UBound(arr())
  arrX() = Split(arr(i), " ")
  If arr(i) <> "" Then
  Select Case arrX(0)
   Case "Window"
    Call OpenWindow(arrX(1), ReadOnly)
   Case "Script"
    Call OpenScript(arrX(1), ReadOnly)
   Case "Image"
    Call OpenImage(arrX(1))
   Case "StartUp"
    STARTUP_OBJ = arrX(1)
   Case "Compiler"
    COMPILER_FILE = arrX(1)
   Case "Method"
    COMPILER_METH = arrX(1)
   Case "Icon"
    ICON_FILE = arrX(1)
   Case "Project"
    PROJECT_FILE = IIf(ReadOnly = True, "Project1", arrX(1))
   Case "Execute"
    EXEC_FILE = IIf(ReadOnly = True, "Project1", arrX(1))
   Case "RunBuild"
    RUN_BUILD = arrX(1)
   Case "WinMain"
    WINMAIN_CODE = Replace(arrX(1), "</cr+lf\>", vbCrLf)
  End Select
  End If
 Next i
 
 PROJECT_FILE = IIf(ReadOnly = True, "\Project1.vpr", File)
End Sub

Public Function StripEndCrLf(ByVal s As String) As String
Dim l As Long

For l = Len(s) To 1 Step -1
 If Mid(s, l, 1) <> Chr(13) And Mid(s, l, 1) <> Chr(10) Then Exit For
Next l

StripEndCrLf = Left(s, l)
End Function

Public Sub OpenScript(ByVal File As String, Optional ByVal ReadOnly As Boolean)
Dim s As String, frmE As New frmEdit
Open File$ For Input As #1
 s$ = Input(LOF(1), #1)
Close #1
  With frmE
  .Caption = Left$(GetFileName$(File$), Len(GetFileName$(File$)) - 4) & " - Script"
  .Tag = GetFileName$(File$) & Chr(0) & IIf(ReadOnly = True, "\" & GetFileName(File), File)

  Call modColor.PrintText(StripEndCrLf(s), .rtbEdit, .rtbx)
  Call ShowWindow(.hwnd, 3)
 End With
 
  Dim nod As Node
 Set nod = mdiMain.tvwFiles.Nodes.Add("scr", tvwChild, , Left$(GetFileName$(File$), Len(GetFileName$(File$)) - 4), 4)
     nod.Parent.Expanded = True
End Sub

Private Function WinCount() As Integer
Dim i As Integer, j As Integer

For i = 0 To Forms.Count - 1
 If Forms(i).Name = "frmWin" And Forms(i).Tag <> "" Then j = j + 1
Next

WinCount = j + 1
End Function

Public Sub OpenWindow(ByVal File As String, Optional ByVal ReadOnly As Boolean)
Dim frmW As New frmWin
Dim frmE As New frmEdit, lW As Integer
Dim s As String, sCode As String

Open File$ For Input As #1
 s$ = Input(LOF(1), #1)
Close #1
s = StripEndCrLf(s)

 If InStr(s$, "!code!") <> 0 Then
  sCode$ = Mid$(s$, InStr(s$, "!code!") + 6)
  s$ = Left$(s$, InStr(s$, "!code!") - 1)
 Else
  MsgBox "Not a valid Visual Ace Window file", vbCritical, "Visual Ace Error: #400"
  Exit Sub
 End If

  Dim arr$(), v As Variant
  Dim arrX() As String
lW = WinCount
 arr$() = Split(s$, vbCrLf)

  With frmE
  .Caption = "Edit (" & GetFileName(File$) & ")"
  .Tag = IIf(ReadOnly = True, "Window" & lW, Left$(GetFileName$(File$), Len(GetFileName$(File$)) - 4)) & Chr(0)
  Call modColor.PrintText(sCode, .rtbEdit, .rtbx)
  '.rtbEdit.Text = sCode$
  .IsAttached = True
  Call ShowWindow(.hwnd, 3)
 End With

Dim i As Integer, j As Integer
Dim obj As Object, k As Integer

If arr(0) <> "Visual Ace Window File" Then MsgBox "Not a valid Visual Ace Window file", vbCritical, "Visual Ace Error: #400": Exit Sub
For i = 1 To UBound(arr())
 s = newTrim(arr(i))
 If s <> "" Then
 arrX() = Split(s, " ")
  If arrX(0) = "New" Then
   Select Case arrX(1)
    Case "Window"
     frmW.Tag = IIf(ReadOnly = True, "Window" & lW, Left$(GetFileName$(File$), Len(GetFileName$(File$)) - 4)) & Chr(0) & IIf(ReadOnly = True, "\Window" & lW & ".vaw", File)
     frmW.Caption = "Window (" & IIf(ReadOnly = True, "Window" & lW & ".vaw", GetFileName$(File$)) & ")"
      Set obj = frmW.picWin
      obj.Tag = arrX(2)
    Case "Button"
     j = frmW.cmdNew.Count - 1
     If j <> 0 Then Call Load(frmW.cmdNew(j))
      Set obj = frmW.cmdNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
    Case "Checkbox"
     j = frmW.chkNew.Count - 1
     If j <> 0 Then Call Load(frmW.chkNew(j))
      Set obj = frmW.chkNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
    Case "Combobox"
     j = frmW.cmbNew.Count - 1
     If j <> 0 Then Call Load(frmW.cmbNew(j))
      Set obj = frmW.cmbNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
    Case "Image"
     j = frmW.imgNew.Count - 1
     If j <> 0 Then Call Load(frmW.imgNew(j))
      Set obj = frmW.imgNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
    Case "Listbox"
     j = frmW.lstNew.Count - 1
     If j <> 0 Then Call Load(frmW.lstNew(j))
      Set obj = frmW.lstNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
    Case "Memo"
     j = frmW.memNew.Count - 1
     If j <> 0 Then Call Load(frmW.memNew(j))
      Set obj = frmW.memNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
    Case "Option"
     j = frmW.optNew.Count - 1
     If j <> 0 Then Call Load(frmW.optNew(j))
      Set obj = frmW.optNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
    Case "Timer"
     j = frmW.tmrNew.Count - 1
     If j <> 0 Then Call Load(frmW.tmrNew(j))
      Set obj = frmW.tmrNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
    Case "Textbox"
     j = frmW.txtNew.Count - 1
     If j <> 0 Then Call Load(frmW.txtNew(j))
      Set obj = frmW.txtNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
    Case "Menu"
     j = frmW.mnuNew.Count - 1
     If j <> 0 Then Call Load(frmW.mnuNew(j))
      frmW.mnuNew(j).Tag = "Menu"
      frmW.mnuNew(j).Visible = True

     k = frmMenu.lstMenu.ListCount
     frmMenu.lstMenu.AddItem arrX(2) & " - "
     Set obj = frmMenu.lstMenu
    Case "SubMenu"
     k = frmMenu.lstMenu.ListCount
     frmMenu.lstMenu.AddItem ". . . " & arrX(2) & " - "
     j = frmMenu.mnuSub.Count - 1
     If j <> 0 Then Call Load(frmMenu.mnuSub(j))
      Set obj = frmMenu.mnuSub(j)
      obj.Tag = arrX(2) & Chr(1) & Left$(GetFileName$(File$), Len(GetFileName$(File$)) - 4)
    Case "Label"
     j = frmW.lblNew.Count - 1
     If j <> 0 Then Call Load(frmW.lblNew(j))
      Set obj = frmW.lblNew(j)
      obj.Tag = arrX(2)
      obj.Visible = True
   End Select
  Else
   Select Case arrX(0)
    Case "Width"
     obj.Width = CInt(arrX(1))
    Case "Height"
     obj.Height = CInt(arrX(1))
    Case "Left"
     obj.Left = CInt(arrX(1))
    Case "Top"
     obj.Top = CInt(arrX(1))
    Case "Text"
     If obj.Name = "mnuSub" Then
      obj.Caption = arrX(1)
      frmMenu.lstMenu.AddItem frmMenu.lstMenu.List(k) & arrX(1) & " ^" & Left(GetFileName$(File$), Len(GetFileName$(File$)) - 4)
      frmMenu.lstMenu.RemoveItem k
     ElseIf obj.Name = "lstMenu" Then
      frmMenu.lstMenu.AddItem frmMenu.lstMenu.List(k) & arrX(1) & " ^" & Left(GetFileName$(File$), Len(GetFileName$(File$)) - 4)
      frmMenu.lstMenu.RemoveItem k
     ElseIf obj.Name <> "picWin" Then
      obj.pText = arrX(1)
     Else
      Call SetText(frmW.picWin.hwnd, arrX$(1))
     End If
    Case "Image"
     If UBound(arrX()) > 0 Then obj.pPicture = arrX(1)
    Case "Interval"
     obj.pInterval = CLng(arrX(1))
   End Select
  End If
 End If
Next i
'Call Unload(frmWin)
Call frmW.onLoad
Call frmMenu.cmdOk_Click
Call ShowWindow(frmW.hwnd, 3)

Dim nodX As Node
 Set nodX = mdiMain.tvwFiles.Nodes.Add("win", tvwChild, , IIf(ReadOnly = True, "Window" & lW, Left$(GetFileName$(File$), Len(GetFileName$(File$)) - 4)), 5)
 nodX.Selected = True
 mdiMain.tvwFiles.Nodes(3).Expanded = True
End Sub

Public Sub SavePrj()
Dim i As Integer
Dim s As String
s = "Visual Ace Project File" & vbCrLf

 For i = 0 To Forms.Count - 1
  If Forms(i).Name = "frmWin" And Forms(i).Tag <> "" Then
   Call SaveWindow(Forms(i))
   s = s & "Window " & Mid(Forms(i).Tag, InStr(Forms(i).Tag, Chr(0)) + 1) & vbCrLf
  ElseIf Forms(i).Name = "frmImg" And Forms(i).Tag <> "" Then
   s = s & "Image " & Mid(Forms(i).Tag, InStr(Forms(i).Tag, Chr(0)) + 1) & vbCrLf
  ElseIf Forms(i).Name = "frmEdit" Then
   If Forms(i).IsAttached = False Then
    Call SaveScript(Forms(i))
    s = s & "Script " & Forms(i).Tag & vbCrLf
   End If
  End If
 Next i

    s = s & "StartUp " & STARTUP_OBJ & vbCrLf
    s = s & "Icon " & ICON_FILE & vbCrLf
    s = s & "Project " & PROJECT_FILE & vbCrLf
    s = s & "Execute " & EXEC_FILE & vbCrLf
    s = s & "Compiler " & COMPILER_FILE & vbCrLf
    s = s & "Method " & COMPILER_METH & vbCrLf
    s = s & "RunBuild " & RUN_BUILD & vbCrLf
    s = s & "WinMain " & Replace(WINMAIN_CODE, vbCrLf, "</cr+lf\>") & vbCrLf


On Error GoTo 1
 With mdiMain.CD
  .Filter = "VA Project File (*.vpr)|*.vpr|"
  .CancelError = True
  .FileName = PROJECT_FILE
  .ShowSave
  PROJECT_FILE = .FileName
 End With

  Open PROJECT_FILE For Output As #1
    Print #1, s
  Close #1

1
End Sub

Public Sub SaveScript(frm As Form, Optional ByVal SaveAs As Boolean)
On Error GoTo 1
Dim sF As String

 sF$ = frm.Tag

If SaveAs = True Or Left$(frm.Tag, 1) = "\" Then 'save as
 With mdiMain.CD
  .Filter = "VA Script File (*.vas)|*.vas|"
  .CancelError = True
  .FileName = frm.Tag
  .ShowSave
  sF$ = .FileName
 End With
End If

  Open sF$ For Output As #1
    Print #1, frm.rtbEdit.Text
  Close #1
  frm.Tag = sF$
1
End Sub

Public Sub SavePrjFile(Optional ByVal File As String, Optional aSave As Boolean)
Dim i As Integer
Dim s As String
s = "Visual Ace Project File" & vbCrLf

 For i = 0 To Forms.Count - 1
  If Forms(i).Name = "frmWin" And Forms(i).Tag <> "" Then
   s = s & "Window " & Mid(Forms(i).Tag, InStr(Forms(i).Tag, Chr(0)) + 1) & vbCrLf
  ElseIf Forms(i).Name = "frmImg" And Forms(i).Tag <> "" Then
   s = s & "Image " & Mid(Forms(i).Tag, InStr(Forms(i).Tag, Chr(0)) + 1) & vbCrLf
  ElseIf Forms(i).Name = "frmEdit" Then
   If Forms(i).IsAttached = False Then
    s = s & "Script " & Forms(i).Tag & vbCrLf
   End If
  End If
 Next i

    s = s & "StartUp " & STARTUP_OBJ & vbCrLf
    s = s & "Icon " & ICON_FILE & vbCrLf
    s = s & "Project " & PROJECT_FILE & vbCrLf
    s = s & "Execute " & EXEC_FILE & vbCrLf

If File = "" And aSave = False Then
On Error GoTo 1
 With mdiMain.CD
  .Filter = "VA Project File (*.vpr)|*.vpr|"
  .CancelError = True
  .FileName = PROJECT_FILE
  .ShowSave
  PROJECT_FILE = .FileName
 End With
 File = PROJECT_FILE
End If

  Open File For Output As #1
    Print #1, s
  Close #1

1
End Sub

Public Sub SaveWindow(frm As Form, Optional ByVal SaveAs As Boolean)
On Error GoTo 1
Dim s$
With frm
 s$ = s$ & "New Window " & .picWin.Tag & vbCrLf
 s$ = s & Chr(9) & "Width " & .picWin.Width & vbCrLf
 s$ = s$ & Chr(9) & "Height " & .picWin.Height & vbCrLf
  s$ = s$ & Chr(9) & "Text " & GetCaption$(.picWin.hwnd) & vbCrLf
End With

Dim con As Control, t$, i%
 For Each con In frm.Controls
  If Right$(con.Name, 3) = "New" And con.Tag <> "" Then
   Select Case Left$(con.Name, 3)
    Case "chk"
     t$ = "Checkbox": i% = 1
    Case "cmb"
     t$ = "Combobox": i% = 1
    Case "cmd"
     t$ = "Button": i% = 1
    Case "img"
     t$ = "Image": i% = 0
    Case "lst"
     t$ = "Listbox": i% = -1
    Case "mem"
     t$ = "Memo": i% = 1
    Case "opt"
     t$ = "Option": i% = 1
    Case "tmr"
     t$ = "Timer": i% = 2
    Case "txt"
     t$ = "Textbox": i% = 1
    Case "lbl"
     t$ = "Label": i% = 1
    Case Else
     t$ = "": i% = -1
   End Select

If t <> "" Then
     s$ = s$ & "  New " & t$ & " " & con.Tag & vbCrLf
     If i% = 0 Then
      s$ = s$ & Chr(9) & "Image " & con.pPicture & vbCrLf
     ElseIf i% = 1 Then
      s$ = s$ & Chr(9) & "Text " & con.pText & vbCrLf
     ElseIf i% = 2 Then
      s$ = s$ & Chr(9) & "Interval " & con.pInterval & vbCrLf
     End If
'     MsgBox con.Left
     s$ = s$ & Chr(9) & "Left " & con.Left & vbCrLf
     s$ = s$ & Chr(9) & "Top " & con.Top & vbCrLf
     s$ = s$ & Chr(9) & "Width " & con.Width & vbCrLf
     s$ = s$ & Chr(9) & "Height " & con.Height & vbCrLf
   i% = 0
     s = s & "  End " & t & vbCrLf
End If
  End If
 Next con
 
For i = 0 To frmMenu.lstMenu.ListCount - 1
 If InStrRev(frmMenu.lstMenu.List(i), "^") <> 0 Then
  If Mid(frmMenu.lstMenu.List(i), InStrRev(frmMenu.lstMenu.List(i), "^") + 1) = frm.picWin.Tag Then
   t = Left(frmMenu.lstMenu.List(i), InStr(frmMenu.lstMenu.List(i), " - "))
   If Left(frmMenu.lstMenu.List(i), 6) = ". . . " Then
    s = s & "  New SubMenu " & Mid(t, 7) & vbCrLf
    s = s & Chr(9) & "Text " & Mid(frmMenu.lstMenu.List(i), InStr(frmMenu.lstMenu.List(i), " - ") + 3, InStrRev(frmMenu.lstMenu.List(i), "^") - InStr(frmMenu.lstMenu.List(i), " - ") - 4) & vbCrLf
    s = s & "  End SubMenu" & vbCrLf
   Else
    s = s & "  New Menu " & t & vbCrLf
    s = s & Chr(9) & "Text " & Mid(frmMenu.lstMenu.List(i), InStr(frmMenu.lstMenu.List(i), " - ") + 3, InStrRev(frmMenu.lstMenu.List(i), "^") - InStr(frmMenu.lstMenu.List(i), " - ") - 4) & vbCrLf
    s = s & "  End Menu" & vbCrLf
   End If
  End If
 End If
Next i
 s = s & "End " & frm.picWin.Tag & vbCrLf
Dim frmX As Form
 For Each frmX In Forms()
  If frmX.Name = "frmEdit" And InStr(frmX.Tag, Chr(0)) <> 0 Then
   If Left$(frmX.Tag, InStr(frmX.Tag, Chr(0)) - 1) = Left$(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) Then Exit For
  End If
 Next frmX

 s$ = s$ & "!code!" & frmX.rtbEdit.Text

Dim sF As String

 sF$ = Mid$(frm.Tag, InStr(frm.Tag, Chr(0)) + 1)

If SaveAs = True Or Left$(Mid$(frm.Tag, InStr(frm.Tag, Chr(0)) + 1), 1) = "\" Then 'save as
 With mdiMain.CD
  .Filter = "VA Window File (*.vaw)|*.vaw|"
  .CancelError = True
  .FileName = Mid$(frm.Tag, InStr(frm.Tag, Chr(0)) + 1)
  .ShowSave
  sF$ = .FileName
 End With
End If

  Open sF$ For Output As #1
    Print #1, "Visual Ace Window File" & vbCrLf & s$
  Close #1
  frm.Tag = Left$(frm.Tag, InStr(frm.Tag, Chr(0))) & sF$
1
End Sub
