Attribute VB_Name = "modScript"
Option Explicit
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Const WM_SETTEXT = &HC

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum UDE_OBJECTS
 eNULL = 1
 eButton = 2
 eTextBox = 3
 eMemo = 4
 eListBox = 5
 eComboBox = 6
 eCheckBox = 7
 eOption = 8
 eTimer = 9
 eImage = 10
 eLabel = 11
 eMenu = 12
End Enum

Public objNew As UDE_OBJECTS, gblCanClose As Boolean, mSetProp As Boolean
Public gblSelObj As Object, gblNotObj As Object, gblSelWinObj As Object
Public gblMenu As String, gblStored() As String, gblStoredCnt As Integer

Public Function FormIndex(ByVal Win As String) As Form
Dim frm As Form
For Each frm In Forms
 If frm.Name = "frmWin" Then
  If frm.picWin.Tag = Win Then Set FormIndex = frm: Exit Function
 End If
Next frm

Set FormIndex = frmProp
End Function

'MsgBox CBool(IsZoomed(picWin.hwnd))
'gets whether a window is maximized or not

Private Function FormatBin(ByVal inp As String) As String
Dim l&, s$

For l& = 1 To Len(inp$)
 s$ = s$ & Mid$(inp$, l&, 1) & Chr(0)
Next l&
 FormatBin$ = s$
End Function

Private Function Insert(ByVal outStr As String, ByVal insStr As String, ByVal lStart As Long) As String
Dim lLen As Long
 lLen& = Len(insStr$)
 Insert$ = Mid$(outStr, 1, lStart& - 1) & insStr$ & Mid$(outStr$, lStart& + lLen&)
End Function

Public Function ReplaceIcon(ByVal icFile As String, ByVal str As String) As String
Dim ico As String
'locate icon
Open icFile For Binary Access Read As #1
 ico$ = Input(LOF(1), #1)
Close #1

 str$ = Insert$(str$, Mid$(ico$, 23), 22125)
 str$ = Insert$(str$, Mid$(ico$, 23), 181013)

'icon saved
ReplaceIcon = str
End Function

Public Sub DrawDots(ByRef obj As Object)
Dim X%, Y%

 If obj.ScaleWidth <= obj.Parent.xX% And obj.ScaleHeight <= obj.Parent.xY% Then Exit Sub

obj.AutoRedraw = True
 For X% = 0 To obj.ScaleWidth / 12
  For Y% = 0 To obj.ScaleHeight / 12
   obj.PSet (X% * 12, Y% * 12), RGB(125, 125, 125)
  Next Y%
 Next X%
 obj.Refresh
 Set obj.Picture = obj.Image
 obj.Parent.xX% = obj.ScaleWidth
 obj.Parent.xY% = obj.ScaleHeight
obj.AutoRedraw = False
End Sub

Public Function GetCaption(WindowHandle As Long) As String
Dim buffer As String, TextLength As Long
    
 TextLength& = GetWindowTextLength(WindowHandle&)
 buffer$ = String(TextLength&, 0&)

  Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)

 GetCaption$ = buffer$
End Function

Public Sub MakeButton(ByRef frm As Form, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
Dim i%

 For i% = 0 To frm.cmdNew.Count - 1
  If frm.cmdNew(i%).Tag = "" Then GoTo Skip
 Next i%

  Call Load(frm.cmdNew(i%))

Skip:

 With frm.cmdNew(i%)
  .Tag = "Button" & i% + 1
  .pText = "Button" & i% + 1
  .Left = x1
  .Top = y1
  .Width = IIf(x2 < 16, 16, x2)
  .Height = IIf(y2 < 16, 16, y2)
  .Visible = True
 End With

  Set gblSelObj = frm.cmdNew(i%)
  Set gblSelWinObj = frm.picWin

 Call frm.DrawFocus(frm.cmdNew(i%), False)
 Call frm.SetProp(frm.cmdNew(i%), , , , , , , , , True)
End Sub

Public Sub MakeImg(ByRef frm As Form, ByVal eType As UDE_OBJECTS, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
Dim i%, obj As Object, s$

   For i% = 0 To frm.imgNew.Count - 1
    If frm.imgNew(i%).Tag = "" Then GoTo Skip
   Next i%

    Call Load(frm.imgNew(i%))

Skip:
   Set obj = frm.imgNew(i%)
   s$ = "Image"

 With obj
  .Tag = s$ & i% + 1
  '.pText = s$ & i% + 1
  .Left = x1
  .Top = y1
  .Width = IIf(x2 < 16, 16, x2)
  .Height = IIf(y2 < 16, 16, y2)
  .Visible = True
 End With

  Set gblSelObj = obj
  Set gblSelWinObj = obj.Parent.picWin

 Call frm.DrawFocus(obj, False)
 Call frm.SetProp(obj, , , , , False, , , , , , True)
'   frmWin.SetProp(,,,,,,,,,,,
End Sub

Public Sub MakeBox(ByRef frm As Form, ByVal eType As UDE_OBJECTS, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
Dim i%, obj As Object, s$

 Select Case eType
  Case eListBox
   For i% = 0 To frm.lstNew.Count - 1
    If frm.lstNew(i%).Tag = "" Then GoTo Skip
   Next i%

    Call Load(frm.lstNew(i%))
  Case eTextBox
   For i% = 0 To frm.txtNew.Count - 1
    If frm.txtNew(i%).Tag = "" Then GoTo Skip
   Next i%

    Call Load(frm.txtNew(i%))
  Case eMemo
   For i% = 0 To frm.memNew.Count - 1
    If frm.memNew(i%).Tag = "" Then GoTo Skip
   Next i%

    Call Load(frm.memNew(i%))
  Case eLabel
   For i% = 0 To frm.lblNew.Count - 1
    If frm.lblNew(i%).Tag = "" Then GoTo Skip
   Next i%

    Call Load(frm.lblNew(i%))

  Case eComboBox
   For i% = 0 To frm.cmbNew.Count - 1
    If frm.cmbNew(i%).Tag = "" Then GoTo Skip
   Next i%

    Call Load(frm.cmbNew(i%))
 End Select

Skip:

 Select Case eType
  Case eListBox
   Set obj = frm.lstNew(i%)
   s$ = "ListBox"
  Case eTextBox
   Set obj = frm.txtNew(i%)
   s$ = "TextBox"
  Case eMemo
   Set obj = frm.memNew(i%)
   s$ = "Memo"
  Case eLabel
   Set obj = frm.lblNew(i%)
   s$ = "Label"
  Case eComboBox
   Set obj = frm.cmbNew(i%)
   s$ = "ComboBox"
 End Select

 With obj
  .Tag = s$ & i% + 1
  .pText = s$ & i% + 1
  .Left = x1
  .Top = y1
  .Width = IIf(x2 < 16, 16, x2)
  .Height = IIf(y2 < 16, 16, y2)
  .Visible = True
 End With

  Set gblSelObj = obj
  Set gblSelWinObj = obj.Parent.picWin

 Call frm.DrawFocus(obj, False)
 Call frm.SetProp(obj, , , , , , , , , IIf(s$ = "ListBox", False, True))
End Sub

Public Sub MakeOpt(ByRef frm As Form, ByVal eType As UDE_OBJECTS, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
Dim i%, obj As Object, s$

 Select Case eType
  Case eCheckBox
   For i% = 0 To frm.chkNew.Count - 1
    If frm.chkNew(i%).Tag = "" Then GoTo Skip
   Next i%

    Call Load(frm.chkNew(i%))
  Case eOption
   For i% = 0 To frm.optNew.Count - 1
    If frm.optNew(i%).Tag = "" Then GoTo Skip
   Next i%

    Call Load(frm.optNew(i%))
 End Select

Skip:

 Select Case eType
  Case eCheckBox
   Set obj = frm.chkNew(i%)
   s$ = "CheckBox"
  Case eOption
   Set obj = frm.optNew(i%)
   s$ = "Option"
 End Select

 With obj
  .Tag = s$ & i% + 1
  .pText = s$ & i% + 1
  .Left = x1
  .Top = y1
  .Width = IIf(x2 < 16, 16, x2)
  .Height = IIf(y2 < 16, 16, y2)
  .Visible = True
 End With

  Set gblSelObj = obj
  Set gblSelWinObj = obj.Parent.picWin

 Call frm.DrawFocus(obj, False)
 Call frm.SetProp(obj, , , , , , , , , True, True)
End Sub

Public Sub MakeDial(ByRef frm As Form, ByVal eType As UDE_OBJECTS, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
Dim i%, obj As Object, s$, m As UDE_DIALOG

 Select Case eType
  Case eTimer
   For i% = 0 To frm.tmrNew.Count - 1
    If frm.tmrNew(i%).Tag = "" Then GoTo Skip
   Next i%

   Call Load(frm.tmrNew(i%))
  Case eMenu
   For i% = 0 To frm.mnuNew.Count - 1
    If frm.mnuNew(i%).Tag = "" Then GoTo Skip
   Next i%

   Call Load(frm.mnuNew(i%))
 End Select

Skip:
Dim b As Boolean, c As Boolean
 Select Case eType
  Case eTimer
   Set obj = frm.tmrNew(i%)
   s$ = "Timer"
   b = True
   c = True
  Case eMenu
   Set obj = frm.mnuNew(i%)
   s$ = "Menu"
 End Select

 With obj
  .Tag = s$ & i% + 1
  .Left = x1
  .Top = y1
  .Width = IIf(x2 < 16, 16, x2)
  .Height = IIf(y2 < 16, 16, y2)
  .Visible = True
 End With

  Set gblSelObj = obj
  Set gblSelWinObj = obj.Parent.picWin

 Call frm.DrawFocus(obj, False)

 Call frm.SetProp(obj, False, False, False, False, c, , , b)
End Sub

Private Sub LoadStored()
Dim s As String
Dim l As Long, t As Long, k As Integer

 Open App.Path & "\res\proc.vas" For Input As #1
  s = Input(LOF(1), #1)
 Close #1
s = StripEndCrLf(s)
If s = "" Then Exit Sub
l = InStr(LCase(s), "!proc")
t = InStr(l + 1, LCase(s), "end!")

Do

 If l = 0 Or t = 0 Then Exit Do

 ReDim Preserve gblStored(1, k) As String
 
 gblStored(1, k) = Mid(s, l, t - l + 4)
 gblStored(0, k) = Mid(gblStored(1, k), InStr(gblStored(1, k), " ") + 1, InStr(InStr(gblStored(1, k), " ") + 1, gblStored(1, k), "(") - InStr(gblStored(1, k), " ") - 1)

k = k + 1
l = InStr(t + 1, LCase(s), "!proc")
t = InStr(l + 1, LCase(s), "end!")
Loop

l = InStr(LCase(s), "!type")
t = InStr(l + 1, LCase(s), "end!")

Do

 If l = 0 Or t = 0 Then Exit Do

 ReDim Preserve gblStored(1, k) As String
 
 gblStored(1, k) = Mid(s, l, t - l + 4)
 gblStored(0, k) = Mid(gblStored(1, k), InStr(gblStored(1, k), " ") + 1, InStr(InStr(gblStored(1, k), " ") + 1, gblStored(1, k), Chr(13)) - InStr(gblStored(1, k), " ") - 1)

k = k + 1
l = InStr(t + 1, LCase(s), "!type")
t = InStr(l + 1, LCase(s), "end!")
Loop
gblStoredCnt = k - 1
End Sub

Public Sub SaveStored()
Dim i As Integer, s As String
For i = 0 To gblStoredCnt
 If gblStored(1, i) <> "" Then s = s & gblStored(1, i) & vbCrLf
Next i

 Open App.Path & "\res\proc.vas" For Output As #1
  Print #1, s
 Close #1
End Sub

Public Function retCode(ByVal s As String) As String
Dim l As Integer

For l = 0 To gblStoredCnt
 If gblStored(0, l) = s Then retCode = gblStored(1, l): Exit Function
Next l
End Function

Public Sub addStored(ByVal s As String)
Dim k As Integer, l As Long, t As Long
k = gblStoredCnt + 1

l = InStr(LCase(s), "!proc")
t = InStr(l + 1, LCase(s), "end!")

 If l = 0 Or t = 0 Then GoTo 1

 ReDim Preserve gblStored(1, k) As String
 
 gblStored(1, k) = Mid(s, l, t - l + 4)
 gblStored(0, k) = Mid(gblStored(1, k), InStr(gblStored(1, k), " ") + 1, InStr(InStr(gblStored(1, k), " ") + 1, gblStored(1, k), "(") - InStr(gblStored(1, k), " ") - 1)
 gblStoredCnt = gblStoredCnt + 1
 frmProc.lstProc.ListItems.Add , , gblStored(0, k)
Exit Sub
1
l = InStr(LCase(s), "!type")
t = InStr(l + 1, LCase(s), "end!")

 If l = 0 Or t = 0 Then Exit Sub

 ReDim Preserve gblStored(1, k) As String
 
 gblStored(1, k) = Mid(s, l, t - l + 4)
 gblStored(0, k) = Mid(gblStored(1, k), InStr(gblStored(1, k), " ") + 1, InStr(InStr(gblStored(1, k), " ") + 1, gblStored(1, k), Chr(13)) - InStr(gblStored(1, k), " ") - 1)
 gblStoredCnt = gblStoredCnt + 1
 frmProc.lstProc.ListItems.Add , , gblStored(0, k)


End Sub

Sub Main()
objNew = eNULL
Call Load(mdiMain)
mdiMain.Show
Call LoadStored
 Call Load(frmProp)
If Command <> "" Then
 If App.PrevInstance = False Then
  Select Case LCase$(Right$(Command, 4))
   Case ".vaw"
    Call OpenWindow(Command)
   Case ".vas"
    Call OpenScript(Command)
   Case ".vpr"
    Call OpenProject(Command)
  End Select
 End If
Else
 frmProjects.Show vbModal
End If
End Sub

Public Sub OnKeyDown(ByRef obj As Object, ByVal KeyCode As Integer, ByRef Parent As Form)
Dim s$, i%
With obj
Select Case KeyCode
 Case vbKeyUp
  Call Parent.BoxHide(False, False)
  .Top = .Top - 4
 Case vbKeyDown
  Call Parent.BoxHide(False, False)
  .Top = .Top + 4
 Case vbKeyLeft
  Call Parent.BoxHide(False, False)
  .Left = .Left - 4
 Case vbKeyRight
  Call Parent.BoxHide(False, False)
  .Left = .Left + 4
 Case vbKeyDelete
  obj.Tag = ""
  obj.Visible = False
  Call Parent.BoxHide(False, False)
  Dim con As Control
   mdiMain.cmbCon.Clear
    For Each con In obj.Parent.Controls()
     If con.Tag <> "" Then mdiMain.cmbCon.AddItem con.Tag & " : " & ControlType$(con.Name)
    Next con
    
   mSetProp = True
    mdiMain.cmbCon.ListIndex = 0
   mSetProp = False
   Call obj.Parent.SetProp(obj.Parent.picWin, , , False, , False, False, False)

  Exit Sub
 Case Else
  Exit Sub
End Select

   s$ = .Left & "," & .Top
  For i% = 0 To 1
   Parent.picTool(i%).Visible = True
   Parent.picTool(i%).Top = .Top + .Height + 37
   Parent.picTool(i%).Left = .Left + .Width + 24
   Parent.picTool(i%).Cls
   Parent.picTool(i%).CurrentX = (Parent.picTool(i%).Width / 2) - (Parent.picTool(i%).TextWidth(s$) / 2) - 2
   Parent.picTool(i%).CurrentY = (Parent.picTool(i%).Height / 2) - (Parent.picTool(i%).TextHeight(s$) / 2)
   Parent.picTool(i%).Print s$
   Parent.picTool(i%).Refresh
  Next i%

End With
End Sub

Public Sub OnKeyUp(ByRef obj As Object, ByVal KeyCode As Integer, ByRef Parent As Form)
Dim i%
Select Case KeyCode
 Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
  Call Parent.DrawFocus(obj, False)
  For i% = 0 To 1
   Parent.picTool(i%).Visible = False
  Next i%
End Select
End Sub

Public Sub SetText(Window As Long, Text As String)
 Call SendMessageByString(Window&, WM_SETTEXT, 0&, Text$)
End Sub

Public Function ControlType(ByVal sName As String) As String
Select Case sName$
 Case "cmdNew"
  ControlType$ = "Button"
 Case "lstNew"
  ControlType$ = "ListBox"
 Case "cmbNew"
  ControlType$ = "ComboBox"
 Case "txtNew"
  ControlType$ = "TextBox"
 Case "memNew"
  ControlType$ = "Memo"
 Case "tmrNew"
  ControlType$ = "Timer"
 Case "chkNew"
  ControlType$ = "CheckBox"
 Case "optNew"
  ControlType$ = "Option"
 Case "picWin"
  ControlType$ = "Window"
 Case "imgNew"
  ControlType$ = "Image"
 Case "lblNew"
  ControlType$ = "Label"
 Case "mnuNew"
  ControlType$ = "Menu"
End Select
End Function
