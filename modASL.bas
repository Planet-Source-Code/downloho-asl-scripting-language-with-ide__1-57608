Attribute VB_Name = "modLan"
Option Explicit

Public Declare Function CopyFile Lib "Kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

#If Win32 Then
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell.dll" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
#End If

Private Type UDT_COLORRGB
  Red   As Long
  Green As Long
  Blue  As Long
End Type

Private Type UDT_LOOP
  Id    As New Collection
  Count As New Collection
End Type

Private Type UDT_VARINFO
  Fun    As New Collection
  Name   As New Collection
  Value  As New Collection
  Type   As New Collection
  Alias  As New Collection
  FAlias As New Collection
  Object As New Collection
End Type

Global SYS_PATH As String, TEMP_PATH As String

Public sString As String, gblEnd As Boolean, gErrorsOn As Boolean, gErrorsEnd As Boolean

Dim Strings As UDT_VARINFO, Loops As UDT_LOOP

Dim iTxt As Integer, iCmd As Integer
Dim iChk As Integer, iCmb As Integer
Dim iMem As Integer, iOpt As Integer
Dim iTmr As Integer, iLbl As Integer, iLblA As Integer
Dim iImg As Integer, iLst As Integer
Dim iMnu As Integer, iSub As Integer
Dim iWc As Integer, iFlb As Integer
Dim iDrv As Integer, iDir As Integer

Public Sub clrStrings()
'basically resets all var's and object count
Dim lA As Long, lB As Long, frmX As Form
gblEnd = True
 For Each frmX In Forms
   If frmX.Name = "frmNew" Then Call Unload(frmX)
 Next frmX
gblEnd = False

 lB = Strings.Name.Count

  For lA = 1 To lB
    Call Strings.Fun.Remove(1)
    Call Strings.Name.Remove(1)
    Call Strings.Value.Remove(1)
    Call Strings.Type.Remove(1)
    Call Strings.Alias.Remove(1)
    Call Strings.FAlias.Remove(1)
    Call Strings.Object.Remove(1)
  Next lA

 lB = Loops.Id.Count

  For lA = 1 To lB
    Call Loops.Id.Remove(1)
    Call Loops.Count.Remove(1)
  Next lA

 iTxt = 1: iCmd = 1
 iChk = 1: iCmb = 1
 iMem = 1: iOpt = 1
 iTmr = 1: iLbl = 1: iLblA = 1
 iImg = 1: iLst = 1
 iMnu = 1: iSub = 1
 iWc = 1: iFlb = 1
 iDir = 1: iDrv = 1
 
End Sub

Public Function combSplit(Data() As String, ByVal sFun As String) As String
'this will go thru an array returning all the values of it's contents
Dim sA As String, v As Variant

   For Each v In Data()
     v = Trim(v)
       If v <> "" Then
         If Left(CStr(v), 1) = Chr(34) And Right(CStr(v), 1) = Chr(34) And Len(v) > 1 Then
           sA = sA & Mid(CStr(v), 2, Len(CStr(v)) - 2)
           'MsgBox v
         Else
           sA = sA & retString(CStr(v), sFun)
         End If
       End If
   Next v

 combSplit = Chr(34) & sA & Chr(34)
End Function

Private Function ctrlIndex(frmX As Form, ByVal ctrlName As String) As Integer
'retreives a control's index
Dim i As Integer, IsThere As Boolean

  For i = 0 To frmX.Controls.Count - 1
     If LCase(frmX.Controls(i).Tag) = LCase(ctrlName) Then IsThere = True: Exit For Else IsThere = False
  Next i

   If IsThere = True Then ctrlIndex = i Else ctrlIndex = -1: Call DoError(frmX.Tag & "." & ctrlName, 302)

End Function

Private Sub delString(ByVal sName As String, ByVal sFun As String, Optional ByVal Index As Long = 0)
'delete's a var
Dim IsThere As Boolean, lA As Long

 If sName = "" Then Exit Sub

 If Index = 0 Then
  For lA = 1 To Strings.Name.Count
    If LCase(Strings.Name.Item(lA)) = LCase(sName) And LCase(Strings.Fun.Item(lA)) = LCase(sFun) Then IsThere = True: Exit For Else IsThere = False
  Next lA
 Else
  If Index > 0 And Index <= Strings.Fun.Count Then
    IsThere = True
    lA = Index
  End If
 End If

    If IsThere = True Then
      Call Strings.Fun.Remove(lA)
      Call Strings.Name.Remove(lA)
      Call Strings.Value.Remove(lA)
      Call Strings.Type.Remove(lA)
      Call Strings.Alias.Remove(lA)
      Call Strings.FAlias.Remove(lA)
      Call Strings.Object.Remove(lA)
   End If

End Sub

Private Function DoError(ByVal sText As String, ByVal Num As Integer) As Boolean
Dim s As String
Select Case Num
 Case 301
  s = "Window not found."
 Case 302
  s = "Control not found."
 Case 901
  s = "Variable not declared."
End Select
If gErrorsOn = True Then MsgBox s & vbCrLf & vbCrLf & sText, vbCritical, "Visual Ace Error# " & Num
If gErrorsEnd = True Then DoError = True 'end
End Function

Private Sub delVars(ByVal sFun As String)
'delete's all var's in a function
Dim IsThere As Boolean, lA As Long, lB  As Long

 If sFun = "" Then Exit Sub
 lB = Strings.Name.Count
  For lA = 1 To lB
   If lA > lB Then Exit For
   If Strings.Name.Count = 0 Then Exit For
   If lA < 1 Then lA = 1
   If lA > Strings.Name.Count Then lA = Strings.Name.Count
    If LCase(Strings.Fun.Item(lA)) = LCase(sFun) And Left(Strings.Name.Item(lA), 1) = "$" Then
     Call Strings.Fun.Remove(lA)
     Call Strings.Name.Remove(lA)
     Call Strings.Value.Remove(lA)
     Call Strings.Type.Remove(lA)
     Call Strings.Alias.Remove(lA)
     Call Strings.FAlias.Remove(lA)
     Call Strings.Object.Remove(lA)
     lA = lA - 1
     lB = lB - 1
    End If
    DoEvents
  Next lA

End Sub

Public Function doCode(ByVal sCode As String, ByVal sFun As String) As String
'goes thru the line of code sending it to the appropriate function
Dim arrInfoX() As String, sA As String, sV As String
Dim lA As Long, lB As Long, lC As Long, lD As Long
Dim v As Variant, arr() As String

  If Left(LCase(sCode), 4) = "set " Then
   'set $var_as_type = $another_var_as_type
   sA = Mid(sCode, 5)
   arrInfoX() = Split(sA, "=")

    sV = retType(Trim(arrInfoX(1)), sFun)
    
    If sV = "" Then
     'If Left(arrInfoX(1), 1) = "%" Then
     ' Call setString(Trim(arrInfoX(0)), "", sFun, , Trim(arrInfoX(1)), , "1")
     'Else
      Call setString(Trim(arrInfoX(0)), "", sFun, , Trim(arrInfoX(1)), sFun)
     'End If
     'set arrays and variables here
     Exit Function
    End If
    
    lA = InStr(LCase(modLan.sString), LCase("!type " & sV))
    lB = InStr(lA + 1, LCase(modLan.sString), LCase("end!"))
    If lA = 0 Or lB = 0 Then Exit Function

     sA = Mid(modLan.sString, InStr(lA + 1, modLan.sString, vbCrLf) + 2, lB - InStr(lA + 1, modLan.sString, vbCrLf) - 2)
     arr() = Split(sA, vbCrLf)

        For Each v In arr()
         If v <> "" Then
          v = Trim(Left(v, InStr(v, "=") - 1))
          Call setString(newTrim(arrInfoX(0)) & "." & newTrim(v), retString(newTrim(arrInfoX(1)) & "." & newTrim(v), sFun), sFun)
         End If
        Next v
   Exit Function
  End If

  If Left(LCase(sCode), 5) = "with " Then
  'with kernel32 { $s=GetSystemDirectory() }
   sCode = Mid(sCode, 6)
   sA = Left(sCode, InStr(sCode, " ") - 1)
   
   sCode = Trim(Mid(sCode, InStr(sCode, "{") + 1))
   sCode = Trim(Left(sCode, InStrRev(sCode, "}") - 1))
   
   lA = InStr(sCode, "=")
   If lA <> 0 Then
    sV = Trim(Left(sCode, lA - 1))
    sCode = Trim(Mid(sCode, lA + 1))
    
    lA = InStr(sCode, "(")
    lB = InStrRev(sCode, ")")
    If lA <> 0 And lB <> 0 Then
     arrInfoX() = newSplit(Mid(sCode, lA + 1, lB - lA - 1), ",")
     sCode = Left(sCode, lA - 1)
     Dim sF As String
     For lA = 0 To UBound(arrInfoX())
      sF = sF & retString(arrInfoX(lA), sFun) & ", "
     Next lA
      sF = Left(sF, Len(sF) - 2)
     doCode = doCode(sV & " = use(""" & sA & """, """ & sCode & """" & IIf(sF <> "", "," & sF, "") & ")", sFun)
    Else
     doCode = doCode(sV & " = use(""" & sA & """, """ & sCode & """)", sFun)
    End If
   Else
    doCode = doCode("use(""" & sA & """, """ & sCode & """)", sFun)
   End If
   Exit Function
  End If

  If Left(LCase(sCode), 4) = "var " Then
     sA = Right(sCode, Len(sCode) - 4)
      arrInfoX() = newSplit(sA, ",")

    For Each v In arrInfoX()
     If InStr(CStr(v), "[") <> 0 And InStr(InStr(CStr(v), "[") + 1, CStr(v), "]") <> 0 Then
     'array
      lA = InStr(CStr(v), "[")
      lB = InStr(lA + 1, CStr(v), "]")

        sA = Trim(Mid(CStr(v), lA + 1, lB - lA - 1))
        If IsNumeric(retString(sA, sFun)) Then
        For lC = 0 To CInt(retString(sA, sFun))
          sCode = Left(CStr(v), lA) & CStr(lC) & Mid(CStr(v), lB)

           lD = InStr(sCode, "=")
            If lD = 0 Then
             lD = InStr(CStr(v), "::")
              If lD = 0 Then
               Call setString(sCode, "", sFun)
              Else
               Call setType(Trim(Left(sCode, lD - 1)), Mid(CStr(v), lD + 2), sFun)
              End If
            Else
             sA = Trim(Right(sCode, Len(sCode) - lD))
              If Left(sA, 1) = Chr(34) And Right(sA, 1) = Chr(34) Then
                sA = Mid(sA, 2, Len(sA) - 2)
                Call setString(Trim(Left(sCode, lD - 1)), sA, sFun)
              Else
                Call setString(Trim(Left(sCode, lD - 1)), retString(sA, sFun), sFun)
              End If
            End If
        Next lC
        Else
         lA = InStr(CStr(v), "=")
          If lA = 0 Then
           Call setString(CStr(v), "", sFun)
          Else
           sA = Trim(Mid(CStr(v), lA + 1))
           If Left(sA, 1) = Chr(34) And Right(sA, 1) = Chr(34) Then
            sA = Mid(sA, 2, Len(sA) - 2)
            Call setString(Trim(Left(CStr(v), lA - 1)), sA, sFun)
           Else
            Call setString(Trim(Left(CStr(v), lA - 1)), retString(sA, sFun), sFun)
           End If
          End If
        End If
     Else
      lA = InStr(CStr(v), "=")
       If lA = 0 Then
        lA = InStr(CStr(v), "::")
         If lA = 0 Then
          Call setString(CStr(v), "", sFun)
         Else
          Call setType(Left(CStr(v), lA - 1), Mid(CStr(v), lA + 2), sFun)
         End If
       Else
       sA = Trim(Right(CStr(v), Len(CStr(v)) - lA))
        If Left(sA, 1) = Chr(34) And Right(sA, 1) = Chr(34) Then
          sA = Mid(sA, 2, Len(sA) - 2)
          Call setString(Trim(Left(CStr(v), lA - 1)), sA, sFun)
        Else
          Call setString(Trim(Left(CStr(v), lA - 1)), retString(sA, sFun), sFun)
        End If
       End If
     End If
    Next v: Exit Function
  End If
  If Left(sCode, 1) = "$" Or Left(sCode, 1) = "@" Then
   If Right(sCode, 2) = "++" Or Right(sCode, 2) = "--" Then
     Select Case Right(sCode, 2) 'Mid(sCode, InStr(sCode, "+"), 2)
       Case "++"
        Call setString(Trim(Left(sCode, Len(sCode) - 2)), retMath(Left(sCode, Len(sCode) - 2) & " + 1", sFun), sFun)
       Case "--"
        Call setString(Trim(Left(sCode, Len(sCode) - 2)), retMath(Left(sCode, Len(sCode) - 2) & " - 1", sFun), sFun)
     End Select
   Else

    If InStr(sCode, "[") <> 0 And InStr(InStr(sCode, "[") + 1, sCode, "]") <> 0 Then
    'array
      lA = InStr(sCode, "[")
      lB = InStr(lA + 1, sCode, "]")

        sA = Trim(Mid(sCode, lA + 1, lB - lA - 1))
        sCode = Left(sCode, lA) & retString(sA, sFun) & Mid(sCode, lB)
    End If

      lA& = InStr(sCode$, "=")
      'If lA& = 0 Then lA& = InStr(sCode$, ".=")
       If lA& <> 0 Then
        sA$ = Trim$(Mid(sCode, lA& + 1))

         If InStr(sA$, " & ") Then
            arrInfoX$() = newSplit(sA$, " & "): sA$ = ""
            sA$ = combSplit$(arrInfoX(), sFun$)
         End If

        If Left(sA, 1) = Chr(34) And Right(sA, 1) = Chr(34) Then
         sA = Mid(sA, 2, Len(sA) - 2)
        Else
         sA = retString(sA, sFun)
        End If
         
          Select Case Mid(sCode, lA - 1, 1)
           Case "."
            Call setString(Trim(Left(sCode, lA - 2)), retString(Trim(Left(sCode, lA - 2)), sFun) & sA, sFun)
           Case "+"
            sV = retString(Trim(Left(sCode, lA - 2)), sFun)
            If IsNumeric(sV) = True And IsNumeric(sA) = True Then _
             Call setString(Trim(Left(sCode, lA - 2)), CSng(sV) + CSng(sA), sFun)
           Case "*"
            sV = retString(Trim(Left(sCode, lA - 2)), sFun)
            If IsNumeric(sV) = True And IsNumeric(sA) = True Then _
             Call setString(Trim(Left(sCode, lA - 2)), CSng(sV) * CSng(sA), sFun)
           Case "/"
            sV = retString(Trim(Left(sCode, lA - 2)), sFun)
            If IsNumeric(sV) = True And IsNumeric(sA) = True Then _
             Call setString(Trim(Left(sCode, lA - 2)), CSng(sV) / CSng(sA), sFun)
           Case "-"
            sV = retString(Trim(Left(sCode, lA - 2)), sFun)
            If IsNumeric(sV) = True And IsNumeric(sA) = True Then _
             Call setString(Trim(Left(sCode, lA - 2)), CSng(sV) - CSng(sA), sFun)
           Case Else
            Call setString(Trim(Left(sCode, lA - 1)), sA, sFun)
          End Select
       End If
    End If: Exit Function
  End If

  If Left$(sCode$, 1) = "%" Then
  'Object
    sA$ = Trim$(Right$(sCode$, Len(sCode$) - 1))
    'sA = Replace(sA, "$_", retString("$_", sFun))
     If InStr(sA$, "=") Then
       arrInfoX$() = Split(Trim$(Left$(sA$, InStr(sA$, "=") - 1)), ".")
     Else
       arrInfoX$() = Split(sA$, ".")
     End If
      lA& = winIndex%(Trim$(CStr(retString(arrInfoX$(0), sFun))))
       If lA& <> -1 Then
          lB& = ctrlIndex%(Forms(lA&), Trim$(retString(arrInfoX$(1), sFun)))
            If lB& <> -1 Then
              lC& = InStr(sA$, "=")
                If lC& <> 0 Then
                  sA$ = Trim$(Right$(sA$, Len(sA$) - lC&))
                   If InStr(sA$, " & ") Then
                    arr() = newSplit(sA$, " & "): sA$ = ""
                    sA$ = combSplit$(arr(), sFun$)
                   End If

                    If Left$(sA$, 1) = Chr(34) And Right$(sA$, 1) = Chr(34) And InStr(sA$, " & ") = 0 Then
                      sA$ = Mid$(sA$, 2, Len(sA$) - 2)
                      Call setObject(Forms(lA&).Controls(lB&), retString(Trim$(CStr(Left$(arrInfoX$(2), lC& - 2))), sFun$), sA$, sFun$, lA&, lB&)
                    Else
                      Call setObject(Forms(lA&).Controls(lB&), retString(Trim$(CStr(Left$(arrInfoX$(2), lC& - 2))), sFun$), retString$(sA$, sFun$), sFun$, lA&, lB&)
                    End If
                Else
                  Call setObject(Forms(lA&).Controls(lB&), retString(Trim$(CStr(arrInfoX$(2))), sFun$), "", "", lA&, lB&)
                End If
            Else
             lC& = InStr(sA$, "=")
             If lC& <> 0 Then
              sA$ = Trim$(Right$(sA$, Len(sA$) - lC&))
               If Left$(sA$, 1) = Chr(34) And Right$(sA$, 1) = Chr(34) And InStr(sA$, " & ") = 0 Then
                sA$ = Mid$(sA$, 2, Len(sA$) - 2)
                Call setObject(Forms(lA&), retString(Trim$(CStr(Left$(arrInfoX$(1), lC& - 2))), sFun), sA$, sFun$, lA&, lB&)
               Else
                Call setObject(Forms(lA&), retString(Trim$(CStr(Left$(arrInfoX$(1), lC& - 2))), sFun), retString$(sA$, sFun$), sFun$, lA&, lB&)
               End If
              End If
            End If
       End If: Exit Function
  End If

 If Left(sCode, 1) = "#" Then
  'InActive Control
    sA$ = Trim$(Right$(sCode$, Len(sCode$) - 1))
     If InStr(sA$, "=") Then
       arrInfoX$() = Split(Trim$(Left$(sA$, InStr(sA$, "=") - 1)), ".")
     Else
       arrInfoX$() = Split(sA$, ".")
     End If

   lA& = winIndex%(Trim$(CStr(retString(arrInfoX$(0), sFun))))
   If lA& <> -1 Then
    lB& = ctrlIndex%(Forms(lA&), Trim$(retString(arrInfoX$(1), sFun)))
      If lB& <> -1 Then
       lC& = InStr(sA$, "=")
        If lC& <> 0 Then
         sA$ = Trim$(Right$(sA$, Len(sA$) - lC&))
         If Left$(sA$, 1) = Chr(34) And Right$(sA$, 1) = Chr(34) And InStr(sA$, " & ") = 0 Then sA$ = Mid$(sA$, 2, Len(sA$) - 2)
         Forms(lA&).Controls(lB).Send "#" & arrInfoX(2) & Chr(2) & retString(Trim$(sA$), sFun)
        Else
         Forms(lA&).Controls(lB).Send "#" & arrInfoX(2)
        End If
      End If
   End If
 End If

doCode$ = doProcedure(sCode$, sFun$)

End Function

Private Function doProcedure(ByVal sCode As String, ByVal sFun As String) As String
'executes a procedure that doesn't return a value
'On Error GoTo 1
Dim sTmp As String
Dim arrInfo() As String, arrtmp() As String, arrInfoX() As String
Dim qA As Integer, iA As Integer, iB As Integer
Dim lA As Long, lB As Long, lC As Long
Dim sA As String, v As Variant, sB$

 iA = InStr(sCode, "(")
 iB = InStrRev(sCode, ")")

   If iA = 0 And iB = 0 Then GoTo 1

 sA = Mid(sCode, iA + 1, iB - iA - 1)
 arrInfo() = newSplit(sA, ",")

If Left(Trim(sCode), 1) = "&" Then
 sB = "do("
 iB = InStr(sCode, "(")
  If iB = 0 Then Call MsgBox("Expected '(' but found nothing!", vbCritical, "Error"): Exit Function
 sB = sB & Mid(sCode, 2, iB - 2) & ""

 For iB = LBound(arrInfo()) To UBound(arrInfo())
  sB = sB & "," & arrInfo(iB)
 Next iB
 sB = sB & ")"
 sCode = sB
 iA = InStr(sCode, "(")
 iB = InStrRev(sCode, ")")

     If iA = 0 Or iB = 0 Then GoTo 1

  sB = Mid(sCode, iA + 1, iB - iA - 1)

 arrInfo() = newSplit(sB, ",")
End If

qA = UBound(arrInfo())

For iB = qA To qA + 6
ReDim Preserve arrInfo(iB)
Next iB

ReDim arrtmp(iB) As String
qA = iB - 1

  If Trim(LCase(Left(sCode, iA - 1))) <> "if" Then
     For iB = 0 To qA
       sA = arrInfo(iB)
         If InStr(sA, " & ") Then
           arrInfoX() = newSplit(sA, " & "): sA = ""
           sA = combSplit(arrInfoX(), sFun)
         End If

       arrInfo(iB) = sA

      If arrInfo(iB) <> "" Then
       If Left(Trim(arrInfo(iB)), 1) = Chr(34) And Right(Trim(arrInfo(iB)), 1) = Chr(34) Then
       arrtmp(iB) = Mid(Trim(arrInfo(iB)), 2, Len(Trim(arrInfo(iB))) - 2)
       Else
       arrtmp(iB) = retString(Trim(arrInfo(iB)), sFun)
       End If
      End If

     Next iB
  End If

Dim IsTrue As Boolean

Select Case Trim(LCase(Left(sCode, iA - 1)))
  Case "code_string_add"
  'code_string_add(code)
   sString = sString & vbCrLf & arrtmp(0)
   Exit Function
  Case "code_string_rem"
  'code_string_rem(code)
   lA = InStr(sString, "!proc " & arrtmp(0) & "(")
   lB = InStr(lA + 1, sString, "end!")
   sString = Left(sString, lA - 1) & Mid(sString, lB + 4)
   Exit Function
  Case "appendfile"
  'appendfile(file path, string)
    If arrtmp(0) <> "" Then
     If LCase(arrtmp(0)) = LCase(App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & "info.dat") Then Exit Function

      Open arrtmp(0) For Input As #1
        sTmp = Input(LOF(1), #1)
      Close #1
      
      Open arrtmp(0) For Output As #1
        Print #1, sTmp & arrtmp(1)
      Close #1
    End If
    Exit Function
  Case "beep"
  'beep()
    Call Beep
    Exit Function
  Case "debug.loops"
   For iA = 1 To Loops.Count.Count
    MsgBox Loops.Id(iA)
   Next iA
  Case "debug.print"
    mdiMain.txtDebug.Text = mdiMain.txtDebug.Text & vbCrLf & arrtmp(0)
    Exit Function
  Case "exec"
  'exec(code)
    Call doCode(arrtmp(0), sFun)
    Exit Function
  Case "nll"
  Case "copyfile"
  'copyfile(variable, string)
    Call CopyFile(arrtmp(0), arrtmp(1), 1)
    'Call doCode(arrtmp(0), sFun)
    Exit Function
  Case "comb"
  'comb(variable, string)
    Call setString(arrInfo(0), arrtmp(0) & arrtmp(1), sFun)
    'Call doCode(arrtmp(0), sFun)
    Exit Function
  Case "getrefbytype"
  'getrefbytype(variable name, second variable, function)
    sA = retType(Trim(arrInfo(1)), sFun)
    
    lA = InStr(LCase(modLan.sString), LCase("!type " & sA))
    lB = InStr(lA + 1, LCase(modLan.sString), LCase("end!"))
    If lA = 0 Or lB = 0 Then Exit Function

     sA = Mid(modLan.sString, InStr(lA + 1, modLan.sString, vbCrLf) + 2, lB - InStr(lA + 1, modLan.sString, vbCrLf) - 2)
     arrInfoX() = Split(sA, vbCrLf)
          If arrtmp(2) = "" Then arrtmp(2) = retString("$_", sFun)
        For Each v In arrInfoX()
         If v <> "" Then
          v = Trim(Left(v, InStr(v, "=") - 1))
          Call setString(newTrim(arrInfo(0)) & "." & newTrim(v), retString(newTrim(arrtmp(1)) & "." & newTrim(v), arrtmp(2)), sFun)
         End If
        Next v
    Exit Function
  Case "setrefbytype"
  'setrefbytype(variable name, second variable, function)
    sA = retType(Trim(arrInfo(1)), sFun)
    
    lA = InStr(LCase(modLan.sString), LCase("!type " & sA))
    lB = InStr(lA + 1, LCase(modLan.sString), LCase("end!"))
    If lA = 0 Or lB = 0 Then Exit Function

     sA = Mid(modLan.sString, InStr(lA + 1, modLan.sString, vbCrLf) + 2, lB - InStr(lA + 1, modLan.sString, vbCrLf) - 2)
     arrInfoX() = Split(sA, vbCrLf)
     If arrtmp(2) = "" Then arrtmp(2) = retString("$_", sFun)
        For Each v In arrInfoX()
         If v <> "" Then
          v = Trim(Left(v, InStr(v, "=") - 1))
          Call setString(newTrim(arrtmp(0)) & "." & newTrim(v), retString(newTrim(arrInfo(1)) & "." & newTrim(v), sFun), arrtmp(2))
         End If
        Next v
    Exit Function
  Case "setref"
  'setref(variable name, variable value, function)
    'MsgBox arrtmp(0) & " - " & arrtmp(1) & " - " & arrtmp(2), , "DD"
    If arrtmp(2) = "" Then arrtmp(2) = retString("$_", sFun)
    Call setString(Trim(arrtmp(0)), arrtmp(1), arrtmp(2))
    'MsgBox retString(arrtmp(0), arrtmp(2))
    Exit Function
  Case "sendkeys"
  'sendkeys(keys string, optional wait integer)
    If arrtmp(1) <> "" Then
     Call SendKeys(arrtmp(0), CInt(arrtmp(1)))
    Else
     Call SendKeys(arrtmp(0))
    End If
    Exit Function
  Case "do"
  'do(procedure,args)
  Dim arrFun() As String, arrInfoY() As String
   lA& = InStr(LCase$(sString$), LCase$("!proc " & arrtmp$(0) & "("))
   lB& = InStr(lA& + 1, sString$, ")" & Chr(13))

    If lA& = 0 Or lB& = 0 Then GoTo 2

      sA$ = Mid$(sString$, lA& + Len("!proc " & arrtmp$(0) & "("), lB& - lA& - Len("!proc " & arrtmp$(0) & "("))

    arrFun$() = newSplit(sA$, ",")

        iA% = 1
        For Each v In arrFun$()
          If v <> "" Then
          Call doCode("var " & Trim(v), arrtmp(0))
           If InStr(arrtmp$(iA%), " & ") Then
             arrInfoY$() = newSplit(arrtmp$(iA%), " & ")
             arrtmp$(iA%) = combSplit$(arrInfoY(), sFun$)
           End If
            If Left$(arrtmp$(iA%), 1) = Chr(34) And Right$(arrtmp$(iA%), 1) = Chr(34) Then
            sA$ = Mid$(arrtmp$(iA%), 2, Len(arrtmp$(iA%)) - 2)
            Else
            sA$ = retString$(arrtmp$(iA%), sFun$)
            End If
            iA% = iA% + 1
            If Left(Trim$(CStr(v)), 1) = ">" Then
             'Call setString(Mid(Trim$(CStr(v)), 2), arrInfo(iA - 1), arrtmp$(0))
             Call setString(Mid(Trim$(CStr(v)), 2), "", arrtmp$(0), "", arrInfo(iA - 1), sFun)
            ElseIf Left(Trim$(CStr(v)), 1) = "?" Then
             Call setString(Mid(Trim$(CStr(v)), 2), arrInfo(iA - 1), arrtmp$(0))
            Else
             If InStr(v, "=") <> 0 And sA <> "" Then
              Call setString(Trim$(Trim(Left(CStr(v), InStr(v, "=") - 1))), sA$, arrtmp$(0))
             ElseIf InStr(v, "::") <> 0 And InStr(v, Chr(34)) = 0 Then
              Call setString(Trim$(Trim(Left(CStr(v), InStr(v, "::") - 1))), sA$, arrtmp$(0))
             Else
              Call setString(Trim$(CStr(v)), sA$, arrtmp$(0))
             End If
            End If
          End If
        Next v
2
    sA$ = ""
    Call setString("$_", sFun, arrtmp$(0))
    sA$ = Execute(sString$, arrtmp$(0))
    Do Until sA$ <> "": DoEvents: Loop
    Exit Function
  Case "exitloop"
  'exitloop(id)
    doProcedure$ = "!exitloop!" & "|" & arrtmp$(0)
    Exit Function
  Case "goto"
  'goto(where)
    doProcedure$ = "!goto!" & arrtmp$(0)
    Exit Function
  Case "hidewin"
  'hidewin(winname)
    lA& = winIndex%(arrtmp$(0))
      If lA& <> -1 Then Forms(lA&).Hide
    Exit Function
  Case "split"
  'split(array, string, delimeter)
    arrInfoX() = Split(arrtmp(1), arrtmp(2))
      For lA = 0 To UBound(arrInfoX)
       Call setString(Trim(arrInfo(0)) & "[" & lA & "]", arrInfoX(lA), sFun)
      Next lA
    Exit Function
  Case "pause"
  'pause(interval)
   If IsNumeric(arrtmp$(0)) = False Then Exit Function
    Call Pause(CInt(arrtmp$(0)))
    Exit Function
  Case "clear_array"
  'clear_array(variable)
  iB = Strings.Name.Count
  If Left(arrInfo(0), 1) = ">" Then arrInfo(0) = retString(Mid(arrInfo(0), 2), sFun): sFun = retString("$_", sFun)
    For iA = 1 To iB
     If iA > iB Then Exit For
     If Left(Strings.Name(iA), Len(arrInfo(0)) + 1) = arrInfo(0) & "[" And LCase(Strings.Fun(iA)) = sFun Then
      'And Strings.Fun(iA) = sFun
      Call Strings.Fun.Remove(iA)
      Call Strings.Name.Remove(iA)
      Call Strings.Value.Remove(iA)
      Call Strings.Alias.Remove(iA)
      Call Strings.FAlias.Remove(iA)
      Call Strings.Object.Remove(iA)

      iA = iA - 1
      iB = iB - 1
     End If
    Next iA
    Exit Function
  Case "shell"
  'shell(path)
    Call Shell(arrtmp$(0))
    Exit Function
  Case "switch"
  'switch(case)
    doProcedure$ = "!switch!" & arrtmp(0)
    Exit Function
  Case "caseelse"
    doProcedure$ = "!caseelse!"
    Exit Function
  Case "case"
  'case(variable)
  sA = ""
    For iA = 0 To UBound(arrInfo())
     If Trim(arrInfo(iA)) <> "" Then sA = sA & arrtmp$(iA) & Chr(0) & "|" & Chr(2) Else Exit For
    Next iA
    doProcedure$ = "!case!" & Left(sA, Len(sA) - 3)
    Exit Function
'  Case "endswitch"
  'endswitch()
  Case "for"
  'for(code,statement,code,id)
   
   'arrX() = Split(arrInfo(0), "=")
   'if trim(arrx(0)) <> trim(arrx(1)) then call setstring(arrx(0),a
   If arrInfo$(1) <> "" Then Call doCode(Trim(arrInfo$(2)), sFun$)
    'Debug.Print arrInfo$(2)
    doProcedure$ = "!goloop!" & Eval(arrInfo$(0), sFun$) & "|" & arrtmp$(3)
    Exit Function
  Case "next"
  'next(id)
    doProcedure$ = "!loop!" & "|" & arrInfo$(0)
    Exit Function
  Case "while"
  'while(statement,id,optional variable)
   If arrInfo$(2) <> "" Then Call doCode(Trim(arrInfo$(2)), sFun$)
    'Debug.Print arrInfo$(2)
    doProcedure$ = "!goloop!" & Eval(arrInfo$(0), sFun$) & "|" & arrtmp$(1)
    Exit Function
  Case "loop"
  'loop(id, optional variable)
  If arrInfo$(1) <> "" Then Call doCode(Trim(arrInfo$(1)), sFun$)
    doProcedure$ = "!loop!" & "|" & arrInfo$(0)
    Exit Function
  Case "ife"
  'ife(statement)
   'Debug.Print Eval(arrInfo$(0), sFun$)
   If Eval(arrInfo$(0), sFun$) = "!iftrue!" Then doProcedure = "_!ExIt!_"
   Exit Function
  Case "if"
  'if(statement)
   Dim arrIf() As String

    If InStr(arrInfo$(0), " && ") <> 0 Then
     arrIf$() = Split(arrInfo$(0), " && ")

      For Each v In arrIf$()
        If Eval(CStr(v), sFun$) = "!iftrue!" Then IsTrue = True Else IsTrue = False: Exit For
      Next v

        doProcedure$ = LCase$("!if" & CStr(IsTrue) & "!")

    ElseIf InStr(arrInfo$(0), " || ") Then
     arrIf$() = newSplit(arrInfo$(0), " || ")

      For Each v In arrIf$()
        If Eval(CStr(v), sFun$) = "!iftrue!" Then IsTrue = True: Exit For Else IsTrue = False
      Next v
     
     doProcedure$ = LCase$("!if" & CStr(IsTrue) & "!")
    Else
'    MsgBox Eval("1==1", sFun$)
    'Debug.Print Eval(arrInfo$(0), sFun$)
     sB$ = Eval(arrInfo$(0), sFun$)

     If sB$ = "!iftrue!" And arrInfo$(1) <> "" Then
      doProcedure$ = doCode$(arrInfo$(1), sFun$)
      'If arrInfo$(2) <> "" Then doProcedure$ = doCode$(arrInfo$(2), sFun$)
      'If arrInfo$(3) <> "" Then doProcedure$ = doCode$(arrInfo$(3), sFun$)
     Else
      doProcedure$ = sB$
     End If
    End If
    Exit Function
  Case "elseif"
  'elseif(statement)
    doProcedure$ = Eval(arrInfo$(0), sFun$)
    Exit Function
  Case "else"
  'else()
    doProcedure$ = "!else!"
    Exit Function
  Case "deletefile"
  'deletefile(file path)
   Call Kill(arrtmp(0))
    Exit Function
  Case "killvar"
  'killvar(var name)
   For lA& = 0 To UBound(arrInfo$())
    If arrInfo$(lA&) = "" Then Exit For
     Call delString(arrInfo$(lA&), IIf(Left$(arrInfo$(0), 1) = "$", sFun$, ""))
   Next lA&
    Exit Function
  Case "killall"
  'killall()
    Call delVars(sFun)
    Exit Function
  Case "msgbox", "alert", "showmessage", "messagebox"
  'msgbox(prompt, style int, title)
    Call MsgBox(arrtmp$(0), CInt(IIf(arrtmp$(1) <> "", arrtmp$(1), 0)), arrtmp$(2))
    Exit Function
  Case "newbutton"
  'newbutton(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).cmdNew(iCmd%))
     With Forms(lA&).cmdNew(iCmd%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iCmd% = iCmd% + 1
    Exit Function
  Case "newcheck"
  'newcheck(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).chkNew(iChk%))
     With Forms(lA&).chkNew(iChk%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iChk% = iChk% + 1
    Exit Function
  Case "newcombo"
  'newcombo(winname,control name,left int,top int,width int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).cmbNew(iCmb%))
     With Forms(lA&).cmbNew(iCmb%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Visible = True
     End With: iCmb% = iCmb% + 1
    Exit Function
  Case "newdrivebox"
  'newdrivebox(winname,control name,left int,top int,width int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).drvNew(iDrv%))
     With Forms(lA&).drvNew(iDrv%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Visible = True
     End With: iDrv% = iDrv% + 1
    Exit Function
  Case "newimage"
  'newimage(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).imgNew(iImg%))
     With Forms(lA&).imgNew(iImg%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iImg% = iImg% + 1
    Exit Function
  Case "newlabela"
  'newlabela(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).lblANew(iLbl%))
     With Forms(lA&).lblANew(iLbl%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iLblA% = iLblA% + 1
    Exit Function
  Case "newlabel"
  'newlabel(winname,control name,left int,top int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).lblNew(iLbl%))
     With Forms(lA&).lblNew(iLbl%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iLbl% = iLbl% + 1
    Exit Function
  Case "newlist"
  'newlist(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).lstNew(iLst%))
     With Forms(lA&).lstNew(iLst%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iLst% = iLst% + 1
    Exit Function
  Case "newfilelist"
  'newfilelist(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).flbNew(iFlb%))
     With Forms(lA&).flbNew(iFlb%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iFlb% = iFlb% + 1
    Exit Function
  Case "newdirlist"
  'newdirlist(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).dirNew(iDir%))
     With Forms(lA&).dirNew(iDir%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iDir% = iDir% + 1
    Exit Function
  Case "newmemo"
  'newmemo(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).memNew(iTxt%))
     With Forms(lA&).memNew(iTxt%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iMem% = iMem% + 1
    Exit Function
  Case "newmenu"
  'newmenu(winname,control name, caption)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    'Call Load(Forms(lA&).mnuNew(iMnu%))
     With Forms(lA&).mnuNew(0)
       .Tag = arrtmp$(1)
       .Caption = arrtmp(2)
       .Visible = True
     End With: 'iMnu% = iMnu% + 1
    Exit Function
  Case "newoption"
  'newoption(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).optNew(iOpt%))
     With Forms(lA&).optNew(iOpt%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iOpt% = iOpt% + 1
    Exit Function
  Case "newsubmenu"
  'newsubmenu(winname,control name, caption)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).subNew(iSub%))
     Forms(lA&).subNew(0).Visible = False
     With Forms(lA&).subNew(iSub%)
       .Tag = arrtmp$(1)
       .Caption = arrtmp(2)
       .Visible = True
     End With: iSub% = iSub% + 1
    Exit Function
  Case "newtext"
  'newtext(winname,control name,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).txtNew(iTxt%))
     With Forms(lA&).txtNew(iTxt%)
       .Tag = arrtmp$(1)
       .Left = CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0))
       .Top = CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0))
       .Width = CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100))
       .Height = CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100))
       .Visible = True
     End With: iTxt% = iTxt% + 1
    Exit Function
  Case "newtimer"
  'newtimer(winname,control name)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).tmrNew(iTmr%))
     With Forms(lA&).tmrNew(iTmr%)
       .Tag = arrtmp$(1)
     End With: iTmr% = iTmr% + 1
    Exit Function
  Case "newwindow"
  'newwindow(winname,caption string,left int,top int,width int,height int)
    Dim frmX As New frmNew
    'Call SetParent(frmX.hwnd, frmMain.desktop.hwnd)
     If ICON_FILE <> "" Then
       Dim frm As Form
       For Each frm In Forms()
        If frm.Name = "frmImg" Then
         If Left(frm.Tag, InStr(frm.Tag, Chr(0)) - 1) = ICON_FILE Then Exit For
        End If
       Next frm
       Set frmX.Icon = frm.imgMain.Picture
     End If
     With frmX
       .Tag = arrtmp$(0)
       .Caption = arrtmp$(1)
       .Left = (CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 0)) * Screen.TwipsPerPixelX)
       .Top = (CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0)) * Screen.TwipsPerPixelY)
       .Width = (CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 100)) * Screen.TwipsPerPixelX)
       .Height = (CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 100)) * Screen.TwipsPerPixelY)
     End With
    Exit Function
  Case "newcontrol"
  'newcontrol(winname,inactive object, control name string,left int,top int,width int,height int)
   lA& = winIndex%(arrtmp$(0))
   If lA& = -1 Then Exit Function
    Call Load(Forms(lA&).wc(iWc%))
     Call Forms(lA&).LoadInactiveCtrl(arrtmp$(1), arrtmp$(2), CInt(IIf(arrtmp$(3) <> "", arrtmp$(3), 0)), CInt(IIf(arrtmp$(4) <> "", arrtmp$(4), 0)), CInt(IIf(arrtmp$(5) <> "", arrtmp$(5), 16)), CInt(IIf(arrtmp$(6) <> "", arrtmp$(6), 16)))
     iWc% = iWc% + 1
    Exit Function
  Case "return", "result"
  'return(string)
    doProcedure$ = "!result!" & arrtmp$(0)
    Exit Function
  Case "showwin"
  'showwin(winname)
    lA& = winIndex%(arrtmp$(0))
      If lA& <> -1 Then Forms(lA&).Show
    Exit Function
  Case "writefile"
  'writefile(file path, string)
    If arrtmp$(0) <> "" Then
     If LCase$(arrtmp$(0)) = LCase$(App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "info.dat") Then Exit Function
      Open arrtmp$(0) For Output As #1
        Print #1, arrtmp$(1)
      Close #1
    End If
    Exit Function
End Select

1
End Function

Private Function Eval(ByVal sCode As String, ByVal sFun As String) As String
'returns wether an statement is true or false
'On Error GoTo 1
Dim lA As Long, lB As Long
Dim sL As String, sR As String
Dim isCase As Boolean

'MsgBox CLng(retString(Mid(sCode, 2), sFun))
If Left(sCode, 1) = "!" Then
 If IsNumeric(retString(Mid(sCode, 2), sFun)) = False Then
  If retString(Mid(sCode, 2), sFun) = "" Then
   Eval$ = "!iftrue!"
   Exit Function
  Else
   Eval$ = "!iffalse!"
   Exit Function
  End If
 Else

  If CLng(retString(Mid(sCode, 2), sFun)) <= 0 Then
   Eval$ = "!iftrue!"
   Exit Function
  Else
   Eval$ = "!iffalse!"
   Exit Function
  End If
 End If
End If
   
If InStr(sCode$, " ") = 0 And InStr(sCode$, "<") = 0 And InStr(sCode$, "=") = 0 And InStr(sCode$, ">") = 0 Then
 If IsNumeric(retString(sCode, sFun)) = False Then
  If retString(sCode, sFun) <> "" Then
    Eval$ = "!iftrue!"
    Exit Function
  Else
   Eval$ = "!iffalse!"
   Exit Function
  End If
 Else
  If CLng(retString(sCode, sFun)) > 0 Then
   Eval$ = "!iftrue!"
   Exit Function
  Else
   Eval$ = "!iffalse!"
   Exit Function
  End If
 End If
End If
  
   
   
   lA& = InStr(sCode$, ">")
 If lA& = 0 Then lA& = InStr(sCode$, "<")
 If lA& = 0 Then lA& = InStr(sCode$, "!")
 If lA& = 0 Then lA& = InStr(sCode$, "=")

  lB& = InStr(lA& + 1, sCode$, "=")

    If lA& = 0 Then Eval$ = "!iffalse!": Exit Function
      If lB& <> 0 Then lA& = lB&
    If lA& = 0 Then Eval$ = "!iffalse!": Exit Function

    sL$ = Trim$(Left$(sCode$, lA& - IIf(lB& <> 0, 2, 1)))
    sR$ = Trim$(Right$(sCode$, Len(sCode$) - IIf(lB& <> 0, lB&, lA&)))

       If Left$(sL$, 1) = Chr(34) And Right$(sL$, 1) = Chr(34) Then
       sL$ = Mid$(sL$, 2, Len(sL$) - 2)
       Else
       sL$ = retString$(sL$, sFun$)
       End If
       
       If Left$(sR$, 1) = Chr(34) And Right$(sR$, 1) = Chr(34) Then
       sR$ = Mid$(sR$, 2, Len(sR$) - 2)
       'MsgBox sL$ & " |-" & sR$ & "-"
       Else
       sR$ = retString$(sR$, sFun$)
       End If

    Select Case Mid$(sCode$, lA& - IIf(lB& <> 0, 1, 0), 1)

        Case "="
            If sL$ = sR$ Then isCase = True Else isCase = False
        Case "!"
            If sL$ <> sR$ Then isCase = True Else isCase = False
        Case "<"
          If lB& <> 0 Then
            If CInt(IIf(sL$ <> "", sL$, 0)) <= CInt(IIf(sR$ <> "", sR$, 0)) Then isCase = True Else isCase = False
          Else
            If CInt(IIf(sL$ <> "", sL$, 0)) < CInt(IIf(sR$ <> "", sR$, 0)) Then isCase = True Else isCase = False
          End If
        Case ">"
          If lB& <> 0 Then
            If CInt(IIf(sL$ <> "", sL$, 0)) >= CInt(IIf(sR$ <> "", sR$, 0)) Then isCase = True Else isCase = False
          Else
            If CInt(IIf(sL$ <> "", sL$, 0)) > CInt(IIf(sR$ <> "", sR$, 0)) Then isCase = True Else isCase = False
          End If
    End Select

  If isCase = True Then
    Eval$ = "!iftrue!"
  Else
    Eval$ = "!iffalse!"
  End If
1
End Function

Private Function Looper(ByVal Id As String, ByVal What As Integer, Optional ByVal Count As Integer) As String
Dim i%
'Debug.Print Id, What, Count

 If What% = 0 Then 'set loop
  For i% = 1 To Loops.Id.Count
   If Loops.Id(i%) = Id$ Then
    Exit Function
    Call Loops.Id.Remove(i)
    Call Loops.Count.Remove(i)
   End If
  Next i%
  Loops.Id.Add Id$
  Loops.Count.Add CStr(Count%)
 ElseIf What% = 1 Then 'get loop
  For i% = 1 To Loops.Id.Count
   If Loops.Id(i%) = Id$ Then Looper$ = Loops.Count(i%): Exit Function
  Next i%
   Looper$ = -1
 Else 'remove loop
  For i% = 1 To Loops.Id.Count
   If Loops.Id(i%) = Id$ Then
    Loops.Count.Remove (i%)
    Loops.Id.Remove (i%)
    Exit For
   End If
  Next i%
 End If

End Function

Public Function Execute(ByVal sCode As String, ByVal sFun As String) As String
'main function that splits the code sends it to docode and
'handle's loops,if/then... and such
Dim s As String
If SYS_PATH = "" Then
 s = Space(255)
 SYS_PATH = Left(s, GetSystemDirectory(s, 255))
End If
If TEMP_PATH = "" Then
 s = Space(255)
 TEMP_PATH = Left(s, GetTempPath(255, s))
End If


'On Error Resume Next
Dim arrLines() As String, sTxt As String
Dim lA As Long, lB As Long, arrtmp() As String
Dim lTm As Long, lLoop As Long, v As Variant
Dim sLId$
sTxt = sCode$

    lA& = InStr(LCase$(sTxt$), "!proc " & LCase$(sFun$)) ' find start of procedure
      If lA& = 0 Then Execute = "Error: " & sFun$ & " not found.": _
         Exit Function ' if not found return an error and exit
    lB& = InStr(lA& + 1, sTxt$, vbCrLf & "end!") ' find end of procedure
      If lB& = 0 Or lB& < lA& Then Execute$ = "Error: " & sFun$ & " close tag not found.": _
         Exit Function ' if not found return an error and exit

'sTxt$ = Mid$(sTxt$, lA& + Len("!proc" & sFun$), lB& - lA& - Len("!proc" & sFun$)) ' grab procedure code
sTxt$ = Mid$(sTxt$, InStr(lA& + 1, sTxt$, ")") + 3, (lB& + 6) - InStr(lA& + 1, sTxt$, ")") + 3) ' grab procedure code

'####- Split_Str -#####
arrLines$() = newSplit(sTxt$, vbCrLf) ' grab each line and assign it to a var in an array

For lA& = LBound(arrLines$()) To UBound(arrLines$()) ' loop thru lines and trim useless chr's
arrLines$(lA&) = newTrim$(arrLines$(lA&))
Next lA&
Execute = "done"

Dim sSwitch As String

For lA& = LBound(arrLines$()) To UBound(arrLines$()) ' loop thru lines
If gblEnd = True Then Exit For
  sTxt$ = newTrim$(arrLines$(lA&)) ' re-trim with our new trim function and assign it to a var
     DoEvents
   If LCase$(sTxt$) = "end!" Then Exit For
         If sTxt$ = "break;" Then
          For lB& = lA& To UBound(arrLines$())
                 If Left$(LCase$(arrLines$(lB&)), 3) = "if(" And lB& <> lA& Then
                    For lTm& = lB& To UBound(arrLines$())
                      If LCase$(arrLines$(lTm&)) = "endif()" Then lB& = lTm& + 1: Exit For
                    Next lTm&
                 End If
                If LCase$(arrLines$(lB&)) = "endif()" Then lA& = lB&: sTxt = "": Exit For
           Next lB&
         End If



     If sTxt$ <> "" Then  ' make sure there is code to execute
       If LCase$(sTxt$) = "exit()" Then Exit For

       sTxt$ = doCode$(sTxt$, sFun$)

       If sTxt = "_!ExIt!_" Then Exit For

       If Left$(sTxt$, 6) = "!case!" And sSwitch <> "" Then
        arrtmp() = Split(Mid(sTxt, 7), Chr(0) & "|" & Chr(2))
        lB = 0
         For Each v In arrtmp()
          If sSwitch = v Then sSwitch = "": lB = 0: Exit For Else lB = 1
         Next v
         If lB = 1 Then
          For lB& = lA& + 1 To UBound(arrLines$())
            If Left(LCase$(arrLines$(lB&)), 5) = "case(" Or LCase$(arrLines$(lB&)) = "endswitch()" Or LCase$(arrLines$(lB&)) = "caseelse()" Then lA& = lB& - 1: Exit For
          Next lB&
         End If
       ElseIf Left$(sTxt$, 6) = "!case!" Or sTxt$ = "!caseelse!" And sSwitch = "" Then
         For lB& = lA& To UBound(arrLines$())
           If LCase$(arrLines$(lB&)) = "endswitch()" Then lA& = lB&: Exit For
         Next lB&
       ElseIf sTxt$ <> "" And sSwitch = "" Then
        
        'MsgBox sTxt$ & " - " & arrLines$(lA&)
         If sTxt$ = "!iffalse!" Then
          For lB& = lA& To UBound(arrLines$())
                 If Left$(LCase$(arrLines$(lB&)), 3) = "if(" And lB& <> lA& Then
                    For lTm& = lB& To UBound(arrLines$())
                      If LCase$(arrLines$(lTm&)) = "endif()" Then lB& = lTm& + 1: Exit For
                    Next lTm&
                 End If
                If Left$(LCase$(arrLines$(lB& + 1)), 6) = "elseif" Then lA& = lB&: Exit For
                If LCase$(arrLines$(lB&)) = "else()" Or LCase$(arrLines$(lB&)) = "endif()" Then lA& = lB&: Exit For
           Next lB&
         End If

         If sTxt$ = "!else!" Then
         'MsgBox sTxt$
           For lB& = lA& To UBound(arrLines$())
                 If Left$(Trim$(arrLines$(lB&)), 3) = "if(" And lB& <> lA& Then
                    For lTm& = lB& To UBound(arrLines$())
                      If arrLines$(lTm&) = "endif()" Then lB& = lTm& + 1: Exit For
                    Next lTm&
                 End If
                 'MsgBox LCase$(arrLines$(lB&)) = "endif()"
             If LCase$(arrLines$(lB&)) = "endif()" Then lA& = lB&: Exit For
           Next lB&
         End If

         If Left$(sTxt$, 8) = "!switch!" Then
           sSwitch = Mid(sTxt, 9)
         End If

         If Left$(sTxt$, 6) = "!goto!" Then
           For lB& = lA& To UBound(arrLines$())
             If LCase$(arrLines$(lB&)) = LCase$(Right$(sTxt$, Len(sTxt$) - 6)) Then lA& = lB&: Exit For
           Next lB&
         End If

         If Left$(sTxt$, 8) = "!result!" Then
           Execute$ = Right$(sTxt$, Len(sTxt$) - 8)
           Exit For
         End If

         If Left$(sTxt$, 8) = "!goloop!" Then
          sLId$ = Mid$(sTxt$, InStrRev(sTxt$, "|") + 1)
           If Mid$(sTxt$, Len("!goloop!") + 1, Len("!iftrue!")) = "!iftrue!" Then
            Call Looper$(sLId$, 0, lA&)
            'MsgBox Looper$(sLId$, 1)
           Else
            Call Looper$(sLId$, -1)
            'Debug.Print arrLines(lA), sTxt, Looper(sLId, 1)
            For lB& = lA& To UBound(arrLines$())
             'Debug.Print "", "-" & LCase$(arrLines$(lB&)) & "-", "loop(" & LCase$(sLId$) & ")"
              If LCase$(arrLines$(lB&)) = "loop(" & LCase$(sLId$) & ")" Then lA& = lB&: sTxt = "": Exit For
            Next lB&
            'Debug.Print "", "-" & LCase$(arrLines$(lA&)) & "-", "loop(" & LCase$(sLId$) & ")", LCase$(arrLines$(lB&)) = "loop(" & LCase$(sLId$) & ")"
            'Debug.Print lA, "false loop"
           End If
         End If

         If Left$(sTxt$, Len("!loop!")) = "!loop!" Then
          sLId$ = Mid$(sTxt$, InStrRev(sTxt$, "|") + 1)
'          MsgBox sTxt$
           If CLng(Looper$(sLId$, 1)) <> -1 Then lA& = CLng(Looper$(sLId$, 1)) - 1
         End If

         If Left$(sTxt$, Len("!exitloop!")) = "!exitloop!" Then
          sLId$ = Mid$(sTxt$, InStrRev(sTxt$, "|") + 1)
          Call Looper$(sLId$, -1)
           For lB& = lA& To UBound(arrLines$())
             If LCase$(arrLines$(lB&)) = "loop(" & LCase$(sLId$) & ")" Then lA& = lB&: Exit For
           Next lB&
         End If
       End If

     End If
Next lA&
Call delVars(sFun)
End Function

Public Function FileExist(ByVal sFile As String) As Boolean
On Error GoTo 1
 If FileLen(sFile$) <> 0 Then FileExist = True Else FileExist = False
Exit Function
1
FileExist = False
End Function

Private Function GetRGB(ByVal CVal As Long) As UDT_COLORRGB
 GetRGB.Blue = Int(CVal / 65536)
 GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
 GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function

Public Function HTML2RGB(Optional ByVal sTxt As String) As Long
On Error GoTo 1
 
 sTxt$ = IIf(Left$(sTxt$, 1) = "#", Mid$(sTxt$, 2), sTxt$) & String(6, "0")

HTML2RGB& = RGB(Val("&H" & Mid$(sTxt$, 1, 2)) _
              , Val("&H" & Mid$(sTxt$, 3, 2)) _
              , Val("&H" & Mid$(sTxt$, 5, 2)))
1
End Function

Private Function Insert(ByVal outStr As String, ByVal insStr As String, ByVal lStart As Long) As String
Dim lLen As Long
lLen& = Len(insStr$)
Insert$ = Mid$(outStr, 1, lStart&) & insStr$ & Mid$(outStr$, lStart& + lLen& + 1)
End Function

Private Function isInt(ByVal sTxt As String) As Boolean
'returns if a string is actually an integer or number
isInt = IsNumeric(sTxt$)
End Function

Public Function newSplit(ByVal sTxt As String, ByVal Del As String)
'new split function that ignores any instance of the delimiter
'in between ()'s
Dim Data() As String, i As Integer, k As Integer
Dim sA As String, t As Integer, c As Integer

  k% = 0: ReDim Data(k%) As String
c% = 0
For i% = 1 To Len(sTxt$)

  sA$ = Mid$(sTxt$, i%, 1)

  If sA$ = Chr(34) Then
   If c% = 0 Then c% = 1 Else c% = 0
  End If

  If sA$ = "(" And c = 0 Then
    t% = t% + 1
  ElseIf sA$ = ")" And c = 0 Then
    t% = t% - 1
    If t% < 0 Then t% = 0
  End If
  
    If Mid$(sTxt$, i%, Len(Del$)) = Del$ And t% = 0 And c% = 0 Then
     i% = i% + Len(Del$) - 1
     k% = k% + 1
     ReDim Preserve Data(k%) As String
    Else
     Data$(k%) = Data$(k%) & sA$
    End If

Next i%

newSplit = Data$
End Function

Public Function newTrim(ByVal sTxt As String) As String
'trims space's,tab's and comments
On Error GoTo 1
Dim sA As String, i As Integer, k As Integer, arr() As String
sA$ = Trim$(sTxt$)

    If sA$ = "" Then newTrim$ = "": Exit Function

i% = InStrRev(sA$, "//") ' find comment marker

If i% <> 0 Then
 arr$() = newSplit(sA$, "//")
 sA$ = Trim$(arr$(0)) ' delete comments
End If

    If sA$ = "" Then newTrim$ = "": Exit Function

For i% = 1 To Len(sA) ' loop thru and strip TAB chr from start to end
  If Mid$(sA$, i%, 1) <> Chr(9) Then i% = i% - 1: Exit For
Next i%

  sA$ = Trim$(Right$(sA$, Len(sA$) - i%)) ' strip tab

    If sA$ = "" Then newTrim$ = "": Exit Function

For i% = Len(sA) To 1 Step -1 ' loop thru and strip TAB chr from end to start
  If Mid$(sA$, i%, 1) <> Chr(9) Then i% = i% + 1: Exit For
Next i%

  sA$ = Trim$(Left$(sA$, i% - 1)) ' strip tab
1
   newTrim$ = sA$
End Function

Private Function StripEndCrLf(ByVal s As String) As String
Dim l As Long

For l = Len(s) To 1 Step -1
 If Mid(s, l, 1) <> Chr(13) And Mid(s, l, 1) <> Chr(10) Then Exit For
Next l
'MsgBox Left(s, l)
StripEndCrLf = Left(s, l)
End Function

Private Function retFunction(ByVal sCode As String, ByVal sFun As String) As String
'handles functions and returns it's value
On Error GoTo 1
Dim sTmp As String
Dim arrInfo() As String, arrtmp() As String, arrInfoX() As String
Dim qA As Integer, iA As Integer, iB As Integer
Dim lA As Long, lB As Long, lC As Long
Dim sA As String, sB As String, v As Variant


 iA% = InStr(sCode$, "(")
 iB% = InStrRev(sCode$, ")")

     If iA% = 0 Or iB% = 0 Then GoTo 1

  sA$ = Mid$(sCode$, iA% + 1, iB% - iA% - 1)
'GoTo 1
arrInfo$() = newSplit(sA$, ",")

If Left$(Trim(sCode$), 1) = "&" Then
 sB$ = "do("
 iB% = InStr(sCode$, "(")
  If iB% = 0 Then Call MsgBox("Expected '(' but found nothing!", vbCritical, "Error"): Exit Function
 sB$ = sB$ & Mid$(sCode$, 2, iB% - 2) & ""

 For iB% = LBound(arrInfo$()) To UBound(arrInfo$())
  sB$ = sB$ & "," & arrInfo$(iB%)
 Next iB%
 sB$ = sB$ & ")"
 sCode$ = sB$
 iA% = InStr(sCode$, "(")
 iB% = InStrRev(sCode$, ")")

     If iA% = 0 Or iB% = 0 Then GoTo 1

  sB$ = Mid$(sCode$, iA% + 1, iB% - iA% - 1)
 arrInfo$() = newSplit(sB$, ",")
End If



qA% = UBound(arrInfo$())

For iB% = qA% To qA% + 6
ReDim Preserve arrInfo$(iB%)
Next iB%

ReDim arrtmp(iB%) As String
qA% = iB% - 1

     For iB% = 0 To qA%
       sA$ = arrInfo$(iB%)
         If InStr(sA$, " & ") Then
           arrInfoX$() = newSplit(sA$, " & "): sA$ = ""
           sA$ = combSplit$(arrInfoX(), sFun$)
         End If

       arrInfo$(iB%) = sA$

      If arrInfo$(iB%) <> "" Then
      'MsgBox arrInfo$(iB%)
       If Left$(Trim$(arrInfo$(iB%)), 1) = Chr(34) And Right$(Trim$(arrInfo$(iB%)), 1) = Chr(34) Then
       arrtmp$(iB%) = Mid$(Trim$(arrInfo$(iB%)), 2, Len(Trim$(arrInfo$(iB%))) - 2)
       Else
       arrtmp$(iB%) = retString$(Trim$(arrInfo$(iB%)), sFun$)
       End If
      End If

     Next iB%

Select Case Trim$(LCase$(Left$(sCode$, iA% - 1)))
  Case "cbool"
  'cbool(value)
   retFunction = CBool(arrtmp(0))
    Exit Function
  Case "html2rgb"
  'html2rgb(color)
   retFunction = HTML2RGB(arrtmp(0))
    Exit Function
  Case "cr"
   retFunction = Chr(13)
    Exit Function
  Case "lf"
   retFunction = Chr(10)
    Exit Function
  Case "crlf"
   retFunction = Chr(13) & Chr(10)
    Exit Function
  Case "code_string_gen"
  'code_string_gen(procedure name, code, optional variables)
   sA = "!proc " & arrtmp(0)
   If arrtmp(2) <> "" Then sA = sA & "(" & arrtmp(2) & ")" Else sA = sA & "()"
   sA = sA & vbCrLf & arrtmp(1) & vbCrLf & "end!"
   retFunction = sA
    Exit Function
  Case "code_string"
  'code_string()
   retFunction = sString
    Exit Function
  Case "eval"
  'eval(condition)
   retFunction = IIf(Eval(arrInfo(0), sFun) = "!iftrue!", True, False)
    Exit Function
  Case "fileexist"
  'fileexist(file path)
   retFunction = FileExist(arrtmp(0))
    Exit Function
  Case "apppath"
  'apppath()
   retFunction = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "")
    Exit Function
  Case "appexename"
  'appexename()
   retFunction = App.EXEName
    Exit Function
  Case "use"
  'uses(nll name, files)
    If modLan.FileExist(App.Path & "\" & arrtmp(0) & ".nll") = True Then
     If modLan.FileExist(TEMP_PATH & "\" & arrtmp(0) & ".exe") = False Then _
        CopyFile App.Path & "\" & arrtmp(0) & ".nll", TEMP_PATH & "\" & arrtmp(0) & ".exe", 0
    ElseIf modLan.FileExist(SYS_PATH & "\" & arrtmp(0) & ".nll") = True Then
     If modLan.FileExist(TEMP_PATH & "\" & arrtmp(0) & ".exe") = False Then _
        CopyFile SYS_PATH & "\" & arrtmp(0) & ".nll", TEMP_PATH & "\" & arrtmp(0) & ".exe", 0
    Else
     MsgBox "Error: Can't find Not-So-Dynamic Link Library", vbCritical, "Error #101"
     Exit Function
    End If
  
  sTmp = CStr(Int(32000 * Rnd))
  sA = App.Path & "\" & sTmp & ".dat" & Chr(1) & ",_," & Chr(1)
  For lA = 1 To UBound(arrtmp()) - 6
   sA = sA & arrtmp(lA) & Chr(1) & ",_," & Chr(1)
  Next lA
  sA = Left(sA, Len(sA) - 5)
  Dim ret&

  ret& = ShellExecute(0, vbNullString, TEMP_PATH & "\" & arrtmp(0) & ".exe", sA, "c:\", 1)

   Do Until FileExist(App.Path & "\" & sTmp & ".dat") = True
   DoEvents
   Loop

   lB = FreeFile()
   
    Open App.Path & "\" & sTmp & ".dat" For Binary Access Read As #lB
     sA = Input(LOF(lB), #lB)
    Close #lB
    'Pause 1
   
    Call Kill(App.Path & "\" & sTmp & ".dat")
    
'    Call Kill(TEMP_PATH & "\" & arrtmp(0) & ".exe")
    
    retFunction = "hello" 'StripEndCrLf(sA)
    Exit Function
  Case "asc"
  'asc(string)
    retFunction$ = CStr(Asc(arrtmp$(0)))
    Exit Function
  Case "chr"
  'chr(int)
    retFunction$ = CStr(Chr(CInt(IIf(arrtmp$(0) <> "", arrtmp$(0), 0))))
    Exit Function
  Case "getref"
  'getref(variable name, function)
    If arrtmp(1) = "" Then arrtmp(1) = retString("$_", sFun)
    retFunction = retString(arrtmp(0), arrtmp(1))
    Exit Function
  Case "count"
  'count(array)
   iB = 0
   If Left(arrInfo(0), 1) = ">" Then arrInfo(0) = retString(Mid(arrInfo(0), 2), sFun): sFun = retString("$_", sFun)
    For iA = 1 To Strings.Name.Count
'fix this, when types are used it counts each type in the array
     If Left(Strings.Name(iA), Len(arrInfo(0)) + 1) = arrInfo(0) & "[" And IIf(Left(arrInfo(0), 1) <> "@", sFun, "") = IIf(Left(arrInfo(0), 1) <> "@", Strings.Fun(iA), "") Then iB = iB + 1
    Next iA
    If iB = 0 Then iB = -1
    If IsNumeric(arrtmp(1)) = True And iB <> -1 Then iB = iB - CInt(arrtmp(1))
    retFunction = CStr(iB)
    Exit Function
  Case "do"
  'do(procedure,args)
  Dim arrFun() As String, arrInfoY() As String
   lA& = InStr(LCase$(sString$), LCase$("!proc " & arrtmp$(0) & "("))
   lB& = InStr(lA& + 1, sString$, ")" & Chr(13))

    If lA& = 0 Or lB& = 0 Then GoTo 2

      sA$ = Mid$(sString$, lA& + Len("!proc " & arrtmp$(0) & "("), lB& - lA& - Len("!proc " & arrtmp$(0) & "("))

    arrFun$() = newSplit(sA$, ",")

        iA% = 1
        For Each v In arrFun$()
          If v <> "" Then
           Call doCode("var " & Trim(v), arrtmp(0))
           If InStr(arrtmp$(iA%), " & ") Then
             arrInfoY$() = newSplit(arrtmp$(iA%), " & ")
             arrtmp$(iA%) = combSplit$(arrInfoY(), sFun$)
           End If
            If Left$(arrtmp$(iA%), 1) = Chr(34) And Right$(arrtmp$(iA%), 1) = Chr(34) Then
             sA$ = Mid$(arrtmp$(iA%), 2, Len(arrtmp$(iA%)) - 2)
            Else
             sA$ = retString$(Trim$(arrtmp$(iA%)), sFun$)
            End If
            iA% = iA% + 1
            If Left(Trim$(CStr(v)), 1) = ">" Then
             Call setString(Mid(Trim$(CStr(v)), 2), "", arrtmp$(0), "", arrInfo(iA - 1), sFun)
            Else
             If InStr(v, "=") <> 0 And sA <> "" Then
              Call setString(Trim$(Trim(Left(CStr(v), InStr(v, "=") - 1))), sA$, arrtmp$(0))
             ElseIf InStr(v, "::") <> 0 And InStr(v, Chr(34)) = 0 Then
              Call setString(Trim$(Trim(Left(CStr(v), InStr(v, "::") - 1))), sA$, arrtmp$(0))
             Else
              Call setString(Trim$(CStr(v)), sA$, arrtmp$(0))
             End If
            End If
          End If
        Next v
2

    sA$ = ""
    Call setString("$_", sFun, arrtmp$(0))
    sA$ = Execute(sString$, arrtmp$(0))
    
    Do Until sA$ <> "": DoEvents: Loop
    retFunction$ = sA$
    Exit Function
  Case "split"
  'split(array, string, delimeter)
    arrInfoX() = Split(arrtmp(1), arrtmp(2))
     If IsNumeric(arrInfo(0)) = True Then
      retFunction = arrInfoX(CLng(arrtmp(0)))
     Else
      For lA = 0 To UBound(arrInfoX)
       Call setString(Trim(arrInfo(0)) & "[" & lA & "]", arrInfoX(lA), sFun)
      Next lA
      retFunction = lA
     End If
    Exit Function
  Case "strdel"
  'strdel(var,start,len)
    retFunction$ = Left$(arrtmp$(0), CLng(IIf(arrtmp$(1) <> "", arrtmp$(1), 1))) & Mid(arrtmp$(0), CLng(IIf(arrtmp$(2) <> "", CLng(arrtmp$(2)) + CLng(IIf(arrtmp$(1) <> "", arrtmp$(1), 1)) + 1, 1)))
    Exit Function
  Case "hex"
  'hex(int)
    retFunction$ = CStr(Hex(CLng(IIf(arrtmp$(0) <> "", arrtmp$(0), 0))))
    Exit Function
  Case "iif"
  'iif(condition, true, false)
  'MsgBox Eval(arrtmp$(0), sFun$) & " - " & arrInfo$(0)
    If Eval(arrInfo$(0), sFun$) = "!iftrue!" Then retFunction$ = arrtmp$(1) Else retFunction$ = arrtmp$(2)
    Exit Function
  Case "implode"
  'implode(array, delimeter)
  If Left(arrInfo(0), 1) = ">" Then arrInfo(0) = retString(Mid(arrInfo(0), 2), sFun): sFun = retString("$_", sFun)
    For iA = 1 To Strings.Name.Count
     If Left(Strings.Name(iA), Len(arrInfo(0)) + 1) = arrInfo(0) & "[" And IIf(Left(arrInfo(0), 1) <> "@", sFun, "") = IIf(Left(arrInfo(0), 1) <> "@", Strings.Fun(iA), "") Then
      sA = sA & Strings.Value(iA) & arrtmp(1)
     End If
    Next iA
    retFunction = Left(sA, Len(sA) - Len(arrtmp(1)))
    Exit Function
  Case "in_array"
  'in_array(array, string, optional case_sensitive boolean)
  iB = Strings.Name.Count
  If Left(arrInfo(0), 1) = ">" Then arrInfo(0) = retString(Mid(arrInfo(0), 2), sFun): sFun = retString("$_", sFun)
    For iA = 1 To iB
     If iA > iB Then Exit For
     If Left(Strings.Name(iA), Len(arrInfo(0)) + 1) = arrInfo(0) & "[" And IIf(arrtmp(2) = "true", LCase(Strings.Value(iA)), Strings.Value(iA)) = arrtmp(1) And IIf(Left(arrInfo(0), 1) <> "@", sFun, "") = IIf(Left(arrInfo(0), 1) <> "@", Strings.Fun(iA), "") Then
      retFunction = "True"
      Exit Function
     End If
    Next iA
    retFunction = "False"
    Exit Function
  Case "strins"
  'insert(string,instring,start)
  'MsgBox arrtmp$(0) & " - " & arrtmp$(1) & " - " & arrtmp$(2)
    retFunction$ = Insert$(arrtmp$(0), arrtmp$(1), CInt(IIf(arrtmp$(2) <> "", arrtmp$(2), 1)))
    Exit Function
  Case "int"
  'int(string or single)
    retFunction$ = CStr(Int(IIf(arrtmp$(0) <> "", arrtmp$(0), 0)))
    Exit Function
  Case "strlen"
  'length(string)
    retFunction$ = CStr(Len(arrtmp$(0)))
    Exit Function
  Case "math"
  'math(equation)
    retFunction$ = CStr(retMath$(arrInfo$(0), sFun$))
    Exit Function
  Case "msgbox", "messagebox"
  'msgbox(prompt, style int, title)
    retFunction$ = CStr(MsgBox(arrtmp$(0), CInt(IIf(arrtmp$(1) <> "", arrtmp$(1), 0)), arrtmp$(2)))
    Exit Function
  Case "openfile"
  'openfile(file path)
    If arrtmp$(0) <> "" Then
     If LCase$(arrtmp$(0)) = LCase$(App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "info.dat") Then Exit Function
     If FileExist(arrtmp$(0)) = False Then Exit Function
      Open arrtmp$(0) For Input As #1
        sTmp$ = Input(LOF(1), #1)
      Close #1
    End If
    retFunction$ = sTmp$
    Exit Function
  Case "prompt"
  'prompt(prompt, title, default)
    retFunction$ = CStr(InputBox(arrtmp$(0), arrtmp$(1), arrtmp$(2)))
    Exit Function
  Case "replace"
  'replace(string,find,replace with)
    retFunction = Replace(arrtmp(0), arrtmp(1), arrtmp(2))
    Exit Function
  Case "replacea"
  'replacea(string,find array,replace with array)
    arrInfo(1) = Trim(arrInfo(1))
    arrInfo(2) = Trim(arrInfo(2))
    lA = CLng(retFunction("count(" & arrInfo(1) & ")", sFun))
    sA = arrtmp(0)
    For iA = 0 To lA - 1
     sA = Replace(sA, retString(arrInfo(1) & "[" & iA & "]", sFun), retString(arrInfo(2) & "[" & iA & "]", sFun))
    Next iA
     retFunction = sA
    Exit Function
  Case "rgb"
  'rgb(int, int, int)
    retFunction = CStr(RGB(CInt(IIf(arrtmp(0) <> "", arrtmp(0), 0)), CInt(IIf(arrtmp(1) <> "", arrtmp(1), 0)), CInt(IIf(arrtmp(2) <> "", arrtmp(2), 0))))
    Exit Function
  Case "rnd", "random"
  'rnd(int)
  Call Randomize
    retFunction = CStr(Rnd * (CSng(IIf(arrtmp(0) <> "", arrtmp(0), 0)) + 1))
    Exit Function
  Case "strreverse"
   retFunction = StrReverse(arrtmp(0))
    Exit Function
  Case "strpos"
  'strpos(search str, search for str,start int)
   If arrtmp(1) = "s" Then sA = "d" Else sA = "s"
   If arrtmp(2) <> "" Then
    lA = InStr(CLng(arrtmp(2)) + 1, sA & arrtmp(0), arrtmp(1))
   Else
    lA = InStr(sA & arrtmp(0), arrtmp(1))
   End If
    retFunction = CStr(lA - 1)
    Exit Function
  Case "strposrev"
  'strposrev(search str, search for str,start int)
   If arrtmp(1) = "s" Then sA = "d" Else sA = "s"
   If arrtmp(2) <> "" Then
    lA = InStrRev(sA & arrtmp(0), arrtmp(1), CLng(arrtmp(2)))
   Else
    lA = InStrRev(sA & arrtmp(0), arrtmp(1))
   End If
    retFunction = CStr(lA - 1)
    Exit Function
  Case "substr"
  'substr(string, start int, length int)
    If arrtmp(2) = "" Then
     retFunction$ = Mid$(arrtmp$(0), CLng(IIf(arrtmp$(1) <> "", arrtmp$(1), 1)))
    Else
     retFunction$ = Mid$(arrtmp$(0), CLng(IIf(arrtmp$(1) <> "", arrtmp$(1), 1)), CLng(IIf(arrtmp$(2) <> "", arrtmp$(2), 1)))
    End If
    Exit Function
  Case "left"
  'left(string, length int)
    If IsNumeric(arrtmp(1)) = False Then
     retFunction$ = Left$(arrtmp$(0), InStr(arrtmp(0), arrtmp(1)))
    Else
     retFunction$ = Left$(arrtmp$(0), CLng(IIf(arrtmp$(1) <> "", arrtmp$(1), 1)))
    End If
    Exit Function
  Case "right"
  'right(string, length int)
    If IsNumeric(arrtmp(1)) = False Then
     retFunction$ = Right$(arrtmp$(0), InStr(arrtmp(0), arrtmp(1)))
    Else
     retFunction$ = Right$(arrtmp$(0), CLng(IIf(arrtmp$(1) <> "", arrtmp$(1), 1)))
    End If
    Exit Function
  Case "string"
  'string(length,charachter)
    retFunction$ = String$(CLng(IIf(IsNumeric(arrtmp$(1)) = False, 0, arrtmp$(1))), arrtmp$(0))
    Exit Function
  Case "trim"
  'trim(string)
    retFunction$ = Trim$(arrtmp$(0))
    Exit Function
  Case "time"
  'time(index)
    Select Case arrtmp(0)
     Case "s"
      retFunction$ = Second(Now)
     Case "m"
      retFunction$ = Minute(Now)
     Case "h"
      retFunction$ = Hour(Now)
     Case Else
      retFunction$ = IIf(arrtmp$(0) = "" Or arrtmp$(0) = "0", Time$, Time)
    End Select
    Exit Function
  Case "date"
  'date(index)
    Select Case arrtmp(0)
     Case "d"
      retFunction$ = Day(Now)
     Case "m"
      retFunction$ = Month(Now)
     Case "y"
      retFunction$ = Year(Now)
     Case Else
      retFunction$ = IIf(arrtmp$(0) = "" Or arrtmp$(0) = "0", Date$, Date)
    End Select
    Exit Function
  Case "ucase"
  'ucase(string)
    retFunction$ = UCase$(arrtmp$(0))
    Exit Function
  Case "lcase"
  'lcase(string)
    retFunction$ = LCase$(arrtmp$(0))
    Exit Function
  Case "val"
  'val(string)
    retFunction$ = CStr(Val(arrtmp$(0)))
    Exit Function
  Case "isnumeric"
  'isnumeric(string)
    retFunction$ = CStr(IsNumeric(arrtmp$(0)))
    Exit Function
  Case ""
  'math(equation)
    retFunction$ = CStr(retMath$(arrInfo$(0), sFun$))
    Exit Function

'GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
'WritePrivateProfileString(lpApplicationName, lpKeyName, lpString, lpFileName)
'SetWindowPos(hwnd, hWndInsertAfter, X, Y, cx, cy, wFlags)
End Select
1
 retFunction$ = sCode$
End Function

Private Function retMath(ByVal sCode As String, ByVal sFun As String) As String
'returns math equations
On Error GoTo 1
Dim arrMath() As String, v As Variant
Dim iI As Integer, iA As Integer, sA As String
Dim sLast As String, sLm As String

arrMath$() = newSplit(sCode$, " ")

  For iI% = LBound(arrMath$()) To UBound(arrMath$())
     sA$ = retString$(arrMath$(iI%), sFun$)
    If sA$ <> "" Then
      If iA% = 0 Then
         If iI% = LBound(arrMath$()) Then
          sLast$ = Trim$(sA$)
         Else
            sA$ = Trim$(sA$)
          Select Case LCase(sLm$)
            Case "+"
                sLast$ = CLng(sLast$) + CLng(sA$)
            Case "++"
                sLast$ = CSng(sLast$) + CSng(sA$)
            Case "-"
                sLast$ = sLast$ - sA$
            Case "*"
                sLast$ = sLast$ * sA$
            Case "/"
                sLast$ = sLast$ / sA$
            Case "\"
                sLast$ = sLast$ \ sA$
            Case "mod"
                sLast$ = sLast$ Mod sA$
            Case "^"
                sLast$ = sLast$ ^ sA$
            Case "and"
                sLast$ = sLast$ And sA$
            Case "or"
                sLast$ = sLast$ Or sA$
            Case "xor"
                sLast$ = sLast$ Xor sA$
          End Select
         End If
        iA% = 1
      Else
            sLm$ = retString$(Trim$(sA$), "")
            'MsgBox sLm
        iA% = 0
      End If
    End If
  Next iI%

retMath$ = sLast$
Exit Function
1
retMath$ = "0"
End Function

Public Function retString(ByVal sName As String, ByVal sFun As String) As String
'returns either a var's value, functions value or an objects info
'On Error Resume Next
Dim IsThere As Boolean, lA As Long, sA As String, sC As String
Dim arrInfoX() As String, lB As Long, lC As Long, sB As String

  If sName$ = "" Then Exit Function

 If Left$(sName$, 1) = "$" Or Left$(sName$, 1) = "@" Then
 
    If InStr(sName$, "[") <> 0 And InStr(InStr(sName$, "[") + 1, sName$, "]") <> 0 Then
    'array
      lA& = InStr(sName$, "[")
      lB& = InStr(lA& + 1, sName$, "]")
      sB = Left(sName, lA - 1)
        
        sA$ = Trim$(Mid$(sName$, lA& + 1, lB& - lA& - 1))
        sName$ = sB & "[" & retString(sA$, sFun$) & Mid$(sName$, lB&)
    End If
    
 lB = -1: lC = -1: sA = "0"
 If sC <> "" Then sFun = sC
 If InStr(sName, "{") <> 0 And InStr(sName, "}") <> 0 Then
   sA = Mid(sName, InStr(sName, "{") + 1, InStrRev(sName, "}") - InStr(sName, "{") - 1)

  If InStr(sA, ">") Then
   lB = CLng(retString(Trim(Left(sA, InStr(sA, ">") - 1)), sFun))
   lC = CLng(retString(Trim(Mid(sA, InStr(sA, ">") + 1)), sFun))
  Else
   lB = CLng(retString(Trim(sA), sFun))
  End If
   sName = Left(sName, InStr(sName, "{") - 1)
 End If

  For lA& = 1 To Strings.Name.Count
    If Left$(sName$, 1) = "$" Then
      If LCase$(Strings.Name.Item(lA&)) = LCase$(sName$) And LCase$(Strings.Fun.Item(lA&)) = LCase$(sFun$) Then IsThere = True: Exit For Else IsThere = False
      If InStr(sName, ".") <> 0 Then
       If LCase$(Strings.Name.Item(lA&)) = LCase$(Left(sName$, InStr(sName, ".") - 1)) And LCase$(Strings.Fun.Item(lA&)) = LCase$(sFun$) Then IsThere = True: Exit For Else IsThere = False
      End If
    Else
      If LCase$(Strings.Name.Item(lA&)) = LCase$(sName$) Then IsThere = True: Exit For Else IsThere = False
    End If
  Next lA&

    If IsThere = True Then
      If lB <> -1 Then
       If lC <> -1 Then
        retString$ = Mid(CStr(Strings.Value.Item(lA&)), lB, lC)
       Else
        retString$ = Mid(CStr(Strings.Value.Item(lA&)), lB, 1)
       End If
      Else
       If Strings.Alias(lA) <> "" And Strings.Object(lA) = "" Then
        retString = retString(Strings.Alias(lA), Strings.FAlias(lA))
       ElseIf Strings.Object(lA) <> "" Then
        'MsgBox retString(Left(sName$, InStr(sName, ".") - 1), sFun) 'sName
        ', sFun) & Mid(sName, InStr(sName, "."))
       Else
        retString$ = CStr(Strings.Value.Item(lA&))
       End If
      End If
    Else
      Call DoError(sFun & " > " & sName, 901)
      retString$ = "" 'sName$
    End If
    
   'GOOD \/
ElseIf Left$(sName$, 1) = "%" Then
    sA$ = Trim$(Right$(sName$, Len(sName$) - 1))
    sA = Replace(sA, "$_", retString("$_", sFun))
    
     arrInfoX$() = Split(sA$, ".")

      lA& = winIndex%(Trim$(CStr(retString$(arrInfoX$(0), sFun$))))
       If lA& <> -1 Then
          lB& = ctrlIndex%(Forms(lA&), Trim$(CStr(retString$(arrInfoX$(1), sFun$))))
            If lB& <> -1 Then
             If InStr(retString$(arrInfoX(2), sFun$), "(") <> 0 And InStr(retString$(arrInfoX(2), sFun$), ")") <> 0 Then
             Dim sF$
               sF$ = Mid$(retString$(arrInfoX$(2), sFun$), InStr(retString$(arrInfoX(2), sFun$), "(") + 1, InStr(retString$(arrInfoX(2), sFun$), ")") - InStr(retString$(arrInfoX(2), sFun$), "(") - 1)
               retString$ = setObject(Forms(lA&).Controls(lB&), Trim$(CStr(Left$(retString$(arrInfoX$(2), sFun$), InStr(retString$(arrInfoX(2), sFun$), "(") - 1))), retString$(sF$, sFun$), "", lA&, lB&, 1)
             Else
               retString$ = setObject(Forms(lA&).Controls(lB&), Trim$(CStr(retString$(arrInfoX$(2), sFun$))), "", "", lA&, lB&, 1) ' retString$(sA$, sFun$), sFun$, 1)
             End If
            End If
       End If: Exit Function
 ElseIf Left(sName, 2) = ">>" And Mid(sName, 3, 1) = Chr(34) And Right(sName, 1) = Chr(34) Then
  'arrInfoX() = Split(Mid(sName, 4, Len(sName) - 5), "$")
  sName = Mid(sName, 4, Len(sName) - 4)
  For lA& = 1 To Strings.Name.Count
   If Left$(Strings.Name(lA), 1) = "$" And LCase$(Strings.Fun.Item(lA&)) = LCase$(sFun$) Or Left$(Strings.Name(lA), 1) = "@" Then
    sName = Replace(sName, Strings.Name(lA), retString(Strings.Name(lA), sFun), , , vbTextCompare)
   End If
  Next lA&
  retString = sName
 Else
   If isInt(sName$) = False Then retString$ = retFunction(sName$, sFun$) Else retString$ = sName$
 End If

End Function

Public Function Rgb2Html(ByVal l As Long) As String
Rgb2Html$ = IIf(Len(Hex(GetRGB(l&).Red)) = 1, "0" & Hex(GetRGB(l&).Red), Hex(GetRGB(l&).Red)) & IIf(Len(Hex(GetRGB(l&).Green)) = 1, "0" & Hex(GetRGB(l&).Green), Hex(GetRGB(l&).Green)) & IIf(Len(Hex(GetRGB(l&).Blue)) = 1, "0" & Hex(GetRGB(l&).Blue), Hex(GetRGB(l&).Blue))
End Function

Private Function setObject(obj As Object, ByVal sProp As String, ByVal sValue As String, ByVal sFun As String, frmIndex As Long, conIndex As Long, Optional ByVal Index As Integer = 0) As String
'actually set's or returns an objects value as requested
On Error GoTo 1

Dim arrInfoX() As String, sA As String
Dim v As Variant, i As Integer, sX As String

 If InStr(sValue$, " & ") Then
  sA$ = sValue$
  arrInfoX$() = newSplit(sA$, " & "): sA$ = ""
  sValue$ = combSplit$(arrInfoX$(), sFun$)
  sValue$ = Mid$(sValue$, 2, Len(sValue$) - 2)
 End If

 If InStr(sProp$, "(") <> 0 Then
  sX$ = Mid$(sProp$, InStr(sProp$, "(") + 1, InStr(sProp$, ")") - InStr(sProp$, "(") - 1)
  sProp$ = Left$(sProp$, InStr(sProp$, "(") - 1)
 End If

Select Case Trim$(LCase$(sProp$))
  Case "add"
    If Index% = 0 Then
     Call obj.AddItem(sValue$)
      If obj.Style = 1 Then
       For i% = 0 To obj.ListCount
         If obj.List(i%) = sValue$ Then obj.Selected(i%) = True
       Next i%
      End If
    End If
  Case "backcolor"
    If Index% = 0 Then obj.BackColor = HTML2RGB&(sValue$) Else setObject$ = Rgb2Html$(obj.BackColor)
  Case "forecolor"
    If Index% = 0 Then obj.ForeColor = HTML2RGB&(sValue$) Else setObject$ = Rgb2Html$(obj.ForeColor)
  Case "caption"
    If Index% = 0 Then obj.Caption = sValue$ Else setObject$ = CStr(obj.Caption)
  Case "clear"
   obj.Clear
  Case "connect"
    If Index% = 0 Then obj.Connect
  Case "close"
  'MsgBox Forms(frmIndex&).sckNew(CInt(Forms(frmIndex&).Controls(conIndex&).Index%)).Tag
    'Call Pause(2)
    'Call Forms(frmIndex&).sckNew(CInt(Forms(frmIndex&).Controls(conIndex&).Index%)).Close
        'If InStr(LCase$(sA$), "</html>") Then
    Call Forms(frmIndex&).CloseD(CInt(Forms(frmIndex&).Controls(conIndex&).Index%))
  Case "enabled"
    If Index% = 0 Then obj.Enabled = CBool(sValue$) Else setObject$ = CStr(obj.Enabled)
  Case "fontsize"
    If Index% = 0 Then obj.FontSize = CInt(IIf(sValue$ <> "", sValue$, 8)) Else setObject$ = CStr(obj.FontSize)
  Case "height"
    If Index% = 0 Then obj.Height = (CInt(IIf(sValue$ <> "", sValue$, 100)) * Screen.TwipsPerPixelY) Else setObject$ = CStr(obj.Height)
  Case "interval"
    If Index% = 0 Then obj.Interval = CInt(IIf(sValue$ <> "", sValue$, 0)) Else setObject$ = CStr(obj.Interval)
  Case "left"
    If Index% = 0 Then obj.Left = (CInt(IIf(sValue$ <> "", sValue$, 0)) * Screen.TwipsPerPixelX) Else setObject$ = CStr(obj.Left)
  Case "list"
    If Index% <> 0 Then setObject$ = CStr(obj.List(CInt(IIf(sValue$ <> "", sValue$, 0))))
  Case "listindex"
    If Index% = 0 Then obj.ListIndex = CInt(IIf(sValue$ <> "", sValue$, 0)) Else setObject$ = obj.ListIndex
  Case "listcount"
    If Index% <> 0 Then setObject$ = obj.ListCount
  Case "locked"
    If Index% = 0 Then obj.Locked = CBool(IIf(sValue$ <> "", sValue$, True)) Else setObject$ = CStr(obj.Locked)
  Case "picture"
    If Index% = 0 Then
      sX = sValue
      For i = 0 To Forms.Count - 1 ' UBound(arrImage()) - 1
       If Forms(i).Name = "frmImg" And InStr(Forms(i).Tag, Chr(0)) <> 0 Then
        If sValue = Left(Forms(i).Tag, InStr(Forms(i).Tag, Chr(0)) - 1) Then
       ' sValue = arrImage(1, i): Exit For Else sValue = ""
         Set obj.Picture = Forms(i).imgMain.Picture
         Exit For
        End If
       End If
      Next i
    Else
     setObject$ = CStr(obj.Picture)
    End If
  Case "remotehost"
   If Index% = 0 Then obj.RemoteHost = sValue$ Else setObject$ = CStr(obj.RemoteHost)
  Case "pattern"
   If Index% = 0 Then obj.Pattern = sValue$ Else setObject$ = CStr(obj.Pattern)
  Case "path"
   If Index% = 0 Then obj.Path = sValue$ Else setObject$ = CStr(obj.Path)
  Case "print"
   If Index% = 0 Then obj.Print sValue$: obj.Refresh
  Case "refresh"
   If Index% = 0 Then obj.Refresh
  Case "remoteport"
   If Index% = 0 Then obj.RemotePort = CInt(IIf(sValue$ <> "", sValue$, 0)) Else setObject$ = CStr(obj.RemotePort)
  Case "remove"
    If Index% = 0 Then Call obj.RemoveItem(CInt(IIf(sValue$ <> "", sValue$, 0)))
  Case "selected"
    If Index% = 0 Then obj.Selected(CInt(sX$)) = CBool(sValue$) Else setObject$ = obj.Selected(CInt(sValue$))
  Case "senddata"
    If Index% = 0 Then Call Forms(frmIndex&).Send(sValue$, CInt(Forms(frmIndex&).Controls(conIndex&).Index%))
  Case "tooltip"
    If Index% = 0 Then obj.ToolTipText = sValue$ Else setObject$ = CStr(obj.ToolTipText)
  Case "text"
    If Index% = 0 Then obj.Text = sValue$ Else setObject$ = CStr(obj.Text)
  Case "top"
    If Index% = 0 Then obj.Top = (CInt(IIf(sValue$ <> "", sValue$, 0)) * Screen.TwipsPerPixelY) Else setObject$ = CStr(obj.Top)
  Case "value"
   If LCase$(sValue$) = "true" Then sValue$ = 1 Else sValue$ = 0
   sValue$ = CInt(IIf(sValue$ <> "", sValue$, 0))
    If Index% = 0 Then obj.Value = IIf(obj.Name = "optNew", CBool(sValue$), sValue$) Else setObject$ = CStr(obj.Value)
  Case "visible"
    If Index% = 0 Then obj.Visible = CBool(sValue$) Else setObject$ = CStr(obj.Visible)
  Case "width"
    If Index% = 0 Then obj.Width = (CInt(IIf(sValue$ <> "", sValue$, 100)) * Screen.TwipsPerPixelX) Else setObject$ = CStr(obj.Width)
End Select
'Exit Function
1
'Call MsgBox("Error : " & Obj.Tag & " does not support the method " & sProp$, vbCritical, "Error")
End Function

Public Sub setString(ByVal sName As String, ByVal sValue As String, ByVal sFun As String, Optional ByVal sType As String, Optional sAlias As String, Optional sFAlias As String, Optional sObject As String)
'set's a var's value
Dim IsThere As Boolean, lA As Long

 If sName$ = "" Then Exit Sub

 If Left$(Trim(sName$), 1) = "$" Or Left$(Trim(sName$), 1) = "@" Then

  For lA& = 1 To Strings.Name.Count
    If Left$(sName$, 1) = "$" Then
      If LCase$(Strings.Name.Item(lA&)) = LCase$(sName$) And LCase$(Strings.Fun.Item(lA&)) = LCase$(sFun$) Then IsThere = True: Exit For Else IsThere = False
    Else
      If LCase$(Strings.Name.Item(lA&)) = LCase$(sName$) Then sFun$ = "": IsThere = True: Exit For Else IsThere = False
    End If
  Next lA&

  If IsThere = True Then
   If Strings.Alias(lA) <> "" Then
    sName = Strings.Alias(lA)
    sFun = Strings.FAlias(lA)
    Call doCode(sName & " = " & sValue, sFun)
    'Call setString(sName, sValue, sFun)
    Exit Sub
   End If
   Call delString(sName$, sFun$, lA&)
  End If

'MsgBox sName & " - " & sFun & sValue
      Call Strings.Name.Add(sName$)
      Call Strings.Fun.Add(sFun$)
      Call Strings.Value.Add(sValue$)
      Call Strings.Type.Add(sType$)
      Call Strings.Alias.Add(sAlias)
      Call Strings.FAlias.Add(sFAlias)
      Call Strings.Object.Add(sObject)
 End If
End Sub

Private Function retType(ByVal sName As String, ByVal sFun As String) As String
Dim l As Long

For l = 1 To Strings.Name.Count
 If LCase(Strings.Name(l)) = LCase(sName) And LCase(Strings.Fun(l)) = LCase(sFun) Then retType = Strings.Type(l): Exit Function
Next l
End Function

Public Sub setType(ByVal sName As String, ByVal sType As String, ByVal sFun As String, Optional sVar As String, Optional sFunX As String)
Dim lA As Long, lB As Long, a As String
a = sType

lA = InStr(LCase(modLan.sString), LCase("!type " & sType))
lB = InStr(lA + 1, LCase(modLan.sString), LCase("end!"))

If lA = 0 Or lB = 0 Then Exit Sub

sType = Mid(modLan.sString, InStr(lA + 1, modLan.sString, vbCrLf) + 2, lB - InStr(lA + 1, modLan.sString, vbCrLf) - 2)
Dim arr() As String, v As Variant

arr() = Split(sType, vbCrLf)
For Each v In arr()
 v = CVar(newTrim(CStr(v)))
 If v <> "" Then
  lA = InStr(v, "=")
   If lA <> 0 Then
    sType = Trim(Mid(v, lA + 1))
    If Left(sType, 1) = Chr(34) And Right(sType, 1) = Chr(34) Then
     sType = Mid(sType, 2, Len(sType) - 2)
     Call setString(sName & "." & Trim(Left(v, lA - 1)), sType, sFun)
    Else
     Call setString(sName & "." & Trim(Left(v, lA - 1)), sType, sFun)
    End If
   Else
    Call setString(sName & "." & Trim(Left(v, lA - 1)), "", sFun)
   End If
 End If
Next v
Call setString(sName, "", sFun, a)
End Sub

Public Function SynChk(ByVal sTxt As String) As String
Dim i%, sA$, c%, t%, s$, v As Variant, arr$(), q&, sProc$, k&
Dim w As Integer, p As String, l As Integer, z As String, b As Integer

arr$() = newSplit(sTxt$, vbCrLf)

For Each v In arr$()
q& = q& + 1
c% = 0: t% = 0
v = CVar(newTrim$(v))

If LCase$(Left$(v, 5)) = "!proc" Then p = sProc: sProc$ = Trim$(Mid$(v, 6, InStrRev(v, "(") - 6)): k& = q&: w = w + 1
If LCase$(Left$(v, 4)) = "end!" Then w = w - 1
If LCase$(Left$(v, 6)) = "while(" Then z = sProc: l = l + 1
If LCase$(Left$(v, 5)) = "loop(" Then l = l - 1

 For i% = 1 To Len(v)
  sA$ = Mid$(v, i%, 1)
  If sA$ = Chr(34) Then
   If c% = 0 Then c% = 1 Else c% = 0
  End If

  If sA$ = "(" And c = 0 Then
    t% = t% + 1
  ElseIf sA$ = ")" And c = 0 Then
    t% = t% - 1
    If t% < 0 Then t% = 0
  End If
 
  If sA$ = "{" And c = 0 Then
    b = b + 1
  ElseIf sA$ = "}" And c = 0 Then
    b = b - 1
    If b < 0 Then b = 0
  End If
 
 Next i%
 
If c <> 0 Then SynChk$ = "q" & ">> Procedure: " & sProc$ & vbCrLf & ">> Line Num: " & (q& - k&) + 1: Exit Function
If b <> 0 Then SynChk$ = "b" & ">> Procedure: " & sProc$ & vbCrLf & ">> Line Num: " & (q& - k&) + 1: Exit Function
If t <> 0 Then SynChk$ = "p" & ">> Procedure: " & sProc$ & vbCrLf & ">> Line Num: " & (q& - k&) + 1: Exit Function
If w > 1 Then SynChk$ = "e" & ">> Procedure: " & p & vbCrLf & ">> Line Num: " & (q& - k&) + 1: Exit Function

DoEvents
Next v

If l > 0 Then SynChk$ = "l" & ">> Procedure: " & z & vbCrLf: Exit Function

SynChk$ = "a"
End Function

Private Function winIndex(ByVal winName As String) As Integer
'returns a form's index in "forms()"
Dim i As Integer, IsThere As Boolean

  For i% = 0 To Forms.Count - 1
     If LCase$(Forms(i%).Tag) = LCase$(winName$) Then IsThere = True: Exit For Else IsThere = False
  Next i%

   If IsThere = True Then winIndex% = i% Else winIndex% = -1: Call DoError(winName, 301)

End Function

Public Sub Pause(ByVal duration As Long)
'this will Pause your program for givin amount of time
'Usage Pause 1
    
    Dim Current As Long
    Current& = Timer
    Do Until Timer - Current& >= duration&
       DoEvents
    Loop
    
End Sub

