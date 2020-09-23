Attribute VB_Name = "modColor"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const EM_CHARFROMPOS = &HD7

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const LB_FINDSTRING = &H18F

Private Type UDT_PRINTTEXT
 SelBold As Boolean
 SelColor As Long
 SelItalic As Boolean
 SelFade As Long
 SelFont As String
 SelFontSize As Long
 SelStrikeThru As Boolean
 SelUnderLine As Boolean
End Type

Public Type UDT_COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type

Dim m_IgnoreURL As Boolean, m_IgnoreStyles As Boolean
Dim m_IgnoreFade As Boolean, m_IgnoreAlt As Boolean
Dim m_IgnoreFont As Boolean, m_IgnoreTags As Boolean
Dim m_ptLast As UDT_PRINTTEXT, m_LastUrl As String
Dim m_IsMouseRTB(1) As Boolean

Private Function AltYahooStyle(sAlt As String, sColors As String) As String
' fading yahoo style
' sColors must not be parsed
Dim arrColors() As String
Dim iI As Integer, sQ As String
Dim iK As Integer, iAlt As Integer
Dim sA As String

arrColors$() = Split(Trim$(sColors$), ",") ' get colors

    For iI% = LBound(arrColors$()) To UBound(arrColors$())
        arrColors$(iI%) = CStr(Trim$(arrColors$(iI%))) ' make sure the colors ae just the colors
    Next iI%

For iK% = 1 To Len(sAlt$)

    If Mid$(sAlt$, iK%, 1) = Chr(27) Then
        sA$ = sA$ & Mid$(sAlt$, iK%, InStr(iK% + 1, sAlt$, "m") - iK% + 1)

        iK% = IIf(InStr(iK% + 1, sAlt$, "m") = 0 >= Len(sAlt$), Len(sAlt$), InStr(iK% + 1, sAlt$, "m"))
    Else
        sA$ = sA$ & Chr(27) & "[" & arrColors$(iAlt%) & "m" & Mid$(sAlt$, iK%, 1)
    End If

    iAlt% = iAlt% + 1
    If iAlt% > UBound(arrColors$()) Then
        iAlt% = 0
    End If
Next iK%

AltYahooStyle$ = sA$

End Function

Private Sub CheckRTBPosURL(X As Single, Y As Single, rtb1 As RichTextBox) 'change cursor
Dim lP As POINTAPI
Dim iI As Long, lK As Long
Dim iA As Long, iB As Integer
Dim sTemp As String
    
    lP.X = X \ Screen.TwipsPerPixelX
    lP.Y = Y \ Screen.TwipsPerPixelY
    
        lK& = SendMessage(rtb1.hwnd, EM_CHARFROMPOS, 0, lP)

    For iI& = lK& To 1 Step -1
    If Mid$(rtb1.Text, iI&, 1) = " " Or Mid$(rtb1.Text, iI&, 1) = Chr(13) Or Mid$(rtb1.Text, iI&, 1) = Chr(10) Then Exit For
    Next iI&

    For iA& = lK& To Len(rtb1.Text)
    If iA& = 0 Then iA& = 1
    If Mid$(rtb1.Text, iA&, 1) = " " Or Mid$(rtb1.Text, iA&, 1) = Chr(13) Or Mid$(rtb1.Text, iA&, 1) = Chr(10) Then Exit For
    Next iA&

If iA& = 0 Or iA& = iI& Then rtb1.MousePointer = 0: Exit Sub

   sTemp$ = Mid$(rtb1.Text, iI& + 1, iA& - iI& - 1)

If InStr(LCase$(sTemp$), "http://") Or InStr(LCase$(sTemp$), "www.") Or InStr(LCase$(sTemp$), "ftp://") Or InStr(LCase$(sTemp$), "aio://") Then
    rtb1.MousePointer = 99
    'rtb1.MouseIcon = pic.Picture
    m_LastUrl$ = sTemp$
    rtb1.ToolTipText = "Goto this url"
    m_IsMouseRTB(0) = True
Else
    If m_IsMouseRTB(1) <> True Then rtb1.MousePointer = 0
    If m_IsMouseRTB(1) <> True Then rtb1.ToolTipText = "Main chat screen": m_LastUrl$ = ""
    m_IsMouseRTB(0) = False
End If
End Sub

Private Function FadeYahooStyle(ByVal sFade As String, ByVal sColors As String) As String
' fading yahoo style
' sColors must not be parsed
If sFade$ = "" Or sColors$ = "" Then Exit Function
Dim arrColors() As String, NumFades As Long
Dim iI As Integer, sQ As String, NumLetters As Long, sA$

Dim ColorR&, ColorG&, ColorB&
Dim RDelta&, GDelta&, BDelta&
Dim LettersPerFade&, FadeLetter&, CurrFade&, CurrLetter$
Dim Letter As Long

arrColors$() = Split(Trim$(sColors$), ",") ' get colors

    For iI% = LBound(arrColors$()) To UBound(arrColors$())
        arrColors$(iI%) = CStr(HTML2RGB(Trim$(arrColors$(iI%)))) ' make sure the colors ae just the colors
    Next iI% ' and also get the max color len : assigned to iI%

  NumFades = iI% - 1

  NumLetters = Len(sFade$)

  ColorR = GetRGB(CLng(arrColors$(0))).Red
  ColorG = GetRGB(CLng(arrColors$(0))).Green
  ColorB = GetRGB(CLng(arrColors$(0))).Blue

      LettersPerFade = Len(StrIp$(sFade$)) / NumFades
If LettersPerFade = 0 Then Exit Function
  RDelta = (GetRGB(CLng(arrColors$(1))).Red - ColorR) / LettersPerFade
  GDelta = (GetRGB(CLng(arrColors$(1))).Green - ColorG) / LettersPerFade
  BDelta = (GetRGB(CLng(arrColors$(1))).Blue - ColorB) / LettersPerFade

'MsgBox sFade$
FadeLetter = 0: CurrFade = 0

  sA$ = ""
  For Letter = 1 To NumLetters
    CurrLetter = Mid$(sFade$, Letter, 1)

   If CurrLetter = Chr(27) Then
    sA$ = sA$ & Mid$(sFade$, Letter, InStr(Letter + 1, sFade$, "m") - Letter + 1)
    Letter = InStr(Letter + 1, sFade$, "m")
   Else
    If ColorR > 255 Then ColorR = 255
    If ColorR < 0 Then ColorR = 0
    If ColorG > 255 Then ColorG = 255
    If ColorG < 0 Then ColorG = 0
    If ColorB > 255 Then ColorB = 255
    If ColorB < 0 Then ColorB = 0

    sA$ = sA$ & Chr(27) & "[#" & IIf(Len(Hex((Int(ColorR)))) = 1, "0", "") & Hex((Int(ColorR))) & IIf(Len(Hex((Int(ColorG)))) = 1, "0", "") & Hex((Int(ColorG))) & IIf(Len(Hex((Int(ColorB)))) = 1, "0", "") & Hex((Int(ColorB))) & "m" & CurrLetter
    ColorR = ColorR + RDelta
    ColorG = ColorG + GDelta
    ColorB = ColorB + BDelta

    If FadeLetter >= LettersPerFade Then
      FadeLetter = FadeLetter - LettersPerFade
      CurrFade = CurrFade + 1

      RDelta = (GetRGB(CLng(arrColors$(CurrFade + 1))).Red - GetRGB(CLng(arrColors$(CurrFade))).Red) / LettersPerFade
      GDelta = (GetRGB(CLng(arrColors$(CurrFade + 1))).Green - GetRGB(CLng(arrColors$(CurrFade))).Green) / LettersPerFade
      BDelta = (GetRGB(CLng(arrColors$(CurrFade + 1))).Blue - GetRGB(CLng(arrColors$(CurrFade))).Blue) / LettersPerFade

      ColorR = GetRGB(CLng(arrColors$(CurrFade))).Red
      ColorG = GetRGB(CLng(arrColors$(CurrFade))).Green
      ColorB = GetRGB(CLng(arrColors$(CurrFade))).Blue
    Else
      FadeLetter = FadeLetter + 1
    End If
   End If ''''''
  Next Letter

FadeYahooStyle$ = sA$
End Function

Public Function GetRGB(ByVal CVal As Long) As UDT_COLORRGB
 GetRGB.Blue = Int(CVal / 65536)
 GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
 GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function

Private Function HTML2RGB(Optional ByVal sTxt As String) As Long
On Error GoTo 1
 
 sTxt$ = IIf(Left$(sTxt$, 1) = "#", Mid$(sTxt$, 2), sTxt$) & String(6, "0")

HTML2RGB& = RGB(Val("&H" & Mid$(sTxt$, 1, 2)) _
              , Val("&H" & Mid$(sTxt$, 3, 2)) _
              , Val("&H" & Mid$(sTxt$, 5, 2)))
1
End Function

Public Sub PrintText(ByVal sTxt As String, ByRef rtb1 As RichTextBox, ByRef rtb2 As RichTextBox)
Dim lA As Long, lB As Long, lC As Long
Dim lColor As Long, Last As UDT_PRINTTEXT
Dim sMid As String, sLeft As String

rtb2.Text = ""
sTxt$ = sTxt$ & "</font></u></i></b></s>"
'm_IgnoreStyles = False
 'sTxt$ = Chr(27) & "[0m" & IIf(m_IgnoreStyles = False, sTxt$, StrIp$(StripHTML$(sTxt$)))
 
 'If m_IgnoreURL = False Then Call TrapUrl(sTxt$): _
                           Call TrapWWW(sTxt$)
 'Call TrapAlt(sTxt$, m_IgnoreAlt)
 'Call TrapFade(sTxt$, m_IgnoreFade)
 'Call TrapFont(sTxt$, m_IgnoreFont)

 'Call ReplaceSub(sTxt$)
 
 Call TrapTags(sTxt$, m_IgnoreTags)
'MsgBox sTxt$
Dim s$

s$ = Replace$(sTxt$, "!proc", Chr(27) & "[bm!proc" & Chr(27) & "[/bm")
s$ = Replace$(s$, "end!", Chr(27) & "[bmend!" & Chr(27) & "[/bm")
s$ = Replace$(s$, "var", Chr(27) & "[bmvar" & Chr(27) & "[/bm")
s$ = Replace$(s$, "!type", Chr(27) & "[bm!type" & Chr(27) & "[/bm")

s$ = Replace$(s$, "$", Chr(27) & "[" & Rgb2Html$(RGB(200, 55, 55)) & "m$" & Chr(27) & "[#000000m")
s$ = Replace$(s$, "@", Chr(27) & "[" & Rgb2Html$(RGB(200, 55, 55)) & "m@" & Chr(27) & "[#000000m")
s$ = Replace$(s$, "&", Chr(27) & "[" & Rgb2Html$(RGB(200, 55, 55)) & "m&" & Chr(27) & "[#000000m")
s$ = Replace$(s$, "%", Chr(27) & "[" & Rgb2Html$(RGB(200, 55, 55)) & "m%" & Chr(27) & "[#000000m")

s$ = Replace$(s$, "(", Chr(27) & "[" & Rgb2Html$(RGB(0, 55, 255)) & "m(" & Chr(27) & "[#000000m")
s$ = Replace$(s$, ")", Chr(27) & "[" & Rgb2Html$(RGB(0, 55, 255)) & "m)" & Chr(27) & "[#000000m")
s$ = Replace$(s$, "//", Chr(27) & "[" & Rgb2Html$(RGB(0, 196, 0)) & "m//" & Chr(27) & "[#000000m")

sTxt$ = s$

Do
 DoEvents
  lA& = InStr(lB& + 1, sTxt$, Chr(27) & "[")
  lB& = InStr(lA& + 1, sTxt$, "m")

   If lA& = 0 Or lB& = 0 Then Exit Do
  
  sMid$ = Mid$(sTxt$, lA& + 2, lB& - lA& - 2)
  
  lColor& = HTML2RGB&(sMid$)
  
   lC& = InStr(lB& + 1, sTxt$, Chr(27) & "[")

  If lC& <> 0 Then sLeft$ = Mid$(sTxt$, lB& + 1, lC& - lB& - 1) Else sLeft$ = Mid$(sTxt$, lB& + 1)

  Select Case LCase$(sMid$)
   Case "b", "1"
    Last.SelBold = True
    rtb2.SelBold = True
    rtb2.SelText = sLeft$
   Case "/b", "x1"
    Last.SelBold = False
    rtb2.SelBold = False
    rtb2.SelText = sLeft$
   Case "i", "2"
    Last.SelItalic = True
    rtb2.SelItalic = True
    rtb2.SelText = sLeft$
   Case "/i", "x2"
    Last.SelItalic = False
    rtb2.SelItalic = False
    rtb2.SelText = sLeft$
   Case "s", "3"
    Last.SelStrikeThru = True
    rtb2.SelStrikeThru = True
    rtb2.SelText = sLeft$
   Case "/s", "x3"
    Last.SelStrikeThru = False
    rtb2.SelStrikeThru = False
    rtb2.SelText = sLeft$
   Case "u", "4"
    Last.SelUnderLine = True
    rtb2.SelUnderLine = True
    rtb2.SelText = sLeft$
   Case "/u", "x4"
    Last.SelUnderLine = False
    rtb2.SelUnderLine = False
    rtb2.SelText = sLeft$
   Case "fade", "alt"
    Last.SelFade = rtb2.SelColor
   Case "/fade", "/alt"
    rtb2.SelColor = Last.SelFade
    rtb2.SelText = sLeft$
   Case "/font"
    rtb2.SelFontName = Last.SelFont
    rtb2.SelFontSize = Last.SelFontSize
    rtb2.SelText = sLeft$
   Case "30"
    Last.SelColor = HTML2RGB("000000")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "31"
    Last.SelColor = HTML2RGB("0000FF")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "32"
    Last.SelColor = HTML2RGB("00FFFF")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "33"
    Last.SelColor = HTML2RGB("808080")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "34"
    Last.SelColor = HTML2RGB("00CC00")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "35"
    Last.SelColor = HTML2RGB("FFC0FF")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "36"
    Last.SelColor = HTML2RGB("C000C0")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "37"
    Last.SelColor = HTML2RGB("C0C000")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "38"
    Last.SelColor = HTML2RGB("FF0000")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "39"
    Last.SelColor = HTML2RGB("808000")
    rtb2.SelColor = Last.SelColor
    rtb2.SelText = sLeft$
   Case "url"
    rtb2.SelColor = vbBlue
    rtb2.SelUnderLine = True
    rtb2.SelText = sLeft$
   Case "/url"
    rtb2.SelColor = Last.SelColor
    rtb2.SelUnderLine = Last.SelUnderLine
    rtb2.SelText = sLeft$
   Case Else
    If Left$(LCase$(sMid$), 4) = "font" Then
     Last.SelFont = rtb2.SelFontName
     rtb2.SelFontName = Replace$(Mid$(sMid$, 5), "^^", "m")
     rtb2.SelText = sLeft$
    ElseIf Left$(LCase$(sMid$), 4) = "size" Then
     Last.SelFontSize = rtb2.SelFontSize
     rtb2.SelFontSize = Mid$(sMid$, 5)
     rtb2.SelText = sLeft$
    Else
     Last.SelColor = lColor&
     rtb2.SelColor = lColor&
     rtb2.SelText = sLeft$
    End If
  End Select
Loop
m_ptLast = Last
rtb1.SelStart = Len(rtb1.Text)
rtb1.SelRTF = rtb2.TextRTF

End Sub

Private Sub ReplaceSub(ByRef s As String)
Dim l&, arrA, arrB

arrA = Array("0;0", "0;31", "0;34", "0;35")
arrB = Array("30", "38", "31", "FF68B8")

 For l& = 0 To UBound(arrA)
  s$ = Replace$(s$, "[" & arrA(l&) & "m", "[" & arrB(l&) & "m", , , vbTextCompare)
 Next l&

'fonts
s$ = Replace$(s$, "À", "", , , vbTextCompare)
s$ = Replace$(s$, "€", "", , , vbTextCompare)
End Sub

Private Function StrIp(sHTML As String) As String

Dim sTemp As String, lSpot1 As Long, lSpot2 As Long, lSpot3 As Long

sTemp$ = sHTML$
Do
  lSpot1& = InStr(lSpot3& + 1, sTemp$, Chr(27) & "[")
  lSpot2& = InStr(lSpot1& + 1, sTemp$, "m")
  
    If lSpot1& = lSpot3& Or lSpot1& < 1 Then Exit Do
    If lSpot2& < lSpot1& Then lSpot2& = lSpot1& + 1
    
  sTemp$ = Left$(sTemp$, lSpot1& - 1) + Right$(sTemp$, Len(sTemp$) - lSpot2&)
  lSpot3& = lSpot1& - 1
Loop

StrIp$ = sTemp$

End Function

Private Function StripHTML(sHTML As String) As String

Dim sTemp As String, lSpot1 As Long, lSpot2 As Long, lSpot3 As Long

sTemp$ = sHTML$
Do
  lSpot1& = InStr(lSpot3& + 1, sTemp$, "<")
  lSpot2& = InStr(lSpot1& + 1, sTemp$, ">")
  
    If lSpot1& = lSpot3& Or lSpot1& < 1 Then Exit Do
    If lSpot2& < lSpot1& Then lSpot2& = lSpot1& + 1
    
  sTemp$ = Left$(sTemp$, lSpot1& - 1) + Right$(sTemp$, Len(sTemp$) - lSpot2&)
  lSpot3& = lSpot1& - 1
Loop

StripHTML$ = sTemp$

End Function

Private Sub TrapAlt(ByRef s As String, ByVal b As Boolean)
On Error Resume Next
Dim l&, k&, c&, t&

Do
 l& = InStr(c& + 1, LCase$(s$), "<alt ")
 k& = InStr(l& + 1, LCase$(s$), ">")
 If k& = 0 Then k& = InStr(l& + 1, LCase$(s$), Chr(13))
 t& = InStr(k& + 1, LCase$(s$), "</alt>")
 If t& = 0 Then t& = Len(s$)

  If l& = 0 Then Exit Do
  If k& <= l& Then k& = Len(s$) + 4

  s$ = Left$(s$, l& - 1) & IIf(b = False, Chr(27) & "[altm" & AltYahooStyle$(Mid$(s$, k& + 1, t& - k& - 1), Mid$(s$, l& + 4, k& - l& - 4)) & Chr(27) & "[/altm", Mid$(s$, k& + 1, t& - k& - 1)) & Mid$(s$, t& + 6)

 c& = InStr(l& + 1, LCase$(s$), " ")

 If c& = 0 Then Exit Do
DoEvents
Loop
End Sub

Private Sub TrapFade(ByRef s As String, ByVal b As Boolean)
On Error Resume Next
Dim l&, k&, c&, t&

Do
 l& = InStr(c& + 1, LCase$(s$), "<fade ")
 k& = InStr(l& + 1, LCase$(s$), ">")
 If k& = 0 Then k& = InStr(l& + 1, LCase$(s$), Chr(13))
 t& = InStr(k& + 1, LCase$(s$), "</fade>")
 If t& = 0 Then t& = Len(s$)

  If l& = 0 Then Exit Do
  If k& <= l& Then k& = Len(s$) + 4

  s$ = Left$(s$, l& - 1) & IIf(b = False, Chr(27) & "[fadem" & FadeYahooStyle$(Mid$(s$, k& + 1, t& - k& - 1), Mid$(s$, l& + 5, k& - l& - 5)) & Chr(27) & "[/fadem", Mid$(s$, k& + 1, t& - k& - 1)) & Mid$(s$, t& + 7)

 c& = InStr(l& + 1, LCase$(s$), " ")

 If c& = 0 Then Exit Do
DoEvents
Loop
End Sub

Private Sub TrapFont(ByRef s As String, ByVal b As Boolean)
On Error Resume Next
Dim l&, k&, c&, t&, n&, v As Variant
Dim arr$(), arrX$(), iI%, sPre$
iI% = 0
Do
 l& = InStr(c& + 1, LCase$(s$), "<font ")
 k& = InStr(l& + 1, LCase$(s$), ">")
 If k& = 0 Then k& = InStr(l& + 1, LCase$(s$), Chr(13))
 t& = InStr(k& + 1, LCase$(s$), "</font>")
 If t& = 0 Then t& = Len(s$)

  If l& = 0 Then Exit Do
  If k& <= l& Then k& = Len(s$) + 4
   arr$() = Split(Trim$(Mid$(s$, l& + 5, k& - l& - 5)), "=")

   For Each v In arr$()
    n& = InStrRev(Trim$(v), " ")
    If n& = 0 Then
     ReDim Preserve arrX$(iI%)
      arrX$(iI%) = Trim$(Replace$(v, """", ""))
      iI% = iI% + 1
    Else
     ReDim Preserve arrX$(iI%)
      arrX$(iI%) = Trim$(Replace$(Left$(v, n& - 1), """", ""))
      iI% = iI% + 1
     ReDim Preserve arrX$(iI%)
      arrX$(iI%) = Trim$(Replace$(Mid$(v, n& + 1), """", ""))
      iI% = iI% + 1
    End If
    
   Next v

   For iI% = 0 To UBound(arrX$())
    If arrX$(iI%) <> "" Then
     If LCase$(Trim$(arrX$(iI%))) = "face" Then
      sPre$ = sPre$ & Chr(27) & "[font" & Replace$(arrX$(iI% + 1), "m", "^^") & "m"
      iI% = iI + 1
     ElseIf Trim$(LCase$(arrX$(iI%))) = "size" And IsNumeric(arrX$(iI% + 1)) Then
      sPre$ = sPre$ & Chr(27) & "[size" & arrX$(iI% + 1) & "m"
      iI% = iI + 1
     End If
    End If
   Next iI%

  ReDim arrX(0)
  
   s$ = Left$(s$, l& - 1) & IIf(b = False, sPre$ & Mid$(s$, k& + 1, t& - k& - 1) & Chr(27) & "[/fontm", Mid$(s$, k& + 1, t& - k& - 1)) & Mid$(s$, t& + 7)
  sPre$ = ""
 
 c& = InStr(l& + 1, LCase$(s$), " ")

 If c& = 0 Then Exit Do
DoEvents
Loop
End Sub

Private Sub TrapTags(ByRef s As String, ByVal b As Boolean)
On Error Resume Next
Dim l&, k&, c&

 Dim arrA, arrB
  arrA = Array("red", "blue", "green", "yellow", "black", "white", "orange", "grey", "gray", "aqua", "pink", "b", "/b", "i", "/i", "u", "/u", "s", "/s", "/fade", "/font")
  arrB = Array("38", "31", "34", "37", "30", "#FFFFFF", "#00FFFF", "33", "33", "32", "36", "b", "/b", "i", "/i", "u", "/u", "s", "/s", "/fade", "/font")

  For l& = 0 To UBound(arrA)
   s$ = IIf(b = False, Replace$(s$, "<" & arrA(l&) & ">", "[" & arrB(l&) & "m", , , vbTextCompare), Replace$(s$, "<" & arrA(l&) & ">", "", , , vbTextCompare))
  Next l&
Do
 l& = InStr(c& + 1, LCase$(s$), "<")
 k& = InStr(l& + 1, LCase$(s$), ">")
 
  If k& = 0 Then Exit Do
  If l& = 0 Then Exit Do

  s$ = Left$(s$, l& - 1) & IIf(Left$(Mid$(s$, l& + 1, k& - l& - 1), 1) = "#", IIf(b = False, Chr(27) & "[" & Trim(Mid$(s$, l& + 1, k& - l& - 1)) & "m", ""), "<" & Mid$(s$, l& + 1, k& - l& - 1) & ">") & Mid$(s$, k& + 1)
 c& = InStr(l& + 1, LCase$(s$), " ")

 If c& = 0 Then Exit Do
DoEvents
Loop

End Sub

Private Sub TrapUrl(ByRef s As String)
On Error Resume Next
Dim l&, k&, c&

Do
 l& = InStr(c& + 1, LCase$(s$), "http://")
 k& = InStr(l& + 1, LCase$(s$), " ")
 If k& = 0 Then k& = InStr(l& + 1, LCase$(s$), Chr(13))

  If l& = 0 Then Exit Do
  If k& <= l& Then k& = Len(s$) + 4

  s$ = Left$(s$, l& - 1) & Chr(27) & "[urlm" & Mid$(s$, l&, k& - l&) & Chr(27) & "[/urlm" & Mid$(s$, k&)
 c& = InStr(l& + 1, LCase$(s$), " ")

 If c& = 0 Then Exit Do
DoEvents
Loop
End Sub

Private Sub TrapWWW(ByRef s As String)
On Error Resume Next
Dim l&, k&, c&

Do
 l& = InStr(c& + 1, LCase$(s$), "www.")
 k& = InStr(l& + 1, LCase$(s$), " ")
 If k& = 0 Then k& = InStr(l& + 1, LCase$(s$), Chr(13))

  If l& = 0 Then Exit Do
  If k& <= l& Then k& = Len(s$) + 4

  s$ = Left$(s$, l& - 1) & Chr(27) & "[urlm" & Mid$(s$, l&, k& - l&) & Chr(27) & "[/urlm" & Mid$(s$, k&)
 c& = InStr(l& + 1, LCase$(s$), " ")

 If c& = 0 Then Exit Do
DoEvents
Loop
End Sub

'Private Sub rtb1_Click()
'If m_LastUrl$ <> "" And m_IsMouseRTB(0) = True Then _
' Call ShellExecute(Parent.hwnd, vbNullString, m_LastUrl$, vbNullString, "c:\", 1)

'End Sub

Private Sub rtb1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call CheckRTBPosURL(X, Y)
End Sub
