Attribute VB_Name = "modWinChange"
'###################################################
'#       coded: J.Huber                            #
'#       name: modWinChange.bas                    #
'#       changes a window properties via API       #
'###################################################
' Uses: SetWindowLong() API to change the windows
'       properties.
'
' I included more Styles then actually used so you
' can modify to your needs.

Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)

Public Const MF_BYPOSITION = &H400&

Public Const SC_CLOSE = &HF060&
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_MINIMIZE = &HF020&
Public Const SC_MOVE = &HF010&
Public Const SC_RESTORE = &HF120&
Public Const SC_SIZE = &HF000&

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1

Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_CLIENTEDGE = 512
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOOLWINDOW = 128
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_maximize = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

Public Sub ChangeWin(ByVal lhWnd As Long, Optional ByVal Border As Boolean = True, Optional ByVal TitleBar As Boolean = True, Optional ByVal Maximize As Boolean = True, Optional ByVal Minimize As Boolean = True, Optional ByVal SystemMenu As Boolean = True, Optional ByVal ThickFrame As Boolean = True, Optional ByVal VScroll As Boolean = False, Optional ByVal HScroll As Boolean = False, Optional ByVal ExTransparent As Boolean = False, Optional ByVal ExClientEdge As Boolean = False, Optional ByVal ExToolWindow As Boolean = False)
'Dim lStyle As Long, lExStyle As Long
' lExStyle& = IIf(ExTransparent, WS_EX_TRANSPARENT, 0) Or IIf(ExTopMost, WS_EX_TOPMOST, 0)
Dim lStyle As Long

Const swpFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE

 lStyle& = GetWindowLong(lhWnd&, GWL_STYLE)
   
 If ThickFrame = True Then lStyle& = lStyle& Or WS_THICKFRAME Else lStyle& = lStyle& And Not WS_THICKFRAME
 If Border = True Then lStyle& = lStyle& Or WS_BORDER Else lStyle& = lStyle& And Not WS_BORDER
 If TitleBar = True Then lStyle& = lStyle& Or WS_CAPTION Else lStyle& = lStyle& And Not WS_CAPTION
 If Maximize = True Then lStyle& = lStyle& Or WS_MAXIMIZEBOX Else lStyle& = lStyle& And Not WS_MAXIMIZEBOX
 If Minimize = True Then lStyle& = lStyle& Or WS_MINIMIZEBOX Else lStyle& = lStyle& And Not WS_MINIMIZEBOX
 If SystemMenu = True Then lStyle& = lStyle& Or WS_SYSMENU Else lStyle& = lStyle& And Not WS_SYSMENU
 If VScroll = True Then lStyle& = lStyle& Or WS_VSCROLL Else lStyle& = lStyle& And Not WS_VSCROLL
 If HScroll = True Then lStyle& = lStyle& Or WS_HSCROLL Else lStyle& = lStyle& And Not WS_HSCROLL

   Call SetWindowLong(lhWnd&, GWL_STYLE, lStyle&)
   Call SetWindowPos(lhWnd&, 0, 0, 0, 0, 0, swpFlags)
 lStyle& = 0
 lStyle& = GetWindowLong(lhWnd&, GWL_EXSTYLE)
   
 If ExTransparent = True Then lStyle& = lStyle& Or WS_EX_TRANSPARENT Else lStyle& = lStyle& And Not WS_EX_TRANSPARENT
 If ExClientEdge = True Then lStyle& = lStyle& Or WS_EX_CLIENTEDGE Else lStyle& = lStyle& And Not WS_EX_CLIENTEDGE
 If ExToolWindow = True Then lStyle& = lStyle& Or WS_EX_TOOLWINDOW Else lStyle& = lStyle& And Not WS_ToolWindow

   Call SetWindowLong(lhWnd&, GWL_EXSTYLE, lStyle&)
   Call SetWindowPos(lhWnd&, 0, 0, 0, 0, 0, swpFlags)

End Sub

Private Function FlipProp(ByVal lhWnd As Long, ByVal lBit As Long, ByVal bValue As Boolean, Optional dStyle As Long = GWL_STYLE) As Boolean
Dim lStyle As Long

Const swpFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE

 lStyle& = GetWindowLong(lhWnd&, dStyle&)
   
 If bValue = True Then lStyle& = lStyle& Or lBit& Else lStyle& = lStyle& And Not lBit&

   Call SetWindowLong(lhWnd&, dStyle&, lStyle&)
   Call SetWindowPos(lhWnd&, 0, 0, 0, 0, 0, swpFlags)
   
   FlipProp = lStyle& = GetWindowLong(lhWnd&, dStyle&)
End Function

Public Sub ClearSysMenu(ByVal lhWnd As Long)
Dim hMenu As Long, lId As Long, lL As Long

 hMenu = GetSystemMenu(lhWnd&, False)

  Do
   DoEvents
    lId& = GetMenuItemID(hMenu, lL&)
   If lId& <> SC_CLOSE Then Call RemoveMenu(hMenu, lL&, MF_BYPOSITION): lL& = lL& - 1
    lL& = lL& + 1
   If lL& >= GetMenuItemCount(hMenu&) - 1 Then Exit Do
  Loop
'  Call RemoveMenu(hMenu, 0, MF_BYPOSITION)
End Sub
