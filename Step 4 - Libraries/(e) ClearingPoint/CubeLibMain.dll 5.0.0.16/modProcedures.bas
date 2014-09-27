Attribute VB_Name = "modProcedures"
Option Explicit

Public g_objDataSourceProperties As CDataSourceProperties

Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type

Public Declare Function PathFileExists _
    Lib "shlwapi.dll" Alias "PathFileExistsA" ( _
    ByVal pszPath As String _
    ) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Public Declare Function IsUserAnAdmin Lib "shell32" () As Long

Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ScrollWindowByNum& Lib "user32" Alias "ScrollWindow" (ByVal hwnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, ByVal lpRect As Long, ByVal lpClipRect As Long)
Private Declare Function GetWindowRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)

Private Const GWL_STYLE = (-16)
Private Const GWL_WNDPROC = (-4)
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SB_BOTH = 3
Private Const SB_LINEDOWN = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGEUP = 2
Private Const SB_THUMBTRACK = 5
Private Const SB_ENDSCROLL = 8

Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_DESTROY = &H2
Private Const SIF_ALL = &H17
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_RANGE = &H1&
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3

Dim s As SCROLLINFO
Dim OriginHeight As Long, OriginWidth As Long
Dim OldProc As Long

Public Sub SetScrollBar(hObj As Long, sbPos As ScrollBarConstants, Optional bShowAlways As Boolean = False)
  Dim lStyle As Long, rc As RECT, OldProc As Long
  lStyle = sbPos * &H100000
  SetWindowLong hObj, GWL_STYLE, GetWindowLong(hObj, GWL_STYLE) Or lStyle
  SetWindowPos hObj, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
  Call GetWindowRect(hObj, rc)
  OriginHeight = rc.Bottom - rc.Top + GetSystemMetrics(SM_CYHSCROLL) * (sbPos And vbHorizontal)
  OriginWidth = rc.Right - rc.Left + GetSystemMetrics(SM_CXVSCROLL) * (sbPos And vbVertical) / 2
  s.cbSize = Len(s)
  s.fMask = SIF_ALL
  If bShowAlways Then s.fMask = s.fMask Or SIF_DISABLENOSCROLL
  s.nMin = 0
  s.nPos = 0
  OldProc = SetWindowLong(hObj, GWL_WNDPROC, AddressOf WndProc)
  SetProp hObj, "OLDPROC", OldProc
  SetProp hObj, "SB_POS", sbPos
  SetProp hObj, "ORIGIN_WIDTH", OriginWidth
  SetProp hObj, "ORIGIN_HEIGHT", OriginHeight
End Sub

Public Sub AdjustScrollInfo(hObj As Long)

    Dim sb As Long
    Dim rc As RECT
    Dim OldProc As Long
    
    sb = GetProp(hObj, "SB_POS")
    Call GetWindowRect(hObj, rc)
    
    If (sb And vbVertical) = vbVertical Then
        s.nMax = (frmPane.tvwMain.Height + frmPane.sccViews.Height + _
               frmPane.fraViews.Height) / Screen.TwipsPerPixelY
        s.nMin = 0
        s.nPos = 0
        s.nPage = rc.Bottom - rc.Top - 1
        SetScrollInfo hObj, SB_VERT, s, True
        
        OldProc = SetWindowLong(hObj, GWL_WNDPROC, AddressOf WndProc)

    End If

End Sub

Public Function WndProc(ByVal hOwner As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim nOldPos As Long, n As Long
    Select Case wMsg
        Case WM_VSCROLL, WM_HSCROLL
            GetScrollInfo hOwner, wMsg - WM_HSCROLL, s
            nOldPos = s.nPos

            Select Case GetLoWord(wParam)
                Case SB_LINEDOWN
                     s.nPos = s.nPos + s.nPage \ 10
                Case SB_LINEUP
                     s.nPos = s.nPos - s.nPage \ 10
                Case SB_PAGEDOWN
                     s.nPos = s.nPos + s.nPage
                Case SB_PAGEUP
                     s.nPos = s.nPos - s.nPage
                Case SB_THUMBTRACK
                     s.nPos = GetHiWord(wParam)
                Case SB_ENDSCROLL
                     If s.nPos = 0 Then
                        AdjustScrollInfo hOwner
                        Exit Function
                     End If
            End Select
            SetScrollInfo hOwner, wMsg - WM_HSCROLL, s, True
            GetScrollInfo hOwner, wMsg - WM_HSCROLL, s
            If wMsg = WM_VSCROLL Then
               ScrollWindowByNum hOwner, 0, nOldPos - s.nPos, 0, 0
            Else
               ScrollWindowByNum hOwner, nOldPos - s.nPos, 0, 0, 0
            End If
        Case WM_DESTROY
           RemoveProp hOwner, "SB_POS"
           RemoveProp hOwner, "ORIGIN_WIDTH"
           RemoveProp hOwner, "ORIGIN_HEIGHT"
           Call SetWindowLong(hOwner, GWL_WNDPROC, GetProp(hOwner, "OLDPROC"))
        Case Else
    End Select
    WndProc = CallWindowProc(GetProp(hOwner, "OLDPROC"), hOwner, wMsg, wParam, lParam)
End Function

Private Function GetHiWord(dw As Long) As Long
  If dw And &H80000000 Then
     GetHiWord = (dw \ 65535) - 1
  Else
     GetHiWord = dw \ 65535
  End If
End Function

Private Function GetLoWord(dw As Long) As Long
   If dw And &H8000& Then
      GetLoWord = &H8000 Or (dw And &H7FFF&)
   Else
      GetLoWord = dw And &HFFFF&
   End If
End Function
