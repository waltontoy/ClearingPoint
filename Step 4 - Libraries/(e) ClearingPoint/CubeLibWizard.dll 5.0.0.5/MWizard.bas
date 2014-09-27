Attribute VB_Name = "MWizard"
Option Explicit
DefLng A-Z

'// a public instance of the controlling class
Public g_Controller As IWizardController

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
    x As Long
    y As Long
End Type

'// Window attribute functions
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'// Z-order and placement APIs
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'// RECT Functions
Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long

'// Focus and activation functions
Declare Function winSetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function winGetFocus Lib "user32" Alias "GetFocus" () As Long

'// used to create the sunken edge around the splash screen
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

'// Cursor position functions
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Const GWL_STYLE = (-16)
Public Const WS_CHILD = &H40000000

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'// DrawEdge() constants
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public G_conDatabase As ADODB.Connection

Public g_blnConfigCancelled As Boolean

'
'   Simple helper function that determines if the passed
'   variant is either a valid numeric value or a non-empty
'   string. Many times when designing components you will
'   need to make these types of checks.
'
Public Function IsValidVariant(vData As Variant) As Boolean

If Not IsMissing(vData) Then
    If Not IsNull(vData) Then
        If Not IsEmpty(vData) Then
            If Not IsArray(vData) Then
                If IsNumeric(vData) Or Len(vData) > 0 Then
                    IsValidVariant = True
                    
                End If
            
            End If
                        
        End If
        
    End If
    
End If




End Function


'
'   Simple function similar to the max() macro defined
'   in the Win32 SDK. Returns the largest of two values.
'
Public Function Max(ByVal p1 As Long, ByVal p2 As Long) As Long

If p1 > p2 Then
    Max = p1
Else
    Max = p2
End If

End Function

'
'   This function converts an IWizardPage handle to
'   a standard VB from reference, by simply casting it.
'
Public Function CForm(hPage As IWizardPage) As VB.Form
    
    Set CForm = hPage

End Function

'
'   Just centers a form on the screen
'
Public Sub Center(hForm As Form)
    
    On Error Resume Next
    
    With hForm
        
        .Move (Screen.Width \ 2) - (.Width \ 2), (Screen.Height \ 2) - (.Height \ 2)
        
    End With
    
End Sub
'
'   Dummy entry point used to create a global
'   instance of the controller class.
'
Sub Main()

Set g_Controller = New CWizardController

End Sub


