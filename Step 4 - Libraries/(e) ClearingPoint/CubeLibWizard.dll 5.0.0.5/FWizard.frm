VERSION 5.00
Begin VB.Form FWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ClearingPoint Configuration Wizard"
   ClientHeight    =   5145
   ClientLeft      =   4350
   ClientTop       =   2745
   ClientWidth     =   7800
   ControlBox      =   0   'False
   Icon            =   "FWizard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   4530
      Left            =   2520
      ScaleHeight     =   302
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   344
      TabIndex        =   4
      Tag             =   "-1"
      Top             =   0
      Width           =   5160
   End
   Begin VB.PictureBox picSplash 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   3390
      Left            =   0
      Picture         =   "FWizard.frx":000C
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "-1"
      Top             =   0
      Width           =   1800
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   345
      Left            =   180
      TabIndex        =   2
      Tag             =   "0"
      Top             =   4740
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdStepBack 
      Caption         =   "<  Back"
      Height          =   345
      Left            =   2850
      TabIndex        =   3
      Tag             =   "-1"
      Top             =   4740
      Width           =   1200
   End
   Begin VB.CommandButton cmdStepFwd 
      Caption         =   "Next  >"
      Height          =   345
      Left            =   4080
      TabIndex        =   0
      Tag             =   "-1"
      Top             =   4740
      Width           =   1200
   End
   Begin VB.CommandButton cmdTerminate 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   6420
      TabIndex        =   1
      Tag             =   "-1"
      Top             =   4740
      Width           =   1200
   End
   Begin VB.Line lnBevel 
      BorderColor     =   &H80000010&
      Index           =   0
      Tag             =   "-1"
      X1              =   0
      X2              =   520
      Y1              =   312
      Y2              =   312
   End
   Begin VB.Line lnBevel 
      BorderColor     =   &H80000014&
      Index           =   1
      Tag             =   "-1"
      X1              =   12
      X2              =   410
      Y1              =   341
      Y2              =   341
   End
End
Attribute VB_Name = "FWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

DefLng A-Z

Private m_bFinishedLoading As Boolean   '// form load completion flag
Private m_bFrozen As Boolean            '// whether or not the wizard is frozen
Private m_hWndLastClicked As Long       '// handle of last clicked button
Private m_bAutoCenterOnResize As Boolean '// Do we center automatically on the screen if the wizard's size changes because of a page transition?
Private m_bStepping As Boolean          '// Internal flag used to know when we're in the middle of a page transition

'// A type used to hold information about
'// a given page. The type array declared
'// below (m_Steps) holds all pages added
'// to the wizard
Private Type tagStep
    Key As Variant
    Handle As IWizardPage
    hWndParent As Long
    bInited As Boolean
End Type

Private m_Steps() As tagStep

Private m_hClient As IWizardController      '// the class that controls us
Private m_intCurrentStep As Integer         '// the current step
Private m_intLastStep As Integer            '// the previous step
Private m_intNextStep As Integer            '// the next step (only valid between transitions)
Private m_intStepCount As Integer           '// the count of all pages

'
'   Read only. Returns the wizard's last active page. If the
'   current page is the first one, (not the first one in the
'   sequence, but rather the first one visited) the value
'   returned is -1
'
Public Property Get LastStep() As Variant
    
    LastStep = m_intLastStep
    
End Property


'
'   This always returns -1 unless the wizard is
'   in the middle of a transition.
'
Public Property Get NextStep() As Variant
    
    NextStep = m_intNextStep
    
End Property


'   This property can only be set from the controller's
'   BeforeStep method as called from the wizard's Step
'   method. Note that, while most other methods in this
'   class check for m_bStepping being true and exit if so,
'   NextStep actually requires the flag to be True.
'
Public Property Let NextStep(ByVal vData As Variant)
    
    Dim intIndex As Integer
    
    On Error Resume Next
    
    If Not m_bStepping Then Exit Property
    
    intIndex = ResolvePageIndexByParam(vData)
    
    If intIndex = -1 Then
        Debug.Assert 0
        Exit Property
    
    End If
    
    '// save this
    m_intNextStep = intIndex
    
End Property


'
'   Because the Property Let is typed as a variant, this
'   one must be typed as a variant too, but it always returns
'   an Integer-subtype variant.
'
Public Property Get CurrentStep() As Variant
    
    CurrentStep = m_intCurrentStep
    
End Property
'
'   Sets/returns whether or not the wizard's
'   UI is frozen and inaccessible
'
Public Property Get Frozen() As Boolean
    
    Frozen = m_bFrozen
    
End Property

Public Sub AddPage(hPage As IWizardPage, Optional ByVal Key As Variant)
'   Adds a new page to the wizard, represented by a VB form that
'   implements the IWizardPage interface.
'   This method should be called at least twice (because a wizard has
'   to have at least two pages) before the "Start" method is called.
    
    Dim lngNewIndex As Long
    
    On Error Resume Next
    
    '// check that the variant is valid, and re-check that it is
    '// non-numeric, also. If we allowed numeric keys, we could
    '// not differentiate between an index into the pages array and
    '// the value of a key held on one of them.
    If IsValidVariant(Key) Then
        If Not IsNumeric(Key) Then
            If (VarType(Key) = vbString And Len(Key) > 0) Then
                '// we should not find it here, unless a duplicate key
                '// was passed
                If (FindPageByKey(Key) <> -1) Then
                    Debug.Assert 0
                    Exit Sub
                End If
            End If
        End If
    End If
    
    lngNewIndex = UBound(m_Steps) + 1
    
    '// See if this is the fisrt page we're adding
    If lngNewIndex = 0 Then
        lngNewIndex = 1
        ReDim m_Steps(1 To lngNewIndex)
    Else
        ReDim Preserve m_Steps(1 To lngNewIndex)
    
    End If
    
    '// Set the new page's information
    With m_Steps(lngNewIndex)
    
        Set .Handle = hPage
        .Key = Key
    
    End With
    
    '// save our step count.
    m_intStepCount = lngNewIndex
    
End Sub


'
'   Sets/returns the wizard's current page
'
Public Property Let CurrentStep(ByVal vData As Variant)

Dim intIndex As Integer

On Error Resume Next

If m_bStepping Then Exit Property

intIndex = ResolvePageIndexByParam(vData)

If intIndex = -1 Then
    Debug.Assert 0
    Exit Property

End If

Call StepTo(intIndex)

End Property

'
'   Searches through the array of steps to locate the
'   one whose Key property matches the passed Key argument.
'   Returns -1 of it finds no match.
'
Private Function FindPageByKey(Key As Variant) As Integer

Dim a%

On Error Resume Next

'// Set initially to an invalid return value
FindPageByKey = -1

For a% = 1 To UBound(m_Steps)

    If m_Steps(a%).Key = Key Then
        FindPageByKey = a%
        Exit For

    End If

Next a%


End Function

'
'   Read only. Returns the IWizardPage interface instance
'   of the Form page identified by the specified Index
'
Public Property Get PageHandle(ByVal Index As Variant) As IWizardPage

Dim intIndex As Integer

intIndex = ResolvePageIndexByParam(Index)

If intIndex = -1 Then
    Debug.Assert 0
    Exit Property

End If

Set PageHandle = m_Steps(intIndex).Handle

End Property

'
'   Given a variant value, this function resolves an
'   index into the m_Steps array.
'
Private Function ResolvePageIndexByParam(vData As Variant) As Integer

Dim intIndex As Integer

On Error Resume Next

'// Initialize this to an invalid return
intIndex = -1

'// sanity check
If Not IsValidVariant(vData) Then
    Debug.Assert 0
    Exit Function

End If

'// Do we have a valid, non numeric key?
If Not IsNumeric(vData) Then

    intIndex = FindPageByKey(vData)

Else        '// or an actual page index?

    intIndex = vData
    '// Can't set the page to an out of bounds index!
    If Not (intIndex >= 1 And intIndex <= m_intStepCount) Then intIndex = -1

End If

ResolvePageIndexByParam = intIndex

End Function

'
'   Retrieves the number of pages in the wizard sequence
'
Public Property Get StepCount() As Integer

    StepCount = m_intStepCount

End Property




'
'   Sets/Retrieves the handle to the IWizardController
'   implementation that acts as controller
'
Public Property Get ClientHandle() As IWizardController

Set ClientHandle = m_hClient

End Property



'
'
'
Public Property Set ClientHandle(vData As IWizardController)

If m_bStepping Then Exit Property

Set m_hClient = vData
If m_hClient Is Nothing Then Call finish

End Property


'
'   Called from the step cycling routines. Can
'   be used to further modify the wizard's behavior
'
Private Sub OnLastStep()

cmdStepFwd.Enabled = False
cmdTerminate.Caption = "Finish"
cmdTerminate.Default = True

End Sub

'
'   Called from the step cycling routines. Can
'   be used to further modify the wizard's behavior
'
Private Sub OnFirstStep()

cmdStepBack.Enabled = False

End Sub

'
'   Steps to any page in the wizard sequence. This
'   method is used to skip pages.
'
Public Sub StepTo(ByVal Step As Variant)

Dim intIndex As Integer
Dim intDirection As Integer

On Error Resume Next

If m_bStepping Then Exit Sub

intIndex = ResolvePageIndexByParam(Step)

If intIndex = -1 Then
    Debug.Assert 0
    Exit Sub

End If

m_intNextStep = intIndex
Call Me.Step(0)

End Sub



'
'   Called to initialize the wizard sequence
'
Public Sub Start(Optional FirstStep As Variant)

Dim intIndex As Integer

If m_bFinishedLoading Then
    Debug.Assert 0
    Exit Sub

End If

If m_intStepCount >= 1 Then

    Load Me

    If IsMissing(FirstStep) Then
        Call Step(wizForward)

    Else
        intIndex = ResolvePageIndexByParam(FirstStep)
        If intIndex = -1 Then
            Debug.Assert 0
            Exit Sub

        End If

        Call StepTo(intIndex)

    End If
    
    Me.Visible = True
    m_bFinishedLoading = True

End If

End Sub


'
'   Called from the Step[] method, this procedure takes
'   care of loading the page and placing it inside the
'   wizard picture frame.
'
Private Function InitStep(intCurrentStep As Integer, intNextStep As Integer) As Boolean

Dim hPage As IWizardPage
Dim hwndPage As Long
Dim hForm As VB.Form
Dim bCancel As Boolean
Dim hWndBin As Long
Dim dwStyle As Long

'// Is there a current step?
If intCurrentStep >= 1 Then

    Set hPage = m_Steps(intCurrentStep).Handle
    '// notify the current page
    Call hPage.BeforePageHide(Me, intNextStep, bCancel)

End If

'// see if the client canceled
If Not bCancel Then

    '// hPage is not set unless there's a current step
    If Not hPage Is Nothing Then CForm(hPage).Visible = False

    '// get the handle of the page...
    Set hPage = m_Steps(intNextStep).Handle
    '// and cast it to a Form object
    Set hForm = CForm(hPage)

    '// see if we actually need to initialize the page
    If Not m_Steps(intNextStep).bInited Then

        Load hForm
        hwndPage = hForm.hwnd
        hWndBin = picPage.hwnd

        '// set the form's window style
        dwStyle = GetWindowLong(hwndPage, GWL_STYLE)
        dwStyle = dwStyle Or WS_CHILD
        Call SetWindowLong(hwndPage, GWL_STYLE, dwStyle)

        '// and tell Windows that we're the new parent.
        m_Steps(intNextStep).hWndParent = SetParent(hwndPage, hWndBin)

        '// move the form flush inside the picturebox frame
        Call SetWindowPos(hwndPage, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE)

        m_Steps(intNextStep).bInited = True

    End If

    '// notify the new page
    Call hPage.BeforePageShow(Me, intCurrentStep)

    '// show the page
    hForm.Visible = True

    InitStep = True

End If


Set hForm = Nothing
Set hPage = Nothing


End Function


'
'   Retrieves the height and width of the current page
'   in the sequence.
'
Private Function FetchPageSize(ByRef cx As Long, ByRef cy As Long) As Boolean

Dim hForm As VB.Form
Dim intCurStep As Integer

On Error Resume Next

intCurStep = m_intCurrentStep

If (intCurStep < 1) Then intCurStep = 1

Set hForm = m_Steps(intCurStep).Handle

If Not hForm Is Nothing Then

    cx = hForm.Width \ Screen.TwipsPerPixelX
    cy = hForm.Height \ Screen.TwipsPerPixelY

    FetchPageSize = True

End If

Set hForm = Nothing

End Function

'
'   Terminates the wizard sequence
'
Public Sub finish()

Dim hPage As VB.Form
Dim a%

If m_bStepping Then Exit Sub

Set hPage = CForm(m_Steps(m_intCurrentStep).Handle)

hPage.Visible = False

'// Make sure we unload all our child pages.
For a% = LBound(m_Steps) To UBound(m_Steps)

    If m_Steps(a%).bInited = True Then
        Set hPage = CForm(m_Steps(a%).Handle)
        Call SetParent(hPage.hwnd, m_Steps(a%).hWndParent)
        Unload hPage
        Set hPage = Nothing
        Set m_Steps(a%).Handle = Nothing

    End If

Next a%

Set hPage = Nothing

Set m_hClient = Nothing


End Sub


'
'   This is the procedure that contains all the logic for
'   the sequence transitions.
'
Public Sub Step(ByVal Direction As WizardStepDirections)

Dim intCurStepTemp As Integer
Dim intNextStepTemp As Integer
Dim bCancel As Boolean

On Error Resume Next

'// Don't recurse
If m_bStepping Then Exit Sub

'// invalid step, end or start reached.
If ((Direction < (-1)) Or (Direction > 1)) Then Exit Sub

If m_intCurrentStep < 1 Then        '// first iteration?
    If Direction <> 0 Then m_intNextStep = 1
    m_intCurrentStep = 0

End If

'// special case: we're doing a jump
If Direction = 0 Then
    intNextStepTemp = m_intNextStep
Else
    intNextStepTemp = m_intCurrentStep + Direction
End If

'// sanity check
If ((intNextStepTemp > m_intStepCount) Or (intNextStepTemp < 1)) Then Exit Sub

m_bStepping = True

'// The NextStep property can be changed only from inside the
'// controller's class BeforeStep method. Why? because it's
'// only at this point (when we're doing a transition) that we
'// actually know what the direction is, so we can calculate
'// the next step.
m_intNextStep = intNextStepTemp

'// notify the controller that we're changing the page
Call m_hClient.BeforeStep(Direction, bCancel)

'// if he didn't cancel, continue
If Not bCancel Then

    intCurStepTemp = m_intCurrentStep

    '// this may have changed at this point
    intNextStepTemp = m_intNextStep

    '// if the page initialization is succcessful, save
    '// the current and last steps and invalidate the
    '// NextStep property, and do any necessary rearranging
    If InitStep(intCurStepTemp, intNextStepTemp) = True Then
        m_intLastStep = m_intCurrentStep
        m_intCurrentStep = intNextStepTemp
        m_intNextStep = -1

        '// make sure we are resized and arranged
        Call NotifyPageChange

        '// see if we're at the last or first step
        If m_intCurrentStep = m_intStepCount Then Call OnLastStep
        If m_intCurrentStep = 1 Then Call OnFirstStep

        '// tell the controller we finished the step transition
        Call m_hClient.AfterStep

    End If

End If      '// the client didn't cancel

'// unset this
m_bStepping = False


End Sub

'
'   Centers the mouse pointer on a given
'   control that has a hWnd property.
'
Private Sub CenterCursor(hwnd As Long)

Dim lppt As POINTAPI
Dim lprc As RECT

If hwnd = 0 Then Exit Sub

Call GetWindowRect(hwnd, lprc)
lppt.x = (lprc.Left + lprc.Right) \ 2
lppt.y = (lprc.Top + lprc.Bottom) \ 2

Call SetCursorPos(lppt.x, lppt.y)

End Sub


'
'   Can be used to query the wizard on the
'   load state of the UI.
'
Public Property Get FinishedLoading() As Boolean

FinishedLoading = m_bFinishedLoading

End Property


'
'   Sets/returns whether or not the wizard's
'   UI is frozen and inaccessible
'
Public Property Let Frozen(ByVal vData As Boolean)

m_bFrozen = vData

'// note that you can customize this. For example,
'// you might wish to leave the 'Cancel' button enabled
'// in case the user wants to abort your lengthy process.

cmdHelp.Enabled = (Not m_bFrozen)
cmdStepBack.Enabled = (Not m_bFrozen) And (m_intCurrentStep <> 1)
cmdStepFwd.Enabled = (Not m_bFrozen) And (m_intCurrentStep <> m_intStepCount)
cmdTerminate.Enabled = (Not m_bFrozen)

picPage.Enabled = (Not m_bFrozen)

DoEvents    '// make sure our changes are reflected visually

End Property

'
'   Used internally to signal a page change
'
Private Sub NotifyPageChange()

On Error Resume Next

Call RecalculateLayout

cmdStepBack.Enabled = True
cmdStepFwd.Enabled = True
cmdStepFwd.Default = True
If cmdTerminate.Caption = "Finish" Then cmdTerminate.Caption = "Cancel"
cmdTerminate.Default = (m_intCurrentStep = m_intStepCount)

picPage.TabIndex = 5
'picPage.SetFocus

End Sub

'
'   This procedure takes care of resizing the wizard form
'   to correctly fit each page form as the step sequence
'   changes. It correctly resizes and rearranges all UI
'   elements on the form.
'
Private Function RecalculateLayout() As Boolean

Dim cxOffset As Long
Dim cyOffset As Long
Dim cxPage As Long
Dim cyPage As Long
Dim xButton As Long
Dim cxButton As Long
Dim xForm As Long
Dim yForm As Long
Dim cxForm As Long
Dim cyForm As Long

Dim lprcBefore As RECT
Dim lprcAfter As RECT

On Error Resume Next

'// Flag this so we don't have UI change notifications
'// while we're moving things around
m_bFinishedLoading = False

'// Calculate the collective widths of the buttons
cxButton = cxButton + (cmdStepBack.Width + cmdStepFwd.Width + cmdTerminate.Width)

'// We have a margin of 4 px into the form
cxOffset = 4

picSplash.Move cxOffset, 4
cxOffset = (picSplash.Left + picSplash.Width)

'// do some padding
'cxOffset = cxOffset + 6

'// Get the size of the current page
If Not FetchPageSize(cxPage, cyPage) Then
    m_bFinishedLoading = True
    Exit Function

End If

'// Move the bin to fit the form page
picPage.Move cxOffset, 0, cxPage + 2, cyPage + 2


'// See what's bigger (splash or Page), and position the page
'// accordingly
If picSplash.Height > picPage.Height Then
    'picPage.Move cxOffset, ((picSplash.Height + picSplash.Top) \ 2) - (picPage.Height \ 2)
    'picPage.Move cxOffset, + picSplash.Top) \ 2) - (picPage.Height \ 2)
    picPage.Move cxOffset, 2

Else
    picPage.Move cxOffset, 2

End If

'Picture1.Move picPage.Left, picPage.Top, picPage.Width, picSplash.Height

cxOffset = picPage.Left + picPage.Width

'// See what's taller and adjust the Y offset to that
cyOffset = Max((picPage.Top + picPage.Height), (picSplash.Height + picSplash.Top))

cyOffset = cyOffset + 6

'// See what's wider, the width of all the buttons plus padding,
'// or the width of the Splash plus the page.
cxButton = cxButton + 30
cxOffset = Max(cxButton, cxOffset)

'// Adjust the bevel lines
lnBevel(0).Y1 = cyOffset
lnBevel(0).Y2 = cyOffset
lnBevel(0).X1 = 4
lnBevel(0).X2 = cxOffset + 3

lnBevel(1).Y1 = cyOffset + 1
lnBevel(1).Y2 = cyOffset + 1
lnBevel(1).X1 = 4
lnBevel(1).X2 = cxOffset + 3

'// Make some additional space
cxOffset = cxOffset + 4
cyOffset = cyOffset + 6

'// move the help button
cmdHelp.Move 4, cyOffset

'// each button is 82 px wide, so we offset that
xButton = cxOffset - 82

cmdTerminate.Move xButton, cyOffset
xButton = (cmdTerminate.Left - 6) - 82     '// Six px separation between the cancel bt and the other two

cmdStepFwd.Move xButton, cyOffset
cxButton = cxButton + cmdStepFwd.Width

'// Again, see where the offset is and move the button back
xButton = cmdStepFwd.Left - 2
cmdStepBack.Move xButton - 80, cyOffset
cxButton = cxButton + cmdStepBack.Width

'// Finally, offset 4 px room between the bottom edge of the buttons
'// and the client edge of the form
cyOffset = (cyOffset + 4) + cmdTerminate.Height

'// The 120 and 410 are magic numbers used for padding. We are converting the
'// offsets into twips before calculating where the form should be.
cxForm = ((cxOffset * Screen.TwipsPerPixelX) + 120)
cyForm = ((cyOffset * Screen.TwipsPerPixelY) + 410)

'// Get the rect of the current layout
Call GetWindowRect(Me.hwnd, lprcBefore)

'// Set the height and width
Me.Height = cyForm + 100
Me.Width = cxForm

'// See if we're supposed to center ourselves in each layout recalc when
'// the size changes
If m_bAutoCenterOnResize Then

    '// Get the rect of the window after the adjustment
    Call GetWindowRect(Me.hwnd, lprcAfter)

    '// See if the two rects are equal. If not, we need to center
    If EqualRect(lprcBefore, lprcAfter) = 0 Then
        xForm = (Screen.Width \ 2) - (cxForm \ 2)
        yForm = (Screen.Height \ 2) - (cyForm \ 2)

        '// Move to the center
        Me.Move xForm, yForm

        '// And center the cursor over the last button clicked, if any.
        Call CenterCursor(m_hWndLastClicked)

    End If

End If

'// Unflag
m_bFinishedLoading = True

End Function


'
'   This method can be used by form pages in conjuntion
'   with control's GotFocus or LostFocus events to transfer
'   the keyboard focus to the FWizard form.
'
Public Sub TransferFocus()

If m_bStepping Then Exit Sub

'// make sure we have a valid focus target
If m_hWndLastClicked = 0 Then m_hWndLastClicked = cmdStepFwd.hwnd

'// We need to re-assign the tab index order because
'// VB gets confused once the form regains the focus.
cmdStepFwd.TabIndex = 0
cmdTerminate.TabIndex = 1
cmdHelp.TabIndex = 2
cmdStepBack.TabIndex = 3
picPage.TabIndex = 4

Call winSetFocus(m_hWndLastClicked)


End Sub

'
'   This method is designed to be called from a page's
'   KeyDown or KeyUp event. Note that, regardless of how
'   we implement these workarounds, certain combinations
'   won't work. For example, a form will not catch an
'   Enter key when a command button has the focus.
'
Public Sub TranslateKey(KeyCode As Integer, Shift As Integer)

Dim fKillKey As Boolean

If m_bStepping Then Exit Sub

'// For each case, we need to check that we can actually do something
'// with the keystroke
Select Case KeyCode

    Case vbKeyReturn
        If Shift = 0 Then
            If (cmdStepFwd.Default And cmdStepFwd.Enabled) Then
                Call cmdStepFwd_Click
                fKillKey = True
            ElseIf (cmdStepBack.Default And cmdStepBack.Enabled) Then
                Call cmdStepBack_Click
                fKillKey = True
            ElseIf (cmdTerminate.Default And cmdTerminate.Enabled) Then
                Call cmdTerminate_Click
                fKillKey = True
            Else
                Beep
            End If
        End If

    Case vbKeyEscape
        If Shift = 0 Then
            If cmdTerminate.Enabled Then
                Call cmdTerminate_Click
                fKillKey = True
            End If

        End If

    Case vbKeyLeft      '// Alt+LeftArrow
        If ((Shift = 4) And cmdStepBack.Enabled) Then _
            Call Step(wizBack)

    Case vbKeyRight '// Alt+RightArrow
        If ((Shift = 4) And cmdStepFwd.Enabled) Then _
            Call Step(wizForward)

End Select

'// if we processed the keystroke, don't let
'// it return to the form page.
If fKillKey Then KeyCode = 0

End Sub

'
'
'
Private Sub cmdHelp_Click()

On Error Resume Next

m_hWndLastClicked = cmdHelp.hwnd
Call m_hClient.OnRequestHelp

End Sub

'
'
'
Private Sub cmdStepBack_Click()

m_hWndLastClicked = cmdStepBack.hwnd
Step wizBack

End Sub

'
'
'
Private Sub cmdStepFwd_Click()

m_hWndLastClicked = cmdStepFwd.hwnd
Step wizForward

End Sub


'
'
'
Private Sub cmdTerminate_Click()

m_hWndLastClicked = cmdTerminate.hwnd
Unload Me


End Sub

'
'
'
Private Sub Form_Initialize()

ReDim m_Steps(-1 To -1)        '// initialize the array to an invalid dimension

'// These are initialized to invalid values too.
m_intCurrentStep = -1
m_intLastStep = -1
m_intNextStep = -1
m_intStepCount = -1
m_bAutoCenterOnResize = True

End Sub



Private Sub Form_Paint()

'// give the "splash" picture a recessed look.

Dim lprc As RECT

With lprc

    .Left = picSplash.Left
    .Top = picSplash.Top
    .Bottom = .Top + picSplash.Height
    .Right = .Left + picSplash.Width

End With

Call InflateRect(lprc, 1, 1)
Call DrawEdge(Me.hdc, lprc, BDR_SUNKENOUTER, BF_RECT)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim bCancel As Boolean

Call m_hClient.OnCancel(bCancel)
Cancel = bCancel

If Not bCancel Then Call finish

End Sub


Private Sub Form_Unload(Cancel As Integer)

'// clean up
Erase m_Steps
Set FWizard = Nothing

End Sub


Private Sub picPage_GotFocus()

On Error Resume Next

'// once the container picturebox gets the focus, we just
'// transfer it to the page.
Call winSetFocus(CForm(m_Steps(m_intCurrentStep).Handle).hwnd)

End Sub


