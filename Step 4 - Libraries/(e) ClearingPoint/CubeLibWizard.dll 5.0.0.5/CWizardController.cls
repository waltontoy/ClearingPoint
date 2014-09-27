VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWizardController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'// CWizardController -- A custom class used to implement the IWizardController
'//                      base interface. This class could have been named anything.
'//                      The type name does not matter as long as the class
'//                      the proper interface.
Option Explicit
DefLng A-Z

'// Local handle to an instance of the FWizard form class
Private m_hWizard As FWizard


'// This tells VB that we can act as an IWizardController too.
Implements IWizardController
'local variable(s) to hold property value(s)

Public Property Set DatabaseConnection(ByVal vData As ADODB.Connection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DatabaseConnection = Form1
    Set G_conDatabase = vData
End Property


Public Property Get DatabaseConnection() As ADODB.Connection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DatabaseConnection
    If IsObject(G_conDatabase) Then
        Set DatabaseConnection = G_conDatabase
    Else
        DatabaseConnection = G_conDatabase
    End If
End Property




Private Sub Class_Initialize()

'// Create a new FWizard object...
Set m_hWizard = New FWizard

'// ...and tell it who's in charge.
Set m_hWizard.ClientHandle = Me

''// Add some pages.
'With m_hWizard
'    .AddPage New FCompanySetup_1, "One"
'    .AddPage New FCompanySetup_2, "Two"
'    .AddPage New FCompanySetup_3, "Three"
'    .AddPage New FCompanySetup_4, "Four"
'    .AddPage New FCompanySetup_5, "Five"
'    .AddPage New FCompanySetup_6, "Six"
'End With

''// And go!
'm_hWizard.Start

End Sub

Public Sub AddPage(hPage As IWizardPage, Optional ByVal Key As Variant)
    ''// Add some pages.
With m_hWizard
    .AddPage hPage, Key
'    .AddPage New FCompanySetup_2, "Two"
'    .AddPage New FCompanySetup_3, "Three"
'    .AddPage New FCompanySetup_4, "Four"
'    .AddPage New FCompanySetup_5, "Five"
'    .AddPage New FCompanySetup_6, "Six"
End With
End Sub

Public Sub Start(Optional FirstStep As Variant)
    '// And go!
    m_hWizard.Start
End Sub

Private Sub IWizardController_AfterStep()

'// We're not doing anything here, but we need to add something
'// or VB will complain about unimplemented interfaces.

End Sub

Private Sub IWizardController_BeforeStep(ByVal Direction As WizardStepDirections, Cancel As Boolean)

'// We're not doing anything here, but we need to add something
'// or VB will complain about unimplemented interfaces.


End Sub


Private Sub IWizardController_BeforeUILoad()

'// We're not doing anything here, but we need to add something
'// or VB will complain about unimplemented interfaces.


End Sub


Private Sub IWizardController_OnCancel(Cancel As Boolean)

'// Are we on the last step, or not?
If m_hWizard.CurrentStep <> m_hWizard.StepCount Then
    Cancel = Not (MsgBox("Are you sure you want to cancel the wizard?", vbYesNo Or vbQuestion, "Wizard") = vbYes)
    
    ConfigurationCancelled = Not Cancel
End If

End Sub


Private Sub IWizardController_OnRequestHelp()

'// We're not doing anything here, but we need to add something
'// or VB will complain about unimplemented interfaces.

MsgBox "You should also provide instructions for your Wizards!"


End Sub


Private Property Get IWizardController_WizardHandle() As Object

'// return the handle to our instance of FWizard
Set IWizardController_WizardHandle = m_hWizard

End Property

Public Property Let ConfigurationCancelled(ByVal Cancelled As Boolean)
    
    g_blnConfigCancelled = Cancelled
    
End Property

Public Property Get ConfigurationCancelled() As Boolean
    
    ConfigurationCancelled = g_blnConfigCancelled
    
End Property


