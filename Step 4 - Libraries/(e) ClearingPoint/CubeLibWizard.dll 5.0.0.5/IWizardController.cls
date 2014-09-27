VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWizardController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'// IWizardController --  A base class used to define the behavior of objects
'//                       that act as "controllers" in a Wizard project. This
'//                       interface does not do anything -- it's meant to be
'//                       implemented.
Option Explicit
DefLng A-Z

'// possible values for the Step() method and others
Public Enum WizardStepDirections
    wizForward = 1
    wizBack = -1
End Enum

Public Sub BeforeUILoad()
Attribute BeforeUILoad.VB_Description = "Called by the UI component of the Wizard before it becomes visible."

'// Called by the FWizard before it finishes loading its graphical interface

End Sub

Public Sub BeforeStep(ByVal Direction As WizardStepDirections, ByRef Cancel As Boolean)
Attribute BeforeStep.VB_Description = "Called by the Wizard before the current page changes."

'// Called by the wizard before a step is loaded and initialized

End Sub

Public Sub AfterStep()
Attribute AfterStep.VB_Description = "Called by the Wizard after the current step has changed."

'// Called by the wizard after a step is loaded and initialized

End Sub

Public Sub OnRequestHelp()
Attribute OnRequestHelp.VB_Description = "Called by the UI component of the Wizard in response to the 'Help' button being pressed."

'// Called by the FWizard when the user clicks on the 'Help' button.

End Sub

Public Sub OnCancel(ByRef Cancel As Boolean)
Attribute OnCancel.VB_Description = "Called by the UI component of the Wizard in response to the 'Cancel' or 'Finish' button being pressed."

'// Called by the FWizard when the user presses the 'Cancel' button on
'// the form. Note that this gets called even on the last step (when the
'// button reads "Finish", so the controller needs to check the wizard's
'// current step to determine what action it takes.

End Sub


'Public Property Get WizardHandle() As FWizard
Public Property Get WizardHandle() As Object

'// This property can be used by other classes to retrieve a handle
'// to the FWizard class. This way you can avoid using more global
'// variables.

End Property


