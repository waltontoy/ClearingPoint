VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWizardPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'// IWizardPage --  A base class used to define the behavior of forms
'//                 that act as "pages" in a Wizard project. This interface
'//                 does not do anything -- it's meant to be implemented.
Option Explicit
DefLng A-Z

Public Sub BeforePageShow(Wizard As Object, ByVal CurrentStep As Integer)
Attribute BeforePageShow.VB_Description = "Called by the Wizard before the form is shown and becomes the current step in the wizard sequence."

'// This method is called each time the wizard is about to show
'// the page.


End Sub

Public Sub BeforePageHide(Wizard As Object, ByVal NextStep As Integer, Cancel As Boolean)
Attribute BeforePageHide.VB_Description = "Called by the Wizard before the form is hidden to be replaced by the next page in the wizard sequence."

'// This method is called each time the wizard is about to hide
'// the page.

End Sub




