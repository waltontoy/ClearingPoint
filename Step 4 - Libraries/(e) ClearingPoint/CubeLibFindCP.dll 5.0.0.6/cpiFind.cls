VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarSelectedItemTag As Variant 'local copy
Private mvarListSubItems As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarSubItems As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarSelectedItemText As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarOpenOnly As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarstrYear As Variant 'local copy
Private m_blnSearchInProgress As Boolean

Private mblnPLDAI As Boolean
Private mblnPLDAE As Boolean
Private mblnPLDAC As Boolean
Private mblnNCTS As Boolean
Private mblnSADI As Boolean
Private mblnSADET As Boolean
Private mblnSADTC As Boolean

Public Property Let LicPLDAI(ByVal bData As Boolean)
    mblnPLDAI = bData
End Property

Public Property Get LicPLDAI() As Boolean
    LicPLDAI = mblnPLDAI
End Property

Public Property Let LicPLDAE(ByVal bData As Boolean)
    mblnPLDAE = bData
End Property

Public Property Get LicPLDAE() As Boolean
    LicPLDAE = mblnPLDAE
End Property

Public Property Let LicPLDAC(ByVal bData As Boolean)
    mblnPLDAC = bData
End Property

Public Property Get LicPLDAC() As Boolean
    LicPLDAC = mblnPLDAC
End Property

Public Property Let LicNCTS(ByVal bData As Boolean)
    mblnNCTS = bData
End Property

Public Property Get LicNCTS() As Boolean
    LicNCTS = mblnNCTS
End Property

Public Property Let LicSADI(ByVal bData As Boolean)
    mblnSADI = bData
End Property

Public Property Get LicSADI() As Boolean
    LicSADI = mblnSADI
End Property

Public Property Let LicSADET(ByVal bData As Boolean)
    mblnSADET = bData
End Property

Public Property Get LicSADET() As Boolean
    LicSADET = mblnSADET
End Property

Public Property Let LicSADTC(ByVal bData As Boolean)
    mblnSADTC = bData
End Property

Public Property Get LicSADTC() As Boolean
    LicSADTC = mblnSADTC
End Property

Public Property Let strYear(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strYear = 5
    mvarstrYear = vData
End Property


Public Property Set strYear(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.strYear = Form1
    Set mvarstrYear = vData
End Property


Public Property Get strYear() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strYear
    If IsObject(mvarstrYear) Then
        Set strYear = mvarstrYear
    Else
        strYear = mvarstrYear
    End If
End Property

Public Property Let OpenOnly(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OpenOnly = 5
    mvarOpenOnly = vData
End Property

Public Property Set OpenOnly(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.OpenOnly = Form1
    Set mvarOpenOnly = vData
End Property

Public Property Get OpenOnly() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OpenOnly
    If IsObject(mvarOpenOnly) Then
        Set OpenOnly = mvarOpenOnly
    Else
        OpenOnly = mvarOpenOnly
    End If
End Property

Public Property Let SelectedItemText(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedItemText = 5
    mvarSelectedItemText = vData
End Property

Public Property Set SelectedItemText(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SelectedItemText = Form1
    Set mvarSelectedItemText = vData
End Property

Public Property Get SelectedItemText() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedItemText
    If IsObject(mvarSelectedItemText) Then
        Set SelectedItemText = mvarSelectedItemText
    Else
        SelectedItemText = mvarSelectedItemText
    End If
End Property

Public Property Let SubItems(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubItems = 5
    mvarSubItems = vData
End Property


Public Property Set SubItems(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SubItems = Form1
    Set mvarSubItems = vData
End Property

Public Property Get SubItems() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubItems
    If IsObject(mvarSubItems) Then
        Set SubItems = mvarSubItems
    Else
        SubItems = mvarSubItems
    End If
End Property

Public Property Let ListSubItems(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ListSubItems = 5
    mvarListSubItems = vData
End Property

Public Property Set ListSubItems(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ListSubItems = Form1
    Set mvarListSubItems = vData
End Property

Public Property Get ListSubItems() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ListSubItems
    If IsObject(mvarListSubItems) Then
        Set ListSubItems = mvarListSubItems
    Else
        ListSubItems = mvarListSubItems
    End If
End Property

Public Property Let SelectedItemTag(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedItemTag = 5
    mvarSelectedItemTag = vData
End Property

Public Property Set SelectedItemTag(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SelectedItemTag = Form1
    Set mvarSelectedItemTag = vData
End Property

Public Property Get SelectedItemTag() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedItemTag
    If IsObject(mvarSelectedItemTag) Then
        Set SelectedItemTag = mvarSelectedItemTag
    Else
        SelectedItemTag = mvarSelectedItemTag
    End If
End Property

Public Sub ShowFind(ByRef OwnerForm As Object, ByRef clsFind As cpiFind, _
                    ByVal lngResourceHandler As Long, ByVal strLanguage As String, _
                    ByVal strAppTitle As String, ByVal strUserID As String, _
                    ByVal About As Object)
    
    Set CallingForm = OwnerForm
    Set clsFindForm = clsFind
    
    AppTitle = strAppTitle
    cAppPath = GetSetting(AppTitle, "Settings", "MdbPath")
    cLanguage = strLanguage
    ResourceHandler = lngResourceHandler
    UserID = strUserID
    Set g_clsAbout = About
    
    m_blnSearchInProgress = False
    
    frm_Find.Show
End Sub

Public Property Get SearchInProgress() As Boolean
    m_blnSearchInProgress = frm_Find.SearchInProgress
    
    SearchInProgress = m_blnSearchInProgress
End Property

Public Sub UnloadFind()
    If Not m_blnSearchInProgress Then
        On Error Resume Next
        Unload frm_Find
        Set frm_Find = Nothing
        On Error GoTo 0
    End If
End Sub
