VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPicklistInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' CONSTANTS
Const CPI_TRUE = -1
Const CPI_FALSE = 0
Const CPI_AUTOCANCEL = 1
    
' EVENTS
Public Event BtnClick(ByRef Record As CRecord, _
                                    ByVal Button As ButtonType, _
                                    ByRef Cancel As Boolean)
  
Public Event BeforeCommit(ByRef SelectedRec As CRecord, ByRef Cancel As Boolean)
  
' FORMS
Private mvarFrmCatalog As Form
Private mvarOwnerForm As Object

' CONTROLS
Private WithEvents mvarCmdOK As CommandButton
Attribute mvarCmdOK.VB_VarHelpID = -1

' CLASSES
Private mvarGridColumns As GridEX16.JSColumns

' CPI CLASSES /COLLECTIONS
Private mvarSelectedRecord As CRecord
Private mvarTransactions As CTransactions
Private mvarPicklistFilter As CPicklistFilter
Private mvarGridSeed As CGridSeed
Private mvarColumns As CGridColumns

' CPI CONSTANTS
Private mvarPicklistStyle As PicklistStyle
Private mvarActiveStatus As cpiActiveStatusConstants
Private mvarActiveKey As cpiKeyPickConstants

'ADO objects
Private mvarActiveConnection As ADODB.Connection
Private mvarGridRecord As ADODB.Recordset

' NATIVE DATA TYPES
Private mvarShowMode As Integer

Private mvarTransactionCtr As Long

Private mvarPKFieldAlias As String
Private mvarPKFieldBaseName As String
Private mvarBaseSQL As String
Private mvarBaseConnection As String
Private mvarCaption As String
Private mvarSearchValue As String
Private mvarSearchField As String

Private mvarLoadAllRecord As Boolean
Private mvarAddButton As Boolean
Private mvarDeleteButton As Boolean
Private mvarModifyButton As Boolean
Private mvarCopyButton As Boolean
Private mvarCancelTrans As Boolean
Private mvarAutoSearch As Boolean
Private mvarPickEnd As Boolean

Private mvarListType As cpiListTypeConstants

Private mvarAutoUnload As Variant

Private Type CPicklistInterface
    IGrid As Object
    IButtonOK As CommandButton
    IButtonCancel As CommandButton
    IButtonAdd As CommandButton
    IButtonModify As CommandButton
    IButtonCopy As CommandButton
    IButtonDelete As CommandButton
    ISearchBox1 As TextBox
    ISearchBox2 As TextBox
    ISearchBox3 As TextBox
    ISearchBox4 As TextBox
    ISearchBox5 As TextBox
End Type

Public Property Get ListType() As cpiListTypeConstants
    ListType = mvarListType
End Property
    
Public Property Let ListType(ByVal vData As cpiListTypeConstants)
    mvarListType = vData
End Property
    
Public Property Get ShowMode() As Integer
    ShowMode = mvarShowMode
End Property
    
Friend Property Let ShowMode(ByVal vData As Integer)
    mvarShowMode = vData
End Property
    
Public Property Get OwnerForm() As Object
    Set OwnerForm = mvarOwnerForm
End Property
    
Public Property Set OwnerForm(ByVal vData As Object)
    Set mvarOwnerForm = vData
End Property
    
Public Property Get ActiveConnection() As ADODB.Connection
    Set ActiveConnection = mvarActiveConnection
End Property
    
Public Property Let ActiveConnection(ByVal vData As ADODB.Connection)
    Set mvarActiveConnection = vData
End Property

Public Property Get PickEnd() As Boolean
    PickEnd = mvarPickEnd
End Property
    
Friend Property Let PickEnd(ByVal vData As Boolean)
    mvarPickEnd = vData
End Property
    
Public Property Get Columns() As CGridColumns
    Set Columns = mvarColumns
End Property
    
Public Property Let Columns(ByVal vData As CGridColumns)
    Set mvarColumns = vData
End Property

Public Property Get ActiveKey() As cpiKeyPickConstants
    ActiveKey = mvarActiveKey
End Property
    
Public Property Let ActiveKey(ByVal vData As cpiKeyPickConstants)
    mvarActiveKey = vData
End Property

Public Property Get ActiveStatus() As cpiActiveStatusConstants
    ActiveStatus = mvarActiveStatus
End Property
    
Public Property Let ActiveStatus(ByVal vData As cpiActiveStatusConstants)
    mvarActiveStatus = vData
End Property

Public Property Get SearchValue() As String
    SearchValue = mvarSearchValue
End Property
    
Public Property Let SearchValue(ByVal vData As String)
    mvarSearchValue = vData
End Property

Public Property Get SearchField() As String
    SearchField = mvarSearchField
End Property
    
Public Property Let SearchField(ByVal vData As String)
    mvarSearchField = vData
End Property

Public Property Get GridRecord() As ADODB.Recordset
    Set GridRecord = mvarGridRecord
End Property

Friend Property Set GridRecord(vData As ADODB.Recordset)
    Set mvarGridRecord = vData
End Property
    
Public Property Get AutoSearch() As Boolean
    AutoSearch = mvarAutoSearch
End Property

Public Property Let AutoSearch(vData As Boolean)
    mvarAutoSearch = vData
End Property
    
Public Property Get CancelTrans() As Boolean
    CancelTrans = mvarCancelTrans
End Property

Friend Property Let CancelTrans(vData As Boolean)
    mvarCancelTrans = vData
End Property
    
Public Property Get Caption() As String
    Caption = mvarCaption
End Property
    
Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property
    
Public Property Get PicklistStyle() As PicklistStyle
    PicklistStyle = mvarPicklistStyle
End Property
    
Public Property Let PicklistStyle(ByVal vData As PicklistStyle)
    mvarPicklistStyle = vData
End Property
    
Public Property Get AutoUnload() As Variant
    AutoUnload = mvarAutoUnload
End Property
    
Public Property Let AutoUnload(ByVal vData As Variant)
    mvarAutoUnload = vData
End Property
    
Public Property Get LoadAllRecord() As Boolean
    LoadAllRecord = mvarLoadAllRecord
End Property
    
Public Property Get GridColumns() As JSColumns
    Set GridColumns = mvarGridColumns
End Property

Public Property Set GridColumns(ByVal vData As JSColumns)
    Set mvarGridColumns = vData
End Property
    
Public Property Get AddButton() As Boolean
    AddButton = mvarAddButton
End Property
    
Public Property Let AddButton(ByVal vData As Boolean)
    mvarAddButton = vData
End Property
    
Public Property Get DeleteButton() As Boolean
    DeleteButton = mvarDeleteButton
End Property
    
Public Property Let DeleteButton(ByVal vData As Boolean)
    mvarDeleteButton = vData
End Property
    
Public Property Get ModifyButton() As Boolean
    ModifyButton = mvarModifyButton
End Property
    
Public Property Let ModifyButton(ByVal vData As Boolean)
    mvarModifyButton = vData
End Property
    
Public Property Get CopyButton() As Boolean
    CopyButton = mvarCopyButton
End Property
    
Public Property Let CopyButton(ByVal vData As Boolean)
    mvarCopyButton = vData
End Property
    
Public Property Let BaseConnection(ByVal vData As String)
    mvarBaseConnection = vData
End Property

Public Property Get BaseConnection() As String
    BaseConnection = mvarBaseConnection
End Property
    
Public Property Let BaseSQL(ByVal vData As String)
    mvarBaseSQL = vData
End Property

Public Property Get BaseSQL() As String
    BaseSQL = mvarBaseSQL
End Property

Friend Property Let PKFieldBaseName(ByVal vData As String)
    mvarPKFieldBaseName = vData
End Property

Friend Property Get PKFieldBaseName() As String
    PKFieldBaseName = mvarPKFieldBaseName
End Property

Private Sub RunPicklist(ByVal ShowMode As String)

    If (ShowMode = "") Then
    
        mvarFrmCatalog.Show
        
    ElseIf (ShowMode <> "") Then
    
        If ((ShowMode = vbModal) Or (ShowMode = vbModeless)) Then
        
            mvarFrmCatalog.Show ShowMode
            
        ElseIf ((ShowMode <> vbModal) And (ShowMode <> vbModeless)) Then
        
            mvarFrmCatalog.Show vbModeless
            
        End If
        
    End If

End Sub

Friend Property Let TransactionCtr(ByVal vData As Long)
    mvarTransactionCtr = vData
End Property

Public Property Get TransactionCtr() As Long
    TransactionCtr = mvarTransactionCtr
End Property

Public Property Get Transactions() As CTransactions
    Set Transactions = mvarTransactions
End Property

Public Property Set Transactions(vData As CTransactions)
    Set mvarTransactions = vData
End Property
    
Public Property Get PicklistFilter() As CPicklistFilter
    Set PicklistFilter = mvarPicklistFilter
End Property
    
Public Property Let PKFieldAlias(ByVal vData As String)
    mvarPKFieldAlias = vData
End Property

Public Property Get PKFieldAlias() As String
    PKFieldAlias = mvarPKFieldAlias
End Property

Friend Function ButtonClick(ByRef CatalogRecord As CRecord, ByVal Button As ButtonType) As Boolean
                            
    Dim blnCancel As Boolean
    Dim frmListType As frmList
    
    blnCancel = False
    
    If mvarListType = cpiUserList Then
        ' the form should be Modal
        RaiseEvent BtnClick(CatalogRecord, Button, blnCancel)
    
    ElseIf mvarListType = cpiSimpleList Then
    
        Set frmListType = New frmList
        ' the form should be Modal
        If (Button <> cpiDelete) Then
            frmListType.ShowForm Me.OwnerForm, CatalogRecord, Button, blnCancel, Me.Caption, "ID", "Name"
        End If
        'frmList.ShowForm
        
        
        Set frmListType = Nothing
        
    End If
    
    ButtonClick = blnCancel
    
End Function

Friend Property Set SelectedRecord(ByVal vData As CRecord)
    Set mvarSelectedRecord = vData
End Property

Public Property Get SelectedRecord() As CRecord
    Set SelectedRecord = mvarSelectedRecord
End Property

Private Function InitPicklist() As CTransactions

    Dim blnContinue As Boolean
    
    Screen.MousePointer = vbHourglass
    
    If (mvarPicklistStyle = cpiSimplePicklist) Then
        mvarAutoUnload = True
    End If
    
    ' check if there is any error before loading the form
    blnContinue = mvarFrmCatalog.ShowPicklist(mvarOwnerForm, _
                                              mvarPicklistStyle, _
                                              mvarActiveConnection, _
                                              mvarBaseSQL, _
                                              mvarPKFieldAlias, _
                                              Me, _
                                              mvarCaption, _
                                              mvarGridSeed, _
                                              mvarPicklistFilter)
                                    
    If (blnContinue = True) Then
    
        Screen.MousePointer = vbDefault
        
        Select Case mvarActiveKey
        
            Case cpiKeyEnter, cpiKeyTabEnter
        
                If ((mvarActiveStatus = cpiManyRecord) Or (mvarActiveStatus = cpiNotFound) Or _
                        (mvarActiveStatus = cpiOneRecord)) Then
                
                    Call RunPicklist(mvarShowMode)
                
                ElseIf (mvarActiveStatus = cpiOneRecordExact) Then
                
                    mvarPickEnd = True
                    
                End If
            
            Case cpiKeyF2
            
                Call RunPicklist(mvarShowMode)
    
            Case cpiKeyTabNoAction
            
                ' do nothing
                
        End Select
       
        ' repeat until frm catalog ends
        Do
          DoEvents
        Loop While (mvarPickEnd = False)
    
        ' return the sekected transactions
        Set InitPicklist = mvarTransactions
    
    ElseIf (blnContinue = False) Then
    
        mvarCancelTrans = True
        
    End If
    
    Screen.MousePointer = vbDefault
    
End Function

Public Function Pick(ByRef OwnerForm As Object, _
                    ByVal Style As PicklistStyle, _
                    ByRef DBConnection As ADODB.Connection, _
                    ByRef RecordsListSQL As String, _
                    ByVal PKFieldAliasInSQL As String, _
                    ByRef PluralEntity As String, _
                    ByVal ShowMode As Integer, _
                    Optional ByRef GridSeed As CGridSeed, _
                    Optional ByRef Filter As CPicklistFilter = Nothing, _
                    Optional ByVal LoadOldRst As Boolean = True, _
                    Optional ByVal blniIsAutoUnload As Boolean, _
                    Optional ByVal KeyPick As cpiKeyPickConstants = cpiKeyEnter) _
                    As CTransactions
                                      
    Dim clsReturn As CTransactions
                                      
    ' check if form
    If ((TypeOf OwnerForm Is Form) = False) Then
        GoTo ERROR_TYPE_MISMATCH
    End If
                                      
    Select Case KeyPick
        
        Case cpiKeyEnter, cpiKeyF2, cpiKeyTabEnter
            
            ' initialize picklist values
            Set mvarOwnerForm = OwnerForm
            mvarPicklistStyle = Style
            
            Debug.Assert False
            '#####
            Set mvarActiveConnection = New ADODB.Connection
            mvarActiveConnection.Open DBConnection.ConnectionString
            mvarBaseSQL = RecordsListSQL
            mvarPKFieldAlias = PKFieldAliasInSQL
            mvarCaption = PluralEntity
            
            ' temporary
            'mvarShowMode = vbModeless
            ' orig
            mvarShowMode = ShowMode
            
            Set mvarGridSeed = GridSeed
            Set mvarPicklistFilter = Filter
            mvarLoadAllRecord = LoadOldRst
            mvarAutoUnload = blniIsAutoUnload
            mvarActiveKey = KeyPick
            
            ' other initialization
            Debug.Assert False
            '#####
            mvarBaseConnection = DBConnection.ConnectionString
            SearchValue = Trim$(SearchValue)
            mvarCancelTrans = False
            mvarPickEnd = False
            If (Trim$(SearchField) = "") Then
                SearchField = mvarPKFieldAlias
            End If

            ' create new picklist form
            Set mvarFrmCatalog = New frmCatalog
            Set mvarCmdOK = mvarFrmCatalog.cmdOK
            Set mvarGridColumns = mvarFrmCatalog.jgxPicklist.Columns
            Set mvarColumns = mvarGridSeed.GridColumns
            
            ' open transactions
            Set mvarTransactions = New CTransactions
    
            ' run picklist
            Set clsReturn = InitPicklist
            Set Pick = clsReturn
            Set clsReturn = Nothing
                        
            ' remove active objects
            Call DestroyObject
                        
        Case cpiKeyTabNoAction
        
            ' do nothing
            
    End Select
    
    Exit Function
    
ERROR_TYPE_MISMATCH:
    Err.Raise 1001, , "Type Mismatch.~Pick~"
    
End Function

Private Sub DestroyObject()
    
    ' initialize picklist values
    Set mvarTransactions = Nothing
    Set mvarColumns = Nothing
    Set mvarGridColumns = Nothing
    Set mvarCmdOK = Nothing
    Set mvarFrmCatalog = Nothing
    
    Set mvarPicklistFilter = Nothing
    Set mvarGridSeed = Nothing
    
    Set mvarActiveConnection = Nothing

    Set mvarOwnerForm = Nothing
    
End Sub

Public Function CreateFilter(ByVal FilterType As enuPicklistFilter, _
                                                ParamArray Filter() As Variant) As CPicklistFilter
    
    Dim intFilterCtr As Integer
    Dim lngDivisor As Long
    
    Set mvarPicklistFilter = New CPicklistFilter
    
    lngDivisor = 3
    
    If (IsEmpty(Filter) = False) Then
    
        If (FilterType = cpiComboRecords) Then
            lngDivisor = 2
        End If
        
        mvarPicklistFilter.FilterType = FilterType
        mvarPicklistFilter.FilterCount = (UBound(Filter) + 1) / lngDivisor
        
        For intFilterCtr = 0 To (mvarPicklistFilter.FilterCount - 1)
        
            If (FilterType <> cpiComboRecords) Then
            
                mvarPicklistFilter.PicklistFilters.Add "Key" & intFilterCtr, _
                CStr(Filter(0 + (intFilterCtr * lngDivisor))), _
                CStr(Filter(1 + (intFilterCtr * lngDivisor))), _
                CBool(Filter(2 + (intFilterCtr * lngDivisor)))
            
            ElseIf (FilterType = cpiComboRecords) Then
            
                mvarPicklistFilter.PicklistFilters.Add "Key" & intFilterCtr, _
                CStr(Filter(0 + (intFilterCtr * lngDivisor))), _
                CStr(Filter(1 + (intFilterCtr * lngDivisor))), True
            
            End If
        
        Next intFilterCtr
    
    End If
    
    Set CreateFilter = mvarPicklistFilter
    
End Function

Private Sub Class_Initialize()
    
    mvarAddButton = True
    mvarDeleteButton = True
    mvarModifyButton = True
    mvarCopyButton = True
    
    mvarAutoUnload = CPI_FALSE

End Sub

Private Sub Class_Terminate()

'---------------------------------------> DESTROY ALL OBJECTS <----------------------------------------------

'5---------------------------------------> CUBEPOINT   OBJECTS <----------------------------------------------

    Set mvarTransactions = Nothing
    Set mvarPicklistFilter = Nothing
    Set mvarColumns = Nothing
    Set mvarGridSeed = Nothing
    
'4--------------------------------------->  THIRD PARTY  OBJECTS <----------------------------------------------
    'remove the columns collection
    Set mvarGridColumns = Nothing

'3---------------------------------------> VB  OBJECTS <----------------------------------------------
    Set mvarCmdOK = Nothing

'--------------------------------------->  ADO objects
    Set mvarActiveConnection = Nothing

'2---------------------------------------> VB FORMS <----------------------------------------------

    Set mvarFrmCatalog = Nothing
    Set mvarOwnerForm = Nothing
    
'1---------------------------------------> MISCELLANEOUS <----------------------------------------------

    Set mvarSelectedRecord = Nothing
    Set mvarTransactions = Nothing
    Set mvarGridRecord = Nothing

End Sub

Public Function SeedGrid(ParamArray Columns() As Variant) As CGridSeed
    
    Dim clsReturn As CGridSeed
    Dim intColumnCtr As Integer
    
    Set clsReturn = New CGridSeed
    
    ' clear the columns
    For intColumnCtr = 1 To ((UBound(Columns) + 1) / 3)
    
        clsReturn.GridColumns.Add CStr(Columns(0 + ((intColumnCtr - 1) * 3))) _
                                            , CLng(Columns(1 + ((intColumnCtr - 1) * 3))) _
                                            , CStr(Columns(2 + ((intColumnCtr - 1) * 3))) _
                                            , CStr(Columns(0 + ((intColumnCtr - 1) * 3)))
        
    Next intColumnCtr
    
    Set mvarColumns = clsReturn.GridColumns
    Set SeedGrid = clsReturn
    Set clsReturn = Nothing
    
End Function

Private Sub Initialize_Form()

   Dim ctlControl As Control
    
    ' 1- hide all controls in frmCatalog
    For Each ctlControl In mvarFrmCatalog.Controls
      ctlControl.Visible = False
    Next

End Sub

Private Sub mvarCmdOK_Click()
    
    RaiseEvent BeforeCommit(mvarSelectedRecord, mvarCancelTrans)

End Sub

Public Sub Search(Optional ByVal AutoSearch As Boolean, _
                                    Optional ByVal FieldAlias As String, _
                                    Optional ByVal SearchValue As String)

    mvarAutoSearch = AutoSearch
    
    If (FieldAlias <> "") Then
    
        mvarSearchField = FieldAlias
        
    End If
    
    If (SearchValue <> "") Then
    
        mvarSearchValue = SearchValue
        
    End If

End Sub





