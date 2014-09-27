VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEntrepotDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cEntrepotDetail"
Attribute VB_Ext_KEY = "Member0" ,"cEntrepotDetail"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(Stock_ID As Long, In_ID As Long, ByVal OrigQty As String, _
                    ByVal OrigGrossWgt As String, ByVal OrigNetWgt As String, _
                    ByVal sTabCaption As String) As cEntrepotDetail
Dim lngCtr As Long
Dim blnKeyFound As Boolean
''    'create a new object
    Dim objNewMember As cEntrepotDetail
    Set objNewMember = New cEntrepotDetail
    
    If mCol.Count = 0 Then
        objNewMember.Stock_ID = Stock_ID
        objNewMember.In_ID = In_ID
        objNewMember.OrigQty = OrigQty
        objNewMember.OrigGrossWgt = OrigGrossWgt
        objNewMember.OrigNetWgt = OrigNetWgt
        objNewMember.Key = sTabCaption
        
        If Len(sTabCaption) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, sTabCaption
        End If
    Else
        For lngCtr = 1 To mCol.Count
            If mCol(lngCtr).Key = sTabCaption Then
                blnKeyFound = True
            End If
        Next
        If blnKeyFound Then
            mCol.Item(sTabCaption).Stock_ID = Stock_ID
            mCol.Item(sTabCaption).In_ID = In_ID
            mCol.Item(sTabCaption).Key = sTabCaption
            mCol.Item(sTabCaption).OrigQty = OrigQty
            mCol.Item(sTabCaption).OrigGrossWgt = OrigGrossWgt
            mCol.Item(sTabCaption).OrigNetWgt = OrigNetWgt
        
        Else

            objNewMember.Stock_ID = Stock_ID
            objNewMember.In_ID = In_ID
            objNewMember.Key = sTabCaption
            objNewMember.OrigQty = OrigQty
            objNewMember.OrigGrossWgt = OrigGrossWgt
            objNewMember.OrigNetWgt = OrigNetWgt
            
            If Len(sTabCaption) = 0 Then
                mCol.Add objNewMember
            Else
                mCol.Add objNewMember, sTabCaption
            End If

        End If
    End If

''
''
''    'set the properties passed into the method
''    objNewMember.Stock_ID = Stock_ID
''    objNewMember.In_ID = In_ID
''    If Len(sKey) = 0 Then
''        mCol.Add objNewMember
''    Else
''        mCol.Add objNewMember, sKey
''    End If
''
''
''    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(ByVal vntIndexKey As Variant) As cEntrepotDetail
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(ByVal sTabCaption As String, Optional ByVal blnReAdjust As Boolean = True)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    If DoesExistInCollection(sTabCaption) Then
        mCol.Remove sTabCaption
        
        If blnReAdjust = True Then
            ReAdjustCollection sTabCaption
        End If
    End If
    
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


Private Function ReAdjustCollection(sTabCaption)
Dim intPeriodPos As Integer
Dim intCtr As Integer
Dim strHeader As String
Dim strNextTabCaption As String

    intPeriodPos = InStr(1, sTabCaption, ".")
    
    strHeader = Mid(sTabCaption, 2, 1)
    
    If intPeriodPos > 0 Then
        intCtr = CInt(Mid(sTabCaption, intPeriodPos + 1))
        'intCtr = intCtr + 1
        
        For intCtr = intCtr To mCol.Count
        'Do While Not mCol(Left(sTabCaption, intPeriodPos) & intCtr) Is Nothing
            'mCol(Left(sTabCaption, intPeriodPos) & CStr(intCtr)).Stock_ID = mCol(Left(sTabCaption, intPeriodPos) & CStr(intCtr + 1)).Stock_ID
            'mCol(Left(sTabCaption, intPeriodPos) & CStr(intCtr)).In_ID = mCol(Left(sTabCaption, intPeriodPos) & CStr(intCtr + 1)).In_ID
            'mCol(Left(sTabCaption, intPeriodPos) & CStr(intCtr)).Key = mCol(Left(sTabCaption, intPeriodPos) & CStr(intCtr + 1)).Key
            
            If Mid(Left(sTabCaption, intPeriodPos) & CStr(intCtr + 1), 2, 1) = strHeader Then
                    
                strNextTabCaption = Left(sTabCaption, intPeriodPos) & CStr(intCtr + 1)
                
                If DoesExistInCollection(strNextTabCaption) Then
                    Me.Add mCol(strNextTabCaption).Stock_ID, _
                        mCol(strNextTabCaption).In_ID, mCol(strNextTabCaption).OrigQty, _
                        mCol(strNextTabCaption).OrigGrossWgt, _
                        mCol(strNextTabCaption).OrigNetWgt, _
                        Left(sTabCaption, intPeriodPos) & CStr(intCtr)
        
                    Debug.Print mCol(strNextTabCaption).Stock_ID
                    Debug.Print mCol(strNextTabCaption).In_ID
                    Debug.Print mCol(strNextTabCaption).OrigQty
                    Debug.Print mCol(strNextTabCaption).OrigGrossWgt
                    Debug.Print mCol(strNextTabCaption).OrigNetWgt
                    
                    RemoveItem Left(sTabCaption, intPeriodPos) & CStr(intCtr + 1)
                End If
                'intCtr = intCtr + 1
            End If
        'Loop
        Next
    End If

    
End Function

Private Sub RemoveItem(ByVal strToRemove As String)
    mCol.Remove strToRemove
End Sub

Public Function DoesExistInCollection(ByVal strTabCaption As String) As Boolean
Dim intCtr As Integer
    
    For intCtr = 1 To mCol.Count
        If mCol.Item(intCtr).Key = strTabCaption Then
            DoesExistInCollection = True
            Exit For
        End If
    Next
End Function
