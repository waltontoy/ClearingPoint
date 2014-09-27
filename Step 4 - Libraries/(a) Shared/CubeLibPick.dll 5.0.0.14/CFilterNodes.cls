VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'local variable to hold collection
Private mCol As Collection
Dim colFilter As Collection

Public Function Add(ByVal Filter As String, _
                                       Optional ByVal sKey As String) As CFilterNode
    'create a new object
    Dim objNewMember As CFilterNode
    Set objNewMember = New CFilterNode

    'set the properties passed into the method
    objNewMember.Key = sKey
    
      'Filter = Trim$(Filter)
      'Filter = Replace(Filter, "'", "''")
      'Filter = Replace(Filter, "%", "%%")
      
    objNewMember.Filter = Filter

    objNewMember.Index = mCol.Count
    
    If Len(sKey) = 0 Then
        'sKey = CStr(objNewMember.Index)
        mCol.Add objNewMember
    'End If
    Else
        mCol.Add objNewMember, sKey
    End If
    'objNewMember.Index mCol.Count

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As CFilterNode
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

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
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
    Set colFilter = Nothing
End Sub

Public Sub Clear()
   Set mCol = New Collection
End Sub

'public function FilterCount()
Public Property Get FilterCount() As Integer
    
   Dim intFilterCtr As Integer
   Dim intResult As Integer
   
   If mCol.Count > 0 Then
      intResult = 1
      For intFilterCtr = 1 To mCol.Count
         'mcol.Item(intfilterctr).
         intResult = intResult * OrCount(mCol.Item(intFilterCtr).Filter)
      Next intFilterCtr
   Else
      intResult = 0
   End If
   
   FilterCount = intResult
   
End Property

Public Function OrCount(ByVal Filter As String) As Integer

   Dim strSplit() As String
   Dim intReturn As Integer
   
   If InStr(1, UCase$(Filter), " OR ", vbTextCompare) <> 0 Then
      strSplit = Split(Filter, "OR", , vbTextCompare)
      intReturn = UBound(strSplit) + 1
   Else
      intReturn = 1
   End If
   
   OrCount = intReturn
   
End Function

Public Function OrList(ByVal Filter As String, ByVal FilterIndex As Integer) As String

   Dim strSplit() As String
   Dim strReturn As String
   
   If InStr(1, UCase$(Filter), " OR ", vbTextCompare) <> 0 Then
      strSplit = Split(Filter, "OR", , vbTextCompare)
      strReturn = Trim$(strSplit(FilterIndex - 1))
   Else
      strReturn = Filter
   End If
   
   OrList = strReturn
   
End Function

Public Property Get FilterList(Index As Integer) As String
'
   'mcol.Add
   FilterList = colFilter.Item(Index).Filter
   
End Property

Public Sub RefreshFilter()
'
   Dim intFilterCtr As Integer
   Dim intAndCtr As Integer
   Dim intOrCount  As Integer
   Dim intFilterCount As Integer
   Dim intFilterIndex As Integer
   Dim strNewFilter As String
   Dim strSelFilter As String
   Dim intCurFilterCount As Integer
   Dim intSubIndex As Integer
   Dim clsFilter As CFilterNode
   
   Set colFilter = New Collection
   
   intFilterCount = FilterCount
   
   ' initialize the filter count object
   For intFilterCtr = 1 To intFilterCount
      Set clsFilter = New CFilterNode
      clsFilter.Filter = ""
      clsFilter.Index = colFilter.Count
      colFilter.Add clsFilter, "key" & CStr(intFilterCtr)
   Next intFilterCtr
   
   intCurFilterCount = intFilterCount
   For intAndCtr = 1 To mCol.Count
      
      intOrCount = OrCount(mCol.Item(intAndCtr).Filter)
      intCurFilterCount = intCurFilterCount / intOrCount
      intSubIndex = intCurFilterCount * intOrCount
      
      For intFilterCtr = 1 To intFilterCount
         
         strNewFilter = ""
         If intOrCount = 1 Then
            intFilterIndex = 1
            strSelFilter = OrList(mCol.Item(intAndCtr).Filter, intFilterIndex)
            strNewFilter = colFilter.Item(intFilterCtr).Filter & " AND (" & strSelFilter & ")"
            colFilter.Item(intFilterCtr).Filter = strNewFilter
         ElseIf intOrCount > 1 Then
            'intSubIndex = intCurFilterCount * intOrCount
            intSubIndex = intFilterCtr
            Do
              intSubIndex = intSubIndex - (intCurFilterCount * intOrCount)
            Loop While intSubIndex > 0
            intSubIndex = intSubIndex + (intCurFilterCount * intOrCount)
            
               intFilterIndex = ((intSubIndex - 1) \ (intCurFilterCount)) + 1
               'intFilterIndex = intFilterIndex - ((intFilterCtr - 1) \ intSubIndex)
            strSelFilter = OrList(mCol.Item(intAndCtr).Filter, intFilterIndex)
            strNewFilter = colFilter.Item(intFilterCtr).Filter & " AND (" & strSelFilter & ")"
            colFilter.Item(intFilterCtr).Filter = strNewFilter
            'strNewFilter = strNewFilter & " AND (" & strSelFilter & ")"
         End If
         ' get the proper index
'         intFilterIndex = ((intAndCtr - 1) \ (intCurFilterCount / intOrCount)) + 1
'         strSelFilter = OrList(mCol.Item(intAndCtr).Filter, intFilterIndex)
'         strNewFilter = strNewFilter & " AND (" & strSelFilter & ")"
      Next intFilterCtr
      Debug.Print mCol.Item(intAndCtr).Filter
   Next intAndCtr


   For intFilterCtr = 1 To colFilter.Count
      ' remove the first and
      colFilter.Item(intFilterCtr).Filter = Right$(colFilter.Item(intFilterCtr).Filter, Len(colFilter.Item(intFilterCtr).Filter) - 5)
      Debug.Print colFilter.Item(intFilterCtr).Filter
   Next intFilterCtr
   
   Set clsFilter = Nothing


End Sub

Public Property Get FilterGroup() As String
   
   Dim intFilterCtr As Integer
   Dim strReturn As String
   
   For intFilterCtr = 1 To FilterCount
      strReturn = strReturn & " OR (" & FilterList(intFilterCtr) & ")"
   Next intFilterCtr
   
   strReturn = Right$(strReturn, Len(strReturn) - 4)
   
   FilterGroup = strReturn

End Property

