VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringExtension"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarGetStringPositionIndex As Integer 'local copy
Private mvarGetStringPositionCount As Integer 'local copy
Private mvarGetStringPositions(0 To 200) As New CElement
Private mvarCurrentString As String 'local copy
Private mvarResultString As String 'local copy
'local variable(s) to hold property value(s)
Private mvarFindString As String 'local copy
'local variable(s) to hold property value(s)
Private mvarReplacementString As String 'local copy
Public Function ReproduceString(Optional ByVal strCurrentString As String, _
                                                              Optional ByVal intCopyCount As Integer) As String
                                                              
   Dim intCounter As Integer
   Dim intMaxCopy As Integer
   Dim strTemp As String
                                                              
  intMaxCopy = intCopyCount
  If Len(strCurrentString) <> 0 Then mvarCurrentString = strCurrentString
  If intCopyCount = 0 Then intMaxCopy = 2
  
  For intCounter = 1 To intMaxCopy
      strTemp = strTemp & strCurrentString
  Next intCounter
                                                              
   mvarResultString = strTemp
   ReproduceString = mvarResultString
                                                              
End Function

Public Property Let ReplacementString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReplacementString = 5
    mvarReplacementString = vData
End Property


Public Property Get ReplacementString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReplacementString
    ReplacementString = mvarReplacementString
End Property



Public Function ReplaceString(Optional ByVal strCurrentString As String, _
                                                         Optional ByVal strReplacementString As String, _
                                                         Optional ByVal strFindString As String) As String

   Dim intIndex As Integer
   Dim strTemp As String

    If Len(strFindString) > 0 Then mvarFindString = strFindString
    If Len(strCurrentString) > 0 Then mvarCurrentString = strCurrentString
    If Len(strReplacementString) > 0 Then mvarReplacementString = strReplacementString
   
   Call ClearStringPositions  'clear previous strings
   
   GetStringPosition

   strTemp = mvarCurrentString
   For intIndex = 0 To mvarGetStringPositionCount - 1
      strTemp = Replace(strTemp, mvarFindString, mvarReplacementString)
   Next intIndex
   
   mvarResultString = strTemp
   ReplaceString = mvarResultString
   
End Function

Public Sub ClearStringPositions()
    
    Dim intIndex As Integer
    
    For intIndex = 0 To mvarGetStringPositionCount - 1
        mvarGetStringPositions(intIndex).Element = ""
        mvarGetStringPositions(intIndex).ElementType = ""
        mvarGetStringPositions(intIndex).ElementValue = 0
    Next
    mvarGetStringPositionCount = 0
    mvarGetStringPositionIndex = 0
End Sub

Public Function TrimString(Optional ByVal strFindString As String, _
                           Optional ByVal strCurrentString As String) As String
    
    Dim intCounter As Integer
    Dim strLeft As String
    Dim strRight As String
    
    If Len(strFindString) > 0 Then mvarFindString = strFindString
    If Len(strCurrentString) > 0 Then mvarCurrentString = strCurrentString
    
    mvarResultString = mvarCurrentString
    
    For intCounter = 1 To Len(mvarResultString)
        If Mid(mvarResultString, intCounter, Len(strFindString)) = strFindString Then
            strLeft = Left(mvarResultString, CLng(intCounter - 1))
            strRight = Right(mvarResultString, _
                CLng(Len(mvarResultString) - Len(strLeft) - Len(strFindString)))
            mvarResultString = strLeft & strRight
            intCounter = intCounter - 1
        End If
    Next
    TrimString = mvarResultString
End Function

Public Function TrimAll(Optional ByVal strCurrentString As String) As String
    
    If Len(strCurrentString) > 0 Then mvarCurrentString = strCurrentString
    
    mvarResultString = mvarCurrentString
    mvarResultString = TrimString(" ", mvarResultString)
    mvarResultString = TrimString(vbCrLf, mvarResultString)
    mvarResultString = TrimString(vbTab, mvarResultString)
    
    TrimAll = mvarResultString
End Function

Public Property Let FindString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FindString = 5
    mvarFindString = vData
End Property


Public Property Get FindString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FindString
    FindString = mvarFindString
End Property




Public Property Let ResultString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ResultString = 5
    mvarResultString = vData
End Property


Public Property Get ResultString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResultString
    ResultString = mvarResultString
End Property



Public Property Let CurrentString(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentString = 5
    mvarCurrentString = vData
End Property


Public Property Get CurrentString() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentString
    CurrentString = mvarCurrentString
End Property



Public Sub GetStringPosition(Optional ByVal strFindString As String, _
                                                      Optional ByVal strCurrentString As String)
    Dim varOldPosition As Long
    Dim varPosition As Variant
    Dim strTemp As String
    
    If Len(strFindString) > 0 Then mvarFindString = strFindString
    If Len(strCurrentString) > 0 Then mvarCurrentString = strCurrentString
       
    varPosition = 0
    Do
        varOldPosition = varPosition + 1
        varPosition = InStr(varOldPosition, _
            mvarCurrentString, mvarFindString, vbTextCompare)
        If varPosition <> 0 Then
            strTemp = Mid(mvarCurrentString, varOldPosition, _
                    varPosition - varOldPosition)
            mvarGetStringPositions(mvarGetStringPositionIndex).Element = strTemp
            mvarGetStringPositions(mvarGetStringPositionIndex) _
                .ElementValue = varPosition
            mvarGetStringPositionIndex = mvarGetStringPositionIndex + 1
            mvarGetStringPositionCount = mvarGetStringPositionIndex
        End If
    Loop While varPosition <> 0

End Sub

'Public Property Set GetStringPositions(ByVal vData As clsElement)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.GetStringPositions = Form1
 '   Set mvarGetStringPositions = vData
'End Property


Public Property Get GetStringPositions(intStringIndex As Integer) As CElement
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GetStringPositions
    Set GetStringPositions = mvarGetStringPositions(intStringIndex)
End Property



Public Property Let GetStringPositionCount(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GetStringPositionCount = 5
    mvarGetStringPositionCount = vData
End Property


Public Property Get GetStringPositionCount() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GetStringPositionCount
    GetStringPositionCount = mvarGetStringPositionCount
End Property



Public Property Let GetStringPositionIndex(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GetStringPositionIndex = 5
    mvarGetStringPositionIndex = vData
End Property


Public Property Get GetStringPositionIndex() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GetStringPositionIndex
    GetStringPositionIndex = mvarGetStringPositionIndex
End Property



