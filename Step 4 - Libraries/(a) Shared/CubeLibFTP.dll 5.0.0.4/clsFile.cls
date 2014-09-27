VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarFileName As String 'local copy
Private mvarFileSize As Long 'local copy
Private mvarCreationTime As Date 'local copy
Private mvarDownloadSuccess As Boolean 'local copy

Public Property Let DownloadSuccess(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DownloadSuccess = 5
    mvarDownloadSuccess = vData
End Property


Public Property Get DownloadSuccess() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DownloadSuccess
    DownloadSuccess = mvarDownloadSuccess
End Property



Public Property Let CreationTime(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CreationTime = 5
    mvarCreationTime = vData
End Property


Public Property Get CreationTime() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CreationTime
    CreationTime = mvarCreationTime
End Property



Public Property Let FileSize(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileSize = 5
    mvarFileSize = vData
End Property


Public Property Get FileSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSize
    FileSize = mvarFileSize
End Property



Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property



