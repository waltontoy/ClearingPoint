VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarFolderName As String 'local copy

Public Property Let FolderName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FolderName = 5
    mvarFolderName = vData
End Property


Public Property Get FolderName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FolderName
    FolderName = mvarFolderName
End Property



