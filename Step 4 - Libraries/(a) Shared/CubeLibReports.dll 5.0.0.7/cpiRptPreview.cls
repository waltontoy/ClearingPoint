VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cpiRptPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub Show(ByVal Report As Variant, ByVal Title As String, ByVal Recipient As String, ByVal MsgBody As String) 'ByVal Subject As String, ByVal BodyMessage As String, ByVal AttachmentPathName As String
    Dim PreviewForm As frmPreview

    Set PreviewForm = New frmPreview

    With PreviewForm
        .RunReport Report, Title, Recipient, MsgBody
        .arvReportViewer.TOCEnabled = False
        .Caption = Title
        .Show vbModal
    End With

    Set PreviewForm = Nothing
End Sub
