VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSQLQuotes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function ProcessQuotes(ByVal QuotedText As String, _
                                Optional ByVal ASCQuoteEnclosingData As Long = 39) As String
    Dim intFieldPosition As Integer
    Dim blnDone As Boolean
    Dim intLength As Integer
    Dim arrLink
    Dim intLinkCtr As Integer
    Dim strChain As String
    Dim intLoopCtr As Integer
    
    ReDim arrLink(0)
    blnDone = False
    intLinkCtr = 0
    strChain = ""
    
    Do While blnDone = False
        intLength = Len(QuotedText)
        
        If InStr(1, QuotedText, Chr(39)) > 0 Then    '-----> Apostrophe
            intFieldPosition = InStr(1, QuotedText, Chr(39))
            
            intLinkCtr = intLinkCtr + 1
            ReDim Preserve arrLink(intLinkCtr)
            
            If Trim$(Chr(ASCQuoteEnclosingData)) = Chr(39) Then
                arrLink(intLinkCtr - 1) = Left(QuotedText, intFieldPosition) & Chr(39)
            Else
                arrLink(intLinkCtr - 1) = Left(QuotedText, intFieldPosition)
            End If
            
            If intFieldPosition <> Len(QuotedText) Then
                QuotedText = Mid(QuotedText, intFieldPosition + 1)
            Else
                blnDone = True
            End If
        ElseIf InStr(1, QuotedText, Chr(34)) > 0 Then    '-----> Quote
            intFieldPosition = InStr(1, QuotedText, Chr(34))
            
            intLinkCtr = intLinkCtr + 1
            ReDim Preserve arrLink(intLinkCtr)
            
            If Trim$(Chr(ASCQuoteEnclosingData)) = Chr(34) Then
                arrLink(intLinkCtr - 1) = Left(QuotedText, intFieldPosition) & Chr(34)
            Else
                arrLink(intLinkCtr - 1) = Left(QuotedText, intFieldPosition)
            End If

            If intFieldPosition <> Len(QuotedText) Then
                QuotedText = Mid(QuotedText, intFieldPosition + 1)
            Else
                blnDone = True
            End If
        Else
            intLinkCtr = intLinkCtr + 1
            ReDim Preserve arrLink(intLinkCtr)
            
            arrLink(intLinkCtr - 1) = QuotedText
            blnDone = True
        End If
    Loop
                        
    For intLoopCtr = 0 To (UBound(arrLink) - 1)
        If Trim(strChain) = "" Then
            strChain = CStr(arrLink(intLoopCtr))
        Else
            strChain = strChain & CStr(arrLink(intLoopCtr))
        End If
        
    Next
    
    'AQ = strChain
    
    ProcessQuotes = Replace(strChain, " ", Chr(32))
End Function
