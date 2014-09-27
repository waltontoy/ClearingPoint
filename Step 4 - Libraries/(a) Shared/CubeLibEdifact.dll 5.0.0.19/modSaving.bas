Attribute VB_Name = "modSaving"
Option Explicit
    



Public Function GetMessageStatusType(ByVal MessageStatusType As EMessageStatusType)
    Select Case MessageStatusType
        Case EMsgStatusType_Document
            GetMessageStatusType = "Document"
        Case EMsgStatusType_Sent
            GetMessageStatusType = "Sent"
        Case EMsgStatusType_Received
            GetMessageStatusType = "Received"
    End Select
End Function
