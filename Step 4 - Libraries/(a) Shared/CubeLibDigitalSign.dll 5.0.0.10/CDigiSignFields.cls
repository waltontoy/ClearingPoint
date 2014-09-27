VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDigiSignFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_strAssignedValue() As String
Private m_lngFieldCount As Long

Public Property Get FieldCount() As Long
    
    FieldCount = m_lngFieldCount

End Property

Public Function FieldValue(ByVal lCtr As Long) As String
    
    FieldValue = m_strAssignedValue(lCtr)

End Function

Public Function StringToSign(ByRef DigiSignData As ADODB.Recordset) As String
    
    Dim strTempField As String
    Dim strTempString As String
    
    
    
    '*************************************************************************************************************
    'Packing Algorithm
    '*************************************************************************************************************
    If Not DigiSignData.State = adStateClosed Then
        
        If DigiSignData.RecordCount > 0 Then
            'DigiSignData.Sort = "DetailNumber, DigiSignOrdinal, GroupOrdinal, AutoID "
            
            DigiSignData.MoveFirst
            
            Do Until DigiSignData.EOF
                
                strTempField = DigiSignData.Fields("Value").Value
                                                                   
                m_lngFieldCount = m_lngFieldCount + 1
                ReDim Preserve m_strAssignedValue(m_lngFieldCount)
                m_strAssignedValue(m_lngFieldCount) = strTempField

                strTempString = strTempString & strTempField
                
                Debug.Print DigiSignData.Fields("DetailNumber").Value & " - " & _
                            DigiSignData.Fields("DigiSignOrdinal").Value & " - " & _
                            DigiSignData.Fields("GroupOrdinal").Value & " - " & _
                            DigiSignData.Fields("Value").Value
                
                DigiSignData.MoveNext
            Loop
         End If
    End If
    '*************************************************************************************************************
       
    StringToSign = strTempString

End Function

