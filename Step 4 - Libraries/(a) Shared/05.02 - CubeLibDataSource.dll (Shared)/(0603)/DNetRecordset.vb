<ComClass(DNetRecordset.ClassId, DNetRecordset.InterfaceId, DNetRecordset.EventsId)> _
Public Class DNetRecordset

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "9e69cfd6-de85-4b63-90da-7f9b424f0c7f"
    Public Const InterfaceId As String = "ef0bdae6-439c-4d20-8f48-654b9b28cf71"
    Public Const EventsId As String = "5da34e22-80c7-466f-b1d8-7e9a30b73458"
#End Region

    Private m_rstADO As ADODB.Recordset
    Private m_bookMark As Double

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    'Parameterized constructor for unit testing
    Public Sub New(ByRef rstADO As ADODB.Recordset, ByVal BookMark As Double)
        MyBase.New()
        m_rstADO = rstADO
        m_bookMark = BookMark
    End Sub

    'Offline recordset to update
    Public Function Recordset() As ADODB.Recordset
        Return m_rstADO
    End Function

    'Connection String used to populate the offline recordset
    Public Function BookMark() As Double
        Return m_bookMark
    End Function

    'Parameterized constructor does not work in VB6, this is a workaround
    Public Sub InitializeClass(ByRef rstADO As ADODB.Recordset, ByVal bookMark As Double)
        m_rstADO = rstADO
        m_bookMark = bookMark
    End Sub

End Class


