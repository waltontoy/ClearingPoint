VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Enum LicenseTypeConstants
    Floating = 0
    Fixed = 1
End Enum

Private Const LICENSE_TYPE_BIT = 25

Private m_varImage As Variant

Private m_strApplicationName As String

Private m_strVersionMajor As String
Private m_strVersionMinor As String
Private m_strVersionRevision As String

Private m_strCopyrightCompany As String
Private m_strCopyrightYearStart As String
Private m_strCopyrightYearEnd As String

Private m_blnIsDemoVersion As Boolean
Private m_strComputerID As String

Private m_blnIsExpired As Boolean
' Always in M/d/yyyy' format so that it will
' be displayed correctly all the time
Private m_dteExpiryDate As String
Private m_dteCorrectExpiryDate As DTPicker
Private m_lngDaysLeft  As Long

Private m_strLicensee As String
Private m_strSerialNumber As String
Private m_enuLicenseType As LicenseTypeConstants
Private m_lngAllowedUsers As Long
Private m_blnLicenseExpires As Boolean
Private m_colActiveFeatures As CFeatures

Private m_strTechSupportURL As String

Private m_strLicPath As String
Private m_strLicFName As String
Private m_conADO As New ADODB.Connection
Private m_objOwnerForm As Object

Private Const G_CONST_USER_NAME1 = "CANDSMASTER"
Private Const G_CONST_IP_ADDRESS1 = "212.190.192.9"
Private Const G_CONST_PASSWORD1 = "FT121DXL03"

Private Const G_CONST_USER_NAME2 = "maximumasp\v001u50tzd"
Private Const G_CONST_IP_ADDRESS2 = "www.cubepoint.be"
Private Const G_CONST_PASSWORD2 = "UZjZeDGx42Dx"

Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const CSIDL_DESKTOP = &H0

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Property Get ApplicationName() As String
    ApplicationName = m_strApplicationName
End Property

Public Property Get CopyrightCompany() As String
    CopyrightCompany = m_strCopyrightCompany
End Property

Public Property Get CopyrightStart() As String
    CopyrightStart = m_strCopyrightYearStart
End Property

Public Property Get CopyrightEnd() As String
    CopyrightEnd = m_strCopyrightYearEnd
End Property

Public Property Set Image(ByVal NewValue As Variant)
    If (IsObject(NewValue) = False) Then
        
        Err.Raise 777001, , "Only picture objects are accepted."
    
    ElseIf (IsObject(NewValue) = True) Then
        Set m_varImage = NewValue
    
    End If
End Property

Public Property Get Image() As Variant
    If (IsObject(m_varImage) = True) Then
        Set Image = m_varImage
    
    ElseIf (IsObject(m_varImage) = False) Then
        Image = m_varImage
    End If
End Property

Public Property Get Licensee() As String
    Licensee = m_strLicensee
End Property

Public Property Get VersionMajor() As String
    VersionMajor = m_strVersionMajor
End Property

Public Property Get VersionMinor() As String
    VersionMinor = m_strVersionMinor
End Property

Public Property Get VersionRevision() As String
    VersionRevision = m_strVersionRevision
End Property

Public Property Get TechSupportURL() As String
    TechSupportURL = m_strTechSupportURL
End Property

Public Sub Show(ByRef OwnerForm As Object, _
                ByRef Style As Variant, _
       Optional ByVal CanActivateUpdate As Boolean = True)
    
    Dim blnFlag As Boolean
    Dim frmAboutForm As Form
    
    
    If (CanActivateUpdate = True) Then
        Set frmAboutForm = FAbout
    Else
        Set frmAboutForm = frmAbout
    End If
    
    
    If ((TypeOf OwnerForm Is Form) = False) Then
        GoTo ERROR_TYPE_MISMATCH
    End If
    
    Set m_objOwnerForm = OwnerForm
    
    frmAboutForm.SetAbout OwnerForm, Me
    
    If (Len(Trim(Style)) = 0) Then
        
        frmAboutForm.Show
        
    Else
        
        blnFlag = ((Style <> vbModal) And (Style <> vbModeless))
        
        If (blnFlag = True) Then
            frmAboutForm.Show vbModeless
        ElseIf (blnFlag = False) Then
            frmAboutForm.Show Style
        End If
        
    End If
    
    
    Set frmAboutForm = Nothing
    
    Exit Sub
    
    
ERROR_TYPE_MISMATCH:
    
    Err.Raise 1001, , "Type Mismatch.~About.Show~"
    
End Sub

Public Sub Initialize(ByRef ADOConnection As ADODB.Connection, _
                      ByRef Application As Object, _
                      ByRef LFLicense As Object, _
                      ByVal CopyrightYearStart As String, _
                      ByVal CopyrightYearEnd As String, _
                      ByVal TechSupportURL As String, _
             Optional ByRef ProductIcon As Object)
    
    Dim rstLicense As ADODB.Recordset
    Dim arrPath() As String
    
    ' Ensure that this is the format so that when displayed,
    ' the values are correct for all system date formats
    LFLicense.dateformat = "M/d/yyyy"
    Set m_dteCorrectExpiryDate = FDTPicker.DTPickerMdyyyy
    
    Set m_conADO = ADOConnection
    
    Set m_varImage = ProductIcon
    
    m_strApplicationName = Application.ProductName
    m_strVersionMajor = Application.Major
    m_strVersionMinor = Application.Minor
    m_strVersionRevision = Application.Revision
    
    m_strCopyrightCompany = Application.CompanyName
    m_strCopyrightYearStart = CopyrightYearStart
    m_strCopyrightYearEnd = CopyrightYearEnd
    
    m_strComputerID = LFLicense.CPCompNo
    m_blnIsDemoVersion = LFLicense.IsDemo
    
    arrPath = Split(LFLicense.LFName, "\")
    m_strLicPath = Left(LFLicense.LFName, Len(LFLicense.LFName) - (Len(arrPath(UBound(arrPath))) + 1))  'remove "\<licfilename>.lf"
    m_strLicFName = arrPath(UBound(arrPath))
    
    If (m_blnIsDemoVersion = True) Then
        
        ' If demo version, retrieve check expiry date
        If (LFLicense.ExpireDateSoft <> "0/0/0") And (IsDate(LFLicense.ExpireDateSoft) = True) Then
            ' This is walways in 'M/d/yyyy' format
            m_dteExpiryDate = LFLicense.ExpireDateSoft
            SetCorrectExpiryDate m_dteExpiryDate
            m_blnIsExpired = (m_dteCorrectExpiryDate.Value <= Date)
            If (m_blnIsExpired = False) Then
                m_lngDaysLeft = DateDiff("d", Format(Date, LFLicense.dateformat), m_dteCorrectExpiryDate.Value)
            End If
        End If
        
    Else
        
        ' Retrieve serial number from database
        ADORecordsetOpen "SELECT [Lic_SerialNumber] FROM [Licensee] ", ADOConnection, rstLicense, adOpenKeyset, adLockOptimistic
        If (rstLicense.RecordCount > 0) Then
            m_strSerialNumber = Trim(FNullField(rstLicense![Lic_SerialNumber]))
        End If
        ADORecordsetClose rstLicense
        
        m_strLicensee = LFLicense.RegCompany
        
        ' Check if license expires and get expiry date
        If (LFLicense.ExpireDateSoft <> "0/0/0") And (IsDate(LFLicense.ExpireDateSoft) = True) Then
            m_blnLicenseExpires = True
            ' This is walways in 'M/d/yyyy' format
            m_dteExpiryDate = LFLicense.ExpireDateSoft
            SetCorrectExpiryDate m_dteExpiryDate
            m_blnIsExpired = (m_dteCorrectExpiryDate.Value < Date)
            If (m_blnIsExpired = False) Then
                m_lngDaysLeft = DateDiff("d", Format(Date, LFLicense.dateformat), m_dteCorrectExpiryDate.Value)
            End If
        Else
            m_blnLicenseExpires = False
            m_blnIsExpired = False
        End If
        
        ' Retrieve license type and number of allowed network users
        If (LFLicense.UserOption(LICENSE_TYPE_BIT) = True) Then
            m_enuLicenseType = Fixed
        Else
            m_enuLicenseType = Floating
            m_lngAllowedUsers = LFLicense.UsersAllowed
        End If
        
        ' Retrieve active features
        RetrieveActiveFeatures ADOConnection, LFLicense
    End If
    
    
    m_strTechSupportURL = TechSupportURL
    
End Sub

Private Sub Class_Initialize()
    Load FDTPicker
    
End Sub

Private Sub Class_Terminate()
    Set m_varImage = Nothing
    Set m_colActiveFeatures = Nothing
    
    On Error Resume Next
    Unload FDTPicker
    Set FDTPicker = Nothing
    On Error GoTo 0
End Sub

Public Property Get IsDemoVersion() As Boolean
    IsDemoVersion = m_blnIsDemoVersion
End Property

Public Property Get ComputerID() As String
    ComputerID = m_strComputerID
End Property

Public Property Get IsExpired() As Boolean
    IsExpired = m_blnIsExpired
End Property

Public Property Get ExpiryDate() As Date
    ExpiryDate = m_dteCorrectExpiryDate.Value
End Property

Public Property Get DaysLeft() As Long
    DaysLeft = m_lngDaysLeft
End Property

Public Property Get SerialNumber() As String
    SerialNumber = m_strSerialNumber
End Property

Public Property Get LicenseType() As LicenseTypeConstants
    LicenseType = m_enuLicenseType
End Property

Public Property Get AllowedUsers() As Long
    AllowedUsers = m_lngAllowedUsers
End Property

Public Property Get LicenseExpires() As Boolean
    LicenseExpires = m_blnLicenseExpires
End Property

Public Property Get ActiveFeatures() As CFeatures
    If (m_colActiveFeatures Is Nothing = True) Then
        Set m_colActiveFeatures = New CFeatures
    End If
    
    Set ActiveFeatures = m_colActiveFeatures
End Property

Private Sub RetrieveActiveFeatures(ByRef ADOConnection As ADODB.Connection, ByRef LFLicense As Object)
    Dim rstFeatures As ADODB.Recordset
    Dim lngCtr As Long
    
    Dim lngFeatureCode As Long
    Dim strFeatureName As String
    
    
    ADORecordsetOpen "SELECT * FROM Features", ADOConnection, rstFeatures, adOpenKeyset, adLockOptimistic
    If (rstFeatures.RecordCount > 0) Then
        rstFeatures.MoveFirst
        
        Set m_colActiveFeatures = New CFeatures
        
        For lngCtr = 1 To rstFeatures.RecordCount
            lngFeatureCode = FNullField(rstFeatures![Feature_Code])
            strFeatureName = Trim$(FNullField(rstFeatures![Feature_Name]))
            
            
            If (lngFeatureCode > 0 And Len(strFeatureName) > 0) Then
                If (LFLicense.UserOption(lngFeatureCode) = True) Then
                    m_colActiveFeatures.Add rstFeatures![Feature_ID], lngFeatureCode, strFeatureName
                End If
            End If
            
            rstFeatures.MoveNext
        Next lngCtr
    End If
    ADORecordsetClose rstFeatures
    
End Sub

Public Function GetLicenseFile(ByVal SerialNumber As String, Optional ADOConnection As ADODB.Connection, Optional OwnerForm As Object, Optional LicFileName As String) As Boolean
    Dim oFtp As clsFTP
    Dim oFile As clsFile
    
    Dim blnOpenSuccess As Boolean
    Dim blnFound As Boolean
    Dim blnDownload As Boolean
    
    Dim strCommand As String
    Dim lngCtr As Long
    
    Dim arrIPAddress() As String
    Dim arrUName() As String
    Dim arrPWord() As String
    
    ReDim arrIPAddress(2)
    ReDim arrUName(2)
    ReDim arrPWord(2)
    
    Dim Path As String
    Dim strDesktop As String
    Dim lngIDLid As Long
    Dim IDL As ITEMIDLIST
    
    Dim arrPath() As String
    
    arrIPAddress(0) = G_CONST_IP_ADDRESS1
    arrUName(0) = G_CONST_USER_NAME1
    arrPWord(0) = G_CONST_PASSWORD1
    
    arrIPAddress(1) = G_CONST_IP_ADDRESS2
    arrUName(1) = G_CONST_USER_NAME2
    arrPWord(1) = G_CONST_PASSWORD2
    
    GetLicenseFile = False
    
    If m_strLicPath = "" And m_strLicFName = "" And LicFileName <> "" Then
        arrPath = Split(LicFileName, "\")
        m_strLicPath = Left(LicFileName, Len(LicFileName) - (Len(arrPath(UBound(arrPath))) + 1))
        m_strLicFName = arrPath(UBound(arrPath))
    End If
                        
    If Not (ADOConnection Is Nothing) Then
        Set m_conADO = ADOConnection
    End If
    If Not (OwnerForm Is Nothing) Then
        Set m_objOwnerForm = OwnerForm
    End If
    
    
    For lngCtr = 0 To 1
        Set oFtp = New clsFTP
        
        DoEvents
    
        oFtp.Connections.Add arrIPAddress(lngCtr), arrUName(lngCtr), arrPWord(lngCtr), "General"
        oFtp.Connections("General").Connect
        
        If oFtp.Connections("General").ConnectSuccess Then
            oFtp.Connections("General").BackToRootFolder
            
            blnOpenSuccess = oFtp.Connections("General").OpenFolder("Licenses/CP")
            
            If blnOpenSuccess Then
                blnFound = False
                
                For Each oFile In oFtp.Connections("General").Files
                    If oFile.FileName = SerialNumber & ".lf" Then
                        blnFound = True
                        Exit For
                    End If
                Next
                
                If blnFound Then
                    DoEvents
                    blnDownload = oFtp.Connections("General").GetFile(SerialNumber & ".lf", m_strLicPath & "\" & SerialNumber & ".lf")
                
                    If blnDownload Then
                        GetLicenseFile = True
                        
                        ReplaceLicFile SerialNumber
                        
                            strCommand = ""
                            strCommand = strCommand & "UPDATE "
                            strCommand = strCommand & "Licensee "
                            strCommand = strCommand & "SET "
                            strCommand = strCommand & "Lic_SerialNumber = '" & SerialNumber & "'"
                        ExecuteNonQuery m_conADO, strCommand
                        'm_conADO.Execute strCommand
                        
                        m_strSerialNumber = SerialNumber
                        
                        MsgBox "License has been updated. Please restart your application for changes to apply.", vbInformation, "Update License"
                        
                    Else
                        GetLicenseFile = False
                        
                        'MsgBox "Updating of License File failed.", vbInformation, "Update License"
                    End If
                Else
                    GetLicenseFile = False
                    
                    'MsgBox "Invalid serial number.", vbInformation, "Update License"
                End If
            End If
            
        End If
    
        Do While oFtp.Connections.Count > 0
            ' Remove the first in the collection always
            oFtp.Connections.Remove 1
        Loop
            
        Set oFtp = Nothing
        
        If GetLicenseFile = True Then Exit Function
    Next
    
    lngIDLid = SHGetSpecialFolderLocation(m_objOwnerForm.hwnd, CSIDL_DESKTOP, IDL)
    Path$ = Space$(512)
    lngIDLid = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
    strDesktop = Left$(Path, InStr(Path, Chr$(0)) - 1)
    
    If Dir(strDesktop & "\" & SerialNumber & ".lf") <> "" And _
       Dir(strDesktop & "\" & SerialNumber & ".lf") = SerialNumber & ".lf" Then     'Case sensitive
        ReplaceLicFile SerialNumber, True, strDesktop
        
            strCommand = ""
            strCommand = strCommand & "UPDATE "
            strCommand = strCommand & "Licensee "
            strCommand = strCommand & "SET "
            strCommand = strCommand & "Lic_SerialNumber = '" & SerialNumber & "'"
        ExecuteNonQuery m_conADO, strCommand
        'm_conADO.Execute strCommand
        
        m_strSerialNumber = SerialNumber
        
        GetLicenseFile = True
        
        MsgBox "License has been updated. Please restart your application for changes to apply.", vbInformation, "Update License"
    Else
        GetLicenseFile = False
        
        MsgBox "Serial Number is not valid.", vbInformation, "Update License"
    End If
    
End Function

Private Sub ReplaceLicFile(ByVal NewSerialNumber As String, _
                           Optional ByVal FromDesktop As Boolean = False, _
                           Optional ByVal DesktopPath As String)
    Dim strDate As String
    
    strDate = Format(Date, "mmddyyyy")
    

    'Back-up previous license
    CopyFile m_strLicPath & "\" & m_strLicFName, m_strLicPath & "\" & Left(m_strLicFName, Len(m_strLicFName) - 3) & m_strSerialNumber & strDate & ".lf", False
    
    'Make file unhidden
    SetFileAttributes m_strLicPath & "\" & m_strLicFName, FILE_ATTRIBUTE_NORMAL
    
    'Replace license file
    If FromDesktop Then
        CopyFile DesktopPath & "\" & NewSerialNumber & ".lf", m_strLicPath & "\" & m_strLicFName, False
    Else
        CopyFile m_strLicPath & "\" & NewSerialNumber & ".lf", m_strLicPath & "\" & m_strLicFName, False
    End If
    
    'Make files hidden
    SetFileAttributes m_strLicPath & "\" & m_strLicFName, FILE_ATTRIBUTE_HIDDEN
    SetFileAttributes m_strLicPath & "\" & m_strLicFName & m_strSerialNumber & strDate & ".lf", FILE_ATTRIBUTE_HIDDEN
    
    On Error Resume Next
    If Not FromDesktop Then
        Kill m_strLicPath & "\" & NewSerialNumber & ".lf"
    Else
        Kill DesktopPath & "\" & NewSerialNumber & ".lf"
    End If
    On Error GoTo 0
    
    ' 2dit- delete backup
    On Error Resume Next
    SetFileAttributes m_strLicPath & "\" & Left(m_strLicFName, Len(m_strLicFName) - 3) & m_strSerialNumber & strDate & ".lf", FILE_ATTRIBUTE_NORMAL
    Kill m_strLicPath & "\" & Left(m_strLicFName, Len(m_strLicFName) - 3) & m_strSerialNumber & strDate & ".lf"
    On Error GoTo 0
    
End Sub

Private Sub SetCorrectExpiryDate(ByVal ExpiryDateSoft As String)
    Dim arrDateSegments
    
    arrDateSegments = Split(ExpiryDateSoft, "/")

    m_dteCorrectExpiryDate.Year = arrDateSegments(2)
    m_dteCorrectExpiryDate.Month = arrDateSegments(0)
    m_dteCorrectExpiryDate.Day = arrDateSegments(1)
End Sub
