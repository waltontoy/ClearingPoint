VERSION 5.00
Object = "{F83FB95C-D981-11D2-A80A-00104BF191A4}#1.0#0"; "SKCL.dll"
Begin VB.Form FCubeLibAbout 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   7980
   ClientTop       =   3195
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6060
   Begin VB.CommandButton cmdIEExplore 
      Caption         =   "Codejock"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About..."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin SKCLLibCtl.LFile lfpLicensing 
      Left            =   480
      OleObjectBlob   =   "FCubeLibAbout.frx":0000
      Tag             =   "License"
      Top             =   2040
   End
End
Attribute VB_Name = "FCubeLibAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private m_objConProps As CConnectionProperties
Private m_objTemplateCP As ADODB.Connection

Private G_strLicCompanyName As String
Private G_strLicCompanyAddress  As String

Private Sub cmdAbout_Click()
    Dim objAbout As New CAbout
    
    
    
    objAbout.Initialize m_objTemplateCP, App, lfpLicensing, "1997", "2014", "www.codejock.com"
    
    objAbout.Show Me, vbModal
    
    Set objAbout = Nothing
End Sub

Private Sub cmdIEExplore_Click()
    Dim objIEExplore As New CIEExplore
    
    objIEExplore.OpenURL "www.codejock.com", Me
    
    Set objIEExplore = Nothing
End Sub

Private Sub Form_Load()
    
    
    
    Set m_objConProps = GetDataSourceProperties(App.Path)
    
    If Not m_objConProps Is Nothing Then
        ADOConnectDB m_objTemplateCP, m_objConProps, DBInstanceType_DATABASE_TEMPLATE
    End If
    
    SetupLicense m_objConProps.DataSource
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    lfpLicensing.SemClose
    
    ' Must go before licenseing is closed
    Set m_objConProps = Nothing
End Sub

Private Sub lfpLicensing_Error()
    Dim errBuf As String * 50
    
    Const GENERIC_READ = &H80000000
    Const FILE_SHARE_READ = &H1
    Const OPEN_EXISTING = 3

    Dim hFile As Long
    Dim nSize As Currency
    Dim sSave As String
    
    'open the file
    hFile = CreateFile(m_objConProps.DataSource & "\CPLic.lf", GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    'get the filesize
    GetFileSizeEx hFile, nSize
    
    'close the file
    CloseHandle hFile
    
    ' Something went wrong - display an error message
    pp_errorstr lfpLicensing.LastErrorNumber, errBuf
        
    Select Case lfpLicensing.LastErrorNumber
        Case 15 '>> No license file
            If hFile <= 0 Then
                MsgBox "The license file is missing. Please contact technical support.", vbInformation, G_CONST_APPLICATION_NAME
            ElseIf nSize <= 0 Then
                MsgBox "The license file could be corrupted. Please contact technical support.", vbInformation, G_CONST_APPLICATION_NAME
            Else
                MsgBox "Product has not been activated. Please contact technical support.", vbInformation, G_CONST_APPLICATION_NAME
            End If
                
            'm_blnLicenseOK = False
            
        Case 18 '>> Product has not been activated
            'm_blnLicenseOK = False
        
        Case 9
            ' lfpLicensing.LastErrorNumber      9
            ' lfpLicensing.LastErrorString      'CANNOT CREATE FILE'
            ' Do Nothing
        
        Case 13 ' WRONG PASSWORD
            If hFile <= 0 Then
                MsgBox "The license file is missing. Please contact technical support.", vbInformation, G_CONST_APPLICATION_NAME
            ElseIf nSize <= 0 Then
                MsgBox "The license file could be corrupted. Please contact technical support.", vbInformation, G_CONST_APPLICATION_NAME
            Else
                MsgBox errBuf & " Please contact technical support." & vbCrLf & "Error Number " & lfpLicensing.LastErrorNumber & " (" & lfpLicensing.LastErrorString & ")", vbInformation, G_CONST_APPLICATION_NAME
            End If
            
            'm_blnLicenseOK = False
            
        Case Else
            MsgBox errBuf & " Please contact technical support." & vbCrLf & "Error Number " & lfpLicensing.LastErrorNumber & " (" & lfpLicensing.LastErrorString & ")", vbInformation, G_CONST_APPLICATION_NAME
            
            'm_blnLicenseOK = False
    End Select
End Sub

Private Function SetupLicense(ByVal DatabasePath As String) As Boolean
    Dim blnSyslinkHasOldLic As Boolean
    Dim blnEntrepotHasOldLic As Boolean
    
    'Initialize Protection Plus control
    'm_blnLicenseOK = True
    
    lfpLicensing.UseEZTrigger = True
    
    'rachelle 110906
    lfpLicensing.CPAlgorithm = 128
    lfpLicensing.TCSeed = 400
    lfpLicensing.TCRegKey2Seed = 200
    lfpLicensing.LFPassword = "C" & "ube" & "po" & "int"
    lfpLicensing.SemPath = DatabasePath & "\"
    lfpLicensing.CPAlgorithmDrive = App.Path
    lfpLicensing.LFName = DatabasePath & "\CPLic.lf"
    
    'If (lfpLicensing.IsDemo = False And m_blnLicenseOK = False) Then
    '    Exit Function
    'End If
    
'''''    ' Disable features in license component if demo has expired
'''''    If (lfpLicensing.IsDemo = True And m_blnDemoExpired = True) Or _
'''''       (lfpLicensing.IsDemo = False And m_blnDemoExpired = True) Or _
'''''       (lfpLicensing.IsDemo = False And m_blnLicenseOK = False) Then
'''''
'''''        ' <<< dandan 080807
'''''        ' Commented to show the previously activated features even if the license is already expired
''''''        ' Update license file, deactivate all features
''''''        lfpLicensing.LFLock
''''''            lfpLicensing.UserOption(PLDA_Import_FeatureID) = False
''''''            lfpLicensing.UserOption(PLDA_Export_FeatureID) = False
''''''            lfpLicensing.UserOption(PLDA_Combined_FeatureID) = False
''''''            lfpLicensing.UserOption(EDIFACT_NCTS_FeatureID) = False
''''''            lfpLicensing.UserOption(SADBEL_Import_FeatureID) = False
''''''            lfpLicensing.UserOption(SADBEL_Export_Transit_FeatureID) = False
''''''            lfpLicensing.UserOption(SADBEL_NCTS_Transit_Combined_FeatureID) = False
''''''
''''''            lfpLicensing.UserOption(SysLink_Automatic_In_FeatureID) = blnSyslinkHasOldLic
''''''            lfpLicensing.UserOption(SysLink_Semi_Automatic_FeatureID) = blnSyslinkHasOldLic
''''''            lfpLicensing.UserOption(SysLink_Automatic_Out_FeatureID) = blnSyslinkHasOldLic
''''''
''''''            lfpLicensing.UserOption(Email_Report_Automatic_FeatureID) = False
''''''            lfpLicensing.UserOption(Email_Report_Semi_Automatic_FeatureID) = False
''''''
''''''            lfpLicensing.UserOption(Entrepot_FeatureID) = blnEntrepotHasOldLic
''''''
''''''            lfpLicensing.UserOption(Repertory_FeatureID) = False
''''''            lfpLicensing.UserOption(Remote_Printing_FeatureID) = False
''''''        lfpLicensing.LFUnlock
'''''
'''''        ' <<< dandan 080807
'''''        ' Added to Deactivate features to false if the cplic  file is already expired
'''''        With G_typFeatures
'''''            .PLDA.Import = False
'''''            .PLDA.Export_Transit = False
'''''            .PLDA.Combined = False
'''''            .NCTS.Departure_Arrival = False
'''''            .SADBEL.Import = False
'''''            .SADBEL.Export_Transit = False
'''''            .SADBEL.NCTS_Transit_Combined = False
'''''            .SysLink.Automatic_In = False
'''''            .SysLink.Semi_Automatic = False
'''''            .SysLink.Automatic_Out = False
'''''            .EmailReport.Automatic = False
'''''            .EmailReport.Semi_Automatic = False
'''''            .Entrepot = False
'''''            .Repertory = False
'''''            .RemotePrinting = False
'''''            .Archiving = False
'''''            .NCTS.Departure_FollowUp_Request = False
'''''            .PDFOut = False
'''''            .EnableBackupDB = False
'''''        End With
'''''    Else
'''''        ' Load licensed features from the license file
'''''        With G_typFeatures
'''''            .PLDA.Import = lfpLicensing.UserOption(PLDA_Import_FeatureID)
'''''            .PLDA.Export_Transit = lfpLicensing.UserOption(PLDA_Export_FeatureID)
'''''            .PLDA.Combined = lfpLicensing.UserOption(PLDA_Combined_FeatureID)
'''''            .NCTS.Departure_Arrival = lfpLicensing.UserOption(EDIFACT_NCTS_FeatureID)
'''''            .SADBEL.Import = lfpLicensing.UserOption(SADBEL_Import_FeatureID)
'''''            .SADBEL.Export_Transit = lfpLicensing.UserOption(SADBEL_Export_Transit_FeatureID)
'''''            .SADBEL.NCTS_Transit_Combined = lfpLicensing.UserOption(SADBEL_NCTS_Transit_Combined_FeatureID)
'''''            .SysLink.Automatic_In = lfpLicensing.UserOption(SysLink_Automatic_In_FeatureID)
'''''            .SysLink.Semi_Automatic = lfpLicensing.UserOption(SysLink_Semi_Automatic_FeatureID)
'''''            .SysLink.Automatic_Out = lfpLicensing.UserOption(SysLink_Automatic_Out_FeatureID)
'''''            .EmailReport.Automatic = lfpLicensing.UserOption(Email_Report_Automatic_FeatureID)
'''''            .EmailReport.Semi_Automatic = lfpLicensing.UserOption(Email_Report_Semi_Automatic_FeatureID)
'''''            .Entrepot = lfpLicensing.UserOption(Entrepot_FeatureID)
'''''            .Repertory = lfpLicensing.UserOption(Repertory_FeatureID)
'''''            .RemotePrinting = lfpLicensing.UserOption(Remote_Printing_FeatureID)
'''''            .NCTS.Departure_FollowUp_Request = lfpLicensing.UserOption(EDIFACT_NCTS_FollowUpRequest)
'''''            .PDFOut = lfpLicensing.UserOption(PDF_OUT) 'vag walls
'''''            .EnableBackupDB = lfpLicensing.UserOption(EnableBackupDB)
'''''        End With
'''''
'''''        ' Update Database so we can retrieve last settings in case the
'''''        ' license file becomes corrupted wherein the licensing component
'''''        ' cannot read the file anymore. an example would be that CPLic.lf's size
'''''        ' becomes zero
'''''        SaveLastSuccessfulFeatureActivationStates G_typFeatures
'''''    End If

    G_strLicCompanyName = lfpLicensing.RegCompany
    'p4tric 062409
    G_strLicCompanyAddress = Replace(lfpLicensing.RegAddress1 & vbCrLf & _
                            lfpLicensing.RegAddress2 & ", " & lfpLicensing.RegAddress3, "|||||", " ")


    'SetupLicense = True
End Function

