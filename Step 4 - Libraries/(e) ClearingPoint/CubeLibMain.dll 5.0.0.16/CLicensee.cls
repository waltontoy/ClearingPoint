VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLicensee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarDefaultLicensee As Boolean

Public Event CountryPick(ByRef Cancel As Boolean, ByRef code As String, ByRef CountryName As String)
Public Event CurrencyPick(ByRef Cancel As Boolean, ByRef code As String, ByRef CurrencyName As String)
Public Event LanguagePick(ByRef Cancel As Boolean, ByRef code As String, ByRef LanguageName As String)

Public Event BeforeUpdate(ByRef Cancel As Boolean, ByRef SelectedCountry As String, ByRef SelectedCurrency As String, ByRef SelectedLanguage As String)

Public Event UnloadForm()

Private m_objConnection As ADODB.Connection

Public Sub ShowForm(ByRef OwnerForm As Object, ByRef UserConnection As ADODB.Connection)
    
    Set m_objConnection = UserConnection
    
    If (mvarDefaultLicensee = True) Then
        
        frmLicensee.ShowForm OwnerForm, Me, UserConnection
    End If
    
End Sub

Public Function CheckLicense(ByRef ADOConnection As ADODB.Connection) As Boolean
    Dim rstFeatures As ADODB.Recordset
    Dim rstLicensee As ADODB.Recordset
    
    Dim lngCtr As Long
    
    Dim blnActivate As Boolean
    
    Dim strCommandText As String
    
    
    On Error GoTo Error_Handler
    
    
    '--->Update Features activated
    strCommandText = vbNullString
    strCommandText = strCommandText & " SELECT "
    strCommandText = strCommandText & " * "
    strCommandText = strCommandText & " FROM "
    strCommandText = strCommandText & " Features "
    
    Call ADORecordsetOpen(strCommandText, ADOConnection, rstFeatures, adOpenKeyset, adLockOptimistic)
    
    If (rstFeatures.RecordCount > 0) Then
        
        rstFeatures.MoveFirst
        
        For lngCtr = 1 To rstFeatures.RecordCount
            If (rstFeatures![Feature_Code] <= 28) Or (rstFeatures![Feature_Code] >= 33 And rstFeatures![Feature_Code] <= 64) Then
                rstFeatures![Feature_Activated] = g_typInterface.ILicense.UserOption(rstFeatures![Feature_Code])
            End If
            
            rstFeatures.MoveNext
        Next lngCtr
        
    End If
    
    
    Call ADORecordsetClose(rstFeatures)
    
    
    '--->Update licensee details (name and serial number)
    If (Len(Trim$(g_typInterface.ILicense.RegCompany)) > 0) Then
        strCommandText = ""
        strCommandText = strCommandText & "UPDATE "
        strCommandText = strCommandText & "Licensee "
        strCommandText = strCommandText & "SET "
        strCommandText = strCommandText & "Lic_Name = '" & g_typInterface.ILicense.RegCompany & "'"
        
        'ADOConnection.Execute strCommandText
        
        ExecuteNonQuery ADOConnection, strCommandText
    End If
    
    CheckLicense = True
    
    Exit Function
    
Error_Handler:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function RaiseLicenseError(ByRef OwnerForm As Object, ByVal ErrorNum As Long, ByVal ErrorString As String) As Boolean
    
    Select Case ErrorNum
        Case 15 '>> No license file
            MsgBox "Product has not been activated. Please contact technical support.", vbInformation + vbYesNo, OwnerForm.Caption
            RaiseLicenseError = True
            
        Case Else
            MsgBox "An error has occurred. Please contact technical support." & vbCrLf & "Error Number " & ErrorNum & " (" & ErrorString & ")", vbInformation, OwnerForm.Caption
            RaiseLicenseError = False
    End Select

End Function

Public Sub License_StatusChanged(ByRef OwnerForm As Object, _
                                 ByRef Application As Object, _
                                 ByRef DemoHasExpired As Boolean, _
                                 ByRef LicenseOK As Boolean, _
                                 ByRef LicenseType As String, _
                                 ByVal LicenseTypeBitCode As Long, _
                                 ByVal EncryptPassword As String, _
                                 ByVal Startup As Boolean, _
                                 ByVal DBPathFileName As String, _
                                 ByVal DBPassword As String)
    
    Dim blnLocalDrive As Boolean
    Dim intReply As Integer
    Dim dteServerDate As Date 'used date to syncronize the date that will be written in lic file.-joy
    
    
    If (Startup = False) Then
        Exit Sub
    End If
    
    
    ' retrieve Date of server PC. - joy 5/26/2006
    Call GetFileLastAccessedDate(DBPathFileName, dteServerDate)
    
    
    If (g_typInterface.ILicense.ExpireMode = "D") Then
        
        '>> Demo version only
        g_typInterface.ILicense.UseLastUsedTime = False
        
        LicenseOK = False
        
        If (g_typInterface.ILicense.ExpireDateSoft = "0/0/0") Then
            ' First run, so set expiration to 7 days from the date of first run (today)
            g_typInterface.ILicense.LFLock
                g_typInterface.ILicense.ExpireDateSoft = Format(DateAdd("d", 7, dteServerDate), g_typInterface.ILicense.DateFormat)
                g_typInterface.ILicense.UserDate(5) = Format(CDate(DateAdd("d", 7, dteServerDate)) - 7, g_typInterface.ILicense.DateFormat)
            g_typInterface.ILicense.LFUnlock
        Else
            ' Not first run. Check if demo has expired
            If (ISLicenseExpiredOrClockTurnedBack(dteServerDate) = True) Then
                MsgBox "The demo version of ClearingPoint has expired. All features will be disabled. Please contact technical support to purchase a license.", vbInformation, Application.Title
                DemoHasExpired = True
            End If
        End If
        
        
        If (DemoHasExpired = True) Then
            
            '>> Demo has expired
            
            ' Display info that the demo version has expired into the status bar, 3rd panel
            If (g_typInterface.IStatusbar.Panels.Count = 2) Then
                g_typInterface.IStatusbar.Panels.Add 3, "DemoDaysLeft", "The demo version has expired."
            End If
            
        ElseIf g_typInterface.ILicense.DaysLeft >= 0 Then
            
            '>> Demo is not yet expired
            
            ' Display info on the number of days for the demo version into the status bar, 3rd panel
            If (g_typInterface.IStatusbar.Panels.Count = 2) Then
                g_typInterface.IStatusbar.Panels.Add 3, "DemoDaysLeft", "Demo version: " & g_typInterface.ILicense.DaysLeft & IIf(g_typInterface.ILicense.DaysLeft > 1, " days ", " day ") & "left. "
            End If
            
            ' Update the date product is last used
            g_typInterface.ILicense.LFLock
                g_typInterface.ILicense.LastUsedDate = Format(dteServerDate, g_typInterface.ILicense.DateFormat)
            g_typInterface.ILicense.LFUnlock
            
        End If
        
    ElseIf (g_typInterface.ILicense.ExpireMode = "N") Then
        
        '>> Licensed version
        If (g_typInterface.ILicense.ExpireDateSoft = "0/0/0") Then
            ' The license has no expiry. No need to check the number of days left
            DemoHasExpired = False
        Else
            ' The license has expiry. Check the number of days left
            If (ISLicenseExpiredOrClockTurnedBack(dteServerDate) = True) Then
                ' License has expired or the clock has been turned back
                MsgBox "This version of ClearingPoint has expired. All features will be disabled. Please contact technical support to purchase a license.", vbInformation, Application.Title
                DemoHasExpired = True
            End If
        End If
        
        
        If (DemoHasExpired = True) Then
            
            ' Display info that license has expired into the status bar, 3rd panel
            If (g_typInterface.IStatusbar.Panels.Count = 2) Then
                g_typInterface.IStatusbar.Panels.Add 3, "License", "License expired."
            End If
            
        Else
            
            ' Assume license is OK
            LicenseOK = True
            
            
            ' Display licensee into the status bar, 3rd panel
            If (g_typInterface.IStatusbar.Panels.Count = 2) Then
                g_typInterface.IStatusbar.Panels.Add 3, "License", "Licensed to: " & g_typInterface.ILicense.RegCompany
            End If
            
            
            ' Determine the license type
            If (g_typInterface.ILicense.UserOption(LicenseTypeBitCode) = True) Then
                LicenseType = "FIXED"
            Else
                LicenseType = "FLOATING"
            End If
            
            
            ' Set up license properties depending on the licensing scheme
            If (UCase$(LicenseType) = "FIXED") Then
                
                ' Set copy protection ID algorithm to WINPRODID (128) and HDSERIAL (2); 128+2=130
                g_typInterface.ILicense.CPAlgorithm = 130
                
                ' See if authorized to run on this computer
                If (g_typInterface.ILicense.CPCheck(0) <> 1) Then
                    LicenseOK = False
                    
                    intReply = MsgBox("This machine needs to be activated for the application to run on it. Would you like to activate this computer now?", vbInformation + vbYesNo, Application.ProductName)
                    If (intReply = vbYes) Then
                        FActivate.Activate OwnerForm, Format(g_typInterface.ILicense.ExpireDateSoft, g_typInterface.ILicense.DateFormat), DBPathFileName, g_typInterface.ILicense.LFName, True, g_typInterface.ILicense.CPCompNo
                    End If
                End If
                
            ElseIf (UCase$(LicenseType) = "FLOATING") Then
                
                ' Assume application is running on a local drive (peer server)
                blnLocalDrive = True
                
                ' Check if authorized to run on this computer or network
                If (g_typInterface.ILicense.Redir(Application.Path) = True) Then
                    blnLocalDrive = False
                End If
                
                If (blnLocalDrive = True) Then
                    ' On peer servers, set copy protection ID algorithm to WINPRODID (128) and HDSERIAL (2); 128+2=130
                    g_typInterface.ILicense.CPAlgorithm = 130
                Else
                    ' On peer clients, set copy protection ID algorithm to COMPNO_NETNAME (16)
                    g_typInterface.ILicense.CPAlgorithm = 16
                End If
                
                ' Check if a semaphore file can be opened for this workstation
                If (g_typInterface.ILicense.SemOpen <> 1) Then
                    MsgBox "User limit exceeded. You are allowed to run with " & g_typInterface.ILicense.FloatingUsersAllowed & IIf(g_typInterface.ILicense.FloatingUsersAllowed > 1, " concurrent users ", " user ") & "only. Please purchase more licenses."
                    
                    LicenseOK = False
                End If
                
            End If
        End If
    End If
    
    If (DemoHasExpired = False) Then
        frmLicenseeReminder.LicenseReminder OwnerForm, m_objConnection, DBPathFileName, g_typInterface.ILicense.LFName, DBPassword
    End If
    
End Sub

Public Property Get DefaultLicensee() As Boolean
    DefaultLicensee = mvarDefaultLicensee
End Property
Public Property Let DefaultLicensee(ByVal vNewValue As Boolean)
    mvarDefaultLicensee = vNewValue
End Property

Friend Sub FLicensee_BeforeUpdate(Cancel As Boolean, SelectedCountry As String, SelectedCurrency As String, SelectedLanguage As String)
    RaiseEvent BeforeUpdate(Cancel, SelectedCountry, SelectedCurrency, SelectedLanguage)
End Sub
Friend Sub FLicensee_CountryPicklist(Cancel As Boolean, code As String, CountryName As String)
    RaiseEvent CountryPick(Cancel, code, CountryName)
End Sub
Friend Sub FLicensee_CurrencyPicklist(Cancel As Boolean, code As String, CurrencyName As String)
    RaiseEvent CurrencyPick(Cancel, code, CurrencyName)
End Sub
Friend Sub FLicensee_LanguagePicklist(Cancel As Boolean, code As String, LanguageName As String)
    RaiseEvent LanguagePick(Cancel, code, LanguageName)
End Sub
Friend Sub FLicensee_UnloadForm()
    RaiseEvent UnloadForm
End Sub
