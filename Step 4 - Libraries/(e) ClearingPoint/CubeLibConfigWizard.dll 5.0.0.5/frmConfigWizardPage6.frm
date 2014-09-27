VERSION 5.00
Begin VB.Form frmConfigWizardPage6 
   BorderStyle     =   0  'None
   Caption         =   "ClearingPoint Configuration Wizard"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   Icon            =   "frmConfigWizardPage6.frx":0000
   LinkTopic       =   "ClearingPoint Configuration Wizard"
   ScaleHeight     =   3855
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   5
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   6555
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3555
         Index           =   5
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2040
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "Finish"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   9
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "     Programs"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   8
            Top             =   2544
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "     Connections"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   7
            Top             =   1968
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "     User Account"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   6
            Top             =   1392
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "     Company Data"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   5
            Top             =   816
            Width           =   1815
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "View Readme File"
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   0
         Top             =   1605
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Launch Program File"
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   1
         Top             =   2025
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "ClearingPoint now has all the necessary information to run the program."
         Height          =   720
         Index           =   10
         Left            =   2280
         TabIndex        =   10
         Top             =   360
         Width           =   4035
      End
   End
End
Attribute VB_Name = "frmConfigWizardPage6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    DefLng A-Z
    
    Implements IWizardPage

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbHourglass
    
    If g_blnConfigurationFinished = True Then
        doFinished
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub IWizardPage_BeforePageHide(Wizard As Object, ByVal NextStep As Integer, Cancel As Boolean)
    
End Sub

Private Sub IWizardPage_BeforePageShow(Wizard As Object, ByVal CurrentStep As Integer)

'// We're not doing anything here, but we need to add something
'// or VB will complain about unimplemented interfaces.

End Sub

Private Sub doFinished()
    Dim rstBoxDefaultAdmin As ADODB.Recordset
    
    Dim cPrnPrg As String
    Dim cLogDesc As String
    Dim lngUserID As Long
    Dim c_UserNo As String
    Dim c_Var As String
    Dim i As Integer
    Dim dtmNow As Date
    Dim enuExecuteRecordset As ExecuteRecordsetConstant
    
    cLogDesc = Trim(clsConfigWizard.CompanyName) ' company name
    
    'Set rstPLDAMessages = ExecuteQuery(strSQL, DBInstanceType_DATABASE_SADBEL)
         
    Dim strCommand As String
    
    With g_rstUser
        .AddNew
        
        'lngUserID = ![User_ID]
        
        ![User_Name] = Trim(clsConfigWizard.AdministratorName) 'administrator name
        ![User_Password] = Trim(clsConfigWizard.Password) ' password
        ![ADMINISTRATOR RIGHTS] = True
        ![MAINTAIN TABLES] = True
        ![ALL LOGICAL IDS] = True
        ![SHOW ALL SENT] = True
        ![SHOW ALL WITH ERRORS] = True
        ![SHOW ALL WAITING] = True
        ![SHOW ALL DELETED] = True
        ![CLEAN UP DELETED] = True
        ![EVERY] = 0
        ![DELETE OTHER USERS ITEMS] = True
        ![REFRESH IN SECONDS] = 20
        
        ![LOGID DESCRIPTION] = Trim(cLogDesc)
        ![SHOW ALL TOBEPRINTED] = True
        ![SHOW ALL DRAFTS] = True
        ![LANGUAGE] = 0
        
        .Update
        
        SaveSetting "ClearingPoint", "Settings", "Company Name", cLogDesc
    End With
    
    lngUserID = ExecuteRecordset(ExecuteRecordsetConstant.Insert, g_conTemplate, g_rstUser, "USERS")
    
    With g_rstUserLogID
        For i = 1 To 3
            .AddNew
            
            ![USER NO] = lngUserID
            ![LOGID DESCRIPTION] = Trim(cLogDesc)
            
            ![DTYPE] = i
            ![COMM] = "S"
            ![PRINT] = "I"
            
            .Update
            
            ExecuteRecordset ExecuteRecordsetConstant.Insert, g_conSadbel, g_rstUserLogID, "USER LOGICAL ID"
        Next
    End With
    
    With g_rstSetupSadbel
        
        If (.EOF And .BOF) Then
            enuExecuteRecordset = ExecuteRecordsetConstant.Insert
            .AddNew
        Else
            enuExecuteRecordset = ExecuteRecordsetConstant.Update
            .MoveFirst
        End If
        
        ![USER NO SERIES] = "2"
        ![DUTCH] = True
        ![LAST USER] = Trim(clsConfigWizard.AdministratorName) 'administrator name
        ![WITH SECURITY] = True
        ![FIRSTRUN] = False
        ![NUMBER OF USERS] = 5
        
        ![ENGLISH] = True
        ![FRENCH] = True
        
        dtmNow = Now()
        
        ![EDIT TIME] = dtmNow
        ![SENT TIME] = dtmNow
        ![TREE TIME] = dtmNow
        
        .Update

        ExecuteRecordset enuExecuteRecordset, g_conSadbel, g_rstSetupSadbel, "SETUP"
        
    End With
    
    If Not (g_rstTemplate.BOF And g_rstTemplate.EOF) Then
        g_rstTemplate.Fields("DBProps_DBEmpty").Value = False
        g_rstTemplate.Update
        
        ExecuteRecordset ExecuteRecordsetConstant.Update, g_conTemplate, g_rstTemplate, "DBPROPERTIES"
    End If
    
    For i = 1 To 5

            strCommand = vbNullString
            strCommand = strCommand & "SELECT "
            strCommand = strCommand & "[EMPTY FIELD VALUE], "
            strCommand = strCommand & "[DEFAULT VALUE], "
            strCommand = strCommand & "[BOX CODE] "
            strCommand = strCommand & "FROM "
            strCommand = strCommand & "[BOX DEFAULT " & Choose(i, "IMPORT", "EXPORT", "TRANSIT", "TRANSIT NCTS", "COMBINED NCTS") & " ADMIN] "
        ADORecordsetOpen strCommand, g_conSadbel, rstBoxDefaultAdmin, adOpenKeyset, adLockOptimistic
        
        With rstBoxDefaultAdmin
            
            If Len(Trim(clsConfigWizard.CustomsOffice)) Then 'customs office
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    .Find "[BOX CODE] = '" & "A4" & "'", , adSearchForward

                    If Not .EOF Then
                        
                        ![EMPTY FIELD VALUE] = Trim(clsConfigWizard.CustomsOffice) 'customs office
                        ![DEFAULT VALUE] = Trim(clsConfigWizard.CustomsOffice) 'customs office
                        
                        .Update
                        
                        ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSadbel, rstBoxDefaultAdmin, "BOX DEFAULT " & Choose(i, "IMPORT", "EXPORT", "TRANSIT", "TRANSIT NCTS", "COMBINED NCTS") & " ADMIN"
                    End If
                End If
            End If
            
            If Len(Trim(clsConfigWizard.LanguageofDeclaration)) Then 'language of declaration
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    .Find "[BOX CODE] = '" & "A5" & "'", , adSearchForward

                    If Not .EOF Then
                    
                        ![EMPTY FIELD VALUE] = Trim(clsConfigWizard.LanguageofDeclaration) 'language of declaration
                        ![DEFAULT VALUE] = Trim(clsConfigWizard.LanguageofDeclaration) 'language of declaration
                        
                        .Update
                        
                        ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSadbel, rstBoxDefaultAdmin, "BOX DEFAULT " & Choose(i, "IMPORT", "EXPORT", "TRANSIT", "TRANSIT NCTS", "COMBINED NCTS") & " ADMIN"
                    End If
                End If
            End If
            
            If Len(Trim(clsConfigWizard.PlaceofLoading)) Then 'place of loading
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                    .Find "[BOX CODE] = '" & IIf(i = 1, "B7", "B5") & "'", , adSearchForward

                    If Not .EOF Then
                    
                        ![EMPTY FIELD VALUE] = Trim(clsConfigWizard.PlaceofLoading) 'place of loading
                        ![DEFAULT VALUE] = Trim(clsConfigWizard.PlaceofLoading) 'place of loading
                        
                        .Update
                        
                        ExecuteRecordset ExecuteRecordsetConstant.Update, g_conSadbel, rstBoxDefaultAdmin, "BOX DEFAULT " & Choose(i, "IMPORT", "EXPORT", "TRANSIT", "TRANSIT NCTS", "COMBINED NCTS") & " ADMIN"
                    End If
                End If
            End If
            
        End With
        
        ADORecordsetClose rstBoxDefaultAdmin
    Next
    
    ADORecordsetClose rstBoxDefaultAdmin
    
    With g_rstSetupSched
        
        If (.EOF And .BOF) Then
            enuExecuteRecordset = ExecuteRecordsetConstant.Insert
            .AddNew
        Else
            enuExecuteRecordset = ExecuteRecordsetConstant.Update
            .MoveFirst
        End If
        
        ![EMPTY PRINTBOX] = "2"
        ![UserName] = Trim(clsConfigWizard.UserName) 'user name
        ![GATEWAY] = Trim(clsConfigWizard.GatewayName) 'gateway name
        ![IP ADDRESS] = Trim(clsConfigWizard.IPAddress) 'ip address
        ![WAIT TIME] = "10"
        ![CONTROLPANEL] = False
        ![CUT3CHARS] = True
        ![OPENCLOSE] = False
        
        .Update
        
        ExecuteRecordset enuExecuteRecordset, g_conScheduler, g_rstSetupSched, "SETUP"
        
    End With
    
    With g_rstPrinterDef
    
        If Not (.EOF And .BOF) Then
            .MoveFirst
            .Find "[LOGID] = '" & Trim(clsConfigWizard.LogicalID) & "'", , adSearchForward
        End If
        
        For i = 1 To 2
            .AddNew
            
            ![LogID] = Trim(clsConfigWizard.LogicalID) 'logical id
            ![Mode] = Choose(i, "O", "T")
            ![Printer] = Trim(clsConfigWizard.PrinterName) 'printers
            ![SEPARATOR] = "none"
            ![PRINTING] = "A"
            ![DOWNLOAD] = "A"
            
            .Update
            
            ExecuteRecordset ExecuteRecordsetConstant.Insert, g_conScheduler, g_rstPrinterDef, "PRINTER DEFINITION"
        Next
    End With
    
    With g_rstLogIDSched
        If Not (.EOF And .BOF) Then
            .MoveFirst
            .Find "[LOGID] = '" & Trim(clsConfigWizard.LogicalID) & "'", , adSearchForward
        End If
        
        For i = 1 To 3
            .AddNew
            
            ![LogID] = Trim(clsConfigWizard.LogicalID) 'logical id
            ![Mode] = Choose(i, "O", "T", "D")
            ![SCHEDULE] = 0
            ![Default] = "D"
            ![LAST RUN] = dtmNow
            ![TEMP DEFAULT] = "D"
            
            .Update
            
            ExecuteRecordset ExecuteRecordsetConstant.Insert, g_conScheduler, g_rstLogIDSched, "LOGID SCHEDULE"
        Next
    End With
    
    If clsConfigWizard.PrintingProgram = 0 Then
        cPrnPrg = frmConfigWizardPage5.Option1(0).Caption
    ElseIf clsConfigWizard.PrintingProgram = 1 Then
        cPrnPrg = frmConfigWizardPage5.Option1(1).Caption
    ElseIf clsConfigWizard.PrintingProgram = 2 Then
        cPrnPrg = frmConfigWizardPage5.Option1(2).Caption
    End If
    
    With g_rstLogID
        .AddNew
        
        ![LOGID DESCRIPTION] = Trim(clsConfigWizard.CompanyName) 'company name
        ![SEND OPERATIONAL CORR] = "TPSAD"
        ![SEND OPERATIONAL LOGID] = Trim(clsConfigWizard.LogicalID) 'logical id
        ![SEND OPERATIONAL PASS] = Trim(clsConfigWizard.SendingPassword) 'sending password
        ![SEND TEST CORR] = "TESSAD"
        ![SEND TEST LOGID] = Trim(clsConfigWizard.LogicalID) 'logical id
        ![SEND TEST PASS] = Trim(clsConfigWizard.SendingPassword) 'sending password
        ![PRINT OPERATIONAL CORR] = Trim(Left(cPrnPrg, InStr(1, cPrnPrg, "/") - 1))
        ![PRINT OPERATIONAL LOGID] = Trim(clsConfigWizard.LogicalID) 'logical id
        ![PRINT OPERATIONAL PASS] = Trim(clsConfigWizard.PrintingPassword) 'printing password
        ![PRINT TEST CORR] = Trim(Mid(cPrnPrg, InStr(1, cPrnPrg, "/") + 1))
        ![PRINT TEST LOGID] = Trim(clsConfigWizard.LogicalID) 'logical id
        ![PRINT TEST PASS] = Trim(clsConfigWizard.PrintingPassword) 'printing password
        ![A2] = Trim(clsConfigWizard.Account49) 'account 49
        ![A1] = Trim(clsConfigWizard.CustomsRegNo) 'customs reg number
        ![HISTORY] = "IET "
        ![VAT] = "Enter VAT"
        ![TIN] = "Enter TIN"
        ![USAGE] = 7
        ![LRN USAGE] = 0
        ![Procedure] = 0
        ![PRINT MODE] = 0
        
        .Update
        
        ExecuteRecordset ExecuteRecordsetConstant.Insert, g_conSadbel, g_rstLogID, "LOGICAL ID"
    End With
    
    'g_conScheduler.Execute "UPDATE [TASK SCHEDULE] SET [LAST RUN] = #" & dtmNow & "#, [PROPERTY] = 10", dbFailOnError
    ExecuteNonQuery g_conScheduler, "UPDATE [TASK SCHEDULE] SET [LAST RUN] = #" & dtmNow & "#, [PROPERTY] = 10"
End Sub
