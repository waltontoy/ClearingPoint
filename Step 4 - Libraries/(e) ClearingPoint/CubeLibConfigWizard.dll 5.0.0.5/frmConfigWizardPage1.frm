VERSION 5.00
Begin VB.Form frmConfigWizardPage1 
   BorderStyle     =   0  'None
   Caption         =   "ClearingPoint Configuration Wizard"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   Icon            =   "frmConfigWizardPage1.frx":0000
   LinkTopic       =   "ClearingPoint Configuration Wizard"
   ScaleHeight     =   3855
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Index           =   0
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   6555
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3555
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2040
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
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   7
            Top             =   240
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
            Index           =   25
            Left            =   120
            TabIndex        =   6
            Top             =   816
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
            Index           =   26
            Left            =   120
            TabIndex        =   5
            Top             =   1392
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
            Index           =   27
            Left            =   120
            TabIndex        =   4
            Top             =   1968
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
            Index           =   28
            Left            =   120
            TabIndex        =   3
            Top             =   2544
            Width           =   1815
         End
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
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   2
            Top             =   3120
            Width           =   1815
         End
      End
      Begin VB.Label Label1 
         Caption         =   $"frmConfigWizardPage1.frx":058A
         Height          =   900
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   405
         Width           =   4125
      End
      Begin VB.Label Label1 
         Caption         =   "Click Next to run this configuration wizard."
         Height          =   900
         Index           =   1
         Left            =   2280
         TabIndex        =   8
         Top             =   2010
         Width           =   4125
      End
   End
End
Attribute VB_Name = "frmConfigWizardPage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    DefLng A-Z
    
    Implements IWizardPage
    
    Private m_WizardController As CWizardController
    
    Private m_blnEndConfiguration As Boolean

Private Sub Form_Unload(Cancel As Integer)
    If m_WizardController.ConfigurationCancelled = True Then
        g_blnConfigurationFinished = False
    Else
        g_blnConfigurationFinished = True
    End If
    
    m_blnEndConfiguration = True
    Set m_WizardController = Nothing
End Sub

Private Sub IWizardPage_BeforePageHide(Wizard As Object, ByVal NextStep As Integer, Cancel As Boolean)

End Sub

Private Sub IWizardPage_BeforePageShow(Wizard As Object, ByVal CurrentStep As Integer)

'// We're not doing anything here, but we need to add something
'// or VB will complain about unimplemented interfaces.

End Sub

Public Function ShowConfigUtility(ByRef CallingForm As Object) As Boolean

    Dim strCommand As String
    
    If Not m_WizardController Is Nothing Then
        Set m_WizardController = Nothing
    End If
    Set m_WizardController = New CWizardController
    
    m_WizardController.AddPage frmConfigWizardPage1, "StepOne"
    m_WizardController.AddPage New frmConfigWizardPage2, "StepTwo"
    m_WizardController.AddPage New frmConfigWizardPage3, "StepThree"
    m_WizardController.AddPage New frmConfigWizardPage4, "StepFour"
    m_WizardController.AddPage New frmConfigWizardPage5, "StepFive"
    m_WizardController.AddPage New frmConfigWizardPage6, "StepSix"

    m_WizardController.Start
    
    m_WizardController.ConfigurationCancelled = False
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[LOGID DESCRIPTION], "
        strCommand = strCommand & "[SEND OPERATIONAL CORR], "
        strCommand = strCommand & "[SEND OPERATIONAL LOGID], "
        strCommand = strCommand & "[SEND OPERATIONAL PASS], "
        strCommand = strCommand & "[SEND TEST CORR], "
        strCommand = strCommand & "[SEND TEST LOGID], "
        strCommand = strCommand & "[SEND TEST PASS], "
        strCommand = strCommand & "[PRINT OPERATIONAL CORR], "
        strCommand = strCommand & "[PRINT OPERATIONAL LOGID], "
        strCommand = strCommand & "[PRINT OPERATIONAL PASS], "
        strCommand = strCommand & "[PRINT TEST CORR], "
        strCommand = strCommand & "[PRINT TEST LOGID], "
        strCommand = strCommand & "[PRINT TEST PASS], "
        strCommand = strCommand & "[A2], "
        strCommand = strCommand & "[A1], "
        strCommand = strCommand & "[HISTORY], "
        strCommand = strCommand & "[VAT], "
        strCommand = strCommand & "[TIN], "
        strCommand = strCommand & "[USAGE], "
        strCommand = strCommand & "[LRN USAGE], "
        strCommand = strCommand & "[Procedure], "
        strCommand = strCommand & "[PRINT MODE] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[LOGICAL ID] "
    ADORecordsetOpen "SELECT * FROM [LOGICAL ID]", g_conSadbel, g_rstLogID, adOpenKeyset, adLockOptimistic
        
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[USER NO], "
        strCommand = strCommand & "[LOGID DESCRIPTION], "
        strCommand = strCommand & "[DTYPE], "
        strCommand = strCommand & "[COMM], "
        strCommand = strCommand & "[PRINT] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[USER LOGICAL ID] "
    ADORecordsetOpen strCommand, g_conSadbel, g_rstUserLogID, adOpenKeyset, adLockOptimistic
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[USER NO SERIES], "
        strCommand = strCommand & "[DUTCH], "
        strCommand = strCommand & "[LAST USER], "
        strCommand = strCommand & "[WITH SECURITY], "
        strCommand = strCommand & "[FIRSTRUN], "
        strCommand = strCommand & "[NUMBER OF USERS], "
        strCommand = strCommand & "[ENGLISH], "
        strCommand = strCommand & "[FRENCH], "
        strCommand = strCommand & "[EDIT TIME], "
        strCommand = strCommand & "[SENT TIME], "
        strCommand = strCommand & "[TREE TIME] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[SETUP] "
    ADORecordsetOpen strCommand, g_conSadbel, g_rstSetupSadbel, adOpenKeyset, adLockOptimistic

        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[User_ID], "
        strCommand = strCommand & "[User_Name], "
        strCommand = strCommand & "[User_Password], "
        strCommand = strCommand & "[ADMINISTRATOR RIGHTS], "
        strCommand = strCommand & "[MAINTAIN TABLES], "
        strCommand = strCommand & "[ALL LOGICAL IDS], "
        strCommand = strCommand & "[SHOW ALL SENT], "
        strCommand = strCommand & "[SHOW ALL WITH ERRORS], "
        strCommand = strCommand & "[SHOW ALL WAITING], "
        strCommand = strCommand & "[SHOW ALL DELETED], "
        strCommand = strCommand & "[CLEAN UP DELETED], "
        strCommand = strCommand & "[EVERY], "
        strCommand = strCommand & "[DELETE OTHER USERS ITEMS], "
        strCommand = strCommand & "[REFRESH IN SECONDS], "
        strCommand = strCommand & "[LOGID DESCRIPTION], "
        strCommand = strCommand & "[SHOW ALL TOBEPRINTED], "
        strCommand = strCommand & "[SHOW ALL DRAFTS], "
        strCommand = strCommand & "[LANGUAGE] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[USERS]"
    ADORecordsetOpen strCommand, g_conTemplate, g_rstUser, adOpenKeyset, adLockOptimistic
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[DBProps_DBEmpty] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[DBPROPERTIES]"
    ADORecordsetOpen strCommand, g_conTemplate, g_rstTemplate, adOpenKeyset, adLockOptimistic
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[EMPTY PRINTBOX], "
        strCommand = strCommand & "[UserName], "
        strCommand = strCommand & "[GATEWAY], "
        strCommand = strCommand & "[IP ADDRESS], "
        strCommand = strCommand & "[WAIT TIME], "
        strCommand = strCommand & "[CONTROLPANEL], "
        strCommand = strCommand & "[CUT3CHARS], "
        strCommand = strCommand & "[OPENCLOSE] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[SETUP] "
    ADORecordsetOpen strCommand, g_conScheduler, g_rstSetupSched, adOpenKeyset, adLockOptimistic
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[LOGID], "
        strCommand = strCommand & "[Mode], "
        strCommand = strCommand & "[Printer], "
        strCommand = strCommand & "[SEPARATOR], "
        strCommand = strCommand & "[PRINTING], "
        strCommand = strCommand & "[DOWNLOAD] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[PRINTER DEFINITION] "
    ADORecordsetOpen strCommand, g_conScheduler, g_rstPrinterDef, adOpenKeyset, adLockOptimistic
    
        strCommand = vbNullString
        strCommand = strCommand & "SELECT "
        strCommand = strCommand & "[LogID], "
        strCommand = strCommand & "[Mode], "
        strCommand = strCommand & "[SCHEDULE], "
        strCommand = strCommand & "[Default], "
        strCommand = strCommand & "[LAST RUN], "
        strCommand = strCommand & "[TEMP DEFAULT] "
        strCommand = strCommand & "FROM "
        strCommand = strCommand & "[LOGID SCHEDULE] "
    ADORecordsetOpen strCommand, g_conScheduler, g_rstLogIDSched, adOpenKeyset, adLockOptimistic
    
    g_blnConfigurationFinished = False
    m_blnEndConfiguration = False
    
    Set clsConfigWizard = New CConfigWizard
    
    g_strAdminUserLabel = "Administrator Name"
    
    frmConfigWizardPage1.Show
    
    Do While m_blnEndConfiguration = False
        DoEvents
    Loop
    
    ADORecordsetClose g_rstLogID
    ADORecordsetClose g_rstUserLogID
    ADORecordsetClose g_rstSetupSadbel

    ADORecordsetClose g_rstUser
    ADORecordsetClose g_rstTemplate
    
    ADORecordsetClose g_rstSetupSched
    ADORecordsetClose g_rstPrinterDef
    ADORecordsetClose g_rstLogIDSched
    
    Set clsConfigWizard = Nothing
    
    ShowConfigUtility = g_blnConfigurationFinished
    
End Function

