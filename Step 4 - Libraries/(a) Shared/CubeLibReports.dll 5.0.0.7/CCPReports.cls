VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCPReports"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ReportNCTSCustomsDocumentTypeConstant   'NCTSReportTypeConstant
    [NCTS Departure Volgbriefje] = 1
    [NCTS Departure Accompanying] = 2
    [NCTS Departure Codisheet] = 3
    [NCTS Codisheet] = 4
    [NCTS Combined Codisheet] = 5
    [NCTS EDI Departure Codisheet] = 6
End Enum

Public Enum ReportCustomsDocumentTypeConstant       'CustomsDocumentTypeConstant
    [SADBEL Import Document] = 1
    [SADBEL Export Document] = 2
    [SADBEL Transit Document] = 3
    [SADBEL Import Document Duplicata] = 4
    [SADBEL Export Document Duplicata] = 5
    [SADBEL Transit Document Duplicata] = 6
    [Transit NCTS Document] = 7
    [Transit NCTS Document Duplicata] = 8
    [Combined NCTS Document] = 9
    [Combined NCTS Document Duplicata] = 10
    [NCTS Departure Document] = 11
    [NCTS Arrival Document] = 12
    [PLDA Import Document] = 14
    [PLDA Import Document Duplicata] = 15
    [PLDA Export Document] = 16
    [PLDA Export Document Duplicata] = 17
    [PLDA Combined Document] = 18
    [PLDA Combined Document Duplicata] = 19
    [PLDA Volgbriefje] = 20
    [PLDA EAD] = 21
    [PLDA DV1] = 22
    [PLDA EUR1] = 23
    [PLDA ATR] = 24
    [PLDA Import Codisheet] = 25
    [PLDA Combined Codisheet] = 26
    [PLDA Volgbriefje Import] = 27
End Enum

'Public Enum ERunType
'    cpiShowModal = 0           ' [Show Report Modal]
'    cpiShowNonModal = 1        ' [Show Report Non-Modal]
'    cpiPreview = 2             ' [Generate Preview]
'    cpiPrintWithDialog = 3     ' [Generate Print With Dialog]
'    cpiPrintWithoutDialog = 4  ' [Generate Print Without Dialog]
'    cpiAutomaticSave = 5       ' [Generate PDF Export Pages]
'End Enum

Public Enum GenerateReportModeConstant
    [Generate Preview] = 1                  ' ModePreview, cpiPreview
    [Generate Print Without Dialog] = 2     ' ModePrint, cpiPrintWithoutDialog
    [Generate Print With Dialog] = 3        ' ModePrint, cpiPrintWithDialog
    [Generate PDF Export Pages] = 4         ' ModeCreatePDF, cpiAutomaticSave
    [Create PDF Output File] = 5            ' ModePDFOut
    [Show Report Modal] = 6                 ' cpiShowModal
    [Show Report Non-Modal] = 6             ' cpiShowNonModal
End Enum

Private Type PROCESS_INFORMATION
    hProcess        As Long
    hThread         As Long
    dwProcessID     As Long
    dwThreadID      As Long
End Type

Private Type STARTUPINFO
    cb              As Long
    lpReserved      As String
    lpDesktop       As String
    lpTitle         As String
    dwX             As Long
    dwY             As Long
    dwXSize         As Long
    dwYSize         As Long
    dwXCountChars   As Long
    dwYCountChars   As Long
    dwFillAttribute As Long
    dwFlags         As Long
    wShowWindow     As Integer
    cbReserved2     As Integer
    lpReserved2     As Long
    hStdInput       As Long
    hStdOutput      As Long
    hStdError       As Long
End Type

Private Type udtCPReportsType
    CommandLine      As String               'command line for CPReports
    udtProcess       As PROCESS_INFORMATION
    udtStart         As STARTUPINFO
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


'********************************************************
' Returns the commandline issued for debugging purposes '
'********************************************************
Public Function RunCPReports(ByVal CPReportsExePath As String, _
                             ByVal PersistenceFilePath As String, _
                             ByVal UserID As Long, _
                             ByVal UniqueCode As String, _
                             ByVal DocumentType As ReportCustomsDocumentTypeConstant, _
                             ByVal GenerateReportMode As GenerateReportModeConstant, _
                             ByVal DocumentName As String, _
                             ByVal AutoPrint As Boolean, _
                             ByVal HistoryDBName As String, _
                             ByVal ReleaseMessageFilePath As String, _
                             ByVal PrintData3Filename As String, _
                             ByVal ForAutomaticPrinting As Boolean, _
                    Optional ByVal WaitToFinishReport As Boolean = False) As String
    
    Dim udtCPReports As udtCPReportsType
    Dim strCommandLine As String
    
    Dim lngExecuteCommand As Long
    Dim lngMilliseconds As Long
    
    Const NORMAL_PRIORITY_CLASS As Long = &H20
    Const WAIT_TIMEOUT          As Long = &H102
    
    Const COMMAND_LINE_DELIMITER = "|*||*|"
    
    
    On Error GoTo Early_Exit
    
    Screen.MousePointer = vbHourglass

    '**********************************************************************************************************************
    'COMMAND STRUCTURE:
    '**********************************************************************************************************************
    ' 0 -   PersistenceFilePath     Path of the Persistence.txt              '''''' 0 -   MdbPath            Path of the database
    ' 1 -   User ID                 Used to retrieve sheets to print; used only when printing from codisheet
    ' 2 -   Unique Code             Unique code identifying the document
    ' 3 -   Document Type           Document type constant to determine the type of report
    ' 4 -   GenerateReport Mode     Preview, print, or create pdf?
    ' 5 -   Document Name           Required if Genrate Mode = [Generate PDF], ELSE Preview, ExportPages
    ' 6 -   From AutoPrint          To determine if the call is from TS automatic printing
    ' 7 -   HistoryDBName           Support for printing in Archives
    ' 8 -   ReleaseMessage          To Update Offline Recordset for EAD and LOI
    ' 9 -   PrintData3Filename      Filename of FlatFile where PrintData3 was saved in CP (EP Folder only) - USED FOR EUR1, ATR (???)
    '10 -   ForAutomaticPrinting    ForAutomaticPrinting
    '**********************************************************************************************************************
    
'******************************************************************************************
'    O  R  I  G  I  N  A  L
'******************************************************************************************
    ' 0 -   PersistenceFilePath     Path of the Persistence.txt              '''''' 0 -   MdbPath            Path of the database
    ' 1 -   Unique Code             Unique code identifying the document
    ' 2 -   Document Type           Document type constant to determine the type of report
    ' 3 -   Generate Mode           Preview, print, or create pdf?
    ' 4 -   PrintDialog?            If Genrate Mode = [Generate Print], this will be used to determine whether
    '                               the print dialog should be displayed or not
    ' 5 -   FileName                Required if Genrate Mode = [Generate PDF]
    ' 6 -   From AutoPrint          To determine if the call is from TS automatic printing
    ' 7 -   User ID                 Used to retrieve sheets to print; used only when printing from codisheet
    ' 8 -   ReleaseMessage          To Update Offline Recordset for EAD and LOI
    ' 9 -   FileName                Filename of FlatFile where PrintData3 was saved in CP (EP Folder only)
    '10 -   HistoryDBName           Support for printing in Archives
    '11 -   ForAutomaticPrinting    ForAutomaticPrinting
    '**********************************************************************************************************************
    
    strCommandLine = vbNullString
    strCommandLine = strCommandLine & PersistenceFilePath
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & UserID
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & UniqueCode
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DocumentType
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & GenerateReportMode
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DocumentName
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & IIf(AutoPrint, 1, 0)
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & HistoryDBName
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & ReleaseMessageFilePath
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & PrintData3Filename
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & IIf(ForAutomaticPrinting, 1, 0)
                             
'******************************************************************************************
'    O  R  I  G  I  N  A  L
'******************************************************************************************
'    strCommandLine = MdbPath
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & UniqueCode
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DocumentType
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & GenerateReportMode.ModeCreatePDF
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & "0"
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & PDFName
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & "False"
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & 0
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & strMessageResponse   'CSCLP-142

    With udtCPReports
        .CommandLine = GetShortPath(NoBackSlash(CPReportsExePath)) & "\CPReports.exe " & strCommandLine
        
        ' Initialize the STARTUPINFO structure:
        .udtStart.cb = Len(.udtStart)
        
        ' Start the shelled application:
        lngExecuteCommand = CreateProcessA(0&, .CommandLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, .udtStart, .udtProcess)
        
        If WaitToFinishReport Then
            ' Wait for the shelled application to finish:
            Do While True
                lngExecuteCommand = WaitForSingleObject(.udtProcess.hProcess, lngMilliseconds)
                If lngExecuteCommand <> WAIT_TIMEOUT Then 'Or CheckPDFForEUR1ATR(DocumentType, PDFName) Then
                    Exit Do
                End If
                DoEvents
            Loop
        End If
        
    End With
    
    RunCPReports = strCommandLine
    
    Screen.MousePointer = vbDefault
    
Early_Exit:
    On Error GoTo 0
    
    RunCPReports = strCommandLine
End Function

'********************************************************
' Returns the commandline issued for debugging purposes '
'********************************************************
'Public Sub RunCPReportsNCTS(ByVal MdbPath As String, _
                            ByVal DataNCTSID As Long, _
                            ByVal UserID As Long, _
                            ByVal UniqueCode As String, _
                            ByVal GenerateMode As ERunType, _
                            ByVal DocumentType As NCTSReportTypeConstant, _
                            ByVal PDFNamePath As String, _
                            ByVal Language As String, _
                            ByVal DocumentName As String, _
                   Optional ByVal DType As Long = 0)

Public Function RunCPReportsNCTS(ByVal CPReportsNCTSExePath As String, _
                                 ByVal PersistenceFilePath As String, _
                                 ByVal UserID As Long, _
                                 ByVal DataNCTSID As Long, _
                                 ByVal UniqueCode As String, _
                                 ByVal DocumentType As ReportNCTSCustomsDocumentTypeConstant, _
                                 ByVal GenerateReportMode As GenerateReportModeConstant, _
                                 ByVal DocumentName As String, _
                                 ByVal AutoPrint As Boolean, _
                                 ByVal HistoryDBName As String, _
                                 ByVal Language As String) As String
    Dim udtNCTS As udtCPReportsType
    
    Dim strCommandLine As String
    Dim strHistoryDBName As String
    
    Dim lngExecuteCommand As Long
    Dim lngMilliseconds As Long
                            
    Const NORMAL_PRIORITY_CLASS As Long = &H20
    Const WAIT_TIMEOUT          As Long = &H102
    
    Const COMMAND_LINE_DELIMITER = "|*||*|"
    
    On Error GoTo Early_Exit
    
    Screen.MousePointer = vbHourglass
    
    '**********************************************************************************************************************
    'COMMAND STRUCTURE:
    '**********************************************************************************************************************
    ' 0 -   PersistencePath         Path of the Persistence.txt              '''''' 0 -   MdbPath             Path of the database
    ' 1 -   User ID                 User ID
    ' 2 -   Data NCTS ID            DATA NCTS ID
    ' 3 -   Unique Code             Unique code identifying the document
    ' 4 -   Document Type           Document type constant to determine the type of report
    '                               e.g. Codisheet, Volgbriefje, Accompanying Document
    ' 5 -   Generate Report Mode    e.g. Preview, Print, Save as PDF
    ' 6 -   Document Name           File Name if GenerateMode = Save as PDF, ELSE Preview, ExportPages
    ' 7 -   From Autoprint          To determine if the call is from TS automatic printing
    ' 8 -   HistoryDBName           History DB name when printing from archive
    ' 9 -   Language                Language
    '**********************************************************************************************************************
    
'******************************************************************************************
'    O  R  I  G  I  N  A  L
'******************************************************************************************
    ' 0 -   PersistencePath         Path of the Persistence.txt              '''''' 0 -   MdbPath             Path of the database
    ' 1 -   DataNCTSID              DATA NCTS ID
    ' 2 -   UserID                  User ID
    ' 3 -   UniqueCode              Unique code identifying the document
    ' 4 -   GenerateMode            e.g. Preview, Print, Save as PDF
    ' 5 -   Report Type             e.g. Codisheet, Volgbriefje, Accompanying Document
    ' 6 -   PDFNamePath             File Name if GenerateMode = Save as PDF
    ' 7 -   HistoryDBName           History DB name when printing from archive
    ' 8 -   Language                Language
    ' 9 -   Document Name           Document Name
    '10 -   CodiType                Codisheet Type
    '11 -   ForAutomaticPrinting    ForAutomaticPrinting
    '**********************************************************************************************************************
    
    
    strCommandLine = vbNullString
    strCommandLine = strCommandLine & PersistenceFilePath
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & UserID
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DataNCTSID
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & UniqueCode
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DocumentType
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & GenerateReportMode
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DocumentName
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & IIf(AutoPrint, 1, 0)
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & HistoryDBName
    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & Language
'******************************************************************************************
'    O  R  I  G  I  N  A  L
'******************************************************************************************
'    strCommandLine = MdbPath
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DataNCTSID
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & UserID
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & UniqueCode
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & GenerateMode
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DocumentType
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & PDFNamePath
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & strHistoryDBName
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & Language
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DocumentName
'    strCommandLine = strCommandLine & COMMAND_LINE_DELIMITER & DType
    
    ' SO THAT WE DON'T HAVE TO LOOK FOR THIS EVERYTIME WE ARE DEBUGGING
    ' CSCLP-886 - Start
    Debug.Assert False
    ' CSCLP-886 - End
    
    Screen.MousePointer = vbHourglass
    
    With udtNCTS
        .CommandLine = GetShortPath(CPReportsNCTSExePath) & "\CPReportNCTS.exe " & strCommandLine
        
        ' Initialize the STARTUPINFO structure:
        .udtStart.cb = Len(.udtStart)
        
        ' Start the shelled application:
        lngExecuteCommand = CreateProcessA(0&, .CommandLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, .udtStart, .udtProcess)
        
        If GenerateReportMode = GenerateReportModeConstant.[Generate PDF Export Pages] Then
            Do While True
                lngExecuteCommand = WaitForSingleObject(.udtProcess.hProcess, lngMilliseconds)
                If lngExecuteCommand <> WAIT_TIMEOUT Then
                    Exit Do
                End If
                DoEvents
            Loop
        End If
        
    End With
    
    RunCPReportsNCTS = strCommandLine
    
    Screen.MousePointer = vbDefault
    
Early_Exit:

    On Error GoTo 0
    
    RunCPReportsNCTS = strCommandLine
    
End Function


