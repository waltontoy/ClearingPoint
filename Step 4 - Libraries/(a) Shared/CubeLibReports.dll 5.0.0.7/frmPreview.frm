VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPreview 
   Caption         =   "Preview"
   ClientHeight    =   7455
   ClientLeft      =   2100
   ClientTop       =   2835
   ClientWidth     =   11520
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7336.901
   ScaleMode       =   0  'User
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   9240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   9840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   10440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer arvReportViewer 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13150
      SectionData     =   "frmPreview.frx":058A
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu miFSave 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu miFSendToo 
         Caption         =   "&Mail..."
      End
      Begin VB.Menu miFPDFExport 
         Caption         =   "P&DF Export..."
         Visible         =   0   'False
      End
      Begin VB.Menu miFRTFExport 
         Caption         =   "&RTF Export..."
         Visible         =   0   'False
      End
      Begin VB.Menu miFHTMLExport 
         Caption         =   "&HTML Export..."
         Visible         =   0   'False
      End
      Begin VB.Menu miFTextExport 
         Caption         =   "&Text Export..."
         Visible         =   0   'False
      End
      Begin VB.Menu miFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu miFPrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu miFPrinterSetup 
         Caption         =   "Printer &Setup"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu miFSep3 
         Caption         =   "-"
      End
      Begin VB.Menu miFExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim strBodyMessage As String
Dim strTitle As String
Dim strRecipient As String
Dim clsFolder As CFolders

Public Sub RunReport(ByVal Report As Object, ByVal Title As String, ByVal Recipient As String, ByVal BodyMessage As String)
    Set arvReportViewer.ReportSource = Report

    strBodyMessage = BodyMessage
    strRecipient = Recipient
    strTitle = Title
End Sub

Private Sub Form_Load()
    Set clsFolder = New CFolders

    Me.Caption = strTitle

    Set clsFolder = Nothing
End Sub

Private Sub Form_Resize()
    arvReportViewer.Height = Me.ScaleHeight
    arvReportViewer.Width = Me.ScaleWidth
End Sub

Private Sub miFPrint_Click()
    arvReportViewer.UseSourcePrinter = True
    arvReportViewer.PrintReport True
End Sub

Private Sub miFExit_Click()
    Unload Me
End Sub

Public Sub miFSave_Click()
    Dim strFileTitle As String
    Dim intPdfExt As Integer
    Dim intRtfExt As Integer
    Dim intHtmlExt As Integer
    Dim intTxtExt As Integer
    Dim xptPdf As ActiveReportsPDFExport.ARExportPDF
    Dim xptRtf As ActiveReportsRTFExport.ARExportRTF
    Dim xptHtml As ActiveReportsHTMLExport.HTMLexport
    Dim xptText As ActiveReportsTextExport.ARExportText

    dlgDialog.Filter = "Rich Text Format(*.rtf)|*.rtf|" & _
                    "HTML Format (*.html)|*.html|" & _
                    "Portable Document Format (*.pdf)|*.pdf|" & _
                    "Text Format (*.txt)|*.txt "

    dlgDialog.ShowSave

    If dlgDialog.Filename <> "" Then
        strFileTitle = dlgDialog.FileTitle

        intPdfExt = InStr(strFileTitle, ".pdf")
        intRtfExt = InStr(strFileTitle, ".rtf")
        intHtmlExt = InStr(strFileTitle, ".html")
        intTxtExt = InStr(strFileTitle, ".txt")

        If intPdfExt > 0 Then
            Set xptPdf = New ActiveReportsPDFExport.ARExportPDF

            xptPdf.Filename = dlgDialog.Filename
            If Not arvReportViewer.ReportSource Is Nothing Then
                xptPdf.Export arvReportViewer.ReportSource.Pages
            Else
                xptPdf.Export arvReportViewer.Pages
            End If

            Set xptPdf = Nothing
        ElseIf intRtfExt > 0 Then
            Set xptRtf = New ActiveReportsRTFExport.ARExportRTF

            xptRtf.Filename = dlgDialog.Filename
            If Not arvReportViewer.ReportSource Is Nothing Then
                xptRtf.Export arvReportViewer.ReportSource.Pages
            Else
                xptRtf.Export arvReportViewer.Pages
            End If

            Set xptRtf = Nothing
        ElseIf intHtmlExt > 0 Then
            Set xptHtml = New ActiveReportsHTMLExport.HTMLexport

            xptHtml.Filename = dlgDialog.Filename
            If Not arvReportViewer.ReportSource Is Nothing Then
                xptHtml.Export arvReportViewer.ReportSource.Pages
            Else
                xptHtml.Export arvReportViewer.Pages
            End If

            Set xptHtml = Nothing
        ElseIf intTxtExt > 0 Then
            Set xptText = New ActiveReportsTextExport.ARExportText

            xptText.Filename = dlgDialog.Filename
            If Not arvReportViewer.ReportSource Is Nothing Then
                xptText.Export arvReportViewer.ReportSource.Pages
            Else
                xptText.Export arvReportViewer.Pages
            End If

            Set xptText = Nothing
        Else
        End If
    End If
End Sub

Private Sub miFSendToo_Click()
    Dim xptPdf As ActiveReportsPDFExport.ARExportPDF
    Dim strTempPath As String
    Dim strFileName As String
    Dim strPathFileName As String
    Dim scpScript
    Dim fleFile
    Dim clsMail As cpiEmail

On Error GoTo File_DNE   '----->File Doesn't Exist yet in the temporary folder

    strTempPath = GetTemporaryPath()
    strPathFileName = strTempPath & Trim(strTitle) & ".pdf"

    Set scpScript = CreateObject("Scripting.FileSystemObject")
    Set fleFile = scpScript.getfile(strPathFileName) 'file to be deleted if it exists
    fleFile.Delete

File_DNE:

    On Error GoTo 0

    dlgDialog.Filter = "Portable Document Format (*.pdf)|*.pdf|"
    dlgDialog.Filename = strPathFileName

    Set xptPdf = New ActiveReportsPDFExport.ARExportPDF
    xptPdf.Filename = dlgDialog.Filename
    If Not arvReportViewer.ReportSource Is Nothing Then
        xptPdf.Export arvReportViewer.ReportSource.Pages
    Else
        xptPdf.Export arvReportViewer.Pages
    End If
    Set xptPdf = Nothing

    dlgDialog.Filter = ""
    dlgDialog.Filename = ""

    '----->Start Sending As an email attachment
    Set clsMail = New cpiEmail
    clsMail.Send strRecipient, strTitle, strBodyMessage, strPathFileName
    Set clsMail = Nothing

    Set scpScript = CreateObject("Scripting.FileSystemObject")
    Set fleFile = scpScript.getfile(strPathFileName)
    fleFile.Delete
End Sub

