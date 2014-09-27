VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCombinedReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Public Enum ERunType
'    cpiShowModal = 0
'    cpiShowNonModal = 1
'    cpiPreview = 2
'    cpiPrintWithDialog = 3
'    cpiPrintWithoutDialog = 4
'    cpiAutomaticSave = 5
'End Enum

Private Enum ExportFilterType
    sbpHTML = 0
    sbpPDF = 1
    sbpText = 2
End Enum

Private mstrFileName As String
Private mstrPrinterName As String
Private mlngPaperSize As Long
Private mlngPageLeftMargin As Long
Private mlngPageRightMargin As Long
Private mlngPageTopMargin As Long
Private mlngPageBottomMargin As Long
Private mintZoom As Integer

Public Enum eRemoteDocType
    e_Import = 0
    e_Export
    e_OTS
    e_CombinedNCTS
    e_DV1
    e_EUR1
    e_Volgbriefje
    e_AccDoc_LOI
End Enum

Public Property Let PageButtomMargin(ByVal vData As Long)
    mlngPageBottomMargin = vData
End Property

Public Property Get PageButtomMargin() As Long
    PageButtomMargin = mlngPageBottomMargin
End Property

Public Property Let PageTopMargin(ByVal vData As Long)
    mlngPageTopMargin = vData
End Property

Public Property Get PageTopMargin() As Long
    PageTopMargin = mlngPageTopMargin
End Property

Public Property Let PageRightMargin(ByVal vData As Long)
    mlngPageRightMargin = vData
End Property

Public Property Get PageRightMargin() As Long
    PageRightMargin = mlngPageRightMargin
End Property

Public Property Let PageLeftMargin(ByVal vData As Long)
    mlngPageLeftMargin = vData
End Property

Public Property Get PageLeftMargin() As Long
    PageLeftMargin = mlngPageLeftMargin
End Property

Public Property Let PaperSize(ByVal vData As Long)
    mlngPaperSize = vData
End Property

Public Property Get PaperSize() As Long
    PaperSize = mlngPaperSize
End Property

Public Property Let PrinterName(ByVal vData As String)
    mstrPrinterName = vData
End Property

Public Property Get PrinterName() As String
    PrinterName = mstrPrinterName
End Property

Public Property Let strFileName(ByVal vData As String)
    mstrFileName = vData
End Property

Public Property Get strFileName() As String
    strFileName = mstrFileName
End Property

Public Property Get Zoom() As Integer
    Zoom = mintZoom
End Property

Public Property Let Zoom(ByVal vData As Integer)
    If vData <> -1 And vData <> -2 And (vData < 10 Or vData > 800) Then
        mintZoom = 100    ' Set default zoom to 100%
    Else
        mintZoom = vData
    End If
End Property

Public Sub RunCombination(ByRef PreviewForm As Object, _
                          ByVal GenerateReportMode As GenerateReportModeConstant, _
                          ByVal PreviewCaption As String, _
                          ParamArray Reports() As Variant)
    
    Dim lngReportsCnt As Long
    Dim lngReportsCtr As Long
    Dim lngCombinedPageIndexCtr As Long
    Dim lngReportPageCnt As Long
    Dim lngReportPageCtr As Long
    Dim lngReportDim As Long
    Dim lngReportDimCtr As Long
    
    Dim c As New cCommonDialog
    Dim lngCtr As Long
    
    Screen.MousePointer = vbHourglass
    
    lngCombinedPageIndexCtr = 0
    
    lngReportDim = UBound(Reports)
    
    ' Run the reports to determine how many pages each one has
    For lngReportDimCtr = 0 To lngReportDim
        lngReportsCnt = UBound(Reports(lngReportDimCtr))
        
        For lngReportsCtr = 0 To lngReportsCnt
            
            If Not (lngReportDimCtr = 0 And lngReportsCtr = 0) Then
                On Error Resume Next
                Reports(lngReportDimCtr)(lngReportsCtr).blnDoNotShowMsg = True
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
            End If
            
            Reports(lngReportDimCtr)(lngReportsCtr).Run
            
            lngReportPageCnt = Reports(lngReportDimCtr)(lngReportsCtr).Pages.Count
            
            ' Start adding pages to the combined report from each of the passed reports
            For lngReportPageCtr = 0 To lngReportPageCnt - 1
                RCombinedReports.Pages.Insert lngCombinedPageIndexCtr, Reports(lngReportDimCtr)(lngReportsCtr).Pages(lngReportPageCtr)
                
                lngCombinedPageIndexCtr = lngCombinedPageIndexCtr + 1
            Next
        Next
    Next
    
    RCombinedReports.Pages.Commit
    
    Screen.MousePointer = vbDefault
    
    Set PreviewForm.m_objReload = RCombinedReports
    With RCombinedReports
        
        .PageSettings.BottomMargin = mlngPageBottomMargin
        .PageSettings.TopMargin = mlngPageTopMargin
        .PageSettings.RightMargin = mlngPageRightMargin
        .PageSettings.LeftMargin = mlngPageLeftMargin
        
        
        Select Case GenerateReportMode
            Case GenerateReportModeConstant.[Show Report Modal]
            
                On Error Resume Next
                If LenB(Trim$(mstrPrinterName)) > 0 Then
                    .Printer.DeviceName = mstrPrinterName
                End If
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                
                If .Printer.DeviceName <> "" Then
                    .Printer.PaperSize = mlngPaperSize
                End If
                .Show vbModal
                
            Case GenerateReportModeConstant.[Show Report Non-Modal]
            
                On Error Resume Next
                If LenB(Trim$(mstrPrinterName)) > 0 Then
                    .Printer.DeviceName = mstrPrinterName
                End If
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                
                If .Printer.DeviceName <> "" Then
                    .Printer.PaperSize = mlngPaperSize
                End If
                .Show
                
            Case GenerateReportModeConstant.[Generate Preview]
            
                On Error Resume Next
                .Printer.DeviceName = mstrPrinterName
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                
                If .Printer.DeviceName <> "" Then
                    .Printer.PaperSize = mlngPaperSize
                End If
                
                With PreviewForm
                    .Caption = PreviewCaption
                    .RunReport RCombinedReports
                    
                    .arv.TOCEnabled = False
                    
                    If mintZoom <> -1 And mintZoom <> -2 And (mintZoom < 10 Or mintZoom > 800) Then
                        .arv.Zoom = 100
                    Else
                        .arv.Zoom = mintZoom
                    End If
                    
                    .Show vbModal
                End With
                
                Set PreviewForm = Nothing
                
            Case GenerateReportModeConstant.[Generate Print With Dialog]
                
                On Error Resume Next
                If LenB(Trim$(mstrPrinterName)) > 0 Then
                    .Printer.DeviceName = mstrPrinterName
                End If
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                .Run
                If .Printer.DeviceName <> "" Then
                    .Printer.PaperSize = mlngPaperSize
                    .Printer.FromPage = 1
                    .Printer.ToPage = .Pages.Count
                End If
                
                With c
                    .DialogTitle = "Choose Printer"
                    .hwnd = RCombinedReports.hwnd
                    .PrinterDefault = True
                    .Object = Printer
                    .FromPage = RCombinedReports.Printer.FromPage
                    .ToPage = RCombinedReports.Pages.Count
                    .ShowPrinter
                End With
                RCombinedReports.Printer.DeviceName = c.PrinterName
                RCombinedReports.Printer.FromPage = c.FromPage
                RCombinedReports.Printer.ToPage = c.ToPage
                RCombinedReports.Printer.Copies = c.Copies
                If c.GetAPIReturn <> 0 Then
                    For lngCtr = 1 To c.Copies
                        RCombinedReports.PrintReport False
                        If InStr(1, RCombinedReports.Printer.DeviceName, "pdf") > 0 Then Exit For
                    Next lngCtr
                    RCombinedReports.Printer.FromPage = 1
                End If
                
                '.PrintReport True
                
            Case GenerateReportModeConstant.[Generate Print Without Dialog]
            
                On Error Resume Next
                If LenB(Trim$(mstrPrinterName)) > 0 Then
                    .Printer.DeviceName = mstrPrinterName
                End If
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                
                If .Printer.DeviceName <> "" Then
                    .Printer.PaperSize = mlngPaperSize
                    .Printer.FromPage = 1
                    .Printer.ToPage = .Pages.Count
                    .Printer.DisplayProgressDialog = False
                End If
                .PrintReport False
                
            Case GenerateReportModeConstant.[Generate PDF Export Pages]
                If .Printer.DeviceName <> "" Then
                    .Printer.PaperSize = mlngPaperSize
                End If
                .Run
                
                AutomaticSaveReport RCombinedReports, mstrFileName, sbpPDF
        End Select
    End With
    
    ' Unload all reports
    For lngReportDimCtr = 0 To lngReportDim
        lngReportsCnt = UBound(Reports(lngReportDimCtr))
        
        For lngReportsCtr = 0 To lngReportsCnt
            Unload Reports(lngReportDimCtr)(lngReportsCtr)
            
            Set Reports(lngReportDimCtr)(lngReportsCtr) = Nothing
        Next
    Next
    
    ' Unload combined report
    Unload RCombinedReports
    
    Set RCombinedReports = Nothing
End Sub

Private Sub AutomaticSaveReport(ByRef rptReport As ActiveReport, _
                                ByVal strFileName As String, _
                                ByVal FileType As ExportFilterType)
                                'ByRef strCaptions() As String
    
    Dim ExportFilter As Object
    
'    Dim intFreeFile As Integer
'    Dim intUBoundCaptions As Integer
'    Dim intPageCtr As Integer
    
    Select Case FileType
        Case sbpHTML
            Set ExportFilter = New ActiveReportsHTMLExport.HTMLexport
            
            ExportFilter.Filename = IIf(UCase(Right(strFileName, 5)) = ".HTML", strFileName, strFileName & ".html")
            rptReport.Export ExportFilter
        Case sbpPDF
            Set ExportFilter = New ActiveReportsPDFExport.ARExportPDF
            
            ExportFilter.JPGQuality = 0
            ExportFilter.Filename = IIf(UCase(Right(strFileName, 4)) = ".PDF", strFileName, strFileName & ".pdf")
            rptReport.Export ExportFilter
            
        Case sbpText
            Set ExportFilter = New ActiveReportsTextExport.ARExportText
            
            ExportFilter.Filename = IIf(UCase(Right(strFileName, 4)) = ".TXT", strFileName, strFileName & ".txt")
            rptReport.Export ExportFilter
            
            'intFreeFile = FreeFile()''
            
            'Open strFileName For Output As #intFreeFile'
            
            'intUBoundCaptions = UBound(strCaptions)
            
            'For intPageCtr = 1 To intUBoundCaptions
            '    Print #intFreeFile, strCaptions(intPageCtr)
            '    Print #intFreeFile,
            'Next
            
            'Close #intFreeFile
    End Select
    
    Set ExportFilter = Nothing
End Sub
