VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEntrepotSendError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrepot Validation"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   Icon            =   "frmEntrepotSendError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2210"
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxValidations 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   3
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowUserResizing=   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Tag             =   "2217"
      Top             =   4440
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwValidate 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmEntrepotSendError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub FormatListView()


    lvwValidate.ColumnHeaders.Add , , "Header/Detail", 1185
    lvwValidate.ColumnHeaders.Add , , "Description", 5670
    
        
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub MyLoad2(ByVal cpiSendValidations As cSendValidations, ByVal MyResourceHandler As Long)
    'Dim m_varSendValidations As cSendValidations
    'Dim x As cSendValidations
    'Dim y As cSendValidation
    Dim i As Integer
    'Set x = New cSendValidations
    
    'Set m_varSendValidations = cpiSendValidations
    
    'x.Add "D1.1", "T6", "Requested quantity/weight exceeds available quantity/weight.", ""
    'x.Add "D1.2", "T3", "Selected Stock Card and Stock Card Number does not match.", ""
    'x.Add "D1.3", "", "No selected stock card.", "Re-select stock card from Available Stocks picklist in box L1 or change countries in affected boxes."
    'x.Add "D1.4", "C1/C2", "Country of Origin and/or Export does not match with the selected stock card.", ""
    
    ResourceHandler = MyResourceHandler
    modGlobals.LoadResStrings Me, True
    
    flxValidations.Row = 0
    
    flxValidations.Rows = cpiSendValidations.Count + 1
    
    'format column widths
    flxValidations.ColWidth(0, 0) = 500
    flxValidations.ColWidth(1, 0) = 800
    flxValidations.ColWidth(2, 0) = 4000
    flxValidations.ColWidth(3, 0) = 4000
    
    
    
    flxValidations.Col = 0
    flxValidations.ColHeaderCaption(0, 0) = "H/D"
    
    flxValidations.Col = 1
    flxValidations.ColHeaderCaption(0, 1) = "Box Code"
    
    flxValidations.Col = 2
    flxValidations.ColHeaderCaption(0, 2) = "Error Description"
    
    flxValidations.Col = 3
    flxValidations.ColHeaderCaption(0, 3) = "Proposed Solution"
    
    
    For i = 1 To cpiSendValidations.Count
        flxValidations.RowHeight(i) = 600
        
        flxValidations.Row = i
        flxValidations.Col = 0
        
        flxValidations.Text = cpiSendValidations.Item(i).HeaderDetailNum
        
        
        flxValidations.Row = i
        flxValidations.Col = 1
        
        flxValidations.Text = cpiSendValidations.Item(i).BoxCode
        
        flxValidations.Row = i
        flxValidations.Col = 2
        
        flxValidations.Text = cpiSendValidations.Item(i).Description
        
        flxValidations.Row = i
        flxValidations.Col = 3
        
        flxValidations.Text = cpiSendValidations.Item(i).Solution
        
        
    Next
    Me.Show
    'Set x = Nothing
    
End Sub


