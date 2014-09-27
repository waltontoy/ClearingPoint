VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmShowFields 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Show Fields"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "frmShowFields.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin GridEX16.GridEX GridEX1 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9128
      TabKeyBehavior  =   1
      MethodHoldFields=   -1  'True
      Options         =   -1
      RecordsetType   =   1
      AutomaticArrange=   0   'False
      GroupByBoxVisible=   0   'False
      ColumnCount     =   3
      CardCaption1    =   -1  'True
      ColHeaderAlignment1=   1
      ColKey1         =   "Checked"
      ColWidth1       =   300
      ColumnType1     =   3
      ColEditType1    =   2
      ColCaption2     =   "Fields"
      ColHeaderAlignment2=   1
      ColKey2         =   "Fields"
      ColWidth2       =   1005
      ColCaption3     =   "Description"
      ColHeaderAlignment3=   1
      ColKey3         =   "Field Desc"
      ColWidth3       =   3495
      DataMode        =   1
      GridLines       =   0   'False
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Tag             =   "614"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Index           =   0
      Left            =   6000
      TabIndex        =   3
      Tag             =   "614"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Clear Selection"
         Height          =   345
         Index           =   1
         Left            =   5880
         TabIndex        =   2
         Tag             =   "614"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select All"
         Height          =   345
         Index           =   0
         Left            =   5880
         TabIndex        =   1
         Tag             =   "614"
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmShowFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strFieldShown() As String
Private strFieldToList() As String
Private strDescToList() As String
Private lngFieldCounter As Long

Private Enum enuDocType
    eDocAny = 0
    edocimport = 1
    eDocExport = 2
    eDocOTS = 3
    eDocNCTS = 4
    edoccombined = 5
    eDocEDIDepartures = 6
    eDocEDIARRIVALS = 7
    eDocPLDAImport = 8
    eDocPLDACombined = 9
End Enum

Private rstOfflineRec As ADODB.Recordset

Private Sub cmdOK_Click(Index As Integer)
    
    Select Case Index
        Case 0 'OK
            rstOfflineRec.MoveFirst
            strFields = ""
            Do While Not rstOfflineRec.EOF
                If rstOfflineRec.Fields("Checked") = -1 Then
                    strFields = strFields & "*" & rstOfflineRec.Fields("Fields").Value
                End If
                rstOfflineRec.MoveNext
            Loop
            Unload Me
            'frm_Find.RefreshList
            
        Case 1 'Cancel
            Unload Me
    End Select
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim lngCounter As Long
    Select Case Index
        Case 0 'Select All
            For lngCounter = 6 To GridEX1.ADORecordset.RecordCount
                GridEX1.Row = lngCounter
                GridEX1.Value(GridEX1.Columns(1).Index) = -1
                Call GridEX1_Change
            Next lngCounter
        Case 1 'Clear Selection
            For lngCounter = 6 To GridEX1.ADORecordset.RecordCount
                GridEX1.Row = lngCounter
                GridEX1.Value(GridEX1.Columns(1).Index) = 0
                Call GridEX1_Change
            Next lngCounter
    End Select
End Sub

Private Sub Form_Load()
    Dim lngCounter As Long
    Dim lngCounter2 As Long
    
    Set rstOfflineRec = New ADODB.Recordset
    
    UpdateBoxList
    
    strFields = ""
    lngFieldCounter = 0
    
    With rstOfflineRec
        .CursorLocation = adUseClient
        .Fields.Append "Checked", adVarChar, 10
        .Fields.Append "Fields", adVarChar, 50
        .Fields.Append "Description", adVarChar, 100
        .Open
        
        .AddNew "Fields", Translate(496)
        .Update
        .AddNew "Fields", Translate(635)
        .Update
        .AddNew "Fields", Translate(625)
        .Update
        .AddNew "Fields", Left(Translate(272), Len(Translate(272)) - 1)
        .Update
        .AddNew "Fields", Translate(611)
        .Update
        
        If Not (Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIARRIVALS Or _
           Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIDepartures Or _
           Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocNCTS Or _
           Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocAny) Then
            .AddNew "Fields", Translate(437)
            .Update
            lngFieldCounter = lngFieldCounter + 1
        End If
        
        .AddNew "Fields", Translate(713)
        .Update
        .AddNew "Fields", Translate(715)
        .Update
        .AddNew "Fields", Translate(742)
        .Update
        
        If Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIARRIVALS Or _
           Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIDepartures Then
            .AddNew "Fields", "Date Last Received"
            .Update
            lngFieldCounter = lngFieldCounter + 1
        End If
        
        .AddNew "Fields", "Date Printed"
        .Update
        .AddNew "Fields", "LogID Description"
        .Update
        .AddNew "Fields", "Error String"
        .Update
        .AddNew "Fields", Translate(423)
        .Update
        
        If Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIARRIVALS Or _
           Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocEDIDepartures Or _
           Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocNCTS Or _
           Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.edoccombined Or _
           Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocPLDAImport Or _
           Right(frm_Find.icbType.SelectedItem.Key, 1) = enuDocType.eDocPLDACombined Then
            .AddNew "Fields", "MRN"
            .Update
            lngFieldCounter = lngFieldCounter + 1
        End If
        
        For lngCounter = 0 To UBound(strFieldToList)
            If Not (UBound(strFieldToList) = 0 And LBound(strFieldToList) = 0 And _
                strFieldToList(0) = "") Then
                .AddNew "Fields", strFieldToList(lngCounter)
                .Fields("Description") = strDescToList(lngCounter)
                .Update
            End If
        Next lngCounter
        
        .MoveFirst
        Do While Not .EOF
            .Fields("Checked") = 0
            .Update
            .MoveNext
        Loop
        
        For lngCounter = 1 To UBound(strFieldShown)
            .MoveFirst
            Do While Not .EOF
                If strFieldShown(lngCounter) = .Fields("Fields").Value Then
                    .Fields("Checked").Value = -1
                    .Update
                    Exit Do
                End If
                .MoveNext
            Loop
        Next lngCounter
    End With
    
    Set GridEX1.ADORecordset = rstOfflineRec
    
    GridEX1.Columns(1).ColumnType = jgexCheckBox
    GridEX1.Columns(1).Caption = ""
    GridEX1.Columns(1).Width = 300
    GridEX1.Columns(2).Caption = "Fields"
    GridEX1.Columns(2).Width = 1300
    GridEX1.Columns(2).EditType = jgexEditNone
    GridEX1.Columns(3).Caption = "Description"
    GridEX1.Columns(3).Width = 3600
    GridEX1.Columns(3).EditType = jgexEditNone
    GridEX1.SelectionStyle = jgexEntireRow
    GridEX1.Columns(2).Selectable = False
    GridEX1.Columns(3).Selectable = False
End Sub

Public Sub PreLoad(ByRef strFields() As String)
    strFieldShown = strFields
    Me.Show vbModal
End Sub

Private Sub UpdateBoxList()
    Dim rstBoxDefault As ADODB.Recordset
    Dim strQueryAll As String
    Dim lngCounter As Long
    
    strQueryAll = "Select [BOX DEFAULT IMPORT ADMIN].[BOX CODE], "
    strQueryAll = strQueryAll & "[BOX DEFAULT IMPORT ADMIN].[" & cLanguage & " DESCRIPTION] as Description, "
    strQueryAll = strQueryAll & "[BOX DEFAULT TRANSIT NCTS ADMIN].[" & cLanguage & " DESCRIPTION] as Description1 FROM "
    strQueryAll = strQueryAll & "[BOX DEFAULT IMPORT ADMIN] INNER JOIN ([BOX DEFAULT "
    strQueryAll = strQueryAll & "EXPORT ADMIN] INNER JOIN ([BOX DEFAULT TRANSIT ADMIN] "
    strQueryAll = strQueryAll & "INNER JOIN ([BOX DEFAULT TRANSIT NCTS ADMIN] "
    strQueryAll = strQueryAll & "INNER JOIN ([BOX DEFAULT COMBINED NCTS ADMIN] "
    strQueryAll = strQueryAll & "INNER JOIN ([BOX DEFAULT EDI NCTS ADMIN] "
    strQueryAll = strQueryAll & "INNER JOIN ([BOX DEFAULT EDI NCTS2 ADMIN] "
    strQueryAll = strQueryAll & "INNER JOIN ([BOX DEFAULT EDI NCTS IE44 ADMIN] "
    strQueryAll = strQueryAll & "INNER JOIN ([BOX DEFAULT PLDA IMPORT ADMIN] "
    strQueryAll = strQueryAll & "INNER JOIN [BOX DEFAULT PLDA COMBINED ADMIN] "
    strQueryAll = strQueryAll & "ON [BOX DEFAULT PLDA COMBINED ADMIN].[BOX CODE] = [BOX DEFAULT PLDA IMPORT ADMIN].[BOX CODE]) "
    strQueryAll = strQueryAll & "ON [BOX DEFAULT PLDA IMPORT ADMIN].[BOX CODE] = [BOX DEFAULT EDI NCTS IE44 ADMIN].[BOX CODE]) "
    strQueryAll = strQueryAll & "ON [BOX DEFAULT EDI NCTS IE44 ADMIN].[BOX CODE] = [BOX DEFAULT EDI NCTS2 ADMIN].[BOX CODE]) "
    strQueryAll = strQueryAll & "ON [BOX DEFAULT EDI NCTS ADMIN].[BOX CODE] = [BOX DEFAULT EDI NCTS2 ADMIN].[BOX CODE]) "
    strQueryAll = strQueryAll & "ON [BOX DEFAULT COMBINED NCTS ADMIN].[BOX CODE] = [BOX DEFAULT EDI NCTS ADMIN].[BOX CODE]) "
    strQueryAll = strQueryAll & "ON [BOX DEFAULT TRANSIT NCTS ADMIN].[BOX CODE] = [BOX DEFAULT COMBINED NCTS ADMIN].[BOX CODE]) "
    strQueryAll = strQueryAll & "ON [BOX DEFAULT TRANSIT ADMIN].[BOX CODE] = [BOX DEFAULT TRANSIT NCTS ADMIN].[BOX CODE]) "
    strQueryAll = strQueryAll & "ON [BOX DEFAULT EXPORT ADMIN].[BOX CODE] = [BOX DEFAULT TRANSIT ADMIN].[BOX CODE]) "
    strQueryAll = strQueryAll & "ON [BOX DEFAULT EXPORT ADMIN].[BOX CODE] = [BOX DEFAULT IMPORT ADMIN].[BOX CODE] "

    Select Case frm_Find.icbType.SelectedItem.Key
        Case "D" & enuDocType.edocimport
            
            ADORecordsetOpen "Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT IMPORT ADMIN] order by [BOX CODE] ", _
                                g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT IMPORT ADMIN] order by [BOX CODE] ")
            
        Case "D" & enuDocType.eDocExport
            ADORecordsetOpen "Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT EXPORT ADMIN] order by [BOX CODE] ", _
                                g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT EXPORT ADMIN] order by [BOX CODE] ")
        
        Case "D" & enuDocType.eDocOTS
            ADORecordsetOpen "Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT TRANSIT ADMIN] order by [BOX CODE] ", _
                                g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
                                
            'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT TRANSIT ADMIN] order by [BOX CODE] ")
        
        Case "D" & enuDocType.eDocNCTS
            ADORecordsetOpen "Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT TRANSIT NCTS ADMIN] order by [BOX CODE] ", _
                                g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT TRANSIT NCTS ADMIN] order by [BOX CODE] ")
        
        Case "D" & enuDocType.edoccombined
            ADORecordsetOpen "Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT COMBINED NCTS ADMIN] order by [BOX CODE] ", _
                                g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT COMBINED NCTS ADMIN] order by [BOX CODE] ")
        
        Case "D" & enuDocType.eDocEDIDepartures
            ADORecordsetOpen "Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT EDI NCTS ADMIN] order by [BOX CODE] ", _
                                g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT EDI NCTS ADMIN] order by [BOX CODE] ")
        
        Case "D" & enuDocType.eDocEDIARRIVALS
            ADORecordsetOpen "Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT EDI NCTS2 ADMIN] order by [BOX CODE] ", _
                                g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT EDI NCTS2 ADMIN] order by [BOX CODE] ")
        
        Case "D" & enuDocType.eDocPLDAImport
            ADORecordsetOpen "Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT PLDA IMPORT ADMIN] order by [BOX CODE] ", _
                                g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT PLDA IMPORT ADMIN] order by [BOX CODE] ")
        
        Case "D" & enuDocType.eDocPLDACombined
            ADORecordsetOpen "Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT PLDA COMBINED ADMIN] order by [BOX CODE] ", _
                                g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datSADBEL.OpenRecordset("Select [BOX CODE], [" & cLanguage & " DESCRIPTION] as Description from [BOX DEFAULT PLDA COMBINED ADMIN] order by [BOX CODE] ")
        
        Case "D" & enuDocType.eDocAny
            ADORecordsetOpen strQueryAll, g_conSADBEL, rstBoxDefault, adOpenKeyset, adLockOptimistic
            'Set rstBoxDefault = datSADBEL.OpenRecordset(strQueryAll)
    
    End Select
    
    With rstBoxDefault
        If Not (.BOF And .EOF) Then
            .MoveFirst
            lngCounter = 0
            Do While Not .EOF
                ReDim Preserve strFieldToList(lngCounter)
                ReDim Preserve strDescToList(lngCounter)
            
                strFieldToList(lngCounter) = .Fields("Box Code").Value
                
                If frm_Find.icbType.SelectedItem.Key = "D" & enuDocType.eDocAny Then
                    strDescToList(lngCounter) = .Fields("Description").Value & "/" & .Fields("Description1").Value
                Else
                    strDescToList(lngCounter) = .Fields("Description").Value
                End If
            
                .MoveNext
                lngCounter = lngCounter + 1
            Loop
        Else
            ReDim Preserve strFieldToList(0)
            ReDim Preserve strDescToList(0)
            strFieldToList(0) = ""
            strDescToList(0) = ""
        End If
    End With
    
    rstBoxDefault.Close
    
    Set rstBoxDefault = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ADODisconnectDB rstOfflineRec
End Sub

Private Sub GridEX1_Change()
    If GridEX1.Value(GridEX1.Columns(2).Index) = Translate(496) Or _
       GridEX1.Value(GridEX1.Columns(2).Index) = Translate(635) Or _
       GridEX1.Value(GridEX1.Columns(2).Index) = Translate(625) Or _
       GridEX1.Value(GridEX1.Columns(2).Index) = Left(Translate(272), Len(Translate(272)) - 1) Or _
       GridEX1.Value(GridEX1.Columns(2).Index) = Translate(611) Then
       GridEX1.Value(GridEX1.Columns(1).Index) = -1
       Exit Sub
    End If

    rstOfflineRec.MoveFirst
    rstOfflineRec.Find "Fields = '" & GridEX1.Value(GridEX1.Columns(2).Index) & "'"
    rstOfflineRec.Fields("Checked").Value = GridEX1.Value(GridEX1.Columns(1).Index)
    rstOfflineRec.Update
End Sub
