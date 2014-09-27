VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_picklist 
   Caption         =   "398"
   ClientHeight    =   5130
   ClientLeft      =   5160
   ClientTop       =   3060
   ClientWidth     =   4575
   ClipControls    =   0   'False
   Icon            =   "frm_picklist.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Tag             =   "398"
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -285
      Top             =   4815
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from import"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   345
      Index           =   1
      Left            =   3510
      TabIndex        =   4
      Tag             =   "179"
      Top             =   4665
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
      Height          =   345
      Index           =   0
      Left            =   2475
      TabIndex        =   3
      Tag             =   "426"
      Top             =   4665
      Width           =   960
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   1300
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   1455
      TabIndex        =   2
      Text            =   " "
      Top             =   120
      Width           =   3000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4110
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7250
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).BandIndent=   2
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm_picklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_conData As ADODB.Connection    'Public datData As DAO.Database
Public m_conSADBEL As ADODB.Connection  'Public datSADBEL As DAO.Database
Public m_rstTemp As ADODB.Recordset

Private m_rstBoxDefa As ADODB.Recordset
Private m_rstA As ADODB.Recordset
Private m_rstB As ADODB.Recordset
Private m_rstHFlex As ADODB.Recordset

Private nDataCtr As Integer
Private nAB As Integer
Private cDocument As String
Private cSelVal As String
Private bCode As String

Private LastVal As String
Private A As String
Private B As String
Private docName As String
Private TreeID As String
Private lRelate As Boolean
Private Ltaposna As Boolean
Private LData As Boolean
Private blnOK As Boolean

Private mstrCallingForm As String

Private m_CallingForm As Object

Private Sub Command1_Click(Index As Integer)
    With MSHFlexGrid1
        If Index = 0 And .Row <> 0 Then
            blnOK = True
            If nDataCtr > 0 Then
                If bCode = "L1" Then
                    cSelVal = .TextMatrix(.Row, 0) + "@@"
                    
                    If lRelate Then
                        cSelVal = .TextMatrix(.Row, 0) + "@@" + .TextMatrix(.Row, 1)
                    Else
                        cSelVal = .TextMatrix(.Row, 0) + "@@!"
                    End If
                Else
                    Select Case Me.CallingForm
                        Case "frm_licensee", "DV1"
                            cSelVal = .TextMatrix(.Row, 0) & "@@" & .TextMatrix(.Row, 1)
                        Case Else
                            cSelVal = .TextMatrix(.Row, 0)
                    End Select
                End If
            End If
        Else
            If bCode = "L1" Then
                cSelVal = LastVal + "@@"
            Else
                Select Case Me.CallingForm
                    Case "frm_licensee", "DV1"
                        cSelVal = "0"
                    Case Else
                        cSelVal = LastVal
                End Select
            End If
        End If
    End With
    
    Unload Me
End Sub

Private Sub Command1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cSelVal = LastVal
        Unload Me
        
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    Dim cWiddth As String
    
    Screen.MousePointer = vbDefault
    Me.Refresh
    
    If Ltaposna Then
        
        If Len(LastVal) <= 0 Then
            Set m_rstHFlex = m_rstB
            'Adodc1.RecordSource = B
            nAB = 2
        Else
            Set m_rstHFlex = m_rstA
            'Adodc1.RecordSource = A
            nAB = 1
        End If
        
        Ltaposna = False
        
        'Adodc1.Refresh
        
        Set MSHFlexGrid1.DataSource = m_rstHFlex
        'Set MSHFlexGrid1.DataSource = Adodc1
        
        If m_rstHFlex.EOF And m_rstHFlex.BOF Then
        'If Adodc1.Recordset.BOF And Adodc1.Recordset.EOF Then
        
            nDataCtr = 0
            cWiddth = "1"
        Else
            m_rstHFlex.MoveFirst
            
            nDataCtr = 1
            cWiddth = CStr(m_rstHFlex!Width)
        End If
        
' ********** Repositioned February 28, 2002 **********
' ********** Moved from FormLoad to Form_Activate to allow for translation of picklist fields.
        Select Case Me.CallingForm
            Case "frm_licensee", "DV1"
                If m_rstHFlex.EOF And m_rstHFlex.BOF Then
                    Select Case bCode
                        Case "B7"
                            Me.Caption = " Zip Codes & Cities - " & Translate(398)
                        Case "C1"
                            Me.Caption = " Countries - " & Translate(398)
                        Case "VN"
                            Me.Caption = " " & Translate(1307) & " - " & Translate(398)
                    End Select
                Else
                    m_rstHFlex.MoveFirst
                    
                    Me.Caption = " " & m_rstHFlex.Fields("PICKDESC").Value & " - " & Translate(398)
                End If
            Case Else
                Me.Caption = " " & bCode & " - " & Translate(398) 'Picklist (" + Proper(cDocument) + " - " + Proper(LngeUsed) + ")"
        End Select
' ********** End Reposition **************************
        
        If nDataCtr > 0 Then
            Command1(0).Enabled = True
        Else
            Command1(0).Enabled = False
            
'            If bCode = "L1" Then
'                Command2.Enabled = True
'            Else
            
            If bCode <> "L1" Then
                MsgBox Trim(Translate(427)), vbInformation
                Unload Me
                
                Exit Sub
            End If
        End If
        
        doInitCaption
        
        lRelate = False
        
        If bCode = "L1" Then
            With m_rstBoxDefa
                .Filter = adFilterNone
                .Filter = "[BOX CODE] = '" & bCode & "' "
                
                If Not (.EOF And .BOF) Then
                    .MoveFirst
                '.Index = "Box Code"
                '.Seek "=", bCode
                '
                'If Not .NoMatch Then
                    lRelate = ![Relate L1 To S1]
                End If
            End With
        End If
        
        Me.Refresh
        
        If Len(LastVal) > 0 Then
            If nDataCtr <= 0 And bCode = "L1" Then
                ' Do nothing.
            Else
                If nDataCtr > 0 And bCode = "L1" Then
                    Text5.Text = Left(LastVal, Len(LastVal) - 1)
                    SendKeysEx "{END}"
' ********** Commented November 23, 2000 **********
' ********** Redundant.  [SendKeysEx "{END}"] takes care of appending last character of key code.
'                    SendKeysEx Right(LastVal, 1)
' ********** End Comment **************************
                Else
                    LastVal = IIf(Len(LastVal) > CInt(cWiddth), Left(LastVal, CInt(cWiddth)), LastVal)
                    Text5.Text = LastVal
                    SendKeysEx "{END}"
                End If
                
                'SendKeysEx "{DOWN}"
                Text5.SetFocus
            End If
        Else
            Text6.SetFocus
        End If
        
        'Me.Refresh
        
        Screen.MousePointer = vbDefault
    End If
    
    MSHFlexGrid1.SetFocus
    
End Sub

Private Sub FormLoad()
    Dim cTag As String
    Dim cSheetInfo As String
    Dim nArrayCtr As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim cDocs As String
    Dim lngForeColor As Long
    Dim lngBackColor As Long
    Dim intLanguageIndex As Integer
    Dim strBoxProp As String
    Dim strLanguageUsed As String
    
    Ltaposna = True
    
    strBoxProp = GetSetting(App.Title, "Settings", "BoxProperty")
    
    If OpenMDB(g_objDataSourceProperties, Me, DBInstanceType_DATABASE_SADBEL) Then
        ' Just set m_conSADBEL.
    End If

    'Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cAppPath & "\mdb_sadbel.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=" & G_Main_Password
    
    If Len(Trim(strBoxProp)) > 0 Then
        j = CountChr(strBoxProp, "#*#")
        
        For i = 1 To j
            k = InStr(1, strBoxProp, "#*#")
            
            Select Case i
                Case 1
                    bCode = UCase(Mid(strBoxProp, 1, k - 1))
                Case 2
                    cSheetInfo = UCase(Mid(strBoxProp, 1, k - 1))
                Case 3
                    nArrayCtr = CInt(UCase(Mid(strBoxProp, 1, k - 1)))
                Case 4
                    strLanguageUsed = UCase(Mid(strBoxProp, 1, k - 1))
                Case 5    ' Import/Export/Transit
                    cDocument = Trim(UCase(Mid(strBoxProp, 1, k - 1)))
                Case 7
                    docName = Mid(strBoxProp, 1, k - 1)
                Case 8
                    TreeID = Mid(strBoxProp, 1, k - 1)
                Case 9
                    'IAN For multiple picklist 03-02-05
                    If Not g_blnMultiplePick Then
                        LastVal = Mid(strBoxProp, 1, k - 1)
                    Else
                        LastVal = Mid(Mid(strBoxProp, 1, k - 1), g_lngFrom)
                    End If
                Case 10
                    cTag = Mid(strBoxProp, 1, k - 1)
                Case 11
                    lngBackColor = CLng(Mid(strBoxProp, 1, k - 1))
                Case 12
                    lngForeColor = CLng(Mid(strBoxProp, 1, k - 1))
            End Select
            
            strBoxProp = Right(strBoxProp, Len(strBoxProp) - (k + 2))
        Next
        
        'Text6.ForeColor = lngForeColor: Text6.BackColor = lngBackColor
        
'-----> Added March 22, 2001
'-----> to skip this procedure if the formis called outside codisheet
        If frm_taricdetail.blnTaricDetail Or frm_taricmain.blnTaricMain Then GoTo SkipThis:
'-----> end add
        
        Select Case Me.CallingForm
            Case "frm_licensee", "DV1"
                ' Do nothing.
            Case Else
                If Left(cSheetInfo, 1) = Trim(Translate(710)) Then

                    Select Case UCase(cDocument)
                        Case "IMPORT"
                            If (newform(nArrayCtr) Is Nothing) = False Then
                                newform(nArrayCtr).Text1_GotFocus CInt(Mid(cSheetInfo, 2))
                            End If

                        Case "EXPORT"
                            If (newform(nArrayCtr) Is Nothing) = False Then
                                newformE(nArrayCtr).Text1_GotFocus CInt(Mid(cSheetInfo, 2))
                            End If
                        Case "TRANSIT"
                            If (newform(nArrayCtr) Is Nothing) = False Then
                                newformT(nArrayCtr).Text1_GotFocus CInt(Mid(cSheetInfo, 2))
                            End If
                        Case "TRANSIT NCTS"
                            m_CallingForm.FocusBox eGot_Focus, eTab_Header, CInt(Mid(cSheetInfo, 2))

                        Case "COMBINED NCTS"
                            m_CallingForm.FocusBox eGot_Focus, eTab_Header, CInt(Mid(cSheetInfo, 2))
                            
                        Case "EDI NCTS"
                            m_CallingForm.FocusBox eGot_Focus, eTab_Header, CInt(Mid(cSheetInfo, 2))
                    End Select
                Else
                    Select Case UCase(cDocument)
                        Case "IMPORT"
                        
                            newform(nArrayCtr).Text2_GotFocus CInt(Mid(cSheetInfo, 2))

                        Case "EXPORT"
                            newformE(nArrayCtr).Text2_GotFocus CInt(Mid(cSheetInfo, 2))

                        Case "TRANSIT"
                            newformT(nArrayCtr).Text2_GotFocus CInt(Mid(cSheetInfo, 2))

                        Case "TRANSIT NCTS"
                            m_CallingForm.FocusBox eGot_Focus, eTab_Detail, CInt(Mid(cSheetInfo, 2))

                        Case "COMBINED NCTS"
                            m_CallingForm.FocusBox eGot_Focus, eTab_Detail, CInt(Mid(cSheetInfo, 2))
                        
                        Case "EDI NCTS"
                            m_CallingForm.FocusBox eGot_Focus, eTab_Detail, CInt(Mid(cSheetInfo, 2))
                    End Select
                End If
        End Select
        
'-----> Added March 22, 2001
SkipThis:
'-----> end add
    End If
    
    Select Case Trim(UCase(cDocument))
        Case "IMPORT"
            If NetUse("BOX DEFAULT IMPORT ADMIN", Me, 1) Then
                Set m_rstBoxDefa = m_rstTemp
            End If
        Case "EXPORT"
            If NetUse("BOX DEFAULT EXPORT ADMIN", Me, 1) Then
                Set m_rstBoxDefa = m_rstTemp
            End If
        Case "TRANSIT"
            If NetUse("BOX DEFAULT TRANSIT ADMIN", Me, 1) Then
                Set m_rstBoxDefa = m_rstTemp
            End If
        'ncts
        Case "TRANSIT NCTS"
            If NetUse("BOX DEFAULT TRANSIT NCTS ADMIN", Me, 1) Then
                Set m_rstBoxDefa = m_rstTemp
            End If
        Case "COMBINED NCTS"
            If NetUse("BOX DEFAULT COMBINED NCTS ADMIN", Me, 1) Then
                Set m_rstBoxDefa = m_rstTemp
            End If
        Case "EDI NCTS"
            If NetUse("BOX DEFAULT EDI NCTS ADMIN", Me, 1) Then
                Set m_rstBoxDefa = m_rstTemp
            End If
    End Select
    
'    If bCode = "L1" Then Command2.Enabled = True Else Command2.Enabled = False
    
    If bCode = "L1" Then
        cSelVal = LastVal + "@@"
    Else
        Select Case Me.CallingForm
            Case "frm_licensee", "DV1"
                cSelVal = "0"
            Case Else
                cSelVal = LastVal
        End Select
    End If
    
    Select Case Left(strLanguageUsed, 1)
        Case "E"
            intLanguageIndex = 1
        Case "D", "N"
            intLanguageIndex = 2
        Case "F"
            intLanguageIndex = 3
        Case Else
            intLanguageIndex = 1
    End Select
    
    'edited for NCTS
    'If Left(cDocument, 1) = "I" Then cDocs = "Import" Else cDocs = "Export/Transit"
    Select Case UCase(cDocument)
        Case "IMPORT"
            cDocs = "Import"
        Case "EXPORT/TRANSIT", "EXPORT", "TRANSIT"
            cDocs = "Export/Transit"
        Case "TRANSIT NCTS"
            cDocs = "Transit NCTS"
        Case "COMBINED NCTS"
            cDocs = "Combined NCTS"
        Case "EDI NCTS"
            cDocs = "EDI NCTS"
    End Select
    
    Select Case intLanguageIndex
        Case 2   'dutch
            A = "SELECT [PICKLIST MAINTENANCE DUTCH].CODE," & _
                        "[PICKLIST MAINTENANCE DUTCH].[DESCRIPTION DUTCH]," & _
                        "[PICKLIST DEFINITION].[WIDTH]," & _
                        "[PICKLIST DEFINITION].[PICKLIST DESCRIPTION DUTCH] AS PICKDESC " & _
                        "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE DUTCH] " & _
                        "WHERE " & _
                        "([PICKLIST DEFINITION].[BOX CODE]= " & Chr(39) & ProcessQuotes(bCode) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[DOCUMENT]= " & Chr(39) & ProcessQuotes(cDocs) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE DUTCH].[internal code]) " & _
                        IIf(g_blnMultiplePick, " AND [Picklist Definition].[Internal Code] = '" & g_strInternalCode & "' ", "") & _
                        "ORDER BY [PICKLIST MAINTENANCE DUTCH].CODE"
            B = "SELECT [PICKLIST MAINTENANCE DUTCH].CODE," & _
                        "[PICKLIST MAINTENANCE DUTCH].[DESCRIPTION DUTCH]," & _
                        "[PICKLIST DEFINITION].[WIDTH]," & _
                        "[PICKLIST DEFINITION].[PICKLIST DESCRIPTION DUTCH] AS PICKDESC " & _
                        "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE DUTCH] " & _
                        "WHERE " & _
                        "([PICKLIST DEFINITION].[BOX CODE]= " & Chr(39) & ProcessQuotes(bCode) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[DOCUMENT]= " & Chr(39) & ProcessQuotes(cDocs) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE DUTCH].[internal code]) " & _
                        IIf(g_blnMultiplePick, " AND [Picklist Definition].[Internal Code] = '" & g_strInternalCode & "' ", "") & _
                        "ORDER BY [PICKLIST MAINTENANCE DUTCH].[DESCRIPTION DUTCH]"
        Case 3   'french
            A = "SELECT [PICKLIST MAINTENANCE FRENCH].CODE," & _
                        "[PICKLIST MAINTENANCE FRENCH].[DESCRIPTION FRENCH]," & _
                        "[PICKLIST DEFINITION].[WIDTH]," & _
                        "[PICKLIST DEFINITION].[PICKLIST DESCRIPTION FRENCH] AS PICKDESC " & _
                        "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE FRENCH] " & _
                        "WHERE " & _
                        "([PICKLIST DEFINITION].[BOX CODE]= " & Chr(39) & ProcessQuotes(bCode) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[DOCUMENT]= " & Chr(39) & ProcessQuotes(cDocs) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE FRENCH].[internal code]) " & _
                        IIf(g_blnMultiplePick, " AND [Picklist Definition].[Internal Code] = '" & g_strInternalCode & "' ", "") & _
                        "ORDER BY [PICKLIST MAINTENANCE FRENCH].CODE"
            B = "SELECT [PICKLIST MAINTENANCE FRENCH].CODE," & _
                        "[PICKLIST MAINTENANCE FRENCH].[DESCRIPTION FRENCH]," & _
                        "[PICKLIST DEFINITION].[WIDTH]," & _
                        "[PICKLIST DEFINITION].[PICKLIST DESCRIPTION FRENCH] AS PICKDESC " & _
                        "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE FRENCH] " & _
                        "WHERE " & _
                        "([PICKLIST DEFINITION].[BOX CODE]= " & Chr(39) & ProcessQuotes(bCode) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[DOCUMENT]= " & Chr(39) & ProcessQuotes(cDocs) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE FRENCH].[internal code]) " & _
                        IIf(g_blnMultiplePick, " AND [Picklist Definition].[Internal Code] = '" & g_strInternalCode & "' ", "") & _
                        "ORDER BY [PICKLIST MAINTENANCE FRENCH].[DESCRIPTION FRENCH]"
        Case Else   'english
            A = "SELECT [PICKLIST MAINTENANCE ENGLISH].CODE," & _
                        "[PICKLIST MAINTENANCE ENGLISH].[DESCRIPTION ENGLISH]," & _
                        "[PICKLIST DEFINITION].[WIDTH]," & _
                        "[PICKLIST DEFINITION].[PICKLIST DESCRIPTION ENGLISH] AS PICKDESC " & _
                        "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE ENGLISH] " & _
                        "WHERE " & _
                        "([PICKLIST DEFINITION].[BOX CODE]= " & Chr(39) & ProcessQuotes(bCode) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[DOCUMENT]= " & Chr(39) & ProcessQuotes(cDocs) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE ENGLISH].[internal code]) " & _
                        IIf(g_blnMultiplePick, " AND [Picklist Definition].[Internal Code] = '" & g_strInternalCode & "' ", "") & _
                        "ORDER BY [PICKLIST MAINTENANCE ENGLISH].CODE"
            B = "SELECT [PICKLIST MAINTENANCE ENGLISH].CODE," & _
                        "[PICKLIST MAINTENANCE ENGLISH].[DESCRIPTION ENGLISH]," & _
                        "[PICKLIST DEFINITION].[WIDTH]," & _
                        "[PICKLIST DEFINITION].[PICKLIST DESCRIPTION ENGLISH] AS PICKDESC " & _
                        "FROM [PICKLIST DEFINITION],[PICKLIST MAINTENANCE ENGLISH] " & _
                        "WHERE " & _
                        "([PICKLIST DEFINITION].[BOX CODE]= " & Chr(39) & ProcessQuotes(bCode) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[DOCUMENT]= " & Chr(39) & ProcessQuotes(cDocs) & Chr(39) & ") AND " & _
                        "([PICKLIST DEFINITION].[internal code] = [PICKLIST MAINTENANCE ENGLISH].[internal code]) " & _
                        IIf(g_blnMultiplePick, " AND [Picklist Definition].[Internal Code] = '" & g_strInternalCode & "' ", "") & _
                        "ORDER BY [PICKLIST MAINTENANCE ENGLISH].[DESCRIPTION ENGLISH]"
    End Select
    
    ADORecordsetOpen A, m_conSADBEL, m_rstA, adOpenKeyset, adLockOptimistic
    ADORecordsetOpen B, m_conSADBEL, m_rstB, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Form_Load()
    Ltaposna = True
    
    Call LoadResStrings(Me, True)

    ' by jason 06-APR-2003 13:38 :-)
    Me.Caption = ""
    FormLoad

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Command1_Click (1)
        
        Exit Sub
    End If
End Sub

Private Sub Form_Resize()
    Dim nTop As Long
    
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    MSHFlexGrid1.Move 0.026 * ScaleWidth, 0.088 * ScaleHeight, 0.948 * ScaleWidth, 0.801 * ScaleHeight
    MSHFlexGrid1.ColWidth(0) = MSHFlexGrid1.Width * (1 / 4)
    MSHFlexGrid1.ColWidth(1) = MSHFlexGrid1.Width * (3 / 4)
    
    Text5.Left = MSHFlexGrid1.Left
    Text5.Width = MSHFlexGrid1.Width * (1 / 4)
    Text6.Left = MSHFlexGrid1.Left + MSHFlexGrid1.Width * (1 / 4)
    Text6.Width = MSHFlexGrid1.Width * (3 / 4)
    Text5.Top = MSHFlexGrid1.Top - (15 + Text5.Height)
    Text6.Top = MSHFlexGrid1.Top - (15 + Text6.Height)
    
    nTop = MSHFlexGrid1.Top + MSHFlexGrid1.Height + 105
    Command1(1).Top = nTop
    Command1(0).Top = nTop
    
    Command1(1).Left = (MSHFlexGrid1.Left + MSHFlexGrid1.Width) - Command1(1).Width
    Command1(0).Left = Command1(1).Left - (75 + Command1(0).Width)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim cWiddth As String
    Dim strRegSection As String
    
    If m_rstHFlex.EOF And m_rstHFlex.BOF Then
    'If Adodc1.Recordset.BOF And Adodc1.Recordset.EOF Then
        If Me.CallingForm = "DV1" Or Me.CallingForm = "frm_licensee" Then
            cSelVal = "0"
        End If
        cWiddth = "1"
    Else
        m_rstHFlex.MoveFirst
        cWiddth = CStr(m_rstHFlex!Width)
    End If
    
    Select Case UCase(cDocument)
        Case "IMPORT"
            strRegSection = "CodiSheet"
        Case "EXPORT"
            strRegSection = "ExSheet"
        Case "TRANSIT"
            strRegSection = "TrSheet"
        'ncts
        Case "TRANSIT NCTS"
            strRegSection = G_CONST_NCTS1_SHEET
        Case "COMBINED NCTS"
            strRegSection = G_CONST_NCTS2_SHEET
            
        Case "EDI NCTS"
            strRegSection = G_CONST_EDINCTS1_SHEET
    End Select
    
    If blnOK = False Then
        If bCode = "L1" Then
            cSelVal = LastVal + "@@"
        Else
            Select Case Me.CallingForm
                Case "frm_licensee", "DV1"
                    cSelVal = "0"
                Case Else
                    cSelVal = LastVal
            End Select
        End If
    End If
        
    SaveSetting App.Title, strRegSection, "Pick_" & Trim(docName) & "_" & Trim(TreeID), Trim(cSelVal & "%" & cWiddth)
    
    ADORecordsetClose m_rstA
    ADORecordsetClose m_rstB
    
    UnloadControls Me
    Set frm_picklist = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
    LData = True
    
    If nDataCtr > 0 Then
        If MSHFlexGrid1.Row = 0 Then
            Exit Sub
        End If
        
        Text6.Text = Me.MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
        Text5.Text = Me.MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
    End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
    If nDataCtr > 0 Then
        If MSHFlexGrid1.Row = 0 Then
            Exit Sub
        End If
        
        LData = False
        Command1_Click (0)
    End If
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Command1_Click (0)
            
            Exit Sub
        Case vbKeyEscape, 113
            Command1_Click (1)
            
            Exit Sub
        Case Else
            LData = True
            
            If nDataCtr > 0 Then
                Text6.Text = Me.MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
                Text5.Text = Me.MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0)
            End If
    End Select
End Sub

Private Sub Text5_Change()
    Dim i As Integer
    
    If nDataCtr > 0 Then
        If Not LData And nAB = 1 Then
            If Len(Trim(Text5.Text)) = 0 Then
                MSHFlexGrid1.Row = 1
                MSHFlexGrid1.RowSel = 1
                MSHFlexGrid1.TopRow = 1
                
                Text6.Text = " "
                
                Exit Sub
            End If
            
            i = 0
            
            For i = 1 To MSHFlexGrid1.Rows - 1
                If Trim(UCase(Text5.Text)) = Trim(UCase(Left(MSHFlexGrid1.TextMatrix(i, 0), Len(Trim(Text5.Text))))) Then
                    If nAB = 1 Then
                        Text6.Text = MSHFlexGrid1.TextMatrix(i, 1)
                    End If
                    
                    MSHFlexGrid1.Col = 0
                    MSHFlexGrid1.Row = i
                    
                    MSHFlexGrid1.LeftCol = 0
                    MSHFlexGrid1.TopRow = i
                    
                    MSHFlexGrid1.ColSel = 0
                    MSHFlexGrid1.RowSel = i
                    
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    LData = False
    
    If nAB <> 1 Then
        Set m_rstHFlex = m_rstA
        'Adodc1.RecordSource = A
        'Adodc1.Refresh
        ' TO DO FOR CP.NET
        
        doInitCaption
    End If
    
    nAB = 1
    
    Select Case KeyCode
        Case vbKeyReturn
            Command1_Click (0)
        Case vbKeyEscape, 113
            Command1_Click (1)
        Case 38, 40
            MSHFlexGrid1.SetFocus
    End Select
End Sub

Private Sub Text6_Change()
    Dim i As Integer
    
    If nDataCtr > 0 Then
        If Not LData And nAB = 2 Then
            If Len(Trim(Text6.Text)) = 0 Then
                MSHFlexGrid1.Row = 1
                MSHFlexGrid1.RowSel = 1
                MSHFlexGrid1.TopRow = 1
                
                Text5.Text = " "
                
                Exit Sub
            End If
            
            i = 0
            
            For i = 1 To MSHFlexGrid1.Rows - 1
                If Trim(UCase(Text6.Text)) = Trim(UCase(Left(MSHFlexGrid1.TextMatrix(i, 1), Len(Trim(Text6.Text))))) Then
                    If nAB = 2 Then
                        Text5.Text = MSHFlexGrid1.TextMatrix(i, 0)
                    End If
                    
                    MSHFlexGrid1.Col = 1
                    MSHFlexGrid1.Row = i
                    
                    MSHFlexGrid1.LeftCol = 1
                    MSHFlexGrid1.TopRow = i
                    
                    MSHFlexGrid1.ColSel = 1
                    MSHFlexGrid1.RowSel = i
                    
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    LData = False
    
    If nAB <> 2 Then
        
        Set m_rstHFlex = m_rstB
        'Adodc1.RecordSource = B
        'Adodc1.Refresh
        ' TO DO FOR CP.NET
        
        doInitCaption
    End If
    
    nAB = 2
    
    Select Case KeyCode
        Case vbKeyReturn
            Command1_Click (0)
        Case vbKeyEscape, 113
            Command1_Click (1)
        Case 38, 40
            MSHFlexGrid1.SetFocus
    End Select
End Sub

Private Sub doInitCaption()
    With MSHFlexGrid1
        .Row = 0
        .Col = 0
        .Text = Trim(Translate(627))
        
        .Row = 0
        .Col = 1
        .Text = Trim(Translate(628))
        
        If nDataCtr > 0 Then
            .Row = 1
            .Col = 0
        Else
            .Row = 0
            .Col = 0
        End If
        
        .ColWidth(0) = Text5.Width
        .ColWidth(1) = .Width - .ColWidth(0)
    End With
End Sub

Public Property Get CallingForm() As String
    CallingForm = mstrCallingForm
End Property

Public Property Let CallingForm(ByVal strCallingForm As String)
    mstrCallingForm = strCallingForm
End Property

Public Sub PassCallingForm(ByRef CallingForm As Object)
    Set m_CallingForm = CallingForm
End Sub
