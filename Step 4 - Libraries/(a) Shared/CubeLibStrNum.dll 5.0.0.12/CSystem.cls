VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSystem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum SystemFormats
    cpiCurrecny = 1
    cpiNumber = 2
    cpiShortDate = 3
    cpiLongDate = 4
End Enum

Public Function GetSystemLanguage() As Integer

    ' ********** Language setting must be determined before performing Syslink via SendKeysEx because
    ' ********** menu keystrokes differ among the three languages.
    Dim strLCID As String
    Dim lngLCID As Long
    
    lngLCID = GetThreadLocale()     ' Get current thread locale ID (defaults to user locale ID)
    strLCID = Hex(CStr(lngLCID))    ' Convert to hexadecimal
    
    Select Case UCase(strLCID)
        Case "413", "813"    ' Dutch
            GetSystemLanguage = 2
        
        ' MUCP-155 - Start
        Case "40C", "80C", "C0C", "100C", "140C", "180C", "46E"    ' French
            GetSystemLanguage = 3
        ' MUCP-155 - End
        
        Case Else            ' English: "409", "809", "C09", "1009", "1409", "1809", "1C09", "2009", "2409", "2809", "2C09", "3009", "3409"
            GetSystemLanguage = 1
    End Select
End Function

Public Function cpiDataFormat(ByVal Value As Variant, ByVal SystemFormat As SystemFormats) As Variant

    Dim Buffer As String
    Dim NF As NUMBERFMT
    Dim ST As SYSTEMTIME
    Dim dteDate As Date
    Dim strDummy As String
    
    On Error GoTo ERROR_TYPE_MISMATCH
    
    strDummy = Value
    
    dteDate = CDate(Value)
    Buffer = String(255, 0)
    
    Select Case SystemFormat
        Case cpiCurrecny
            'GetCurrencyFormat ByVal 0&, 0, "-3345.32", ByVal 0&, Buffer, Len(Buffer)
            strDummy = CStr(CDbl(strDummy))
            GetCurrencyFormat ByVal 0&, 0, strDummy, ByVal 0&, Buffer, Len(Buffer)
            Buffer = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
        Case cpiNumber
            With NF
                .NumDigits = 3
                .Grouping = 1
                .lpDecimalSep = "."
                .lpThousandSep = ""
                .NegativeOrder = 0
            End With
            GetNumberFormat ByVal 0&, 0, "300243.24", NF, Buffer, Len(Buffer)
            
        Case cpiShortDate, cpiLongDate
            'Format

            With ST
                .wDay = Day(dteDate)
                .wMonth = Month(dteDate)
                .wYear = Year(dteDate)
            End With
            Buffer = String(255, 0)
            GetDateFormat ByVal 0&, 0, ST, vbNullString, Buffer, Len(Buffer)
            Buffer = Left$(Buffer, InStr(1, Buffer, Chr$(0)) - 1)
            
            If SystemFormat = cpiLongDate Then
                Buffer = Format(Buffer, "Long Date")
            End If
    End Select
            
    cpiDataFormat = IIf(Trim(Buffer) = "", Null, Buffer)
    
    Exit Function
    
ERROR_TYPE_MISMATCH:

    Select Case Err.Number
        Case 13
            cpiDataFormat = Null
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
            Err.Clear
    End Select
End Function

Public Function GetRegionalSetting(Setting As LOCALE_SETTINGS) As String

    Dim Symbol As String
    Dim iRet1 As Long
    Dim iRet2 As Long
    Dim lpLCDataVar As String
    Dim pos As Integer
    Dim Locale As Long

    Locale = GetUserDefaultLCID()

    iRet1 = GetLocaleInfo(Locale, Setting, _
    lpLCDataVar, 0)
    Symbol = String$(iRet1, 0)

    iRet2 = GetLocaleInfo(Locale, Setting, Symbol, iRet1)
    pos = InStr(Symbol, Chr$(0))
    If pos > 0 Then
         Symbol = Left$(Symbol, pos - 1)
    End If

    GetRegionalSetting = Symbol

End Function

Public Sub SetRegionalSetting(Setting As LOCALE_SETTINGS, NewValue As String) 'Change the regional setting

      Dim iRet As Long
      Dim Locale As Long

      Locale = GetUserDefaultLCID() 'Get user Locale ID
      iRet = SetLocaleInfo(Locale, Setting, NewValue)

End Sub





