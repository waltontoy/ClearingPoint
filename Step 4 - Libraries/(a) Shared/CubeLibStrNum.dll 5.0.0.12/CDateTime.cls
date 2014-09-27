VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDateTime"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
'local variable(s) to hold property value(s)
Private mvarcpiYear As Integer 'local copy
Private mvarcpiMonth As Integer 'local copy
Private mvarcpiDayOfWeek As Integer 'local copy
Private mvarcpiDay As Integer 'local copy
Private mvarcpiHour As Integer 'local copy
Private mvarcpiMinute As Integer 'local copy
Private mvarcpiSecond As Integer 'local copy
Private mvarcpiMilliseconds As Integer 'local copy

Public Enum LOCALE_SETTINGS

    LOCALE_ICENTURY = &H24
    LOCALE_ICOUNTRY = &H5
    LOCALE_ICURRDIGITS = &H19
    LOCALE_ICURRENCY = &H1B
    LOCALE_IDATE = &H21
    LOCALE_IDAYLZERO = &H26
    LOCALE_IDEFAULTCODEPAGE = &HB
    LOCALE_IDEFAULTCOUNTRY = &HA
    LOCALE_IDEFAULTLANGUAGE = &H9
    LOCALE_IDIGITS = &H11
    LOCALE_IINTLCURRDIGITS = &H1A
    LOCALE_ILANGUAGE = &H1
    LOCALE_ILDATE = &H22
    LOCALE_ILZERO = &H12
    LOCALE_IMEASURE = &HD
    LOCALE_IMONLZERO = &H27
    LOCALE_INEGCURR = &H1C
    LOCALE_INEGSEPBYSPACE = &H57
    LOCALE_INEGSIGNPOSN = &H53
    LOCALE_INEGSYMPRECEDES = &H56
    LOCALE_IPOSSEPBYSPACE = &H55
    LOCALE_IPOSSIGNPOSN = &H52
    LOCALE_IPOSSYMPRECEDES = &H54
    LOCALE_ITIME = &H23
    LOCALE_ITLZERO = &H25
    LOCALE_NOUSEROVERRIDE = &H80000000
    LOCALE_S1159 = &H28
    LOCALE_S2359 = &H29
    LOCALE_SABBREVCTRYNAME = &H7
    LOCALE_SABBREVDAYNAME1 = &H31
    LOCALE_SABBREVDAYNAME2 = &H32
    LOCALE_SABBREVDAYNAME3 = &H33
    LOCALE_SABBREVDAYNAME4 = &H34
    LOCALE_SABBREVDAYNAME5 = &H35
    LOCALE_SABBREVDAYNAME6 = &H36
    LOCALE_SABBREVDAYNAME7 = &H37
    LOCALE_SABBREVLANGNAME = &H3
    LOCALE_SABBREVMONTHNAME1 = &H44
    LOCALE_SCOUNTRY = &H6
    LOCALE_SCURRENCY = &H14
    LOCALE_SDATE = &H1D
    LOCALE_SDAYNAME1 = &H2A
    LOCALE_SDAYNAME2 = &H2B
    LOCALE_SDAYNAME3 = &H2C
    LOCALE_SDAYNAME4 = &H2D
    LOCALE_SDAYNAME5 = &H2E
    LOCALE_SDAYNAME6 = &H2F
    LOCALE_SDAYNAME7 = &H30
    LOCALE_SDECIMAL = &HE
    LOCALE_SENGCOUNTRY = &H1002
    LOCALE_SENGLANGUAGE = &H1001
    LOCALE_SGROUPING = &H10
    LOCALE_SINTLSYMBOL = &H15
    LOCALE_SLANGUAGE = &H2
    LOCALE_SLIST = &HC
    LOCALE_SLONGDATE = &H20
    LOCALE_SMONDECIMALSEP = &H16
    LOCALE_SMONGROUPING = &H18
    LOCALE_SMONTHNAME1 = &H38
    LOCALE_SMONTHNAME10 = &H41
    LOCALE_SMONTHNAME11 = &H42
    LOCALE_SMONTHNAME12 = &H43
    LOCALE_SMONTHNAME2 = &H39
    LOCALE_SMONTHNAME3 = &H3A
    LOCALE_SMONTHNAME4 = &H3B
    LOCALE_SMONTHNAME5 = &H3C
    LOCALE_SMONTHNAME6 = &H3D
    LOCALE_SMONTHNAME7 = &H3E
    LOCALE_SMONTHNAME8 = &H3F
    LOCALE_SMONTHNAME9 = &H40
    LOCALE_SMONTHOUSANDSEP = &H17
    LOCALE_SNATIVECTRYNAME = &H8
    LOCALE_SNATIVEDIGITS = &H13
    LOCALE_SNATIVELANGNAME = &H4
    LOCALE_SNEGATIVESIGN = &H51
    LOCALE_SPOSITIVESIGN = &H50
    LOCALE_SSHORTDATE = &H1F
    LOCALE_STHOUSAND = &HF
    LOCALE_STIME = &H1E
    LOCALE_STIMEFORMAT = &H1003
End Enum




Public Property Get cpiMilliseconds() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cpiMilliseconds
    cpiMilliseconds = mvarcpiMilliseconds
End Property

Public Property Get cpiSecond() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cpiSecond
    cpiSecond = mvarcpiSecond
End Property

Public Property Get cpiMinute() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cpiMinute
    cpiMinute = mvarcpiMinute
End Property

Public Property Get cpiHour() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cpiHour
    cpiHour = mvarcpiHour
End Property

Public Property Get cpiDay() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cpiDay
    cpiDay = mvarcpiDay
End Property

Public Property Get cpiDayOfWeek() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cpiDayOfWeek
    cpiDayOfWeek = mvarcpiDayOfWeek
End Property

Public Property Get cpiMonth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cpiMonth
    cpiMonth = mvarcpiMonth
End Property

Public Property Get cpiYear() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cpiYear
    cpiYear = mvarcpiYear
End Property

Public Function GetNow() As String
    Dim TimeNow As SYSTEMTIME
    Static StaticCount As Long
    
    StaticCount = StaticCount + 1
    
    Call GetSystemTime(TimeNow)
    With TimeNow
        mvarcpiYear = .wYear
        mvarcpiMonth = .wMonth
        mvarcpiDayOfWeek = .wDayOfWeek
        mvarcpiDay = .wDay
        mvarcpiHour = .wHour
        mvarcpiMinute = .wMinute
        mvarcpiSecond = .wSecond
        mvarcpiMilliseconds = .wMilliseconds
    End With
End Function

Public Function SystemDateFormat() As String
    
    Dim clsSetting As New CSystem
    
    SystemDateFormat = clsSetting.GetRegionalSetting(LOCALE_SSHORTDATE)
    
    Set clsSetting = Nothing
    

End Function
