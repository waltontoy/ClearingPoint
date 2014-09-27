Attribute VB_Name = "modCubeFTPBrowse"
Option Explicit
Public Function FormatSize(lngSize As Long) As String
    Select Case lngSize
        Case 0 To 1023
            FormatSize = CStr(lngSize) & " Bytes"
        Case 1024 To 1048575
            FormatSize = Format(lngSize / 1024#, "###0.00") & " KB"
        Case 1024# ^ 2 To 1043741824
            FormatSize = Format(lngSize / 1024# ^ 2, "###0.00") & " MB"
        Case Is > 1043741824
            FormatSize = Format(lngSize / 1024# ^ 3, "###0.00") & " GB"
    End Select

End Function

