Attribute VB_Name = "modScript"
Public Const SP = vbCrLf & vbCrLf
Public Const AP = """"

'Public Enum SLanguage
'JavaScript = 1
'JavaScript11 = 2
'JavaScript12 = 3
'JScript = 4
'VBScript = 5
'LiveScript = 6
'TCLScript = 7
'PHPScript = 8
'End Enum

Public Function BuildScript(ByVal LanguageName As String, Optional ByVal Content As String) As String
    BuildScript = "<script language=""" & LanguageName & """>" & vbCrLf & "<!--" & vbCrLf & Content & vbCrLf & "-->" & vbCrLf & "</script>"
End Function

