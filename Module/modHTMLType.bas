Attribute VB_Name = "modHTMLType"
Public Const DIV = "<DIV style=""HEIGHT: 30px; POSITION: absolute; WIDTH: 30px"">"
Public Const SPAN = "<SPAN></SPAN>"
Public Const HTMLForm = "<FORM action="""" id=FORM1 method=post name=FORM1></FORM>"
Public Const HTMLLine = "<HR></HR>"
Public Const HTMLButton = "<INPUT type=button value=Button event=onclick >"
Public Const HTMLCheckBox = "<INPUT type=""checkbox"" value=""Check Box"">"
Public Const HTMLRadio = "<INPUT type=""radio"" value=""Radio Button"">"
Public Const HTMLMultiline = "<textarea rows=""2"" name=""S1"" cols=""20""></textarea>"
Public Const HTMLLabel = "<LABEL>Label</LABEL>"
Public Const HTMLCombo = "<select size=""1"" name=""D1""></select>"
Public Const HTMLTextbox = "<input type=""text"" size=""20"">"

Public Type FlashObject
Filename As String
width As Long
height As Long
align As String
Visible As Boolean
End Type

Public Type Marquee
align As String
direction As String
loop As Long
scrollamount As Long
scrolldelay As Long
width As Long
height As Long
behavior As String
End Type

Public Type Table
align As String
Background As String
bgcolor As String
border As Integer
borderColor As String
borderColorLight As String
borderColorDark As String
cellPadding As Integer
cellSpacing As Integer
width As Long
End Type

