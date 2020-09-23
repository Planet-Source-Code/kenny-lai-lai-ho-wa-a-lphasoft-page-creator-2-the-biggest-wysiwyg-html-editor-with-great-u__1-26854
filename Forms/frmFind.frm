VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Find"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.Frame Frame2 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000004&
      Caption         =   "Find Target"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1320
      TabIndex        =   11
      Top             =   480
      Width           =   3255
      Begin VB.OptionButton optIn 
         Caption         =   "Selected &Text"
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   5
         ToolTipText     =   "find in selected text"
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optIn 
         Caption         =   "&Current Document"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Find in the current document"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000004&
      Caption         =   "Direction"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1095
      Begin VB.OptionButton optDirection 
         Caption         =   "&Up"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Find Direction: Up"
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "&Down"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Find Direction: Down"
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   8
      ToolTipText     =   "Find Now!"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4680
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   9
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  '¥­­±
      Height          =   300
      Left            =   600
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.CheckBox ckCases 
      Caption         =   "&Ignore Cases (""A"" same as ""a"")"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "It will Ignore the difference between capital letters and small letters. "
      Top             =   1320
      Value           =   1  '®Ö¨ú
      Width           =   3015
   End
   Begin VB.CheckBox ckTrim 
      Caption         =   "Auto &Delete unuseful spaces from the ""Find"" string"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "It will delete all spaces on the left and right, and also the double spaces in the string you want to find."
      Top             =   1560
      Value           =   1  '®Ö¨ú
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "&Find:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FindDirection As Integer    '0, down    ;1, up
Private Findstr As String
Private fMode As Integer
Public formF As frmMain


Private Sub ckTrim_Click()
txtFind.ToolTipText = " String to Find: " & IIf(ckTrim.Value, SuperTrim(txtFind.text), txtFind.text)
cmdOK.ToolTipText = "String to Find:" & IIf(ckTrim.Value, SuperTrim(txtFind.text), txtFind.text)
End Sub



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
optIn(0).Value = True
optIn(2).Enabled = False
Dim TextFind  As String
TextFind = IIf(ckTrim.Value, SuperTrim(txtFind.text), txtFind.text)
If optIn(0).Value = True Then  'current
    If optDirection(0).Value = True Then   'down and current
        If ckCases.Value = 1 Then 'down,current,ignore cases
            FindCurrentDown TextFind, 2
        Else    'down, current,consider cases
            FindCurrentDown TextFind, 1
        End If
    ElseIf optDirection(1).Value = True Then 'up, current
        If ckCases.Value = 1 Then 'up,current,ignore cases
            FindCurrentUp TextFind, 2
        Else    'up, current,consider cases
            FindCurrentUp TextFind, 1
        End If
    End If
Else  'seltext
    If optDirection(0).Value = True Then   'down and sel
        If ckCases.Value = 1 Then 'down,sel,ignore cases
            FindSelDown TextFind, 2
        Else    'down, sel,consider cases
            FindSelDown TextFind, 1
        End If
    End If
End If
End Sub

Private Sub Form_Load()
If formF.rt1.SelText = "" Then
    optIn(0).Value = True
    optIn(2).Enabled = False
Else
    optIn(2).Value = True
    optDirection(0).Enabled = False
    optDirection(1).Enabled = False
End If
End Sub



Private Sub optIn_Click(Index As Integer)
If Index = 2 Then
    optDirection(0).Enabled = False
    optDirection(1).Enabled = False
Else
    optDirection(0).Enabled = True
    optDirection(1).Enabled = True
End If
On Error Resume Next
    txtFind.SetFocus
End Sub

Private Sub txtFind_Change()
txtFind.ToolTipText = "String to Find:" & IIf(ckTrim.Value, SuperTrim(txtFind.text), txtFind.text)
cmdOK.ToolTipText = "String to Find:" & IIf(ckTrim.Value, SuperTrim(txtFind.text), txtFind.text)
End Sub


Private Sub FindCurrentDown(ByVal FindText As String, ByVal findMode As Integer)
Dim i As Integer
For i = IIf(formF.rt1.SelText = vbNullString, formF.rt1.SelStart + 1, formF.rt1.SelStart + 2) To Len(formF.rt1.text)
    If findMode = 1 Then  '1, binary
        If Mid(formF.rt1.text, i, Len(FindText)) = FindText Then Exit For
    Else  'text compare
        If LCase(Mid(formF.rt1.text, i, Len(FindText))) = LCase(FindText) Then Exit For
    End If
Next i
If i = Len(formF.rt1.text) + 1 Then GoTo NotFound
formF.rt1.SelStart = i - 1
formF.rt1.SelLength = Len(FindText)
Findstr = FindText
FindDirection = 0
fMode = findMode
Exit Sub
NotFound:
MsgBox FindText & " Not Found.", vbInformation
frmFind.txtFind.SetFocus
frmFind.txtFind.SelStart = 0
frmFind.txtFind.SelLength = Len(FindText)
End Sub

Private Sub FindCurrentUp(ByVal FindText As String, ByVal findMode As Integer)
Dim i As Long
For i = formF.rt1.SelStart + 1 - Len(FindText) To 1 Step -1
    If findMode = 1 Then  '1, binary
        If Mid(formF.rt1.text, i, Len(FindText)) = FindText Then Exit For
    Else  'text compare
        If LCase(Mid(formF.rt1.text, i, Len(FindText))) = LCase(FindText) Then Exit For
    End If
Next i
If i = 0 Then GoTo NotFound
formF.rt1.SelStart = i - 1
formF.rt1.SelLength = Len(FindText)
Findstr = FindText
FindDirection = 1
fMode = findMode
Exit Sub
NotFound:
MsgBox FindText & " Not Found.", vbInformation
frmFind.txtFind.SetFocus
frmFind.txtFind.SelStart = 0
frmFind.txtFind.SelLength = Len(FindText)
End Sub

Private Sub FindSelDown(ByVal FindText As String, ByVal findMode As Integer)
Dim Selstr As String
Selstr = formF.rt1.SelText
Dim i As Integer
For i = 1 To Len(Selstr)
    If findMode = 1 Then  '1, binary
        If Mid(Selstr, i, Len(FindText)) = FindText Then Exit For
    Else  'text compare
        If LCase(Mid(Selstr, i, Len(FindText))) = LCase(FindText) Then Exit For
    End If
Next i
If i = Len(Selstr) + 1 Then GoTo NotFound
formF.rt1.SelStart = formF.rt1.SelStart + i - 1
formF.rt1.SelLength = Len(FindText)
Findstr = FindText
FindDirection = 0
fMode = findMode
Exit Sub
NotFound:
MsgBox FindText & " Not Found.", vbInformation
frmFind.txtFind.SetFocus
frmFind.txtFind.SelStart = 0
frmFind.txtFind.SelLength = Len(FindText)
End Sub

Private Function ReadText(strFile)
Dim iFile, sData

On Error Resume Next

iFile = FreeFile
sData = ""

If Dir(strFile) <> "" Then
If Len(strFile) Then
Open strFile For Binary As #iFile

sData = Input(LOF(iFile), #iFile)
DoEvents

Close #iFile

End If

ReadText = sData
Else
ReadText = ""
End If

End Function

Private Function Absolute(Str1 As String) As String
Dim Temp As String
Temp = Replace(Str1, ",", Chr$(32))
Temp = Replace(Temp, ".", Chr$(32))
Temp = Replace(Temp, "/", Chr$(32))
Temp = Replace(Temp, "<", Chr$(32))
Temp = Replace(Temp, ">", Chr$(32))
Temp = Replace(Temp, "?", Chr$(32))
Temp = Replace(Temp, ";", Chr$(32))
Temp = Replace(Temp, "'", Chr$(32))
Temp = Replace(Temp, ":", Chr$(32))
Temp = Replace(Temp, Chr$(34), Chr$(32))
Temp = Replace(Temp, "[", Chr$(32))
Temp = Replace(Temp, "]", Chr$(32))
Temp = Replace(Temp, "{", Chr$(32))
Temp = Replace(Temp, "}", Chr$(32))
Temp = Replace(Temp, "`", Chr$(32))
Temp = Replace(Temp, "~", Chr$(32))
Temp = Replace(Temp, "!", Chr$(32))
Temp = Replace(Temp, "@", Chr$(32))
Temp = Replace(Temp, "#", Chr$(32))
Temp = Replace(Temp, "$", Chr$(32))
Temp = Replace(Temp, "%", Chr$(32))
Temp = Replace(Temp, "^", Chr$(32))
Temp = Replace(Temp, "&", Chr$(32))
Temp = Replace(Temp, "*", Chr$(32))
Temp = Replace(Temp, "(", Chr$(32))
Temp = Replace(Temp, ")", Chr$(32))
Temp = Replace(Temp, "-", Chr$(32))
Temp = Replace(Temp, "_", Chr$(32))
Temp = Replace(Temp, "=", Chr$(32))
Temp = Replace(Temp, "+", Chr$(32))
Temp = Replace(Temp, "\", Chr$(32))
Temp = Replace(Temp, "|", Chr$(32))
Temp = SuperTrim(Temp)
Absolute = Temp
End Function

Private Function SuperTrim(StrToTrim As String) As String
Dim Temp As String
Dim DblSpc As String
Temp = Replace(StrToTrim, vbCrLf, Chr$(32))
Do Until InStr(1, Temp, vbCrLf) = 0
    Temp = Replace(Temp, vbCrLf, Chr$(32))
Loop
Temp = Replace(Temp, vbCr, Chr$(32))
Temp = Replace(Temp, vbLf, Chr$(32))
Temp = Trim(Temp)
DblSpc = Chr$(32) & Chr$(32)
Temp = Replace(Temp, DblSpc, Chr$(32))
Do Until InStr(1, Temp, DblSpc) = 0
    Temp = Replace(Temp, DblSpc, Chr$(32))
Loop
SuperTrim = Temp
End Function

Private Function wordCount(ByVal text As String) As Long
    If text = "" Then GoTo NoChar
    'count char without sapce and crlf
    AbsStr = Absolute(text)
    TempStr = Split(AbsStr, Chr$(32))
    wordCount = UBound(TempStr) + 1
    Exit Function
NoChar:
    wordCount = 0
End Function


