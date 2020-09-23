Attribute VB_Name = "modGeneral"
Option Base 1

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public ThreadID As Long
Public Process As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GdiFlush Lib "gdi32" () As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long

Public Const EXIT_PROCESS_DEBUG_EVENT = 5
Public Const EXIT_THREAD_DEBUG_EVENT = 4
Public Const RESOURCEUSAGE_RESERVED As Long = &H80000000
Public Const MOD_CONTROL = &H2

Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public FontSizePoint(1 To 7) As Integer

Public WordCommand As String

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))


    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Sub AddDirSep(strPathName As String)
    If Right(Trim(strPathName), Len("/")) <> "/" And _
       Right(Trim(strPathName), Len("\")) <> "\" Then
        strPathName = RTrim$(strPathName) & "\"
    End If
End Sub

Function GetWindowsSysDir() As String
    Dim strBuf As String

    strBuf = Space$(255)

    '
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetSystemDirectory(strBuf, 255) > 0 Then
        strBuf = StripTerminator(strBuf)
        AddDirSep strBuf
        
        GetWindowsSysDir = strBuf
    Else
        GetWindowsSysDir = vbNullString
    End If
End Function

Sub Install(ByVal DllName As String)
On Error GoTo 2
FileCopy App.Path & "\" & DllName, GetWindowsSysDir & DllName
2

Shell "regsvr32 " & DllName & " /s"

End Sub

Sub Main()

Install "SSubTmr6.dll"
    
'Make HTML Files can open with Page Creator
bSetRegValue HKEY_CLASSES_ROOT, "htmlfile\shell\Edit with Alphasoft", "", "Edit with Alphasoft Page Creator"
bSetRegValue HKEY_CLASSES_ROOT, "htmlfile\shell\Edit with Alphasoft\command", "", """" & App.Path & "\" & App.EXEName & ".exe" & """" & " ""%1" & AP

'Register my unique file: Page Creator Web Site File(*.pcs)
bSetRegValue HKEY_CLASSES_ROOT, ".pcs", "", "Page Creator Web Site File"
bSetRegValue HKEY_CLASSES_ROOT, "Page Creator Web Site File\shell", "", "open"
bSetRegValue HKEY_CLASSES_ROOT, "Page Creator Web Site File\shell\open", "", "Open with Alphasoft Page Creator"
bSetRegValue HKEY_CLASSES_ROOT, "Page Creator Web Site File\shell\open\command", "", """" & App.Path & "\" & App.EXEName & ".exe" & """" & " ""%1" & AP
bSetRegValue HKEY_CLASSES_ROOT, "Page Creator Web Site File\DefaultIcon", "", App.Path & "\Icons\PCSIcon.ico,0"

    If Command <> "" Then
    
    Dim FName As String
    FName = Mid(Command, 2, InStr(2, Command, """") - 2)
    WordCommand = FName
    MfrmProgram.Show
    'OpenFile Command, Command
    
    Else
    WordCommand = ""
    DoEvents
    
    frmSplash.Show
    
    End If
End Sub

Sub FreeMemory()
'WARNING

'WW                WWWW                WW
' WW              WW  WW              WW
'  WW            WW    WW            WW
'   WW          WW      WW          WW
'    WW        WW        WW        WW
'     WW      WW          WW      WW
'      WW    WW            WW    WW
'       WW  WW              WW  WW
'        WWWW                WWWW    A R N I N G

'Use the command below ONLY when you are finished
'or Release.
'It's used to exit process and free memory, but it also close your VB.

'Kenny Lai
'In Design mode, Please use END
On Error Resume Next
TerminateThread ThreadID, 4
'TerminateProcess Process, 4

End Sub

Function CommandPress(ByVal intC As Integer) As Integer
DoEvents
Select Case intC
Case 3: CommandPress = 0
Case 7: CommandPress = 1
End Select
End Function

Function CommandEnable(ByVal intC As Integer) As Boolean
DoEvents
Select Case intC
Case 1: CommandEnable = False
Case 3, 7, 11: CommandEnable = True
End Select
End Function

Function ButtonPress(ByVal CmdID As DHTMLEDITCMDID) As Integer
DoEvents
Dim i As Integer
i = MfrmProgram.ActiveForm.DHTML1.QueryStatus(CmdID)
ButtonPress = CommandPress(i)
End Function

Function ButtonEnable(ByVal CmdID As DHTMLEDITCMDID) As Boolean
DoEvents
Dim i As Integer
i = MfrmProgram.ActiveForm.DHTML1.QueryStatus(CmdID)
ButtonEnable = CommandEnable(i)
End Function

Sub OpenTemp(ByVal File As String, Title As String)
MfrmProgram.File1.Enabled = False

'Handle the error when the page cannot open

'On Error GoTo OpenErr
'MfrmProgram.DE1.LoadDocument File, True
'GoTo 2
'OpenErr:
'MsgBox "Open this file will cause error!", vbCritical
'MfrmProgram.File1.Enabled = True
'Exit Sub

'2

Dim Editor As frmMain
Set Editor = New frmMain

'MsgBox "New Form Complete", vbSystemModal
DoEvents

With Editor
.txtFile.Text = File

.DHTML1.ZOrder 0
On Error GoTo 1

'MsgBox "Loading: " & File, vbSystemModal

'Load the File into the DHTML control
DoEvents
.DHTML1.LoadDocument File, False

'Get the html code for reference

.FileName = Mid(.txtFile.Text, 1, Len(.txtFile.Text) - 4) & "tmp.htm"
.cIsSave.Value = 0
.Caption = Title
.Flags = "OK"
.Tag = "GetHTML"
'MsgBox "Complete, now show", vbSystemModal
DoEvents

.Show
Exit Sub

1
MsgBox "Error opening file!" & vbCrLf & "ErrNumber: " & Err.Number & " " & Error, vbCritical
.Flags = "EXIT"
Unload Editor
End With

End Sub

Function SaveHistory(ByVal URL As String)
Dim HistoryCount As Currency
Dim HistoryFull As Boolean
Dim num As Long
Dim HistoryReg As Currency

HistoryCount = GetSet(App.ProductName, "history", "historycount", 0)

'Force number of history less than 30
HistoryReg = HistoryCount
Do Until HistoryReg < 100
HistoryReg = HistoryReg - 100
Loop

num = HistoryReg + 1
SaveSet App.ProductName, "history", "file" & num, URL

SaveSet App.ProductName, "history", "historycount", HistoryCount + 1

End Function

Sub OpenFile(ByVal File As String, Title As String)

DoEvents
SaveHistory File

MfrmProgram.File1.Enabled = False

'Handle the error when the page cannot open

'On Error GoTo OpenErr
'MfrmProgram.DE1.LoadDocument File ', True
'GoTo 2

'OpenErr:
'MsgBox "Open this file will cause error!", vbCritical
'MfrmProgram.file1.Enabled = True
'Exit Sub

'2

Dim Editor As frmMain
Set Editor = New frmMain

'MsgBox "New Form Complete", vbSystemModal
DoEvents

With Editor
.txtFile.Text = File

.DHTML1.ZOrder 0
On Error GoTo 1

'MsgBox "Loading: " & File, vbSystemModal

'Load the File into the DHTML control
DoEvents
.DHTML1.LoadDocument File

'Get the html code for reference

.FileName = Mid(.txtFile.Text, 1, Len(.txtFile.Text) - 4) & "tmp.htm"
.cIsSave.Value = 1
.Caption = File
.Flags = "OK"
.Tag = "GetHTML"
'MsgBox "Complete, now show", vbSystemModal
DoEvents

.Show
Exit Sub

1
MsgBox "Error opening file!" & vbCrLf & "ErrNumber: " & Err.Number & " " & Error, vbCritical
.Flags = "EXIT"
Unload Editor
End With

End Sub

Function ReadText(ByVal FileName As String)
Dim iFile, sData
On Error Resume Next
iFile = FreeFile
sData = ""
Open strFile For Binary As #iFile
sData = Input(LOF(iFile), #iFile)
DoEvents
Close #iFile
ReadText = sData
End Function

Sub InsertHTML(ByVal code As String)
    
Select Case MfrmProgram.ActiveForm.SSTab1.Tab
Case 0
    Dim doc As Object
    Dim sel As Object
    Dim tr As Object
    
    ' get the DHTML Document object
    Set doc = MfrmProgram.ActiveForm.DHTML1.DOM
    ' get the IE4 selection object
    Set sel = doc.selection
    ' create a TextRange from the current selection
    Set tr = sel.createRange
    
    ' paste our html into the range
    tr.pasteHTML (code)
Case 1
    MfrmProgram.ActiveForm.rt1.SelText = code
End Select
End Sub


Sub SetToolbar4rtf()
    With MfrmProgram.tbrGeneral
    'DoEvents
    .Buttons(6).Enabled = IIf(MfrmProgram.ActiveForm.rt1.SelLength = 0, False, True)
    .Buttons(8).Enabled = .Buttons(6).Enabled
    .Buttons(10).Enabled = IIf(Clipboard.GetText = "", False, True)
    .Buttons(12).Enabled = .Buttons(6).Enabled
    End With
End Sub


Sub StartProgram()

End Sub

'UI Sub and Functions
'UI Sub and Functions
'UI Sub and Functions
Sub LoadUICore(TabNow As Integer)
Dim i As Integer
Select Case TabNow
Case -1
    On Error Resume Next
    
    With MfrmProgram.tbrGeneral
    For i = 3 To 18
    .Buttons(i).Enabled = False
    Next i
    End With
    
    With MfrmProgram.tbrSimFunction
    .Buttons(1).Enabled = False
    For i = 3 To 9
    .Buttons(i).Enabled = False
    Next
    End With
    
    MfrmProgram.tbrEdit.Enabled = False
    
    With MfrmProgram
    .mnuSave = False
    .mnuSaveAs = False
    .mnuEdit_Top = False
    .mnuPreview = False
    .mnuPrint = False
    .mnuPSetup = False
    .mnuWindow = False
    .mnuDetail = False
    .lstTags.Enabled = False
    .lstCode.Enabled = False
    .tbrForm.Enabled = False
    .mnuTools.Enabled = False
    End With
    
Case 0
On Error Resume Next
    MfrmProgram.tbrEdit.Enabled = True
    With MfrmProgram.tbrGeneral
        .Buttons(3).Enabled = True
        .Buttons(5).Visible = True: .Buttons(7).Visible = True: .Buttons(9).Visible = True: .Buttons(11).Visible = True
        '.Buttons(5).Enabled = True: .Buttons(7).Enabled = True: .Buttons(9).Enabled = True: .Buttons(11).Enabled = True
        .Buttons(6).Visible = False: .Buttons(8).Visible = False: .Buttons(10).Visible = False: .Buttons(12).Visible = False
        '.Buttons(14).Enabled = True: .Buttons(15).Enabled = True:
        .Buttons(17).Enabled = True: .Buttons(18).Enabled = False
    End With
    
    With MfrmProgram.tbrSimFunction
    For i = 1 To .Buttons.Count
    .Buttons(i).Enabled = True
    Next
    End With
    
    With MfrmProgram
    .lstCode.Enabled = True
    .lstTags.Enabled = True
    .mnuEdit_Top = True
    .mnuSave = True
    .mnuSaveAs = True
    .mnuWindow = True
    '.mnuUndo = True
    '.mnuRedo = True
    .mnuDetail = True
    .mnuPrint = False
    .mnuPSetup = False
    .mnuPreview = False
    '.mnuDelete = True
    .tbrForm.Enabled = True
    .mnuTools = True
    End With
    
Case 1
On Error Resume Next
    With MfrmProgram.tbrGeneral
        .Buttons(3).Enabled = False
        .Buttons(5).Visible = False: .Buttons(7).Visible = False: .Buttons(9).Visible = False: .Buttons(11).Visible = False
        .Buttons(6).Visible = True: .Buttons(8).Visible = True:  .Buttons(10).Visible = True:  .Buttons(12).Visible = True
       .Buttons(14).Enabled = False: .Buttons(15).Enabled = False: .Buttons(17).Enabled = True: .Buttons(18).Enabled = False
    End With
    
    With MfrmProgram
    .lstCode.Enabled = True
    .lstTags.Enabled = True
    .tbrEdit.Enabled = False
    .mnuSave = False
    .mnuSaveAs = False
    .mnuUndo = True
    .mnuRedo = False
    .mnuEdit_Top = True
    .mnuWindow = True
    .mnuDetail = False
    .mnuTools = True
    .mnuPrint = False
    .mnuPSetup = False
    .mnuPreview = False
    .tbrForm.Enabled = True
    End With
    
Case 2
On Error Resume Next
    With MfrmProgram.tbrGeneral
        .Buttons(3).Enabled = False: .Buttons(5).Visible = True: .Buttons(7).Visible = True: .Buttons(9).Visible = True: .Buttons(11).Visible = True
        .Buttons(5).Enabled = False: .Buttons(7).Enabled = False: .Buttons(9).Enabled = False: .Buttons(11).Enabled = False
        .Buttons(6).Visible = False: .Buttons(8).Visible = False: .Buttons(10).Visible = False: .Buttons(12).Visible = False
        .Buttons(14).Enabled = False: .Buttons(15).Enabled = False: .Buttons(17).Enabled = False: .Buttons(18).Enabled = True
    End With
    
    With MfrmProgram
    .tbrEdit.Enabled = False
    .mnuEdit_Top = False
    .lstCode.Enabled = False
    .lstTags.Enabled = False
    .tbrForm.Enabled = False
    .mnuTools = False
    .mnuPrint = True
    .mnuPSetup = True
    .mnuPreview = True
End With
    
End Select
End Sub

Public Sub OpenURL(ByVal URL As String)
    ShellExecute MfrmProgram.hwnd, "open", URL, "", "", 1
End Sub

Function DelTree(ByVal strDir As String) As Long
Dim X As Long
Dim intAttr As Integer
Dim strAllDirs As String
Dim strFile As String
DelTree = -1
On Error Resume Next
strDir = Trim$(strDir)
If Len(strDir) = 0 Then Exit Function
If Right$(strDir, 1) = "\" Then strDir = Left$(strDir, Len(strDir) - 1)
If InStr(strDir, "\") = 0 Then Exit Function
intAttr = GetAttr(strDir)
If (intAttr And vbDirectory) = 0 Then Exit Function
strFile = Dir$(strDir & "\*.*", vbSystem Or vbDirectory Or vbHidden)
Do While Len(strFile)
If strFile <> "." And strFile <> ".." Then
  intAttr = GetAttr(strDir & "\" & strFile)
  If (intAttr And vbDirectory) Then
   strAllDirs = strAllDirs & strFile & Chr$(0)
  Else
   If intAttr <> vbNormal Then
    SetAttr strDir & "\" & strFile, vbNormal
    If Err Then DelTree = Err: Exit Function
   End If
   Kill strDir & "\" & strFile
   If Err Then DelTree = Err: Exit Function
  End If
End If
strFile = Dir$
Loop
Do While Len(strAllDirs)
X = InStr(strAllDirs, Chr$(0))
strFile = Left$(strAllDirs, X - 1)
strAllDirs = Mid$(strAllDirs, X + 1)
X = DelTree(strDir & "\" & strFile)
If X Then DelTree = X: Exit Function
Loop
RmDir strDir
If Err Then
DelTree = Err
Else
DelTree = 0
End If
End Function

Public Sub SaveOption(ByVal Section As String, ByVal Value As String)
SaveSet App.ProductName, "option", Section, Value
End Sub

Public Function GetOption(ByVal Section As String, Optional ByVal Default As String) As String
GetOption = GetSet(App.ProductName, "option", Section, Default)
End Function
