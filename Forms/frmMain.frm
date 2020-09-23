VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Unknown Page"
   ClientHeight    =   6720
   ClientLeft      =   1710
   ClientTop       =   1725
   ClientWidth     =   7275
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   7275
   Visible         =   0   'False
   WindowState     =   2  '³Ì¤j¤Æ
   Begin RichTextLib.RichTextBox tpl 
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0442
   End
   Begin VB.TextBox NewPage 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmMain.frx":04EF
      Top             =   6720
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CheckBox cIsTmp 
      Caption         =   "cIsTmp"
      Height          =   300
      Left            =   840
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Value           =   1  '®Ö¨ú
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6000
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10583
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Design Mode"
      TabPicture(0)   =   "frmMain.frx":0652
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DHTML1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Core Mode"
      TabPicture(1)   =   "frmMain.frx":066E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pbSyntax"
      Tab(1).Control(1)=   "rt1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Preview Mode"
      TabPicture(2)   =   "frmMain.frx":068A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblStatus"
      Tab(2).Control(1)=   "webPreview"
      Tab(2).ControlCount=   2
      Begin MSComctlLib.ProgressBar pbSyntax 
         Height          =   255
         Left            =   -74520
         TabIndex        =   10
         Top             =   5160
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin SHDocVwCtl.WebBrowser webPreview 
         CausesValidation=   0   'False
         Height          =   4455
         Left            =   -74760
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   150
         Width           =   6975
         ExtentX         =   12303
         ExtentY         =   7858
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   1
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin RichTextLib.RichTextBox rt1 
         Height          =   4215
         Left            =   -73920
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   7435
         _Version        =   393217
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":06A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DHTMLEDLibCtl.DHTMLEdit DHTML1 
         Height          =   4695
         Left            =   480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   6015
         ActivateApplets =   -1  'True
         ActivateActiveXControls=   -1  'True
         ActivateDTCs    =   -1  'True
         ShowDetails     =   0   'False
         ShowBorders     =   -1  'True
         Appearance      =   0
         Scrollbars      =   -1  'True
         ScrollbarAppearance=   0
         SourceCodePreservation=   -1  'True
         AbsoluteDropMode=   -1  'True
         SnapToGrid      =   -1  'True
         SnapToGridX     =   20
         SnapToGridY     =   20
         BrowseMode      =   0   'False
         UseDivOnCarriageReturn=   0   'False
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  '³æ½u©T©w
         Caption         =   "Browser Statue"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   4800
         Visible         =   0   'False
         Width           =   6975
      End
   End
   Begin VB.CheckBox cIsRO 
      Caption         =   "cIsRO"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox cIsSave 
      Caption         =   "cIsSave"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsTmp As Boolean, IsSave As Boolean, IsRO As Boolean, IsA As Boolean
Dim MyName As String
Public FileName As String
Public Colorized As Boolean
Dim ctxtStdItemCount As Long
Dim SelT As String
Public en As Object, TabN As Integer
Dim WithEvents EleImg As HTMLImg
Attribute EleImg.VB_VarHelpID = -1
Dim WithEvents EleSpan As HTMLSpanElement
Attribute EleSpan.VB_VarHelpID = -1
Dim WithEvents EleA As HTMLAnchorElement
Attribute EleA.VB_VarHelpID = -1
Dim WithEvents EleFont As HTMLFontElement
Attribute EleFont.VB_VarHelpID = -1
Dim WithEvents EleBody As HTMLBody
Attribute EleBody.VB_VarHelpID = -1
Dim WithEvents EleTD As HTMLTableCell
Attribute EleTD.VB_VarHelpID = -1
Dim WithEvents EleTable As HTMLTable
Attribute EleTable.VB_VarHelpID = -1
Dim WithEvents EleInput As HTMLInputElement
Attribute EleInput.VB_VarHelpID = -1
Dim WithEvents EleHR As HTMLHRElement
Attribute EleHR.VB_VarHelpID = -1
Public DocTitle As String
Public DReady As String, Flags As String
Public File2Load As String
Public HTMLString As String

Private Sub cIsRO_Click()
IsRO = cIsRO.Value
End Sub

Private Sub cIsSave_Click()
IsSave = cIsSave.Value

End Sub

Private Sub cIsTmp_Click()
IsTmp = cIsTmp.Value
End Sub


Public Function GetActiveElement() As Object
DoEvents
Dim rg As IHTMLTxtRange
   Dim ctlRg As IHTMLControlRange
   On Error Resume Next
   Select Case DHTML1.DOM.selection.Type
      Case "None", "Text"
         ' This reduces the selection to just the insertion
         ' point. The parentElement method will then return the
         ' element directly under the mouse pointer.
         Set rg = DHTML1.DOM.selection.createRange
         rg.collapse False
         
         Set GetActiveElement = rg.parentElement
      Case "Control"
         ' A form or image is selected. The commonParentElement
         ' will return the site selected element.
         Set ctlRg = DHTML1.DOM.selection.createRange
         
         Set GetActiveElement = ctlRg.commonParentElement
         
         
   End Select
End Function

Private Function ElementP() As Object
DoEvents
    Dim e As IHTMLEventObj
    Set e = DHTML1.DOM.parentWindow.event
    Set ElementP = DHTML1.DOM.elementFromPoint(e.clientX, e.clientY)

End Function

Function MakeRainbowText(ByVal Text As String, ByVal FadeColor As Integer) As String

Dim s As String, st As String, p As String, out As String
s = Text
For i = 1 To Len(s)
st = Mid(s, i, 1)
Select Case FadeColor
Case 0
p = "<font color=""" & RGB2HTML(RGB(Abs(255 - (i * 2)), 0, 0)) & """" & ">" & st & "</font>"
Case 1
p = "<font color=""" & RGB2HTML(RGB(0, Abs(255 - (i * 2)), 0)) & """" & ">" & st & "</font>"
Case 2
p = "<font color=""" & RGB2HTML(RGB(0, 0, Abs(255 - (i * 2)))) & """" & ">" & st & "</font>"
End Select
'p = "<font color=""" & RGB2HTML(RGB(i, i, i)) & """" & ">" & st & "</font>"
out = out & p
Next
MakeRainbowText = out
End Function

Sub InRainbowText(ByVal FadeColor As Integer)
Dim rg As IHTMLTxtRange
Set rg = DHTML1.DOM.selection.createRange
s = rg.Text
rg.Text = ""
InsertHTML MakeRainbowText(s, FadeColor)
End Sub

Private Sub DHTML1_ContextMenuAction(ByVal itemIndex As Long)
MfrmProgram.mnuExt.Visible = False
End Sub

Private Sub DHTML1_DisplayChanged()
On Error GoTo 1
If DReady = "OK" And Flags = "OK" Then
MfrmProgram.RefreshEditBar

With MfrmProgram
On Error Resume Next
'DoEvents
If .cobFonts.Text <> DHTML1.execCommand(DECMD_GETFONTNAME) Then .cobFonts.Text = DHTML1.execCommand(DECMD_GETFONTNAME)
'DoEvents
If .cpBack.Color <> HTML2RGB(DHTML1.execCommand(DECMD_GETBACKCOLOR)) Then .cpBack.Color = HTML2RGB(DHTML1.execCommand(DECMD_GETBACKCOLOR))
'DoEvents
If .cpFore.Color <> HTML2RGB(DHTML1.execCommand(DECMD_GETFORECOLOR)) Then .cpFore.Color = HTML2RGB(DHTML1.execCommand(DECMD_GETFORECOLOR))
.cobFormat.Text = DHTML1.execCommand(DECMD_GETBLOCKFMT)
.cobSize.Text = DHTML1.execCommand(DECMD_GETFONTSIZE) & " (" & FontSizePoint(DHTML1.execCommand(DECMD_GETFONTSIZE)) & " pt)"
End With

End If
1
End Sub

Private Sub DHTML1_DocumentComplete()

If Me.Tag = "GetHTML" Then
    HTMLString = DHTML1.DocumentHTML
    Me.Tag = ""
End If

With MfrmProgram
.Timer1 = True
.Timer3 = True
.Timer4 = True
.Timer5 = True
.File1.Enabled = True
End With

DReady = "OK"

End Sub


Private Sub DHTML1_onclick()

DoEvents

Set en = GetActiveElement
SelT = en.innerText

End Sub

Sub DynaInfo()
DoEvents
On Error GoTo 1
MfrmProgram.staProp.Panels(1).Text = ElementP.tagName
Select Case UCase(ElementP.tagName)
    
    Case "IMG"
    Set EleImg = ElementP
    'MfrmProgram.staProp.Panels(2).Text = "Image Source=''" & EleImg.src & "''"
    
    Case "SPAN"
    Set EleSpan = ElementP
    'MfrmProgram.staProp.Panels(2).Text = "Span Element"
    
    Case "P"
    'MfrmProgram.staProp.Panels(2).Text = "Paragraph"
    
    Case "A"
    Set EleA = ElementP
    'MfrmProgram.staProp.Panels(2).Text = "Links to : " & EleA.href & "        Double Click to open the link."
    
    Case "FONT"
    Set EleFont = ElementP
    'MfrmProgram.staProp.Panels(2).Text = "Font Element"
    
    Case "BODY"
    Set EleBody = ElementP
    'MfrmProgram.staProp.Panels(2).Text = "Document Body Background"
    
    Case "TD"
    Set EleTD = ElementP
    'MfrmProgram.staProp.Panels(2).Text = "Text= " & EleTD.innerText & "       BgColor (R: " & HTML2typergb(EleTD.bgcolor).R & ")" & " (G: " & HTML2typergb(EleTD.bgcolor).G & ")" & " (B: " & HTML2typergb(EleTD.bgcolor).B & ")"
    
    Case "TABLE"
    Set EleTable = ElementP
    'MfrmProgram.staProp.Panels(2).Text = "Table Object"
    
    Case "INPUT"
    Set EleInput = ElementP
    'MfrmProgram.staProp.Panels(2).Text = "Type =" & EleInput.Type & "  Value =" & EleInput.Value & "  Name =" & EleInput.Name
    
    Case "HR"
    Set EleHR = ele
    'MfrmProgram.staProp.Panels(2).Text = "Horizontal Line"
    
End Select
1
End Sub




Private Sub DHTML1_onmousemove()
DoEvents
DynaInfo
End Sub

Private Sub DHTML1_ShowContextMenu(ByVal xPos As Long, ByVal yPos As Long)
'Show different menus with different type of elements
DoEvents
With MfrmProgram
    .extCut.Enabled = ButtonEnable(DECMD_CUT)
    .extCopy.Enabled = ButtonEnable(DECMD_COPY)
    .extPaste.Enabled = ButtonEnable(DECMD_PASTE)
    .extAll.Enabled = ButtonEnable(DECMD_SELECTALL)
    .extDelete.Enabled = ButtonEnable(DECMD_DELETE)
    
    .extAbs.Enabled = ButtonEnable(DECMD_MAKE_ABSOLUTE)
    .extAbs.Checked = ButtonPress(DECMD_MAKE_ABSOLUTE)
    .extDetail.Checked = DHTML1.ShowDetails
    .extSnap.Checked = DHTML1.SnapToGrid
    
    .extInCol.Enabled = ButtonEnable(DECMD_INSERTCOL)
    .extInRow.Enabled = ButtonEnable(DECMD_INSERTROW)
    .extMerge.Enabled = ButtonEnable(DECMD_MERGECELLS)
    .extSplit.Enabled = ButtonEnable(DECMD_SPLITCELL)
    
    .extBack.Enabled = ButtonEnable(DECMD_SEND_TO_BACK)
    .extBack.Visible = True
    .extBackward.Enabled = ButtonEnable(DECMD_SEND_BACKWARD)
    .extBackward.Visible = True
    .extBelowText.Enabled = ButtonEnable(DECMD_SEND_BELOW_TEXT)
    .extBelowText.Visible = True
    
    .extAboveText.Enabled = ButtonEnable(DECMD_BRING_ABOVE_TEXT)
    .extAboveText.Visible = True
    .extForeward.Enabled = ButtonEnable(DECMD_BRING_FORWARD)
    .extForeward.Visible = True
    .extFront.Enabled = ButtonEnable(DECMD_BRING_TO_FRONT)
    .extFront.Visible = True
    
    
End With


Dim e As IHTMLElement
Set e = GetActiveElement
With MfrmProgram
Select Case e.tagName


    Case "IMG"
    
    Case "SPAN"
    
    Case "P"
    
    Case "A"
    
    Case "FONT"
    
    Case "BODY"
    
    Case "TD"
    
    Case "TABLE"
    
    Case "INPUT"
    
    Case "HR"
    

End Select
End With

PopupMenu MfrmProgram.mnuExt

End Sub


Private Function EleA_ondblclick() As Boolean
    ShellExecute Me.hwnd, "open", EleA.href, "", "", 1
End Function

Private Sub EleA_onmouseover()
'Me.Tag = EleA.Style.backgroundColor

'EleA.Style.backgroundColor = RGB2HTML(RGB(255, 0, 0))

    MfrmProgram.staProp.Panels(2).Text = "Links to : " & EleA.href & "        Double Click to open the link."
End Sub

Private Sub EleA_onmouseout()
    MfrmProgram.staProp.Panels(2).Text = ""

'EleA.Style.backgroundColor = Me.Tag
End Sub


Private Sub EleBody_onmouseout()
    MfrmProgram.staProp.Panels(2).Text = ""

End Sub

Private Sub EleBody_onmouseover()
    'MfrmProgram.staProp.Panels(2).Text = "Document Body Background"

End Sub

Private Sub EleFont_onmouseout()
DoEvents
    MfrmProgram.staProp.Panels(2).Text = ""

End Sub

Private Sub EleFont_onmouseover()
DoEvents
    MfrmProgram.staProp.Panels(2).Text = "Font Element"

End Sub

Private Sub EleHR_onmouseout()
    MfrmProgram.staProp.Panels(2).Text = ""

End Sub

Private Sub EleHR_onmouseover()
    MfrmProgram.staProp.Panels(2).Text = "Horizontal Line"

End Sub

Private Function EleImg_ondblclick() As Boolean
DHTML1.execCommand DECMD_IMAGE
End Function

Private Sub EleImg_onmouseout()
    MfrmProgram.staProp.Panels(2).Text = ""

End Sub

Private Sub EleImg_onmouseover()
    MfrmProgram.staProp.Panels(2).Text = "Image Source=''" & EleImg.src & "''"

End Sub

Private Sub EleInput_onmouseout()
    MfrmProgram.staProp.Panels(2).Text = ""

End Sub

Private Sub EleInput_onmouseover()
    MfrmProgram.staProp.Panels(2).Text = "Type =" & EleInput.Type & "  Value =" & EleInput.Value & "  Name =" & EleInput.Name

End Sub

Private Sub EleSpan_onmouseout()

    MfrmProgram.staProp.Panels(2).Text = ""

End Sub

Private Sub EleSpan_onmouseover()
    MfrmProgram.staProp.Panels(2).Text = "Span Element"

End Sub

Private Sub EleTable_onmouseout()
    MfrmProgram.staProp.Panels(2).Text = ""

End Sub

Private Sub EleTable_onmouseover()
    MfrmProgram.staProp.Panels(2).Text = "Table Object"

End Sub

Private Sub EleTD_onmouseout()
    MfrmProgram.staProp.Panels(2).Text = ""

End Sub

Private Sub EleTD_onmouseover()
    MfrmProgram.staProp.Panels(2).Text = "Text= " & EleTD.innerText & "       BgColor (R: " & HTML2typeRGB(EleTD.bgcolor).R & ")" & " (G: " & HTML2typeRGB(EleTD.bgcolor).G & ")" & " (B: " & HTML2typeRGB(EleTD.bgcolor).B & ")"

End Sub

Private Sub Form_Activate()

DoEvents

IsA = True
LoadUICore SSTab1.Tab
DHTML1.ShowDetails = MfrmProgram.mnuDetail.Checked
On Error Resume Next
MfrmProgram.staMain.Refresh
MfrmProgram.staProp.Refresh

If TabN = 2 Then
With Me
.webPreview.Visible = True
.lblStatus.Visible = True
.lblStatus.ZOrder 0
.rt1.Visible = False
.DHTML1.Visible = False
End With
End If
MfrmProgram.FillHTMLTree
End Sub


Private Sub Form_Initialize()
IsA = False
IsTmp = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

Select Case SSTab1.Tab
Case 0
    If DHTML1.DocumentHTML = Me.HTMLString Then
        On Error Resume Next
        DoEvents
        Kill FileName
        MfrmProgram.SetFocus
        Exit Sub
    End If
Case 1
    If rt1.Text = Me.HTMLString Then
        On Error Resume Next
        DoEvents
        Kill FileName
        MfrmProgram.SetFocus
        Exit Sub
    End If
Case 2
End Select

If Me.Flags = "EXIT" Then Exit Sub

Dim m As Integer
m = MsgBox("Do you want to save change to " & Me.Caption & "?", vbQuestion + vbYesNoCancel, App.ProductName)
Select Case m

Case vbYes

    Select Case Me.cIsSave.Value
    Case 0: MfrmProgram.mnuSaveAs_Click
    Case 1: MfrmProgram.mnuSave_Click
    End Select
    
If MfrmProgram.SaveSuc = True Then
    On Error GoTo 1
    DoEvents
    Kill FileName
1
    Exit Sub
Else

    Cancel = 1

End If

Case vbNo
    On Error GoTo 2
    DoEvents
    Kill FileName
2
    Exit Sub

Case vbCancel
GoTo 9

End Select

9          Cancel = 1
End Sub


Public Sub PrepareUI()
On Error GoTo 1
rt1.Text = DHTML1.DocumentHTML
1
End Sub

Sub Resize()

'Resize Tab control
With SSTab1
.Move 0, 0
.width = Me.ScaleWidth
On Error GoTo 1
.height = Me.ScaleHeight

End With

'Resize all tabs
With DHTML1
.Top = 0
.Left = 0
.width = SSTab1.width
On Error GoTo 1
.height = SSTab1.height - SSTab1.TabHeight - 50
End With

With rt1

.Top = 0
.Left = 0
.width = SSTab1.width
On Error GoTo 1
.height = SSTab1.height - SSTab1.TabHeight - pbSyntax.height - 50
End With

pbSyntax.Move 0, rt1.Top + rt1.height, SSTab1.width

With webPreview
.Offline = True
.Top = 0
.Left = 0
.width = SSTab1.width
On Error GoTo 1
.height = SSTab1.height - SSTab1.TabHeight - lblStatus.height - 50
End With

With lblStatus
.Top = webPreview.height
.Left = 0
.width = SSTab1.width
End With

1 End Sub


Private Sub Form_Resize()
On Error GoTo 1
Resize
1
End Sub

Private Sub rt1_KeyDown(KeyCode As Integer, Shift As Integer)

DoEvents

'If KeyCode = vbKeyZ And Shift = 2 Then
'MsgBox "Undo"
'DoEvents
'Exit Sub
'End If

If KeyCode = vbKeyBack And rt1.SelLength <> 0 Then
KeyCode = vbKeyDelete
End If

'If KeyCode = vbKeyHome And Shift = 1 Or _
'    KeyCode = vbKeyEnd And Shift = 1 Or _
'    KeyCode = vbKeyA And Shift = 1 Then
'SetToolbar4rtf
'End If

End Sub

Public Sub rt1Undo()
rt1.SetFocus
SendKeys "^{z}"
End Sub

Private Sub rt1_KeyUp(KeyCode As Integer, Shift As Integer)

'If KeyCode = vbKeyHome And Shift = 1 Or _
'    KeyCode = vbKeyEnd And Shift = 1 Or _
'    KeyCode = vbKeyA And Shift = 1 Then
'SetToolbar4rtf
'End If

End Sub

Private Sub rt1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'SetToolbar4rtf
End Sub

Private Sub rt1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'SetToolbar4rtf
End Sub

Private Sub rt1_SelChange()
SetToolbar4rtf
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

TabN = SSTab1.Tab
If PreviousTab = 0 Then
On Error GoTo SaveError
DHTML1.SaveDocument FileName
SaveError:
If Err.Number = 5 Then MsgBox "Error Number: " & Err.Number & " Error saving file!", vbExclamation
End If


'LoadUICore SSTab1.Tab

Select Case SSTab1.Tab

Case 0

lblStatus.Visible = False
DHTML1.Visible = True
rt1.Visible = False
'webPreview.Visible = False
If PreviousTab = 1 Then CreateTMP
On Error GoTo CodeErr
DHTML1.LoadDocument FileName
CodeErr:

'Entered Error HTML Code
If Err.Number = -2146435068 Then
    MsgBox "Error HTML Code", vbExclamation
    a = rt1.Text
    SSTab1.Tab = 1
    rt1.Text = a
    'With syn
    '.AttribCol = &HC0&
    '.CommentCol = &H8000&
    '.TagCol = &HC00000
    '.TextCol = &H0&
    '.Highlight
    'End With
    DoEvents
    pbSyntax.Visible = True
    pbSyntax.ZOrder 0
    ColorHtml rt1
End If
DHTML1.ZOrder 0

Case 1
DHTML1.Visible = False
'webPreview.Visible = False
rt1.Visible = True
'rt1.Font.Charset = 136
    'With syn
    '.AttribCol = &HC0&
    '.CommentCol = &H8000&
    '.TagCol = &HC00000
    '.TextCol = &H0&
    'End With
    lblStatus.Visible = False
    rt1.LoadFile FileName
    'Set syn.RichTxtBox = rt1
    
    DoEvents
    'syn.Highlight
    pbSyntax.Visible = True
    pbSyntax.ZOrder 0
    ColorHtml rt1
    
    s = IIf(InStr(1, rt1.Text, Trim(SelT), vbTextCompare) = 0, 0, InStr(1, rt1.Text, Trim(SelT), vbTextCompare) - 1)
    L = IIf(InStr(1, rt1.Text, Trim(SelT)) = 0, 0, Len(SelT))
    
    rt1.ZOrder 0
    rt1.SelStart = s
    rt1.HideSelection = False
    rt1.SelLength = L

Case 2
With Me
.webPreview.Visible = True
.webPreview.ZOrder 0
.DHTML1.Visible = False
.rt1.Visible = False
.lblStatus.Visible = True
.lblStatus.ZOrder 0
pbSyntax.Visible = False
End With
If PreviousTab = 1 Then CreateTMP
webPreview.Offline = True
On Error Resume Next
webPreview.navigate FileName
End Select

Exit Sub
1 MsgBox Error, vbCritical
End Sub

Sub CreateTMP()
        rt1.SaveFile FileName, rtfText
        'Dim t As String
        't = FreeFile
        'Open Filename For Output As #t
        'Print #t, rt1.Text
        'Close #t
End Sub

Private Sub webPreview_StatusTextChange(ByVal Text As String)
On Error GoTo 1
lblStatus.Caption = Text
1 End Sub
