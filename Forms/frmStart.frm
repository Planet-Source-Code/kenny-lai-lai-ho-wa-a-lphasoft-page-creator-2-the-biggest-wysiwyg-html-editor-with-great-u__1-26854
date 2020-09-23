VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{9F631458-BEE6-11D3-AFAF-9F131A29873D}#1.7#0"; "Tree.ocx"
Object = "{05B9F8C4-05D2-11D1-A081-444553540000}#1.0#0"; "newTree.ocx"
Begin VB.Form frmStart 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Start your work"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5475
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4200
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":0442
            Key             =   "temp"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":1294
            Key             =   "pcs"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":16E6
            Key             =   "webpage"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":1FC0
            Key             =   "template"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   794
      BackColor       =   -2147483638
      MouseIcon       =   "frmStart.frx":289A
      TabCaption(0)   =   "Create"
      TabPicture(0)   =   "frmStart.frx":28B6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstCreate"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Templates"
      TabPicture(1)   =   "frmStart.frx":3590
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCancel2"
      Tab(1).Control(1)=   "cmdOK2"
      Tab(1).Control(2)=   "lstTemplates"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Open"
      TabPicture(2)   =   "frmStart.frx":426A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdCancel3"
      Tab(2).Control(1)=   "Tree1"
      Tab(2).Control(2)=   "File1"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "History"
      TabPicture(3)   =   "frmStart.frx":4B44
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdCancel4"
      Tab(3).Control(1)=   "cmdOK4"
      Tab(3).Control(2)=   "lstHistory"
      Tab(3).ControlCount=   3
      Begin VB.CommandButton cmdCancel4 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   -71040
         TabIndex        =   14
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK4 
         Caption         =   "&OK"
         Height          =   375
         Left            =   -72360
         TabIndex        =   13
         Top             =   3840
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstHistory 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   12
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Recent File Records"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.CommandButton cmdCancel3 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   -71040
         TabIndex        =   11
         Top             =   3840
         Width           =   1215
      End
      Begin NEWEXLib.ExplorerTree Tree1 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   10
         Top             =   600
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   5318
         _StockProps     =   161
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin ExplorerCtls.asxFileListView File1 
         Height          =   3015
         Left            =   -72840
         TabIndex        =   9
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   5318
         FlatHeaders     =   0
         Pattern         =   "*.htm;*.html"
         IncludeFolders  =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   2
         PictureAlignment=   0
      End
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   -71040
         TabIndex        =   8
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK2 
         Caption         =   "&OK"
         Height          =   375
         Left            =   -72360
         TabIndex        =   7
         Top             =   3840
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstTemplates 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   6
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5318
         View            =   3
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Templates"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   3840
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstCreate 
         Height          =   3015
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5318
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000A&
      Caption         =   "Show this dialog every time."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   0
      Picture         =   "frmStart.frx":6C7E
      ScaleHeight     =   675
      ScaleWidth      =   5235
      TabIndex        =   5
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
SaveSet App.ProductName, "option", "startup", Check1.Value
End Sub

Private Sub cmdCancel_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdCancel3_Click()
cmdCancel_Click
End Sub

Private Sub cmdCancel4_Click()
cmdCancel_Click
End Sub

Private Sub cmdOK_Click()
Select Case lstCreate.SelectedItem.Key

    Case "webpage"
    Me.Hide
    MfrmProgram.SetFocus
    DoEvents
    MfrmProgram.NewBlankPage
    Unload Me
    
    Case "website"
    MsgBox "Sorry, not support yet!", vbInformation
    Exit Sub
    
End Select
End Sub

Private Sub cmdOK2_Click()
DoEvents
On Error Resume Next
Me.Hide
MfrmProgram.SetFocus
DoEvents
OpenTemp lstTemplates.SelectedItem.Key, lstTemplates.SelectedItem.Text
Unload Me
End Sub



Private Sub cmdOK4_Click()
DoEvents
On Error Resume Next
Me.Hide
MfrmProgram.SetFocus
DoEvents
OpenFile lstHistory.SelectedItem.Text, lstHistory.SelectedItem.Text
Unload Me
End Sub

Private Sub File1_FileDblClick(ByVal File As String)
DoEvents
On Error Resume Next
Me.Hide
MfrmProgram.SetFocus
DoEvents
OpenFile File, File
Unload Me
End Sub

Sub ReadHistory()
Dim HistoryCount As Integer
Dim URLtoRead As String
Dim num As Long
HistoryCount = GetSet(App.ProductName, "history", "historycount", 0)
num = HistoryCount
If num > 100 Then
num = 100
End If
On Error Resume Next
For i = 1 To num
URLtoRead = GetSet(App.ProductName, "history", "file" & i, "")
lstHistory.ListItems.Add , URLtoRead, URLtoRead, , "webpage"
Next i
lstHistory.ListItems(lstHistory.ListItems.Count).Selected = True
End Sub

Private Sub Form_Load()

If GetSet(App.ProductName, "option", "startup", 1) = 0 Then
    GoTo 1
End If

Check1.Value = GetSet(App.ProductName, "option", "startup", 1)

AddItem "webpage", "Blank Web Page", "webpage"
AddItem "website", "Web Site Project", "pcs"

AddTemp "Bibliography", "Description of books."
AddTemp "FAQ answers", "A response page of the FAQ section, contain some kinds of solution."
AddTemp "Feedback Confirmation", "Show it after users write feedback to you."
AddTemp "Feedback Form"
AddTemp "Guest Book"
AddTemp "Left Content", "Photo on right and content on left."
AddTemp "One Column Body with SideBar", "Content on the centre, Links on the side. Like that of PSC."
AddTemp "One Column Body"
AddTemp "Right Content", "Photo on left and content on right."
AddTemp "Search Page"
AddTemp "Table of Content"
AddTemp "Three Column Body"
AddTemp "Two Column Body"
AddTemp "User Registration"
AddTemp "Wide Body", "Like that in a news paper."

ReadHistory
Exit Sub
1
Me.Hide
MfrmProgram.SetFocus
Unload Me

End Sub

Sub AddTemp(ByVal Text As String, Optional ByVal Description As String)

lstTemplates.ListItems.Add , App.Path & "\Template\" & Text & ".htm", Text, 1, 1

If IsMissing(Description) = True Then Exit Sub

lstTemplates.ListItems(App.Path & "\Template\" & Text & ".htm").ListSubItems.Add , , Description

End Sub

Sub AddItem(ByVal Key As String, ByVal Text As String, ByVal Icon As String)
lstCreate.ListItems.Add , Key, Text, Icon
End Sub


Private Sub lstCreate_DblClick()
cmdOK_Click
End Sub


Private Sub lstHistory_DblClick()
cmdOK4_Click
End Sub

Private Sub lstTemplates_DblClick()
cmdOK2_Click
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
    Case 3
    lstHistory.SetFocus
End Select
End Sub

Private Sub Tree1_OnDirChanged()
On Error Resume Next
file1.Path = tree1.Path
End Sub
