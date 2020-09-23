VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMSearch 
   Caption         =   "Alphasoft Quick Multi-Searcher"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '¨t²Î¹w³]­È
   WindowState     =   2  '³Ì¤j¤Æ
   Begin MSComctlLib.StatusBar sa1 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2940
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
   Begin TabDlg.SSTab st1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "Altavista"
      TabPicture(0)   =   "frmMSearch.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Astalavista"
      TabPicture(1)   =   "frmMSearch.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Ask"
      TabPicture(2)   =   "frmMSearch.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Excite"
      TabPicture(3)   =   "frmMSearch.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "HotBot"
      TabPicture(4)   =   "frmMSearch.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "InfoSeek"
      TabPicture(5)   =   "frmMSearch.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "MSN"
      TabPicture(6)   =   "frmMSearch.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Lycos"
      TabPicture(7)   =   "frmMSearch.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Yahoo"
      TabPicture(8)   =   "frmMSearch.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2295
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2295
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2295
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2295
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2295
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2295
      Index           =   6
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2295
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2295
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2295
      Index           =   9
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmMSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Search Combo1.Text
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
Combo1_Click
SaveHistory (Combo1.Text)
End If
End Sub
Sub ReadHistory()
Dim HistoryCount As Integer
Dim URLtoRead As String
Dim num As Long
HistoryCount = GetSetting(Me.Caption, "history", "historycount", 0)
num = HistoryCount
If num > 30 Then
num = 30
End If
On Error Resume Next
For i = 1 To num
URLtoRead = GetSetting(Me.Caption, "history", "url" & i, "")
Combo1.AddItem URLtoRead
Next i
End Sub

Function SaveHistory(ByVal URL As String)
Dim HistoryCount As Currency
Dim HistoryFull As Boolean
Dim num As Long
Dim HistoryReg As Currency

HistoryCount = GetSetting(Me.Caption, "history", "historycount", 0)

'Force number of history less than 30
HistoryReg = HistoryCount
Do Until HistoryReg < 30
HistoryReg = HistoryReg - 30
Loop

num = HistoryReg + 1
SaveSetting Me.Caption, "history", "url" & num, URL

SaveSetting Me.Caption, "history", "historycount", HistoryCount + 1

Combo1.AddItem URL
End Function
Private Sub Form_Load()
web(1).ZOrder 0
ReadHistory
End Sub

Private Sub Form_Resize()
On Error GoTo 1
Combo1.Move 0, 0, Me.ScaleWidth
st1.Move 0, Combo1.height, Me.ScaleWidth, Me.ScaleHeight - Combo1.height - sa1.height
For i = 1 To 9
web(i).Move 0, st1.TabHeight + Combo1.height, st1.width, st1.height - st1.TabHeight
Next
1 End Sub

Private Sub st1_Click(PreviousTab As Integer)
web(st1.Tab + 1).ZOrder 0
End Sub

Sub Search(ByVal Text As String)
web(1).navigate "http://www.altavista.com/cgi-bin/query?pg=q&kl=XX&stype=stext&q=" & Text
web(2).navigate "http://astalavista.box.sk/cgi-bin/astalavista/robot?srch=" & Text
web(3).navigate "http://www.ask.com/main/askJeeves.asp?ask=" & Text
web(4).navigate "http://www.excite.com/search.gw?search=" & Text
web(5).navigate "http://www.hotbot.com/?MT=" & Text
web(6).navigate "http://infoseek.go.com/Titles?qt=" & Text
web(7).navigate "http://search.msn.com/spbasic.htm?MT=" & Text
web(8).navigate "http://www.lycos.com/cgi-bin/pursuit?cat=dir&query=" & Text
web(9).navigate "http://ink.yahoo.com/bin/query?p=" & Text
End Sub

Private Sub web_ProgressChange(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo 1
Dim n As Long
n = Int(Progress / ProgressMax * 100)
sa1.Panels(1).Text = "Downloaded " & n & "%"
1 End Sub

Private Sub web_StatusTextChange(Index As Integer, ByVal Text As String)
sa1.Panels(2).Text = Text
End Sub
