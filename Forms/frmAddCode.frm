VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAddCode 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Add Custom Code..."
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   Icon            =   "frmAddCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5085
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Database\EffectLibraryDatabase.mdb"
      DefaultCursorType=   0  '¹w³]ªº¸ê®Æ«ü¼Ð
      DefaultType     =   2  '¨Ï¥Î ODBCDirect
      Exclusive       =   0   'False
      Height          =   285
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '°ÊºA¶°(Dynaset)
      RecordSource    =   "Custom"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin PageCreator.synHighlight99 syn 
      Left            =   -600
      Top             =   4560
      _ExtentX        =   4233
      _ExtentY        =   609
      AttribCol       =   192
      CommentCol      =   32768
      TagCol          =   1.67117e7
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add to Database"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame fam2 
      Caption         =   "Content"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4815
      Begin RichTextLib.RichTextBox rt1 
         DataField       =   "Content"
         DataSource      =   "Data1"
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4260
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmAddCode.frx":014A
      End
   End
   Begin VB.Frame fam1 
      Caption         =   "Code Name and Language"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox Combo1 
         DataField       =   "Language"
         DataSource      =   "Data1"
         Height          =   300
         ItemData        =   "frmAddCode.frx":01F7
         Left            =   1320
         List            =   "frmAddCode.frx":0213
         TabIndex        =   1
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         DataField       =   "Name"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmAddCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Data1.Recordset.Update
Unload Me
End Sub

Private Sub cmdClear_Click()
Text1.Text = ""
rt1.Text = ""
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Database\EffectLibraryDatabase.mdb"
Data1.Refresh
Set syn.RichTxtBox = rt1
syn.AttribCol = RGB(255, 0, 0)
syn.CommentCol = &H8000&
syn.TagCol = &HFF0000
syn.TextCol = &H0&
End Sub


Private Sub rt1_KeyPress(KeyAscii As Integer)
On Error Resume Next

syn.KeyPressEvent KeyAscii
End Sub
