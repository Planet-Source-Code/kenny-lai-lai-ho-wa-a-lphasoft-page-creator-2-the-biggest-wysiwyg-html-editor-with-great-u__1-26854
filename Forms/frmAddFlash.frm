VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Object = "{05B9F8C4-05D2-11D1-A081-444553540000}#1.0#0"; "newTree.ocx"
Object = "{9F631458-BEE6-11D3-AFAF-9F131A29873D}#1.7#0"; "Tree.ocx"
Begin VB.Form frmAddFlash 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Add Flash Movie ..."
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "frmAddFlash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5760
      Top             =   3000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CheckBox chkVisible 
      Caption         =   "Visible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Top             =   5280
      Value           =   1  '®Ö¨ú
      Width           =   855
   End
   Begin VB.ComboBox cobAlign 
      Height          =   300
      ItemData        =   "frmAddFlash.frx":000C
      Left            =   5400
      List            =   "frmAddFlash.frx":002E
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtHeight 
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Text            =   "128"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Text            =   "128"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame fam1 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   5175
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
         Height          =   4095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4935
         _cx             =   4203009
         _cy             =   4201527
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   0   'False
         Base            =   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
   End
   Begin NEWEXLib.ExplorerTree tree1 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   4260
      _StockProps     =   161
      BackColor       =   16777215
   End
   Begin ExplorerCtls.asxFileListView file1 
      Height          =   2415
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4260
      FlatHeaders     =   0
      AllowViewMenu   =   0   'False
      Pattern         =   "*.swf"
      AllowDrives     =   0   'False
      IncludeFolders  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   2
      AllowProperties =   -1  'True
      PictureAlignment=   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '³z©ú
      Caption         =   "Align"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '³z©ú
      Caption         =   "Height"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '³z©ú
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FlashCode As String

Private Type FlashObject
Filename As String
Width As Long
Height As Long
Align As String
Visible As Boolean
End Type

Private Sub cmdCancel_Click()
FlashCode = "NO"
Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim f As FlashObject
With f
.Filename = Flash1.Movie
On Error GoTo 2
.Height = txtHeight.Text
.Width = txtWidth.Text
.Align = Trim(cobAlign.Text)
.Visible = IIf(chkVisible.Value = 0, False, True)
End With
FlashCode = GenerateCode(f)
Me.Hide
Exit Sub
2
MsgBox Error
End Sub

Private Sub file1_FileClick(ByVal File As String)
Flash1.Movie = File
End Sub

Private Sub Timer1_Timer()
If Flash1.Movie = "" Then
txtWidth.Enabled = False
txtHeight.Enabled = False
cobAlign.Enabled = False
chkVisible.Enabled = False
cmdOK.Enabled = False
Else
txtWidth.Enabled = Not False
txtHeight.Enabled = Not False
cobAlign.Enabled = Not False
chkVisible.Enabled = Not False
cmdOK.Enabled = Not False
End If
End Sub

Private Sub tree1_OnDirChanged()
file1.Path = tree1.Path
End Sub

Private Function GenerateCode(Flash As FlashObject) As String
GenerateCode = "<embed scr=" & AP & Flash.Filename & AP & " " & _
                                "width=" & AP & Flash.Width & AP & " " & _
                                "height=" & AP & Flash.Height & AP & " " & _
                                "align=" & AP & Flash.Align & AP & _
                                IIf(Flash.Visible = False, " hidden>", ">")
End Function
