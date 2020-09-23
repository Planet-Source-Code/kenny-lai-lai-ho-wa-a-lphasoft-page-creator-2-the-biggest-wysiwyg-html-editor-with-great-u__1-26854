VERSION 5.00
Object="{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Object="{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object="{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}#4.0#0"; "mshtml.tlb"
Object="{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object="{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object="{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object="{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object="{AE80D66B-8000-11D2-8B17-600109C10000}#5.0#0"; "OutlookBar.ocx"
Object="{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object="{05B9F8C4-05D2-11D1-A081-444553540000}#1.0#0"; "newex.ocx"
Begin VB.Form frmAddCustom 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Add Custom HTML Page..."
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmAddCustom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   30
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   29
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Color Properties"
      Height          =   1335
      Left            =   0
      TabIndex        =   18
      Top             =   3360
      Width           =   5775
      Begin VB.PictureBox picText 
         Appearance      =   0  '¥­­±
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.PictureBox picInUse 
         Appearance      =   0  '¥­­±
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   27
         Top             =   600
         Width           =   1215
      End
      Begin VB.PictureBox picVisited 
         Appearance      =   0  '¥­­±
         BackColor       =   &H00800080&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox picHyperlinks 
         Appearance      =   0  '¥­­±
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   25
         Top             =   840
         Width           =   1215
      End
      Begin VB.PictureBox picBackground 
         Appearance      =   0  '¥­­±
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   1185
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Links in use"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Hyperlinks"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Text"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Background"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Visited Links"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "BackGround"
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   5775
      Begin VB.CheckBox chkFixBPicture 
         Caption         =   "Fix Background"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmdOpenSound 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtLoopTime 
         Height          =   270
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "1"
         Top             =   720
         Width           =   615
      End
      Begin MSComCtl2.UpDown UD1 
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkAlwaysLoop 
         Caption         =   "Always Loop"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   720
         Value           =   1  '®Ö¨ú
         Width           =   1335
      End
      Begin VB.CommandButton cmdOpenPicture 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   10
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtBPicture 
         Height          =   270
         Left            =   1560
         TabIndex        =   9
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtBSound 
         Height          =   270
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label3 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Loop Times "
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Back Sound"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Back Picture"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin VB.TextBox txtTitle 
         Height          =   270
         Left            =   1560
         TabIndex        =   4
         Text            =   "New Page by Alphasoft"
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtDocumentPlace 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Title"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '¾a¥k¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Document Place"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackStyle       =   0  '³z©ú
      Caption         =   "Enter properties of your new page."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmAddCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Place As String

Private Sub chkAlwaysLoop_Click()
txtLoopTime.text = -1 * chkAlwaysLoop.Value
UD1.Enabled = 1 - chkAlwaysLoop.Value
txtLoopTime.Enabled = 1 - chkAlwaysLoop.Value
End Sub

Private Sub cmdOpenPicture_Click()
With cd1
.CancelError = True
.Flags = cdlOFNFileMustExist
.Filter = "All Files|*.*|All Graphic|*.gif;*.jpg;*.bmp;*.tif"
On Error GoTo 1
.ShowOpen
txtBPicture.text = cd1.filename
End With
1
End Sub

Private Sub cmdOpenSound_Click()
With cd1
.CancelError = True
.Flags = cdlOFNFileMustExist
.Filter = "All Sound Files: wav, mid, ram, ra, aif, au|*.wav;*.mid;*.ram;*.ra;*.aif;*.au"
On Error GoTo 1
.ShowOpen
txtBSound.text = cd1.filename
End With
1 End Sub

Private Sub Command1_Click()
AddCustomPage Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
PrepareUI
End Sub

Sub PrepareUI()
chkAlwaysLoop_Click
Me.txtDocumentPlace.text = Place
End Sub




Private Sub picBackground_Click()
With cd1
.CancelError = True
On Error GoTo 1
.ShowColor
picBackground.BackColor = cd1.Color
End With
1
End Sub

Private Sub picHyperlinks_Click()
With cd1
.CancelError = True
On Error GoTo 1
.ShowColor
picHyperlinks.BackColor = cd1.Color
End With
1
End Sub

Private Sub picInUse_Click()
With cd1
.CancelError = True
On Error GoTo 1
.ShowColor
picInUse.BackColor = cd1.Color
End With
1
End Sub

Private Sub picText_Click()
With cd1
.CancelError = True
On Error GoTo 1
.ShowColor
picText.BackColor = cd1.Color
End With
1
End Sub

Private Sub picVisited_Click()
With cd1
.CancelError = True
On Error GoTo 1
.ShowColor
picVisited.BackColor = cd1.Color
End With
1
End Sub

Private Sub UD1_DownClick()
Dim n As Integer
n = txtLoopTime.text
If n <= 0 Then Exit Sub
txtLoopTime.text = n - 1
End Sub

Private Sub UD1_UpClick()
Dim n As Integer
n = txtLoopTime.text
txtLoopTime.text = n + 1
End Sub
