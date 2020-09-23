VERSION 5.00
Begin VB.Form frmProp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Page Style"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin PageCreator.CoolButton cmdApply 
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Apply"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PageCreator.CoolButton cmdCancel 
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PageCreator.CoolButton cmdOK 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Page Margins"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   5775
      Begin VB.TextBox txtRight 
         Height          =   375
         Left            =   4440
         TabIndex        =   14
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtLeft 
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtBottom 
         Height          =   330
         Left            =   1800
         TabIndex        =   12
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtTop 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  '¾a¥k¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Right:"
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  '¾a¥k¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Left:"
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  '¾a¥k¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Bottom:"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  '¾a¥k¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "Top:"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   320
         Width           =   975
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   5655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Background Image"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   5775
      Begin PageCreator.CoolButton cmdBrowse 
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Browse"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "·s²Ó©úÅé"
            Size            =   8.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkFix 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fixed Background"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtBImage 
         Height          =   330
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Colors"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5775
      Begin PageCreator.ColorPick pic1 
         Height          =   345
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         Text            =   "Body Background"
      End
      Begin PageCreator.ColorPick pic2 
         Height          =   345
         Left            =   1320
         TabIndex        =   20
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         Text            =   "Content"
      End
      Begin PageCreator.ColorPick pic3 
         Height          =   345
         Left            =   4080
         TabIndex        =   21
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         Text            =   "Hyperlinks"
      End
      Begin PageCreator.ColorPick pic4 
         Height          =   345
         Left            =   3120
         TabIndex        =   22
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         Text            =   "Visited Links"
      End
      Begin PageCreator.ColorPick pic5 
         Height          =   345
         Left            =   2280
         TabIndex        =   23
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         Text            =   "Links in use"
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '³z©ú
      Caption         =   "Title"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DHTML As DHTMLEdit

Sub GetValue()
txtTitle.Text = DHTML.DocumentTitle
pic1.Color = HTML2RGB(DHTML.DOM.bgcolor)
pic2.Color = HTML2RGB(DHTML.DOM.fgColor)
pic3.Color = HTML2RGB(DHTML.DOM.linkColor)
pic4.Color = HTML2RGB(DHTML.DOM.vlinkColor)
pic5.Color = HTML2RGB(DHTML.DOM.alinkColor)
txtBImage.Text = DHTML.DOM.body.Background
chkFix.Value = IIf(DHTML.DOM.body.bgProperties = "fixed", 1, 0)
txtTop.Text = DHTML.DOM.body.topMargin
txtBottom.Text = DHTML.DOM.body.bottomMargin
txtLeft.Text = DHTML.DOM.body.leftMargin
txtRight.Text = DHTML.DOM.body.rightMargin
End Sub

Sub SetValue()
'With frmMain.DHTML1.DOM
With MfrmProgram.ActiveForm.DHTML1.DOM
.bgcolor = RGB2HTML(pic1.Color)
.fgColor = RGB2HTML(pic2.Color)
.linkColor = RGB2HTML(pic3.Color)
.vlinkColor = RGB2HTML(pic4.Color)
.alinkColor = RGB2HTML(pic5.Color)
.body.Background = txtBImage.Text
.body.bgProperties = IIf(chkFix.Value = 1, "fixed", "")
.body.topMargin = txtTop.Text
.body.bottomMargin = txtBottom.Text
.body.leftMargin = txtLeft.Text
.body.rightMargin = txtRight.Text
End With
End Sub

Private Sub cmdApply_Click()
SetValue
End Sub

Private Sub cmdBrowse_Click()
With MfrmProgram.cd1
.Filter = "GIF Images|*.gif|JPEG Images|*.jpg|Windows Bitmap|*.bmp|TIFF Images|*.tif;*.tiff|All Files|*.*"
.CancelError = True
.DialogTitle = "Select an image..."
On Error GoTo 1
.ShowOpen
txtBImage.Text = .FileName
Exit Sub
1 If Not (Err.Number = 32755) Then
MsgBox Err.Number & ": " & Error, vbCritical
End If
End With
End Sub

Private Sub cmdCancel_Click()
cmdOK.SetFocus
Unload Me
End Sub

Private Sub cmdOK_Click()
SetValue
Unload Me
End Sub

Private Sub Form_Load()
pic1.MakeMeFlat
pic2.MakeMeFlat
pic3.MakeMeFlat
pic4.MakeMeFlat
pic5.MakeMeFlat
GetValue
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtButtom_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtRight_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtTop_GotFocus()
txtTop.SelStart = 0
txtTop.SelLength = Len(txtTop.Text)
End Sub

Private Sub txtLeft_GotFocus()
txtLeft.SelStart = 0
txtLeft.SelLength = Len(txtLeft.Text)
End Sub

Private Sub txtRight_GotFocus()
txtRight.SelStart = 0
txtRight.SelLength = Len(txtRight.Text)
End Sub

Private Sub txtBottom_GotFocus()
txtBottom.SelStart = 0
txtBottom.SelLength = Len(txtBottom.Text)
End Sub
