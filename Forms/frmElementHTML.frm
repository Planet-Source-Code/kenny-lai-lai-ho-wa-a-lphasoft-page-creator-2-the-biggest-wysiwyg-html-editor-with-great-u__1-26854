VERSION 5.00
Begin VB.Form frmElementHTML 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  '¨S¦³®Ø½u
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin PageCreator.CoolButton cmdCancel 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PageCreator.CoolButton cmdOK 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '¥­­±
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  '««ª½±²¶b
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmElementHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Value As String

Private Sub cmdCancel_Click()
cmdCancel1_Click
End Sub

Private Sub cmdCancel1_Click()
Value = "-1"
Me.Hide
MfrmProgram.SetFocus
Unload Me
End Sub

Private Sub cmdOK_Click()
cmdOK1_Click
End Sub

Private Sub cmdOK1_Click()
Value = Text1.Text
Me.Hide
MfrmProgram.SetFocus
Unload Me
End Sub
