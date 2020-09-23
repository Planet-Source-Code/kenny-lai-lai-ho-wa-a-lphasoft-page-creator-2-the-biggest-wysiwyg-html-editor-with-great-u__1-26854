VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5235
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.TextBox picFile 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '««ª½±²¶b
      TabIndex        =   5
      Text            =   "frmAbout_Public.frx":0000
      Top             =   3360
      Width           =   4215
   End
   Begin VB.TextBox picComment 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '««ª½±²¶b
      TabIndex        =   4
      Text            =   "frmAbout_Public.frx":0006
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label lblExit 
      Alignment       =   2  '¸m¤¤¹ï»ô
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Image imaIcon 
      Height          =   855
      Left            =   480
      Top             =   840
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  '³z©ú
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  '¸m¤¤¹ï»ô
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '³z©ú
      Caption         =   "HTML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1695
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BorderColor     =   &H000080FF&
      Height          =   975
      Left            =   120
      Shape           =   4  '¶ê¨¤¯x§Î
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  '¤£³z©ú
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '¤£³z©ú
      BorderColor     =   &H00FFC0C0&
      FillColor       =   &H00FFC0C0&
      Height          =   1215
      Left            =   480
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblProductName 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblExit.BackColor = &H8000000F

End Sub

Private Sub lblExit_Click()
Unload Me
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblExit.BackColor = RGB(255, 100, 100)
End Sub
