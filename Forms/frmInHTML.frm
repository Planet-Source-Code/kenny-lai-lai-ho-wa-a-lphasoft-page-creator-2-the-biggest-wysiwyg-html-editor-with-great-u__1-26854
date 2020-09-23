VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInHTML 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Insert HTML"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel1"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK1 
      Caption         =   "&OK1"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin PageCreator.CoolButton cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rt1 
      CausesValidation=   0   'False
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmInHTML.frx":0000
   End
End
Attribute VB_Name = "frmInHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Form_Resize()
rt1.Move 0, 0, Me.ScaleWidth, Int(Me.ScaleHeight - cmdOK.height)

End Sub
