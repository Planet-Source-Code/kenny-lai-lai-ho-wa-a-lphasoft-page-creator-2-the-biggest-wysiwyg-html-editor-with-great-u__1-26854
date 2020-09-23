VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTool_Rainbow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Rainbow-Fade Text Wizard..."
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   Icon            =   "frmTool_Rainbow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5400
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin PageCreator.CoolButton Command3 
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Cancel"
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
   Begin PageCreator.CoolButton Command2 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Insert Code"
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
   Begin PageCreator.CoolButton Command1 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "Generate Code"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin PageCreator.ColorPick pic2 
         Height          =   345
         Left            =   2400
         TabIndex        =   10
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         Color           =   65280
      End
      Begin PageCreator.ColorPick pic1 
         Height          =   345
         Left            =   1200
         TabIndex        =   9
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   609
         Color           =   255
      End
      Begin RichTextLib.RichTextBox rt1 
         Height          =   975
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1720
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmTool_Rainbow.frx":014A
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '³z©ú
         Caption         =   "to"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '³z©ú
         Caption         =   "Fade from"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '³z©ú
         Caption         =   "Text"
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin RichTextLib.RichTextBox rtc 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2880
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3836
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmTool_Rainbow.frx":057E
   End
End
Attribute VB_Name = "frmTool_Rainbow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
GenerateCode
End Sub

Sub GenerateCode()
rtc.Text = RainbowColorText(rt1.Text, pic1.Color, pic2.Color)
End Sub

Private Sub Command2_Click()
Dim out As String
out = RainbowColorText(rt1.Text, pic1.Color, pic2.Color)
Select Case MfrmProgram.ActiveForm.SSTab1.Tab
    Case 0
    InsertHTML out
    Case 1
    MfrmProgram.ActiveForm.rt1.SelText = out
End Select
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub pic1_Click()
Me.rt1.SetFocus
End Sub

Private Sub pic2_Click()
Me.rt1.SetFocus
End Sub

