VERSION 5.00
Begin VB.Form frmCusMenu 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Customize Menu..."
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCusMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Menu Colors Properties..."
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   4455
      Begin PageCreator.ColorPick cpHighFore 
         Height          =   345
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   609
         Text            =   "Highlight ForeColor"
      End
      Begin PageCreator.ColorPick cpBorder 
         Height          =   345
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   609
         Text            =   "Highlight Border Color"
      End
      Begin PageCreator.ColorPick cpHighlight 
         Height          =   345
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   609
         Text            =   "Highlight Color"
      End
      Begin PageCreator.ColorPick cpFore 
         Height          =   345
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   609
         Text            =   "Menu ForeColor"
      End
      Begin VB.Label Label1 
         Caption         =   "Chaning these colors will change the appearance of your program menu."
         Height          =   2055
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Background Image"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use custom background picture"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmCusMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ApplyValue()

With MfrmProgram.PM

    If Check1.Value = 1 Then
        SaveSet App.ProductName, "option", "MenuStyle", 0
        SaveSet App.ProductName, "option", "CustomMenuBackground", Text1.Text
        Set MfrmProgram.PM.BackgroundPicture = LoadPicture(Text1.Text)
    End If
    
    .ForeColor = cpFore.Color
    .borderColor = cpBorder.Color
    .HighlightColor = cpHighlight.Color
    .HighlightForeColor = cpHighFore.Color
    
    SaveOption "mnuFore", cpFore.Color
    SaveOption "mnuBorder", cpBorder.Color
    SaveOption "mnuHighlight", cpHighlight.Color
    SaveOption "mnuHighFore", cpHighFore.Color
    
End With

End Sub

Public Sub GetValue()
On Error Resume Next
With Me
    .Check1.Value = IIf(GetOption("MenuStyle") = 0, 1, 0)
    .Text1.Text = GetSet(App.ProductName, "option", "CustomMenuBackground", "")
    .cpFore.Color = MfrmProgram.PM.ForeColor
    .cpBorder.Color = MfrmProgram.PM.borderColor
    .cpHighFore.Color = MfrmProgram.PM.HighlightForeColor
    .cpHighlight.Color = MfrmProgram.PM.HighlightColor
End With

End Sub

Private Sub Check1_Click()
Text1.Enabled = Check1.Value
cmdBrowse.Enabled = Check1.Value
End Sub

Private Sub cmdBrowse_Click()

    With MfrmProgram.cd1
        
        .Flags = cdlOFNFileMustExist
        .DialogTitle = "Please select an image as your menu background."
        .Filter = "All Image(*.jpg,*.gif,*.bmp)|*.jpg;*.jpeg;*.gif;*.bmp"
        .CancelError = True
        On Error GoTo EndSub
        .ShowOpen
        Text1.Text = .FileName
        
    End With
    Text1.SetFocus
EndSub:
End Sub

Private Sub cmdCancel_Click()
Me.Hide
MfrmProgram.SetFocus
Unload Me
End Sub


Private Sub cmdOK_Click()
Me.Hide
MfrmProgram.SetFocus
ApplyValue
Unload Me
End Sub

Private Sub Form_Load()
Me.cpBorder.MakeMeFlat
Me.cpFore.MakeMeFlat
Me.cpHighFore.MakeMeFlat
Me.cpHighlight.MakeMeFlat
Me.GetValue
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
