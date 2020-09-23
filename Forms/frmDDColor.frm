VERSION 5.00
Begin VB.Form frmDDColor 
   BorderStyle     =   4  '³æ½u©T©w¤u¨ãµøµ¡
   ClientHeight    =   1305
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmDDColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private K() As cControlFlater, i As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()

'Flatted all commom controls
    Dim CTL As Control
    ReDim Preserve K(0 To Me.Controls.Count)
    For Each CTL In Me.Controls
        Select Case TypeName(CTL)
        Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar", "ListBox"
            Set K(i) = New cControlFlater
            K(i).Attach CTL
            i = i + 1
        End Select
    Next CTL


End Sub
