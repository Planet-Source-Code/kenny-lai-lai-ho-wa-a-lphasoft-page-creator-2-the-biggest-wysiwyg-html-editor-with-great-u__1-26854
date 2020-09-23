VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDown 
   BorderStyle     =   4  '³æ½u©T©w¤u¨ãµøµ¡
   ClientHeight    =   1470
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton picCurrent 
      Enabled         =   0   'False
      Height          =   225
      Left            =   120
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1905
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   15
      Left            =   1800
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   14
      Left            =   1560
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   13
      Left            =   1320
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   12
      Left            =   1080
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   11
      Left            =   840
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   10
      Left            =   600
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   9
      Left            =   360
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   8
      Left            =   120
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   7
      Left            =   1800
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   6
      Left            =   1560
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   5
      Left            =   1320
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   4
      Left            =   1080
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   3
      Left            =   840
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   2
      Left            =   600
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   1
      Left            =   360
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   225
   End
   Begin VB.CommandButton picColor 
      Height          =   225
      Index           =   0
      Left            =   120
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   225
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   -240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "Customs..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private K() As cControlFlater, i As Integer
Public Value As Long

Private Sub cmdCancel_Click()
Value = -1
Me.Hide
End Sub

Private Sub cmdCustom_Click()
With cd1
.CancelError = True
.Flags = cdlCCFullOpen
On Error GoTo 1
.ShowColor
Value = .Color
Me.Hide
End With
Exit Sub
1
MsgBox Error
End Sub

Private Sub Form_Load()
MakeMeFlat

For i = 0 To 15
    picColor(i).BackColor = QBColor(i)
Next

End Sub

Sub MakeMeFlat()
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

Private Sub picColor_Click(Index As Integer)
Value = picColor(Index).BackColor
Me.Hide
End Sub
