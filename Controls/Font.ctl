VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl Font 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageCombo ic1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16777215
      ImageList       =   "iml"
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   4200
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "Font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Click()
Public Event Change()
Dim Flags As String
Private K() As cControlFlater, i As Long
Dim FontSuc As Boolean

Private Sub ic1_Click()
RaiseEvent Click
End Sub

Public Property Let Enabled(ena As Boolean)
ic1.Enabled = ena
End Property
Public Property Get Enabled() As Boolean
Enabled = ic1.Enabled
End Property

Public Property Let Text(str As String)
On Error GoTo 1
    If FontSuc = False Then
        ic1.Text = str
    Else
        ic1.ComboItems(str).Selected = True
    End If
1 End Property

Public Property Get Text() As String
Text = ic1.Text
End Property

Public Sub MakeText()
ic1.ComboItems.Clear
For i = 0 To Screen.FontCount - 1
ic1.ComboItems.Add , , Screen.Fonts(i)
Next
End Sub

Public Sub MakeFonts()

If FontSuc = True Then Exit Sub

ic1.ComboItems.Clear

With ic1
pb1.Move .Left, .Top, .width, .height
End With

pb1.Visible = True
pb1.Max = Screen.FontCount
pb1.Value = 0

iml.ListImages.Clear

With img
        .CurrentX = 0
        .CurrentY = 0
        .width = ic1.width
        .height = ic1.height
End With

For i = 0 To Screen.FontCount - 1

    With img
        .Cls
        .FontItalic = False
        .FontBold = False
        .FontName = Screen.Fonts(i)
        .FontSize = 10
        DoEvents
        img.Print Screen.Fonts(i)
    End With
    
    iml.ListImages.Add , , img.Image
    
    Set ic1.ImageList = iml
    On Error Resume Next
    ic1.ComboItems.Add , Screen.Fonts(i), Screen.Fonts(i), i + 1
    pb1.Value = i + 1
    
Next

pb1.Visible = False

FontSuc = True

End Sub

Private Sub ic1_Dropdown()
If FontSuc = False Then Me.MakeFonts
End Sub

Private Sub ic1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Public Sub MakeMeFlat()

Dim CTL As Control

    ReDim Preserve K(0 To UserControl.Controls.Count)
    
    For Each CTL In UserControl.Controls
        Select Case TypeName(CTL)
        Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar", "ListBox"
            Set K(i) = New cControlFlater
            K(i).Attach CTL
            i = i + 1
        End Select
    Next CTL

End Sub

Private Sub UserControl_InitProperties()
FontSuc = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Me.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
UserControl.height = 300
ic1.height = 300
ic1.width = UserControl.width
ic1.Move 0, 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Enabled", Me.Enabled, True
End Sub
