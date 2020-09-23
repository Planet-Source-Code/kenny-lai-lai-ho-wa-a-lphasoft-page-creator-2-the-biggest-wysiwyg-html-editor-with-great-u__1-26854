VERSION 5.00
Begin VB.UserControl ColorPick 
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ScaleHeight     =   2970
   ScaleWidth      =   2685
   Begin VB.ComboBox CoB 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "ColorPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private K() As cControlFlater, i As Integer
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Dim InitText As String
Public Event DropDown()
Public Event Click()
Private CaptionText As String

Public Property Let Enabled(ena As Boolean)
CoB.Enabled = ena
End Property
Public Property Get Enabled() As Boolean
Enabled = CoB.Enabled
End Property

Public Property Let Text(str As String)

On Error GoTo 2
CaptionText = str
CoB.Text = str
Exit Property
2
'i = Err.Number
'Err.Raise i
End Property
Public Property Get Text() As String
Text = CoB.Text
End Property

Private Sub CoB_Change()
CoB.Text = CaptionText
End Sub

Private Sub CoB_DropDown()

RaiseEvent DropDown

Dim rt As RECT
GetWindowRect CoB.hwnd, rt
frmDown.Top = rt.Bottom * Screen.TwipsPerPixelY
frmDown.Left = rt.Left * Screen.TwipsPerPixelX
frmDown.picCurrent.BackColor = CoB.BackColor
frmDown.Show vbModal
If frmDown.Value = -1 Then
    Exit Sub
Else
    Color = frmDown.Value
End If
RaiseEvent Click
End Sub

Public Property Let Color(Col As OLE_COLOR)

On Error Resume Next

CoB.BackColor = Col

Dim Negative As TypeRGB
Dim Positive As TypeRGB
Positive = RGB2typeRGB(Col)
With Negative
.R = 255 - Positive.R
.G = 255 - Positive.G
.B = 255 - Positive.B
CoB.ForeColor = RGB(.R, .G, .B)
End With

End Property
Public Property Get Color() As OLE_COLOR
Color = CoB.BackColor
End Property

Private Sub CoB_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Me.Color = PropBag.ReadProperty("Color", RGB(255, 255, 255))
Me.Text = PropBag.ReadProperty("Text", "")
Me.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
UserControl.height = CoB.height
CoB.Move 0, 0, UserControl.width
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

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Color", Me.Color, RGB(255, 255, 255)
PropBag.WriteProperty "Text", Me.Text, ""
PropBag.WriteProperty "Enabled", Me.Enabled, True
End Sub
