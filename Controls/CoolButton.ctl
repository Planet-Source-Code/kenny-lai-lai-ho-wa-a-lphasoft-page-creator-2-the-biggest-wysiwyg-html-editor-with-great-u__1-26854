VERSION 5.00
Begin VB.UserControl CoolButton 
   BackStyle       =   0  '³z©ú
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   ScaleHeight     =   3120
   ScaleWidth      =   4545
   ToolboxBitmap   =   "CoolButton.ctx":0000
   Begin VB.Label lblCaption 
      Alignment       =   2  '¸m¤¤¹ï»ô
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Cool Button"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1395
      TabIndex        =   2
      Top             =   1080
      Width           =   1365
      WordWrap        =   -1  'True
   End
   Begin VB.Label Front 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '³æ½u©T©w
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Back 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00800000&
      BorderStyle     =   1  '³æ½u©T©w
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "CoolButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Horizontal shadow ratio: 6%
'vertical shadow ratio: 14%
'
'This is a Text-only Button
'Made by Kenny Lai
'Not for release
'---------------------------------------------------------------------------------------------

Dim TFont As New StdFont
Dim TShadow As Boolean
Dim TEnabled As Boolean
Dim TCaption As String

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub Front_Click()
RaiseEvent Click
End Sub

Private Sub Front_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
Front.BackColor = &HFFFF80
End Sub

Private Sub Front_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
Front.BackColor = vbWhite
End Sub

Private Sub lblCaption_Click()
RaiseEvent Click
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
Front.BackColor = &HFFFF80
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
Front.BackColor = vbWhite
End Sub


Private Sub UserControl_InitProperties()
Me.Enabled = True
Me.Caption = "CoolButton"
Me.Shadow = True

End Sub

Private Sub UserControl_Resize()
    
    With Back
        .Top = Int(UserControl.height * 0.14)
        .Left = Int(UserControl.width * 0.06)
        .height = Int(UserControl.height * 0.84)
        .width = Int(UserControl.width * 0.94)
    End With
    
    With Front
        .Top = 0
        .Left = 0
        .height = Int(UserControl.height * 0.84)
        .width = Int(UserControl.width * 0.94)
    End With
    
    With lblCaption
        .Top = (Front.height - .height) \ 2
        .Left = (Front.width - .width) \ 2
    End With
    
End Sub

Public Property Let Caption(ByVal str As String)
TCaption = str
lblCaption.Caption = str
UserControl_Resize
End Property

Public Property Get Caption() As String
Caption = TCaption
End Property

Public Property Let Enabled(bln As Boolean)
    TEnabled = bln
lblCaption.Enabled = bln
Front.Enabled = bln

    If bln = False Then
        Front.BackColor = &HE0E0E0
    Else
    Front.Tag = "True"
        Front.BackColor = vbWhite
    End If
    
End Property

Public Property Get Enabled() As Boolean
Enabled = TEnabled
End Property

Public Property Let Shadow(bln As Boolean)
TShadow = bln
Back.Visible = bln
End Property

Public Property Get Shadow() As Boolean
Shadow = TShadow
End Property

Public Property Set Font(fnt As StdFont)
Set TFont = fnt
Set lblCaption.Font = fnt
End Property

Public Property Get Font() As StdFont
Set Font = TFont
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Caption", TCaption, ""
    .WriteProperty "Enabled", TEnabled, True
    .WriteProperty "Font", TFont, lblCaption.Font
    .WriteProperty "Shadow", TShadow, True
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    Me.Caption = .ReadProperty("Caption", "")
    Me.Enabled = .ReadProperty("Enabled", True)
    Me.Font = .ReadProperty("Font", lblCaption.Font)
    Me.Shadow = .ReadProperty("Shadow", True)
End With
End Sub
