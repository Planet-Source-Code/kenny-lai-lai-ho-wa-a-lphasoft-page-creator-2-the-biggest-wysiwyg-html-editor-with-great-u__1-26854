VERSION 5.00
Begin VB.UserControl synHighlight99 
   CanGetFocus     =   0   'False
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   375
   ScaleWidth      =   2385
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "syntaxHighlight"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "synHighlight99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Default Property Values:
Const m_def_TextCol = 0
Const m_def_AttribCol = 0
Const m_def_CommentCol = 0
Const m_def_TagCol = 0
Const m_def_Enabled = 0
Const m_def_TextBold = 0
Const m_def_TextFont = ""
Const m_def_TextItalics = 0
Const m_def_AttribItalics = 0
Const m_def_AttribBold = 0
Const m_def_AttribFont = ""
Const m_def_CommentItalics = 0
Const m_def_CommentBold = 0
Const m_def_CommentFont = ""
Const m_def_TagItalics = 0
Const m_def_TagBold = 0
Const m_def_TagFont = ""

'Const m_def_MarkedCol = 0

Public RichTxtBox As RichTextBox

Private CurrInTag As Boolean
Private CurrInAttrib As Boolean

Public Event HighlightProgress(CurrProgress As Single, TotalRTFLen As Single)


Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Sub Refresh()

End Sub

Public Property Get TextBold() As Boolean
    TextBold = m_TextBold
End Property

Public Property Let TextBold(ByVal New_TextBold As Boolean)
    m_TextBold = New_TextBold
    PropertyChanged "TextBold"
End Property

Private Property Get TextFont() As String
    TextFont = m_TextFont
End Property

Private Property Let TextFont(ByVal New_TextFont As String)
    m_TextFont = New_TextFont
    PropertyChanged "TextFont"
End Property


Public Property Get TextItalics() As Boolean
    TextItalics = m_TextItalics
End Property

Public Property Let TextItalics(ByVal New_TextItalics As Boolean)
    m_TextItalics = New_TextItalics
    PropertyChanged "TextItalics"
End Property

Public Property Get AttribItalics() As Boolean
    AttribItalics = m_AttribItalics
End Property

Public Property Let AttribItalics(ByVal New_AttribItalics As Boolean)
    m_AttribItalics = New_AttribItalics
    PropertyChanged "AttribItalics"
End Property

Public Property Get AttribBold() As Boolean
    AttribBold = m_AttribBold
End Property

Public Property Let AttribBold(ByVal New_AttribBold As Boolean)
    m_AttribBold = New_AttribBold
    PropertyChanged "AttribBold"
End Property

Private Property Get AttribFont() As String
    AttribFont = m_AttribFont
End Property

Private Property Let AttribFont(ByVal New_AttribFont As String)
    m_AttribFont = New_AttribFont
    PropertyChanged "AttribFont"
End Property


Public Property Get CommentItalics() As Boolean
    CommentItalics = m_CommentItalics
End Property

Public Property Let CommentItalics(ByVal New_CommentItalics As Boolean)
    m_CommentItalics = New_CommentItalics
    PropertyChanged "CommentItalics"
End Property

Public Property Get CommentBold() As Boolean
    CommentBold = m_CommentBold
End Property

Public Property Let CommentBold(ByVal New_CommentBold As Boolean)
    m_CommentBold = New_CommentBold
    PropertyChanged "CommentBold"
End Property

Private Property Get CommentFont() As String
    CommentFont = m_CommentFont
End Property

Private Property Let CommentFont(ByVal New_CommentFont As String)
    m_CommentFont = New_CommentFont
    PropertyChanged "CommentFont"
End Property


Public Property Get TagItalics() As Boolean
    TagItalics = m_TagItalics
End Property

Public Property Let TagItalics(ByVal New_TagItalics As Boolean)
    m_TagItalics = New_TagItalics
    PropertyChanged "TagItalics"
End Property

Public Property Get TagBold() As Boolean
    TagBold = m_TagBold
End Property

Public Property Let TagBold(ByVal New_TagBold As Boolean)
    m_TagBold = New_TagBold
    PropertyChanged "TagBold"
End Property

Private Property Get TagFont() As String
    TagFont = m_TagFont
End Property

Private Property Let TagFont(ByVal New_TagFont As String)
    m_TagFont = New_TagFont
    PropertyChanged "TagFont"
End Property





'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_TextBold = m_def_TextBold
    m_TextFont = m_def_TextFont
    m_TextItalics = m_def_TextItalics
    m_AttribItalics = m_def_AttribItalics
    m_AttribBold = m_def_AttribBold
    m_AttribFont = m_def_AttribFont
    m_CommentItalics = m_def_CommentItalics
    m_CommentBold = m_def_CommentBold
    m_CommentFont = m_def_CommentFont
    m_TagItalics = m_def_TagItalics
    m_TagBold = m_def_TagBold
    m_TagFont = m_def_TagFont
    m_TextCol = m_def_TextCol
    m_AttribCol = m_def_AttribCol
    m_CommentCol = m_def_CommentCol
    m_TagCol = m_def_TagCol
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_TextBold = PropBag.ReadProperty("TextBold", m_def_TextBold)
    m_TextFont = PropBag.ReadProperty("TextFont", m_def_TextFont)
    m_TextItalics = PropBag.ReadProperty("TextItalics", m_def_TextItalics)
    m_AttribItalics = PropBag.ReadProperty("AttribItalics", m_def_AttribItalics)
    m_AttribBold = PropBag.ReadProperty("AttribBold", m_def_AttribBold)
    m_AttribFont = PropBag.ReadProperty("AttribFont", m_def_AttribFont)
    m_CommentItalics = PropBag.ReadProperty("CommentItalics", m_def_CommentItalics)
    m_CommentBold = PropBag.ReadProperty("CommentBold", m_def_CommentBold)
    m_CommentFont = PropBag.ReadProperty("CommentFont", m_def_CommentFont)
    m_TagItalics = PropBag.ReadProperty("TagItalics", m_def_TagItalics)
    m_TagBold = PropBag.ReadProperty("TagBold", m_def_TagBold)
    m_TagFont = PropBag.ReadProperty("TagFont", m_def_TagFont)
    m_TextCol = PropBag.ReadProperty("TextCol", m_def_TextCol)
    m_AttribCol = PropBag.ReadProperty("AttribCol", m_def_AttribCol)
    m_CommentCol = PropBag.ReadProperty("CommentCol", m_def_CommentCol)
    m_TagCol = PropBag.ReadProperty("TagCol", m_def_TagCol)
    'm_MarkedCol = PropBag.ReadProperty("MarkedCol", m_def_MarkedCol)
End Sub

Private Sub UserControl_Resize()
    UserControl.width = lbl.width
    UserControl.height = lbl.height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("TextBold", m_TextBold, m_def_TextBold)
    Call PropBag.WriteProperty("TextFont", m_TextFont, m_def_TextFont)
    Call PropBag.WriteProperty("TextItalics", m_TextItalics, m_def_TextItalics)
    Call PropBag.WriteProperty("AttribItalics", m_AttribItalics, m_def_AttribItalics)
    Call PropBag.WriteProperty("AttribBold", m_AttribBold, m_def_AttribBold)
    Call PropBag.WriteProperty("AttribFont", m_AttribFont, m_def_AttribFont)
    Call PropBag.WriteProperty("CommentItalics", m_CommentItalics, m_def_CommentItalics)
    Call PropBag.WriteProperty("CommentBold", m_CommentBold, m_def_CommentBold)
    Call PropBag.WriteProperty("CommentFont", m_CommentFont, m_def_CommentFont)
    Call PropBag.WriteProperty("TagItalics", m_TagItalics, m_def_TagItalics)
    Call PropBag.WriteProperty("TagBold", m_TagBold, m_def_TagBold)
    Call PropBag.WriteProperty("TagFont", m_TagFont, m_def_TagFont)
    Call PropBag.WriteProperty("TextCol", m_TextCol, m_def_TextCol)
    Call PropBag.WriteProperty("AttribCol", m_AttribCol, m_def_AttribCol)
    Call PropBag.WriteProperty("CommentCol", m_CommentCol, m_def_CommentCol)
    Call PropBag.WriteProperty("TagCol", m_TagCol, m_def_TagCol)
    'Call PropBag.WriteProperty("MarkedCol", m_MarkedCol, m_def_MarkedCol)
End Sub

Public Property Get TextCol() As OLE_COLOR
    TextCol = m_TextCol
End Property

Public Property Let TextCol(ByVal New_TextCol As OLE_COLOR)
    m_TextCol = New_TextCol
    PropertyChanged "TextCol"
End Property

Public Property Get AttribCol() As OLE_COLOR
    AttribCol = m_AttribCol
End Property

Public Property Let AttribCol(ByVal New_AttribCol As OLE_COLOR)
    m_AttribCol = New_AttribCol
    PropertyChanged "AttribCol"
End Property

'Public Property Get MarkedCol() As OLE_COLOR
'    MarkedCol = m_MarkedCol
'End Property
'
'Public Property Let MarkedCol(ByVal New_MarkedCol As OLE_COLOR)
'    m_MarkedCol = New_MarkedCol
'    PropertyChanged "MarkedCol"
'End Property
Public Property Get CommentCol() As OLE_COLOR
    CommentCol = m_CommentCol
End Property

Public Property Let CommentCol(ByVal New_CommentCol As OLE_COLOR)
    m_CommentCol = New_CommentCol
    PropertyChanged "CommentCol"
End Property

Public Property Get TagCol() As OLE_COLOR
    TagCol = m_TagCol
End Property

Public Property Let TagCol(ByVal New_TagCol As OLE_COLOR)
    m_TagCol = New_TagCol
    PropertyChanged "TagCol"
End Property

Public Sub Highlight()
    Dim rtf As String

    On Error Resume Next
    m_TextFont = RichTxtBox.Font.Name
    m_TagFont = RichTxtBox.Font.Name
    m_AttribFont = RichTxtBox.Font.Name
    m_CommentFont = RichTxtBox.Font.Name
    
    If InStr(1, RichTxtBox.TextRTF, "\'") <> 0 Then
    'Exit Sub
    End If
    
    On Error Resume Next
    LoadStrOptions
    rtf = InsertHighLightInfo(RichTxtBox)
    rtf = HighlightHTML(rtf)

    RichTxtBox.TextRTF = rtf

End Sub

'Colorizes HTML while typing
Public Function KeyPressEvent(ByVal KeyAscii As Integer) As Integer
    Static cInAttrib As Boolean, cInTag As Boolean
    Static cInAttribQuote As Boolean, cTypedIn As Boolean
    Static cInComment As Boolean
    
    Dim cChar As String

    With RichTxtBox
        cChar = Chr$(KeyAscii)

        If cInTag = False And cInAttrib = False And cInComment = False Then
            .SelColor = m_TextCol
        End If

        If cInTag = True And (cInAttrib = True Or cInAttribQuote = True) Then
            .SelColor = m_AttribCol
        End If

        If cChar = "<" Then
            .SelColor = m_TagCol
            cInTag = True
            cTypedIn = True
        End If

        If cChar = "=" And cInTag = True Then
            cInAttrib = True
        End If

        If cChar = Chr$(34) And cInAttrib = True And cInAttribQuote = True Then
            '.SelColor = m_TagCol
            cInAttrib = False
            cInAttribQuote = False
        ElseIf cChar = Chr$(34) And cInAttrib = True And cInAttribQuote = False Then
            cInAttribQuote = True
        End If

        If cChar = " " And (cInAttribQuote = False And cInTag = True) Then
            .SelColor = m_TagCol
            cInAttrib = False
        End If




        If cChar = ">" Then
            If cInComment = False Then
                .SelColor = m_TagCol
            Else
                .SelColor = m_CommentCol
            End If
            cInTag = False
            cInComment = False
            cTypedIn = False
        End If

    End With

    KeyPressEvent = KeyAscii
ErrExit:
    Exit Function
End Function


Public Function HighlightHTML(ByVal rtf As String) As String
    Dim ePos As Long
    Dim after As String

    Dim InTag As Boolean
    Dim Total As String, curr As String
    Dim nPos As Single
    Dim infoRTF As String
    Dim Pos As Long
    Dim bef As String

If InStr(1, rtf, "\'") <> 0 Then
'HighlightHTML = RTF
'Exit Function
End If

    '\cf1 = TagCol
    '\cf2 = AttribCol
    '\cf3 = CommentCol
    '\cf4 = TextCol

    'Set the initial vars...
    Total = rtf
    infoRTF = "\plain\f3\fs20"
    nPos = 1

    'Execute the loop to color the tags...
    Do While Len(rtf) > 0

        'See if there's still any tags left...
        Pos = InStr(nPos, rtf, "<")

        'If so, then color it...
        If Pos <> 0 Then 'Found a tag
            InTag = True 'In a tag
            'Get everything before the tag we're on...
            bef = Mid$(rtf, 1, Pos - 1)
            'Find the end of the next tag...
            ePos = InStr(Pos, rtf, ">")

            'Get the current HTML tag...
            curr = Mid$(rtf, Pos, ePos - Pos + 1)
            If ePos <> 0 Then

                'Check to see if it's a comment or not...
                'If it is, then it requires different
                'handling than a normal tag...
                If Left$(curr, 4) <> "<!--" Then
                    curr = cycleAttrib(curr)
                Else 'It's a comment...
                    ePos = InStr(Pos, rtf, "-->") + 2
                    curr = Mid$(rtf, Pos, ePos - Pos + 1)
                    curr = CycleComment(curr)
                End If

                nPos = ePos + Len(infoRTF & TextInfo) + 1
                'Get the HTML after the current tag...
                after = Mid$(rtf, ePos + 1)

                rtf = bef & curr & infoRTF & TextInfo & after
            End If
            RaiseEvent HighlightProgress(Len(bef & curr), Len(rtf))
        Else
            Exit Do
        End If
    Loop
    RaiseEvent HighlightProgress(Len(rtf), Len(rtf))
    HighlightHTML = rtf
End Function

