Attribute VB_Name = "modSynHighlight"
Option Explicit

Public Enum synHighRGBType
    Red = 0
    Green = 1
    Blue = 2
End Enum

Public Enum ColReturnType
    AtCol = 0
    ComCol = 1
    TgCol = 2
    TxtCol = 3
    MarkCol = 4
End Enum

'Property Variables:
Public m_TextCol As Single
Public m_AttribCol As Single
Public m_CommentCol As Single
Public m_TagCol As Single
Public m_Enabled As Boolean
Public m_TextBold As Boolean
Public m_TextFont As String
Public m_TextItalics As Boolean
Public m_AttribItalics As Boolean
Public m_AttribBold As Boolean
Public m_AttribFont As String
Public m_CommentItalics As Boolean
Public m_CommentBold As Boolean
Public m_CommentFont As String
Public m_TagItalics As Boolean
Public m_TagBold As Boolean
Public m_TagFont As String

'Public m_MarkedCol As Single

Public TagBold As String
Public TagItalics As String
Public AttribBold As String
Public AttribItalics As String
Public CommentBold As String
Public CommentItalics As String
Public TextBold As String
Public TextItalics As String

Public TagInfo As String
Public AttribInfo As String
Public CommentInfo As String
Public TextInfo As String



Public Sub LoadStrOptions()
    If m_TagBold = True Then
        TagBold = "\b"
    Else
        TagBold = ""
    End If

    If m_TagItalics = True Then
        TagItalics = "\i"
    Else
        TagItalics = ""
    End If

    If m_AttribBold = True Then
        AttribBold = "\b"
    Else
        AttribBold = ""
    End If

    If m_AttribItalics = True Then
        AttribItalics = "\i"
    Else
        AttribItalics = ""
    End If

    If m_CommentBold = True Then
        CommentBold = "\b"
    Else
        CommentBold = ""
    End If

    If m_CommentItalics = True Then
        CommentItalics = "\i"
    Else
        CommentItalics = ""
    End If

    If m_TextBold = True Then
        TextBold = "\b"
    Else
        TextBold = ""
    End If

    If m_TextItalics = True Then
        TextItalics = "\i"
    Else
        TextItalics = ""
    End If

    TagInfo = TagBold & TagItalics & "\cf1 "
    AttribInfo = AttribBold & AttribItalics & "\cf2 "
    CommentInfo = CommentBold & CommentItalics & "\cf3 "
    TextInfo = TextBold & TextItalics & "\cf4 "
End Sub

Public Function InsertHighLightInfo(txtCode As RichTextBox) As String
    Dim sRTF As String, curr As String
    Dim after As String, bef As String
    Dim tblPos As Long, secTblPos As Long

    '{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Tahoma;}{\f3\froman\fprq2 Times New Roman;}}
    '{\colortbl\red0\green0\blue0;\red0\green0\blue160;\red0\green128\blue0;}
    '\deflang1033\pard\plain\f3\fs20
    '\par \plain\f3\fs20\cf1 blue\plain\f3\fs20\cf2 green\plain\f3\fs20
    '\par }
    '\cf1 = blue
    '\cf2 = green

    'Insert color RTF info.
    sRTF = txtCode.TextRTF
    tblPos = InStr(1, sRTF, "{\colortbl")
    If tblPos <> 0 Then
        secTblPos = InStr(tblPos, sRTF, "}")
        curr = Mid$(sRTF, tblPos, secTblPos - tblPos + 1)
        bef = Mid$(sRTF, 1, tblPos - 1)
        after = Mid$(sRTF, secTblPos + 1)

        sRTF = bef & "{\colortbl\red0\green0\blue0;" & ReturnRTFColorStr(TgCol) & ReturnRTFColorStr(AtCol) & ReturnRTFColorStr(ComCol) & ReturnRTFColorStr(TxtCol) & "}" & after
    End If

    '\cf1 = TagCol
    '\cf2 = AttribCol
    '\cf3 = CommentCol
    '\cf4 = TextCol

    InsertHighLightInfo = sRTF
End Function

Public Function ReturnRGBValue(Col As Single, CurrType As synHighRGBType) As Integer
    Dim color&
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    Dim Ret As Integer

    color = Col
    Red = color Mod 256
    Green = (color And &HFF00FF00) / 256
    Blue = Int(color / 65536)

    On Error Resume Next
    Select Case CurrType
        Case 0 'Red
            Ret = Red
        Case 1 'Green
            Ret = Green
        Case 2 'Blue
            Ret = Blue
    End Select

    ReturnRGBValue = Ret
End Function

'Returns RTF code for colortbl switch
Public Function ReturnRTFColorStr(ColNumType As ColReturnType) As String
    Dim EndStr As String

    Select Case ColNumType
        Case 0 'AttribCol
            EndStr = "\red" & Trim(CStr(ReturnRGBValue(m_AttribCol, Red))) & "\green" & Trim(CStr(ReturnRGBValue(m_AttribCol, Green))) & "\blue" & Trim(CStr(ReturnRGBValue(m_AttribCol, Blue))) & ";"
        Case 1 'CommentCol
            EndStr = "\red" & Trim(CStr(ReturnRGBValue(m_CommentCol, Red))) & "\green" & Trim(CStr(ReturnRGBValue(m_CommentCol, Green))) & "\blue" & Trim(CStr(ReturnRGBValue(m_CommentCol, Blue))) & ";"
        Case 2 'TagCol
            EndStr = "\red" & Trim(CStr(ReturnRGBValue(m_TagCol, Red))) & "\green" & Trim(CStr(ReturnRGBValue(m_TagCol, Green))) & "\blue" & Trim(CStr(ReturnRGBValue(m_TagCol, Blue))) & ";"
        Case 3 'TextCol
            EndStr = "\red" & Trim(CStr(ReturnRGBValue(m_TextCol, Red))) & "\green" & Trim(CStr(ReturnRGBValue(m_TextCol, Green))) & "\blue" & Trim(CStr(ReturnRGBValue(m_TextCol, Blue))) & ";"
        Case 4 'MarkedCol
            'EndStr = "\red" & Trim(cstr(ReturnRGBValue(m_MarkedCol, Red))) & "\green" & Trim(cstr(ReturnRGBValue(m_MarkedCol, Green))) & "\blue" & Trim(cstr(ReturnRGBValue(m_MarkedCol, Blue))) & ";"
    End Select

    ReturnRTFColorStr = EndStr
End Function



'Goes through each tag and inserts the correct colors
Public Function cycleAttrib(CurrTag As String) As String
    'Used to find position of =, quotes, and spaces...
    Dim fPos As Long, sPos As Long, sPos2 As Long

    'Finds the position of quotes...
    Dim qPos As Long, qnPos As Long

    'Checkes to see if it's the first cycle through
    'the loop and then change some info. if it is...
    Dim isFirstCycle As Boolean

    'Used to hold all the RTF info.
    Dim bef As String, infoRTF As String

    'Holds the progressively smaller tag and the
    'attrib value as it cycles through the loop...
    Dim eTag As String, sPosTxt As String

    '**********************************************
    '       THE MAIN CYCLE/LOOPING CODE...
    '**********************************************

    'This is the RTF that sets it back to normal
    'before we add in the formatting info...
    infoRTF = "\plain\f3\fs20"

    'Use another variable so we don't modify the
    'original...This var will also get progressively
    'smaller as it goes through the loop by eliminating
    'each attrib and its value one by one as the
    'loop cycles
    eTag = CurrTag

    'Makes sure it doesn't exit after the first round
    isFirstCycle = True

    'Begin the loop to insert the RTF info.
    Do While Len(eTag) > 0
        'Find the first instance of an = sign...
        fPos = InStr(1, eTag, "=")
        'There are no attributes so return the entire
        'tag as the colored tag by itself...
        'i.e. <html> -- no attributes
        If (fPos = 0 And isFirstCycle = True) Then
            cycleAttrib = infoRTF & TagInfo & CurrTag
            Exit Function
        ElseIf fPos <> 0 Then 'Put in the color info...
            If Left$(eTag, 1) = "<" Then
                'Gets the info. before the first = sign...
                'i.e. <body bgcolor="#FFFFFF" vlink=#000000>
                '-- it would be <body bgcolor=
                bef = bef & infoRTF & TagInfo & Mid$(eTag, 1, fPos) & infoRTF & AttribInfo

                'Truncates eTag so it would now be (using the prev.
                'example) "#FFFFFF" vlink=#000000>
                eTag = Mid$(eTag, fPos + 1)
            End If
        End If

        'Find the first instance of a space in the
        'part of the tag that we have left...
        sPos = InStr(1, eTag, Chr$(32))

        'Gets the text up to the next space...
        sPosTxt = Mid$(eTag, 1, sPos)

        'Checks to see if there's a quote in the text...
        qPos = InStr(1, sPosTxt, Chr$(34))

        'If there's a quote found, then we need to find
        'its end...
        If qPos <> 0 Then
            'Look for the next quote...
            qnPos = InStr(2, eTag, Chr$(34))

            'If the quote is found, then we need to
            'get the text all the way up to the next
            'quote...we need to do this since it might
            'contain spaces...and if it's in a quote, then
            'those spaces need to be included in the
            'attrib value...
            If qnPos <> 0 Then
                sPosTxt = Mid$(eTag, 1, qnPos)
            End If
        End If
        'Adds the attrib value to the tag...
        bef = bef & infoRTF & AttribInfo & Mid$(eTag, 1, Len(sPosTxt))
        'Truncates the tag so there's no attrib value left...
        eTag = Mid$(eTag, Len(sPosTxt) + 1)

        'Find the next position of an equal sign...
        sPos = InStr(1, eTag, "=")

        'If there's no =, then we know we're on the last
        'attrib value, so we need to put in some final
        'info...all that's left is something like:
        '"#ffffff">
        If sPos = 0 Then
            'Put in the attrib color before the ">"
            'if it's the last attribute...
            eTag = Mid$(eTag, 1, Len(eTag) - 1)

            'Insert the RTF info...
            bef = bef & infoRTF & AttribInfo & eTag

            'Truncate the end...
            sPos = Len(eTag)
            Exit Do
        End If

        'We're not on the last attrib value so we need
        'to get it ready for the next cycle by putting
        'in the attrib and setting the RTF info...
        bef = bef & infoRTF & TagInfo & Mid$(eTag, 1, sPos) & infoRTF & AttribInfo

        'Truncates the tag appropriately...
        eTag = Mid$(eTag, sPos + 1)

        'Some of the code is dependent on if it's the
        'first time through the loop...if it is the
        'first time, then this will set it to false
        'so the code isn't executed the 2nd time around
        isFirstCycle = False

        'If there's nothing left, then we need to exit
        'the loop so it doesn't loop infinitely...
        If sPos = 0 And qPos = 0 Then Exit Do
    Loop

    '**********************************************
    '              LOOP/CYCLING IS DONE
    '**********************************************
    'Insert some ending info...The ">" was taken off
    'during the cycling so we could put in the correct
    'color at the end of the loop and save some time...
    cycleAttrib = bef & infoRTF & TagInfo & ">" & TextInfo

    'Loop's done!!  The value returned is the
    'individual tag w/ all the RTF info. inserted...
    Exit Function


End Function

Public Function CycleComment(CurrTag As String) As String
    CycleComment = "\plain\f3\fs20" & CommentInfo & CurrTag
End Function

Public Sub SynHighlightCleanUp()
    m_TextCol = 0
    m_AttribCol = 0
    m_CommentCol = 0
    m_TagCol = 0
    m_TextFont = ""
    m_AttribFont = ""
    m_CommentFont = ""
    m_TagFont = ""

    TagBold = ""
    TagItalics = ""
    AttribBold = ""
    AttribItalics = ""
    CommentBold = ""
    CommentItalics = ""
    TextBold = ""
    TextItalics = ""

    TagInfo = ""
    AttribInfo = ""
    CommentInfo = ""
    TextInfo = ""
End Sub

Function ColorHtml(ByRef rtf As RichTextBox)


DoEvents
rtf.Visible = False
Dim ss As Long
ss = rtf.SelStart
rtf.SelStart = 0
rtf.SelLength = Len(rtf.Text)
rtf.SelColor = RGB(0, 0, 0)
Dim rtfstart As Long
rtfstart = rtf.SelStart
If rtf.SelLength < 1 Then
    MsgBox "No text selected"
Exit Function
End If

MfrmProgram.ActiveForm.pbSyntax.Max = Len(rtf.Text)
MfrmProgram.ActiveForm.pbSyntax.Value = 0
'Screen.MousePointer = 11
    Dim regEx, Match, Matches     ' Create variable.
    Set regEx = New RegExp            ' Create a regular expression.
    regEx.Pattern = "<[^>]*>"       ' Set pattern.
    'regEx.IgnoreCase = regcase          ' Set case insensitivity.
    regEx.Global = True           ' Set global applicability.
    Set Matches = regEx.Execute(rtf.SelText)    ' Execute search.
    For Each Match In Matches     ' Iterate Matches collection.
        'if use * (match 0 or more of previous pattern) then can return empty string
        'bug ?
        If Match.Value <> "" Then 'used to stop empty string match return
            rtf.SelStart = rtfstart + Match.FirstIndex
            'rtf.SelRTF = rtf.SelRTF & "\cf1 "
            rtf.SelLength = Len(Match.Value)
            rtf.SelColor = vbBlue
            'rtf.SelStart = rtf.SelStart + Len(Match.Value)
            'rtf.SelRTF = rtf.SelRTF & "\cf0 "
            
            MfrmProgram.ActiveForm.pbSyntax.Value = rtf.SelStart
        End If
    Next
rtf.Visible = True
'Screen.MousePointer = 0

MfrmProgram.ActiveForm.pbSyntax.Value = 0

rtf.SelStart = ss


End Function
