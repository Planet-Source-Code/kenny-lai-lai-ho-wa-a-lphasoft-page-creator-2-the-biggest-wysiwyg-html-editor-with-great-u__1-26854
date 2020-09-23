Attribute VB_Name = "modColor"
Type TypeRGB        'This user type is used to convert long decimals to Bytes
    R As Byte
    G As Byte
    B As Byte
End Type

Private Type TypeLong       'This user type is used to convert long decimals to Bytes
    RGBColor As Long
End Type

Function RGB2HTML(ByVal RGBColor As Long) As String

Dim tmpCol As TypeRGB

With tmpCol
.R = RGBColor Mod 256
.G = (RGBColor \ 256) Mod 256
.B = RGBColor \ 65536
End With

Dim RString As String, GString As String, BString As String

RString = Trim(Hex$(tmpCol.R))
GString = Trim(Hex$(tmpCol.G))
BString = Trim(Hex$(tmpCol.B))

If Len(RString) = 1 Then RString = "0" & RString
If Len(GString) = 1 Then GString = "0" & GString
If Len(BString) = 1 Then BString = "0" & BString

RGB2HTML = "#" & RString & GString & BString

End Function

Function HTML2RGB(ByVal sHexColor As String) As Long
    Dim lCol As Long, i, n
    If Left(sHexColor, 1) = "#" Then sHexColor = Mid(sHexColor, 2)
    sHexColor = UCase(sHexColor)
    
    For i = 1 To Len(sHexColor) Step 2
        lCol = lCol + Dec(Mid(sHexColor, i, 2)) * 256 ^ n
        n = n + 1
    Next i
    HTML2RGB = lCol
End Function

Function RGB2typeRGB(ByVal RGBColor As Long) As TypeRGB
With RGB2typeRGB
.R = RGBColor Mod 256
.G = (RGBColor \ 256) Mod 256
.B = RGBColor \ 65536
End With
End Function

Function HTML2typeRGB(ByVal HTMLColor As String) As TypeRGB
Dim temp As String
temp = HTML2RGB(HTMLColor)
Dim out As TypeRGB
out = RGB2typeRGB(temp)
With HTML2typeRGB
.R = out.R
.G = out.G
.B = out.B
End With
End Function

Private Function Dec(ByVal sHex As String) As Long 'Converts Hex to Decimal
    Const HVal = "0123456789ABCDEF"
    Dim iPos As Byte, i As Integer, lDec As Long
    Dim L As Integer, x As Byte
    L = Len(sHex)
    If L > 255 Then Exit Function
    lDec = 0
    For i = L To 1 Step -1
        x = InStr(1, HVal, Mid(sHex, i, 1), vbTextCompare)
        If x = 0 Then Exit Function Else x = x - 1
        lDec = lDec + x * 16 ^ (L - i)
    Next i
    Dec = lDec
End Function

Public Function RainbowColorText(ByVal str As String, ByVal Color1 As Long, ByVal Color2 As Long) As String
Dim Ccount As Long, i As Long
Ccount = Len(str) + 1: If Ccount = 0 Then Exit Function
Dim r1 As Integer, r2 As Integer, g1 As Integer, g2 As Integer, b1 As Integer, b2 As Integer
Dim Col1 As TypeRGB, Col2 As TypeRGB

Col1 = RGB2typeRGB(Color1)
Col2 = RGB2typeRGB(Color2)

r1 = Col1.R: g1 = Col1.G: b1 = Col1.B
r2 = Col2.R: g2 = Col2.G: b2 = Col2.B

Dim rd As Integer, gd As Integer, bd As Integer
Dim rf As Integer, gf As Integer, bf As Integer

Dim s As String, code As String, out As String

rd = r1 - r2: gd = g1 - g2: bd = b1 - b2

    For i = 1 To Ccount
        s = Mid(str, i, 1)
        
        'rf = Int(IIf(rd < 0, r1 + Int(Abs(rd / Ccount * i)), r1 - Int(Abs(rd / Ccount * i)))) Mod 255
        'gf = Int(IIf(gd < 0, r1 + Int(Abs(gd / Ccount * i)), r1 - Int(Abs(gd / Ccount * i)))) Mod 255
        'bf = Int(IIf(bd < 0, r1 + Int(Abs(bd / Ccount * i)), r1 - Int(Abs(bd / Ccount * i)))) Mod 255
        
        rf = Int(r1 - (rd / Ccount * i)) 'Mod 255
        gf = Int(g1 - (gd / Ccount * i)) 'Mod 255
        bf = Int(b1 - (bd / Ccount * i)) 'Mod 255

        If rd = 0 Then rf = r1
        If gd = 0 Then gf = g1
        If bd = 0 Then bf = b1
        
        'On Error Resume Next
    code = "<FONT Color=""" & RGB2HTML(RGB(rf, gf, bf)) & """>" & s & "</FONT>"
    
    out = out & code
    
    Next
    
    RainbowColorText = out

End Function
