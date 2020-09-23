VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   ClientHeight    =   5910
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4079.187
   ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.TextBox txtDisclaim 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '««ª½±²¶b
      TabIndex        =   7
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   4575
      Begin VB.Label Label3 
         Caption         =   " Written by Kenny Lai"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  '³z©ú
         Caption         =   "title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3885
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  '³z©ú
         Caption         =   "des"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   930
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   3885
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  '³z©ú
         Caption         =   "company name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   240
         MousePointer    =   2  '¤Q¦r§Îª¬
         TabIndex        =   4
         Top             =   840
         Width           =   3975
      End
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   2400
      TabIndex        =   0
      Top             =   5400
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
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
      Left            =   3720
      TabIndex        =   2
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "For all purpose and Open Source"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Alphasoft Page Creator 3 (Public Freeware)"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Revis 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   660
      WordWrap        =   -1  'True
   End
   Begin VB.Label Minor 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   540
      WordWrap        =   -1  'True
   End
   Begin VB.Label Major 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   540
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '¤º¹ê½u
      Index           =   1
      X1              =   0
      X2              =   5408.938
      Y1              =   2650.436
      Y2              =   2650.436
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5408.938
      Y1              =   2650.436
      Y2              =   2650.436
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long



Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub



Private Sub CoolButton1_Click()
MsgBox "Coolbutton1 Clicked"
End Sub


Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    Me.Icon = MfrmProgram.Icon
    picIcon.Move picIcon.Left, picIcon.Top, Me.Icon.width, Me.Icon.height
    picIcon.Picture = MfrmProgram.Icon
    lblTitle.Caption = App.Title
    Me.lblDescription.Caption = App.Comments
    Me.txtDisclaim.Text = App.LegalTrademarks
    Me.lblCompany.Caption = App.CompanyName
    Major = "Major: " & vbCrLf & App.Major
    Minor = "Minor: " & vbCrLf & App.Minor
    Revis = "Revision: " & Format(App.Revision, "0000")
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        Else
            GoTo SysInfoErr
        End If
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Cannot provide System Information.", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long
    Dim rc As Long
    Dim hKey As Long
    Dim hDepth As Long
    Dim KeyValType As Long
    Dim tmpVal As String
    Dim KeyValSize As Long
    
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    
    

    
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    '
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
        tmpVal = Left(tmpVal, KeyValSize - 1)
    Else
        tmpVal = Left(tmpVal, KeyValSize)
    End If
    Select Case KeyValType
    Case REG_SZ
        KeyVal = tmpVal
    Case REG_DWORD
        For i = Len(tmpVal) To 1 Step -1
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
        Next
        KeyVal = Format$("&h" + KeyVal)
    End Select
    
    GetKeyValue = True
    rc = RegCloseKey(hKey)
    Exit Function
    
GetKeyError:
    KeyVal = ""
    GetKeyValue = False
    rc = RegCloseKey(hKey)
End Function


Private Sub lblCompany_DblClick()
On Error Resume Next
MsgBox "Thank you for using my software." & vbCrLf & vbCrLf & _
        "Please mail me on assw@hkem.com" & vbCrLf & vbCrLf & _
        "Kenny Lai", vbInformation, "Kenny Lai"
Shell "start mailto:assw@hkem.com", vbNormalFocus
End Sub
