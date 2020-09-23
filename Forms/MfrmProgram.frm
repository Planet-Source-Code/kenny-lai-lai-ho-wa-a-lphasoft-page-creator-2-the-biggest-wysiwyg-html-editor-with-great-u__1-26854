VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{9F631458-BEE6-11D3-AFAF-9F131A29873D}#1.7#0"; "TREE.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{05B9F8C4-05D2-11D1-A081-444553540000}#1.0#0"; "NEWTREE.OCX"
Object = "{AE80D66B-8000-11D2-8B17-600109C10000}#5.0#0"; "OUTLOOKBAR.OCX"
Object = "{C9FF5F4F-78AB-4799-A8B8-EA9191E3BBA7}#1.0#0"; "CPOPMENU.OCX"
Begin VB.MDIForm MfrmProgram 
   Appearance      =   0  '¥­­±
   BackColor       =   &H80000003&
   Caption         =   "Alphasoft Page Creator 3"
   ClientHeight    =   5175
   ClientLeft      =   1935
   ClientTop       =   -60
   ClientWidth     =   6240
   Icon            =   "MfrmProgram.frx":0000
   LinkTopic       =   "MDIForm1"
   Visible         =   0   'False
   WindowState     =   2  '³Ì¤j¤Æ
   Begin cPopMenu.PopMenu PM 
      Left            =   3240
      Top             =   1680
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      HighlightStyle  =   2
      HighlightColor  =   16777215
      BorderColor     =   8388608
      HForeColor      =   8388608
   End
   Begin VB.PictureBox pic2 
      Align           =   4  '¹ï»ôªí³æ¥k¤è
      BackColor       =   &H00800000&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   6195
      Left            =   10920
      ScaleHeight     =   6195
      ScaleWidth      =   150
      TabIndex        =   32
      Top             =   750
      Width           =   150
      Begin VB.CommandButton SB2 
         Height          =   1695
         Left            =   0
         Picture         =   "MfrmProgram.frx":0442
         Style           =   1  '¹Ï¤ù¥~Æ[
         TabIndex        =   33
         ToolTipText     =   "Project Manager"
         Top             =   1560
         Width           =   150
      End
   End
   Begin VB.PictureBox picProjectMan 
      Align           =   4  '¹ï»ôªí³æ¥k¤è
      Appearance      =   0  '¥­­±
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   11070
      ScaleHeight     =   6195
      ScaleWidth      =   810
      TabIndex        =   31
      Top             =   750
      Width           =   815
      Begin Outlook_bar.OutlookBar OutBar 
         Height          =   6135
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   10821
         MenusMax        =   2
         MenuCur         =   2
         MenuStartup     =   2
         MenuCaption1    =   "Project"
         MenuItemIcon11  =   "MfrmProgram.frx":0808
         MenuItemCaption11=   "Manager"
         MenuCaption2    =   "Library"
         MenuItemsMax2   =   2
         MenuItemIcon21  =   "MfrmProgram.frx":0C5A
         MenuItemCaption21=   "Cool !"
         MenuItemIcon22  =   "MfrmProgram.frx":1EDC
         MenuItemCaption22=   "Custom"
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   3  '¹ï»ôªí³æ¥ª¤è
      BackColor       =   &H00800000&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   6195
      Left            =   2670
      ScaleHeight     =   6195
      ScaleWidth      =   150
      TabIndex        =   24
      Top             =   750
      Width           =   150
      Begin VB.CommandButton SB1 
         Height          =   1695
         Left            =   0
         Picture         =   "MfrmProgram.frx":4016
         Style           =   1  '¹Ï¤ù¥~Æ[
         TabIndex        =   25
         ToolTipText     =   "Toolbox"
         Top             =   1560
         Width           =   150
      End
   End
   Begin MSComctlLib.ImageList imlHTML 
      Left            =   5040
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":43E2
            Key             =   "img"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":453C
            Key             =   "a"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":4696
            Key             =   "p"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":47F0
            Key             =   "span"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":494A
            Key             =   "input"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":4AA4
            Key             =   "script"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":4BFE
            Key             =   "embed"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":4D58
            Key             =   "font"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":4EB2
            Key             =   "object"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":500C
            Key             =   "all"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":5166
            Key             =   "hr"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":52C0
            Key             =   "html"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":541A
            Key             =   "meta"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":5574
            Key             =   "div"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":56CE
            Key             =   "label"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":5828
            Key             =   "li"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":5982
            Key             =   "applet"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":5ADC
            Key             =   "form"
         EndProperty
      EndProperty
   End
   Begin VB.Data DatLib 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Caption         =   "Script Library"
      Connect         =   "Access"
      DatabaseName    =   "Database\EffectLibraryDatabase.mdb"
      DefaultCursorType=   0  '¹w³]ªº¸ê®Æ«ü¼Ð
      DefaultType     =   2  '¨Ï¥Î ODBCDirect
      Exclusive       =   0   'False
      Height          =   450
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '°ÊºA¶°(Dynaset)
      RecordSource    =   "Effect"
      Top             =   7350
      Visible         =   0   'False
      Width           =   11880
   End
   Begin VB.Data DatCustom 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Caption         =   "Custom Script"
      Connect         =   "Access"
      DatabaseName    =   "D:\data\Alphasoft Series\NetWalker Suit\Copy -Page Creator\Database\EffectLibraryDatabase.mdb"
      DefaultCursorType=   0  '¹w³]ªº¸ê®Æ«ü¼Ð
      DefaultType     =   2  '¨Ï¥Î ODBCDirect
      Exclusive       =   0   'False
      Height          =   405
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '°ÊºA¶°(Dynaset)
      RecordSource    =   "Custom"
      Top             =   6945
      Visible         =   0   'False
      Width           =   11880
   End
   Begin MSComctlLib.ImageList imlSide 
      Left            =   5760
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":5F2E
            Key             =   "button"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":6088
            Key             =   "textbox"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":61E2
            Key             =   "multi"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":633C
            Key             =   "check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":6496
            Key             =   "radio"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":65F0
            Key             =   "form"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":6A42
            Key             =   "label"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":6B9C
            Key             =   "combo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":6CF6
            Key             =   "line"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":6E50
            Key             =   "div"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":6FAA
            Key             =   "marquee"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":7104
            Key             =   "java"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":725E
            Key             =   "script"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":73B8
            Key             =   "scontrol"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":7512
            Key             =   "plugin"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":766C
            Key             =   "lib"
         EndProperty
      EndProperty
   End
   Begin PageCreator.synHighlight99 syn 
      Left            =   3480
      Top             =   1080
      _extentx        =   1693
      _extenty        =   609
      attribcol       =   192
      commentcol      =   32768
      tagcol          =   1.67117e7
   End
   Begin VB.Timer TmrStart 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3480
      Top             =   2400
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   3360
   End
   Begin MSComctlLib.ImageList imlImage 
      Left            =   5760
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   187
      ImageHeight     =   438
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":7AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":ACA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":AD1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":B274
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":B7CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":BD20
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":ECAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":11BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":13D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1448C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":17535
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1A469
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1C95B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   2880
   End
   Begin MSComctlLib.StatusBar staProp 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   7800
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   12356
            MinWidth        =   12347
         EndProperty
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbar3 
      Left            =   5040
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1F44B
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1F5A7
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1F703
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1F85F
            Key             =   "left"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1F9BB
            Key             =   "centre"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1FB17
            Key             =   "right"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1FC73
            Key             =   "front"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1FD87
            Key             =   "back"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1FE9B
            Key             =   "all"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":1FFAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":207E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":21017
            Key             =   "absolute"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2146B
            Key             =   "foreground"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":219BD
            Key             =   "background"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":21F0F
            Key             =   "indent"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":22073
            Key             =   "outdent"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":221D7
            Key             =   "bullet"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2233B
            Key             =   "number"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbar2 
      Left            =   5040
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2249F
            Key             =   "find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":225F9
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":22753
            Key             =   "search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":228AD
            Key             =   "image"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":22A09
            Key             =   "refresh server"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":22E5D
            Key             =   "server setup"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":22FB9
            Key             =   "script"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":23115
            Key             =   "link"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":23271
            Key             =   "table"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":236C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":239ED
            Key             =   "doc"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   8055
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "NUM"
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "CAP"
            TextSave        =   "CAP"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1005
            MinWidth        =   882
            Text            =   "ROW "
            TextSave        =   "ROW "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   8811
            MinWidth        =   8819
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   2400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   1920
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4560
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4080
      Tag             =   "For Menu Action"
      Top             =   1440
   End
   Begin VB.PictureBox picToolbox 
      Align           =   3  '¹ï»ôªí³æ¥ª¤è
      Appearance      =   0  '¥­­±
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   6195
      Left            =   0
      ScaleHeight     =   6195
      ScaleWidth      =   2670
      TabIndex        =   2
      Top             =   750
      Width           =   2675
      Begin VB.PictureBox picTest 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   15
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
      Begin DHTMLEDLibCtl.DHTMLEdit DE1 
         Height          =   30
         Left            =   -30
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   30
         ActivateApplets =   0   'False
         ActivateActiveXControls=   0   'False
         ActivateDTCs    =   -1  'True
         ShowDetails     =   0   'False
         ShowBorders     =   0   'False
         Appearance      =   1
         Scrollbars      =   -1  'True
         ScrollbarAppearance=   1
         SourceCodePreservation=   -1  'True
         AbsoluteDropMode=   0   'False
         SnapToGrid      =   0   'False
         SnapToGridX     =   50
         SnapToGridY     =   50
         BrowseMode      =   0   'False
         UseDivOnCarriageReturn=   0   'False
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6495
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   11456
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   529
         TabMaxWidth     =   132
         MouseIcon       =   "MfrmProgram.frx":23F87
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   " "
         TabPicture(0)   =   "MfrmProgram.frx":23FA3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fam(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   " "
         TabPicture(1)   =   "MfrmProgram.frx":240FD
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fam(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " "
         TabPicture(2)   =   "MfrmProgram.frx":24257
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fam(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   " "
         TabPicture(3)   =   "MfrmProgram.frx":243B1
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fam(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   " "
         TabPicture(4)   =   "MfrmProgram.frx":2450B
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "fam(4)"
         Tab(4).ControlCount=   1
         Begin VB.Frame fam 
            Height          =   4815
            Index           =   4
            Left            =   -74760
            TabIndex        =   21
            Top             =   480
            Width           =   1575
            Begin MSComctlLib.TreeView tv1 
               Height          =   2655
               Left            =   120
               TabIndex        =   22
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   4683
               _Version        =   393217
               LabelEdit       =   1
               LineStyle       =   1
               Sorted          =   -1  'True
               Style           =   7
               ImageList       =   "imlHTML"
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame fam 
            Height          =   5775
            Index           =   3
            Left            =   -74880
            TabIndex        =   17
            Top             =   480
            Width           =   2415
            Begin MSComctlLib.Toolbar tbrForm 
               Height          =   5610
               Left            =   0
               TabIndex        =   18
               Top             =   120
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   9895
               ButtonWidth     =   4577
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Appearance      =   1
               Style           =   1
               TextAlignment   =   1
               ImageList       =   "imlSide"
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   17
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "                                      Button"
                     ImageKey        =   "button"
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Text Box"
                     ImageKey        =   "textbox"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Multiline Textbox"
                     ImageKey        =   "multi"
                  EndProperty
                  BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Checkbox"
                     ImageKey        =   "check"
                  EndProperty
                  BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Radio Button"
                     ImageKey        =   "radio"
                  EndProperty
                  BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Form"
                     ImageKey        =   "form"
                  EndProperty
                  BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Label"
                     ImageKey        =   "label"
                  EndProperty
                  BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Combo Box"
                     ImageKey        =   "combo"
                  EndProperty
                  BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Horizontal Line"
                     ImageKey        =   "line"
                  EndProperty
                  BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Caption         =   "Floating Text"
                     ImageKey        =   "div"
                  EndProperty
                  BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Style           =   3
                  EndProperty
                  BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Caption         =   "Moving Text"
                     ImageKey        =   "marquee"
                  EndProperty
                  BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Caption         =   "Custom Script Block"
                     ImageKey        =   "script"
                  EndProperty
                  BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Caption         =   "Control with Script"
                     ImageKey        =   "scontrol"
                  EndProperty
                  BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Caption         =   "Java Applet"
                     ImageKey        =   "java"
                  EndProperty
                  BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Caption         =   "Plug-ins"
                     ImageKey        =   "plugin"
                  EndProperty
                  BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Caption         =   "Effect Library ..."
                     ImageKey        =   "lib"
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Frame fam 
            Height          =   5295
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1935
            Begin ExplorerCtls.asxFileListView file1 
               Height          =   2415
               Left            =   240
               TabIndex        =   14
               Top             =   2280
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   4260
               FlatHeaders     =   0
               Pattern         =   "*.htm; *.html"
               AllowDrives     =   0   'False
               IncludeFolders  =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               View            =   2
               AllowDelete     =   -1  'True
               AllowProperties =   -1  'True
               GridLines       =   -1  'True
               PictureAlignment=   0
            End
            Begin NEWEXLib.ExplorerTree tree1 
               Height          =   1695
               Left            =   240
               TabIndex        =   11
               Top             =   360
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   2990
               _StockProps     =   161
               BackColor       =   16777215
               Appearance      =   1
            End
         End
         Begin VB.Frame fam 
            Height          =   3495
            Index           =   2
            Left            =   -74880
            TabIndex        =   12
            Top             =   1260
            Width           =   1695
            Begin VB.ListBox lstTags 
               ForeColor       =   &H00800000&
               Height          =   1500
               ItemData        =   "MfrmProgram.frx":24665
               Left            =   240
               List            =   "MfrmProgram.frx":247AD
               TabIndex        =   13
               Top             =   600
               Width           =   1455
            End
         End
         Begin VB.Frame fam 
            Height          =   4575
            Index           =   1
            Left            =   -74760
            TabIndex        =   9
            Top             =   600
            Width           =   1695
            Begin VB.CommandButton cmdEditCode 
               Height          =   375
               Left            =   840
               Picture         =   "MfrmProgram.frx":24E51
               Style           =   1  '¹Ï¤ù¥~Æ[
               TabIndex        =   27
               Top             =   240
               Width           =   615
            End
            Begin VB.CommandButton cmdAddCode 
               Height          =   375
               Left            =   240
               Picture         =   "MfrmProgram.frx":24F9B
               Style           =   1  '¹Ï¤ù¥~Æ[
               TabIndex        =   26
               Top             =   240
               Width           =   495
            End
            Begin MSDBCtls.DBList lstCode 
               Bindings        =   "MfrmProgram.frx":250E5
               Height          =   900
               Left            =   120
               TabIndex        =   20
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1588
               _Version        =   393216
               MatchEntry      =   -1  'True
               ListField       =   "Name"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "·s²Ó©úÅé"
                  Size            =   9
                  Charset         =   136
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RichTextLib.RichTextBox rtc 
               DataField       =   "Content"
               DataSource      =   "DatCustom"
               Height          =   2775
               Left            =   0
               TabIndex        =   19
               Top             =   1800
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   4895
               _Version        =   393217
               ScrollBars      =   3
               FileName        =   "D:\data\Alphasoft Series\NetWalker Suit\Copy -Page Creator\code.rtf"
               TextRTF         =   $"MfrmProgram.frx":250FD
            End
         End
      End
      Begin VB.Image imlDown 
         Height          =   1725
         Left            =   1560
         Picture         =   "MfrmProgram.frx":25811
         Top             =   120
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Image imlUp 
         Height          =   1725
         Left            =   1920
         Picture         =   "MfrmProgram.frx":25BDD
         Top             =   120
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   5040
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":25FA3
            Key             =   "play"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":260FD
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":26257
            Key             =   "new"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":267F3
            Key             =   "open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":26D8F
            Key             =   "save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2732B
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":278C7
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":27E63
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":283FF
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2899B
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":28F37
            Key             =   "redo"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1323
      _CBWidth        =   11880
      _CBHeight       =   750
      _Version        =   "6.7.8862"
      Child1          =   "tbrGeneral"
      MinHeight1      =   330
      Width1          =   4320
      NewRow1         =   0   'False
      Child2          =   "tbrSimFunction"
      MinHeight2      =   330
      Width2          =   1395
      NewRow2         =   0   'False
      Child3          =   "tbrEdit"
      MinHeight3      =   330
      Width3          =   1335
      NewRow3         =   -1  'True
      Begin MSComctlLib.Toolbar tbrEdit 
         Height          =   330
         Left            =   165
         TabIndex        =   5
         Top             =   390
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlToolbar3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   22
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   4600
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Bold"
               ImageKey        =   "bold"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Italic"
               ImageKey        =   "italic"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Underline"
               ImageKey        =   "underline"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Left"
               ImageKey        =   "left"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Centre"
               ImageKey        =   "centre"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Right"
               ImageKey        =   "right"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Indent"
               ImageKey        =   "indent"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Outdent"
               ImageKey        =   "outdent"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Bullet"
               ImageKey        =   "bullet"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Number"
               ImageKey        =   "number"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Bring foreward"
               ImageKey        =   "front"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Bring backward"
               ImageKey        =   "back"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Foreground color"
               ImageKey        =   "foreground"
               Style           =   4
               Object.Width           =   750
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Background color"
               ImageKey        =   "background"
               Style           =   4
               Object.Width           =   750
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Make Absolute"
               ImageKey        =   "absolute"
            EndProperty
         EndProperty
         Begin PageCreator.Font cobFonts 
            Height          =   315
            Left            =   1320
            TabIndex        =   30
            Top             =   0
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
         End
         Begin PageCreator.ColorPick cpBack 
            Height          =   300
            Left            =   10080
            TabIndex        =   29
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
            Text            =   "B"
         End
         Begin PageCreator.ColorPick cpFore 
            Height          =   300
            Left            =   9240
            TabIndex        =   28
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
            Text            =   "F"
         End
         Begin VB.ComboBox cobFormat 
            Height          =   300
            ItemData        =   "MfrmProgram.frx":294D3
            Left            =   0
            List            =   "MfrmProgram.frx":294D5
            Style           =   2  '³æ¯Â¤U©Ô¦¡
            TabIndex        =   15
            Top             =   0
            Width           =   1335
         End
         Begin VB.ComboBox cobSize 
            CausesValidation=   0   'False
            Height          =   300
            Left            =   3480
            TabIndex        =   6
            Top             =   0
            Width           =   1095
         End
      End
      Begin MSComctlLib.Toolbar tbrSimFunction 
         Height          =   330
         Left            =   4515
         TabIndex        =   4
         Top             =   30
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlToolbar2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Find"
               ImageKey        =   "find"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Search"
               ImageKey        =   "search"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Image"
               ImageKey        =   "image"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insert HTML"
               ImageKey        =   "script"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Link"
               ImageKey        =   "link"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Table"
               ImageKey        =   "table"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Document Properties"
               ImageKey        =   "doc"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrGeneral 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "New"
               ImageKey        =   "new"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Open"
               ImageKey        =   "open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Save"
               ImageKey        =   "save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cut"
               ImageKey        =   "cut"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Cut "
               ImageKey        =   "cut"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copy"
               ImageKey        =   "copy"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Copy "
               ImageKey        =   "copy"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Paste"
               ImageKey        =   "paste"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Paste "
               ImageKey        =   "paste"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Delete"
               ImageKey        =   "delete"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Delete "
               ImageKey        =   "delete"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Undo"
               ImageKey        =   "undo"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Redo"
               ImageKey        =   "redo"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Play"
               ImageKey        =   "play"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Stop"
               ImageKey        =   "stop"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlMenu 
      Left            =   5760
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   51
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":294D7
            Key             =   "mnunew_top"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":29A19
            Key             =   "mnunew_design"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":29B73
            Key             =   "mnuopen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2A0B5
            Key             =   "mnusave"
            Object.Tag             =   "Save (&S)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2A5F7
            Key             =   "mnucut"
            Object.Tag             =   "&Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2A709
            Key             =   "mnustart"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2A863
            Key             =   "mnuview_projectman"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2A9BD
            Key             =   "mnucopy"
            Object.Tag             =   "&Copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2AACF
            Key             =   "mnupaste"
            Object.Tag             =   "&Paste"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2ABE1
            Key             =   "mnudelete"
            Object.Tag             =   "&Delete"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2ACF3
            Key             =   "mnuselectall"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2AE07
            Key             =   "mnuundo"
            Object.Tag             =   "Undo (&U)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2AF19
            Key             =   "mnuredo"
            Object.Tag             =   "Redo (&R)"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2B02B
            Key             =   "play"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2B193
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2B5E7
            Key             =   "mnuview_toolbox"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2BA39
            Key             =   "mnuexit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2BFD3
            Key             =   "mnupreview"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2C12D
            Key             =   "mnuprint"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2C287
            Key             =   "mnupsetup"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2C3E1
            Key             =   "mnudetail"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2C53B
            Key             =   "mnuwindowcascade"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2C695
            Key             =   "mnuwindowtilehorizontal"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2C7EF
            Key             =   "mnuwindowtilevertical"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2C949
            Key             =   "mnuabout"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2CC63
            Key             =   "extcut"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2CDBD
            Key             =   "extcopy"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2CF17
            Key             =   "extpaste"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2D071
            Key             =   "extdelete"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2D1CB
            Key             =   "extall"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2D2DD
            Key             =   "extpagepro"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2D437
            Key             =   "extabs"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2D659
            Key             =   "extsnap"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2D973
            Key             =   "extdetail"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2DACD
            Key             =   "extinrow"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2DC27
            Key             =   "extincol"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2DD81
            Key             =   "extmerge"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2DEDB
            Key             =   "extsplit"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2E035
            Key             =   "extapro"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2E18F
            Key             =   "extimagepro"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2E2E9
            Key             =   "extinhtml"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2E443
            Key             =   "extproperties"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2E59D
            Key             =   "exttablegeneral"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2E9EF
            Key             =   "extabsgeneral"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2EE41
            Key             =   "extback"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2EF53
            Key             =   "extbackward"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2F065
            Key             =   "extbelowtext"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2F177
            Key             =   "extfront"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2F289
            Key             =   "extforeward"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2F39B
            Key             =   "extabovetext"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MfrmProgram.frx":2F4AD
            Key             =   "mnutool_rainbow"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile_Top 
      Caption         =   "File (&F)"
      Begin VB.Menu mnuNew_Top 
         Caption         =   "New Document...(&N)"
         Begin VB.Menu mnuNew_Design 
            Caption         =   "in Design Mode (&D)"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open... (&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save (&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As... (&A)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Startup Manager"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print (&P)"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit (&X)"
      End
   End
   Begin VB.Menu mnuEdit_Top 
      Caption         =   "Edit (&E)"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo (&U)"
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo (&R)"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools (&T)"
      Begin VB.Menu mnuTool_Rainbow 
         Caption         =   "Color-Fading Text"
      End
   End
   Begin VB.Menu mnuView_Top 
      Caption         =   "View (&V)"
      Begin VB.Menu mnuView_menustyle 
         Caption         =   "Menu Style (&M)"
         Begin VB.Menu mnuView_MenuOption 
            Caption         =   "Menu Option... (&O)"
            Begin VB.Menu mnuView_CusMenu 
               Caption         =   "Customize... (&C)"
            End
         End
         Begin VB.Menu mnu5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuView_style 
            Caption         =   "Sky Blue"
            Index           =   1
         End
         Begin VB.Menu mnuView_style 
            Caption         =   "Clear XP Style"
            Index           =   2
         End
         Begin VB.Menu mnuView_style 
            Caption         =   "Blue"
            Index           =   3
         End
         Begin VB.Menu mnuView_style 
            Caption         =   "Green"
            Index           =   4
         End
         Begin VB.Menu mnuView_style 
            Caption         =   "Orange"
            Index           =   5
         End
         Begin VB.Menu mnu6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuView_ImageStyle 
            Caption         =   "Brown Lines"
            Index           =   6
         End
         Begin VB.Menu mnuView_ImageStyle 
            Caption         =   "Flowers"
            Index           =   7
         End
         Begin VB.Menu mnuView_ImageStyle 
            Caption         =   "Moon Surface"
            Index           =   8
         End
         Begin VB.Menu mnuView_ImageStyle 
            Caption         =   "Woody"
            Index           =   9
         End
         Begin VB.Menu mnuView_ImageStyle 
            Caption         =   "Rock"
            Index           =   10
         End
         Begin VB.Menu mnuView_ImageStyle 
            Caption         =   "Sea"
            Index           =   11
         End
         Begin VB.Menu mnuView_ImageStyle 
            Caption         =   "Water Life"
            Index           =   12
         End
         Begin VB.Menu mnuView_ImageStyle 
            Caption         =   "Page Creator Special"
            Index           =   13
         End
      End
      Begin VB.Menu mnuView_Toolbox 
         Caption         =   "Tools Box"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_ProjectMan 
         Caption         =   "Project Manager"
      End
      Begin VB.Menu mnuDetail 
         Caption         =   "Show Detail (&D)"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window (&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Cascade (&C)"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Horizontal (&H)"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Vertical (&V)"
      End
   End
   Begin VB.Menu mnuHelp_Top 
      Caption         =   "Help (&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuPSC 
         Caption         =   "Connect to Planet-Source-Code..."
      End
   End
   Begin VB.Menu mnuExt 
      Caption         =   ""
      Enabled         =   0   'False
      Begin VB.Menu extCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu extCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu extPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu extDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu extAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu ext1 
         Caption         =   "-"
      End
      Begin VB.Menu extProperties 
         Caption         =   "Properties..."
         Begin VB.Menu extPagePro 
            Caption         =   "Page Properties"
         End
         Begin VB.Menu extImagePro 
            Caption         =   "Image Properties"
            Enabled         =   0   'False
         End
         Begin VB.Menu extTablePro 
            Caption         =   "Table Properties"
            Enabled         =   0   'False
         End
         Begin VB.Menu extTDPro 
            Caption         =   "Table Cell Properties"
            Enabled         =   0   'False
         End
         Begin VB.Menu extAPro 
            Caption         =   "Hyperlink Properties"
            Enabled         =   0   'False
         End
         Begin VB.Menu extFontsPro 
            Caption         =   "Fonts Properties"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu ext2 
         Caption         =   "-"
      End
      Begin VB.Menu extAbs 
         Caption         =   "Absolute Position"
         Enabled         =   0   'False
      End
      Begin VB.Menu extDetail 
         Caption         =   "Show Detail"
      End
      Begin VB.Menu extSnap 
         Caption         =   "Snap to Grid"
      End
      Begin VB.Menu extInHTML 
         Caption         =   "Insert Custom HTML"
      End
      Begin VB.Menu extTableGeneral 
         Caption         =   "Table..."
         Begin VB.Menu extInCol 
            Caption         =   "Insert Colume"
         End
         Begin VB.Menu extInRow 
            Caption         =   "Insert Row"
         End
         Begin VB.Menu extMerge 
            Caption         =   "Merge Cells"
         End
         Begin VB.Menu extSplit 
            Caption         =   "Split Cell"
         End
      End
      Begin VB.Menu extAbsGeneral 
         Caption         =   "Absolute object..."
         Begin VB.Menu extFront 
            Caption         =   "Bring to Front"
         End
         Begin VB.Menu extForeward 
            Caption         =   "Bring Foreward"
         End
         Begin VB.Menu extAboveText 
            Caption         =   "Bring above text"
         End
         Begin VB.Menu extBack 
            Caption         =   "Send to Back"
         End
         Begin VB.Menu extBackward 
            Caption         =   "Send Backward"
         End
         Begin VB.Menu extBelowText 
            Caption         =   "Send below text"
         End
      End
      Begin VB.Menu extGetTag 
         Caption         =   "Get Element TagName"
      End
   End
   Begin VB.Menu mnuCustom 
      Caption         =   "Custom"
      Visible         =   0   'False
      Begin VB.Menu cusDel 
         Caption         =   "Delete Record"
      End
   End
End
Attribute VB_Name = "MfrmProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Public Ready As Boolean
Dim EditorCount As Integer
Public SaveSuc As Boolean
Private K() As cControlFlater, i As Integer

Private Sub FlatControls()
Dim CTL As Control

    ReDim Preserve K(0 To Me.Controls.Count)
    
    For Each CTL In Me.Controls
    On Error Resume Next
        Select Case TypeName(CTL)
        Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar", "ListBox"
            Set K(i) = New cControlFlater
            K(i).Attach CTL
            i = i + 1
        End Select
    Next CTL
End Sub

Private Sub cmdAddCode_Click()

With frmAddCode
.Data1.Recordset.AddNew
End With
frmAddCode.Show vbModal
DatCustom.Refresh
lstCode.Refresh
lstCode.ReFill
End Sub

Private Sub cmdEditCode_Click()
lstCode_Click
On Error GoTo 1
DatCustom.Recordset.Delete
DatCustom.Refresh
lstCode.Refresh
lstCode.ReFill
1
End Sub

Private Sub cobFonts_Click()
ActiveForm.DHTML1.execCommand DECMD_SETFONTNAME, , cobFonts.Text
ActiveForm.DHTML1.SetFocus
End Sub

Private Sub cobFormat_Change()
On Error GoTo 1
ActiveForm.DHTML1.execCommand DECMD_SETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER, cobFormat.Text
ActiveForm.DHTML1.SetFocus
1
End Sub

Private Sub cobFormat_Click()

On Error Resume Next
If ActiveForm.DHTML1.execCommand(DECMD_GETBLOCKFMT) = cobFormat.Text Then Exit Sub

On Error GoTo 1
ActiveForm.DHTML1.execCommand DECMD_SETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER, cobFormat.Text
ActiveForm.DHTML1.SetFocus
1

End Sub

Private Sub cobSize_Click()
'DoEvents
On Error GoTo 1
ActiveForm.DHTML1.execCommand DECMD_SETFONTSIZE, , Left(cobSize.Text, 1)
ActiveForm.DHTML1.SetFocus
1
End Sub

Sub InTable()
frmTable.Show vbModal
If frmTable.Value = True Then
ActiveForm.DHTML1.execCommand DECMD_INSERTTABLE, , frmTable.tableParam
Else
End If
End Sub








Private Sub cpBack_Click()
ActiveForm.DHTML1.execCommand DECMD_SETBACKCOLOR, , RGB2HTML(cpBack.Color)
ActiveForm.DHTML1.SetFocus
End Sub

Private Sub cpFore_Click()
ActiveForm.DHTML1.execCommand DECMD_SETFORECOLOR, , RGB2HTML(cpFore.Color)
ActiveForm.DHTML1.SetFocus
End Sub

Private Sub DE1_DocumentComplete()

DisplayFormats

End Sub

Private Sub extAboveText_Click()
ActiveForm.DHTML1.execCommand DECMD_BRING_ABOVE_TEXT
End Sub

Private Sub extAbs_Click()
ActiveForm.DHTML1.execCommand DECMD_MAKE_ABSOLUTE
End Sub

Private Sub extAll_Click()
'On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_SELECTALL
Case 1: .rt1.SelStart = 0: .rt1.SelLength = Len(.rt1.Text)
End Select
End With
Timer2_Timer
1
End Sub

Private Sub extBack_Click()
ActiveForm.DHTML1.execCommand DECMD_SEND_TO_BACK
End Sub

Private Sub extBackward_Click()
ActiveForm.DHTML1.execCommand DECMD_SEND_BACKWARD
End Sub

Private Sub extBelowText_Click()
ActiveForm.DHTML1.execCommand DECMD_SEND_BELOW_TEXT
End Sub

Private Sub extCopy_Click()
On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_COPY
Case 1: Clipboard.SetText .rt1.SelText
End Select
End With
1
End Sub

Private Sub extCut_Click()
On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_CUT
Case 1: Clipboard.Clear: Clipboard.SetText .rt1.SelText: rt1.SelText = ""
End Select
End With
1
End Sub

Private Sub extDelete_Click()
On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_DELETE
Case 1: .rt1.SelText = ""
End Select
End With
1
End Sub

Private Sub extDetail_Click()
extDetail.Checked = Not extDetail.Checked
mnuDetail.Checked = extDetail.Checked
ActiveForm.DHTML1.ShowDetails = extDetail.Checked
End Sub

Private Sub extForeward_Click()
ActiveForm.DHTML1.execCommand DECMD_BRING_FORWARD
End Sub

Private Sub extFront_Click()
ActiveForm.DHTML1.execCommand DECMD_BRING_TO_FRONT
End Sub

Private Sub extGetTag_Click()
MsgBox "Element Tag Name: " & ActiveForm.GetActiveElement.tagName, vbInformation
End Sub

Private Sub extInCol_Click()
ActiveForm.DHTML1.execCommand DECMD_INSERTCOL
End Sub

Private Sub extInRow_Click()
ActiveForm.DHTML1.execCommand DECMD_INSERTROW
End Sub

Private Sub extMerge_Click()
ActiveForm.DHTML1.execCommand DECMD_MERGECELLS
End Sub

Private Sub extPagePro_Click()
Set frmProp.DHTML = ActiveForm.DHTML1
frmProp.Show vbModal
End Sub

Private Sub extPaste_Click()
On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_PASTE
Case 1: .rt1.SelText = Clipboard.GetText
End Select
End With
1
End Sub



Private Sub extSnap_Click()
extSnap.Checked = Not extSnap.Checked
ActiveForm.DHTML1.SnapToGrid = extSnap.Checked
End Sub

Private Sub extSplit_Click()
ActiveForm.DHTML1.execCommand DECMD_SPLITCELL
End Sub




Private Sub lstCode_Click()
On Error GoTo 2
DatCustom.Recordset.FindFirst "Name='" & lstCode.Text & "'"

Exit Sub
2
MsgBox Err.Number & " " & Error
End Sub


Private Sub MDIForm_Resize()
On Error Resume Next
Me.Move 0, 0
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuPSC_Click()
OpenURL "http://www.planet-source-code.com"
End Sub

Private Sub mnuSelectAll_Click()
'On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_SELECTALL
Case 1: .rt1.SelStart = 0: .rt1.SelLength = Len(.rt1.Text)
End Select
End With
Timer2_Timer
1
End Sub

Private Sub mnuCopy_Click()
On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_COPY
Case 1: Clipboard.SetText .rt1.SelText
End Select
End With
1
End Sub

Private Sub mnuCut_Click()
On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_CUT
Case 1: Clipboard.Clear: Clipboard.SetText .rt1.SelText: rt1.SelText = ""
End Select
End With
1
End Sub

Private Sub mnuDelete_Click()
On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_DELETE
Case 1: .rt1.SelText = ""
End Select
End With
1
End Sub

Private Sub mnuPaste_Click()
On Error GoTo 1
With ActiveForm
Select Case ActiveForm.SSTab1.Tab
Case 0: .DHTML1.execCommand DECMD_PASTE
Case 1: .rt1.SelText = Clipboard.GetText
End Select
End With
1
End Sub

Private Sub File1_FileDblClick(ByVal File As String)
Dim RO As Integer

On Error GoTo 1
DoEvents
OpenFile File, File
1 End Sub

Private Sub lstTags_DblClick()
Select Case ActiveForm.SSTab1.Tab
Case 1
ActiveForm.rt1.SelText = lstTags.Text
Case 0
InsertHTML lstTags.Text
End Select
End Sub

Private Sub MDIForm_Activate()
staMain.Refresh
staProp.Refresh

End Sub
Private Sub DisplayFormats()
'Dim fmt As DEGetBlockFmtNamesParam
'    Set f = CreateObject("DEGetBlockFmtNamesParam.DEGetBlockFmtNamesParam")
'    On Error Resume Next
'    DHTMLEdit1.execCommand DECMD_GETBLOCKFMTNAMES, , f
'    For Each fmtName In f.Names
'       cobFormat.AddItem fmtName
'    Next
    
        Dim fmt As DEGetBlockFmtNamesParam
        Dim i As Long
        Dim fmtName As Variant
        
        ' Create the block fmt names holder
        Set fmt = CreateObject("DEGetBlockFmtNamesParam.DEGetBlockFmtNamesParam.1")
        
        ' Get the localized strings for the DECMD_SETBLOCKFMT command
        DE1.execCommand DECMD_GETBLOCKFMTNAMES, OLECMDEXECOPT_DONTPROMPTUSER, fmt
        
        ' Put the strings into the Format menu
        i = 0
        For Each fmtName In fmt.Names
        cobFormat.AddItem fmtName
'            FormatSub(i).Caption = fmtName
'            i = i + 1
        Next
End Sub

Private Sub MDIForm_Initialize()

DoEvents
SetLoadText "Initializing..."
SetMenu
Ready = False
Process = GetCurrentProcess
ThreadID = App.ThreadID

FontSizePoint(1) = 8
FontSizePoint(2) = 10
FontSizePoint(3) = 12
FontSizePoint(4) = 14
FontSizePoint(5) = 18
FontSizePoint(6) = 24
FontSizePoint(7) = 36

With syn
Set .RichTxtBox = rtc
.AttribCol = &HC0&
.CommentCol = &H8000&
.TagCol = &HFF0000
.TextCol = &H800000
End With

'SetTabs

staMain.Refresh
staProp.Refresh
End Sub
 


Sub SetMenu()
DoEvents
With PM
.ImageList = imlMenu.hImageList
.SubClassMenu MfrmProgram

Dim m As Control
For Each m In Me
On Error Resume Next
.ItemIcon(m.Name) = imlMenu.ListImages(LCase(m.Name)).Index - 1
Next
End With
i = GetSet(App.ProductName, "option", "MenuStyle", 2)

If i = 0 Then
    Set PM.BackgroundPicture = LoadPicture(GetOption("CustomMenuBackground"))
Else
    Set PM.BackgroundPicture = imlImage.ListImages(i).Picture
End If

With PM
    .HighlightStyle = cspHighlightXP
    .ShadowXPHighlight = True
    .ShadowXPHighlightTopMenu = True
    
    .ForeColor = GetOption("mnuFore", &H0&)
    .borderColor = GetOption("mnuBorder", &H800000)
    .HighlightColor = GetOption("mnuHighlight", &HFFC0C0)
    .HighlightForeColor = GetOption("mnuHighFore", &H0&)
End With

End Sub


Private Sub MDIForm_Unload(Cancel As Integer)

Timer1 = False
Timer2 = False
Timer3 = False
Timer4 = False
Timer5 = False

GdiFlush
On Error Resume Next
For Each Form In Forms
Unload Form
Next

FreeMemory
End Sub



Private Sub mnuDelete_P_Click()

End Sub



Private Sub mnuDetail_Click()
mnuDetail.Checked = Not mnuDetail.Checked
ActiveForm.DHTML1.ShowDetails = mnuDetail.Checked
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOpen_Click()
With cd1
.Flags = cdlOFNCreatePrompt
.CancelError = True
.Filter = "HTML Document *.htm, *.html|*.htm;*.html"
.DialogTitle = "Open HTML Document"
On Error GoTo 1
.ShowOpen

DoEvents
OpenFile cd1.FileName, cd1.FileName
End With
1 End Sub

Private Sub mnuPreview_Click()
On Error GoTo 1
ActiveForm.webPreview.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
1
End Sub

Private Sub mnuPrint_Click()
On Error GoTo 1
ActiveForm.webPreview.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
1 End Sub

Private Sub mnuPSetup_Click()
On Error GoTo 1
ActiveForm.webPreview.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
1
End Sub

Public Function SaveAs() As Boolean
'Save New File

        'Prepare cd1
        On Error GoTo 1
        With cd1
        .Filter = "HTML Document *.htm|*.htm"
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt
        On Error GoTo 1
        .ShowSave
        
        If ActiveForm.cIsSave.Value = 1 Then
        
        'Set it to "Never Read Only type" !! Ha Ha!!
        SetAttr .FileName, vbReadOnly = False
        ActiveForm.DHTML1.SaveDocument .FileName
        Else
        
        'Just save it
        ActiveForm.DHTML1.SaveDocument .FileName
        End If
        
        ActiveForm.cIsSave.Value = 1
        With ActiveForm
            .txtFile.Text = cd1.FileName
            .Caption = cd1.FileName
        End With
        End With
Me.SaveSuc = True
SaveAs = True
ActiveForm.HTMLString = ActiveForm.DHTML1.DocumentHTML
Exit Function
1
Me.SaveSuc = False
SaveAs = False
End Function
Public Function Save() As Boolean
If ActiveForm.cIsSave.Value = 0 Then
2
'Save New File

        'Prepare cd1
        On Error GoTo 1
        With cd1
        .Filter = "HTML Document *.htm|*.htm"
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt
        On Error GoTo 1
        .ShowSave
        
        If ActiveForm.cIsSave.Value = 1 Then
        
        'Set it to "Never Read Only type !! Ha Ha!!
        On Error GoTo sets
        SetAttr .FileName, vbReadOnly = False
sets:         ActiveForm.DHTML1.SaveDocument .FileName
        Else
        
        'Just save it
        ActiveForm.DHTML1.SaveDocument .FileName
        End If
        
        ActiveForm.cIsSave.Value = 1
        With ActiveForm
            .txtFile.Text = cd1.FileName
            .FileName = Mid(.txtFile.Text, 1, Len(.txtFile.Text) - 4) & "tmp.htm"
            .Caption = cd1.FileName
        End With
        End With

ElseIf ActiveForm.cIsSave.Value = 1 Then
'Save old file
On Error GoTo sets1
        SetAttr ActiveForm.txtFile.Text, vbReadOnly = False
sets1:
ActiveForm.FileName = Mid(ActiveForm.txtFile.Text, 1, Len(ActiveForm.txtFile.Text) - 4) & "tmp.htm"
ActiveForm.DHTML1.SaveDocument ActiveForm.txtFile.Text
End If

SaveSuc = True
Save = True
ActiveForm.HTMLString = ActiveForm.DHTML1.DocumentHTML
Exit Function
1
SaveSuc = False
Save = False

End Function

Public Sub mnuSave_Click()
Save
End Sub

Public Sub mnuSaveAs_Click()
SaveAs
End Sub

Private Sub mnuStart_Click()
SaveSet App.ProductName, "option", "startup", 1
frmStart.Show vbModal
End Sub

Private Sub mnuTool_Rainbow_Click()
frmTool_Rainbow.Show vbModal
End Sub

Sub HideToolbox(ByVal Visible As Boolean)
SaveSet App.ProductName, "option", "viewToolbox", Visible
    mnuView_Toolbox.Checked = Visible
If Visible = False Then
        'Animation
        For i = 0 To picToolbox.width Step 150
            On Error Resume Next
                picToolbox.width = picToolbox.width - 150
            Next
        picToolbox.Visible = False
        SB1.Picture = imlUp.Picture
    Else
        picToolbox.Visible = True
        
        Do Until picToolbox.width >= 2675
        picToolbox.width = picToolbox.width + 150
        Loop
        
        picToolbox.width = 2675
        SB1.Picture = imlDown.Picture
End If
End Sub

Sub HideProject(ByVal Visible As Boolean)
SaveSet App.ProductName, "option", "viewProject", Visible

    mnuView_ProjectMan.Checked = Visible
    
If Visible = False Then
        'Animation
        For i = 0 To picProjectMan.width Step 100
            On Error Resume Next
                picProjectMan.width = picProjectMan.width - 100
            Next
        picProjectMan.Visible = False
        SB2.Picture = imlDown.Picture
    Else
        picProjectMan.Visible = True
        
        Do Until picProjectMan.width >= 815
        picProjectMan.width = picProjectMan.width + 100
        Loop
        
        picProjectMan.width = 815
        SB2.Picture = imlUp.Picture
End If
End Sub

Private Sub mnuView_CusMenu_Click()

    frmCusMenu.Show vbModal

End Sub

Private Sub mnuView_ImageStyle_Click(Index As Integer)
SaveSet App.ProductName, "option", "MenuStyle", Index
Set PM.BackgroundPicture = imlImage.ListImages(Index).Picture
End Sub

Private Sub mnuView_ProjectMan_Click()
mnuView_ProjectMan.Checked = Not mnuView_ProjectMan.Checked
HideProject mnuView_ProjectMan.Checked
On Error Resume Next
Select Case ActiveForm.SSTab1.Tab
    Case 0: ActiveForm.DHTML1.SetFocus
    Case 1: ActiveForm.rt1.SetFocus
    Case 2: ActiveForm.webPreview.SetFocus
End Select
End Sub

Private Sub mnuView_style_Click(Index As Integer)
SaveSet App.ProductName, "option", "MenuStyle", Index
Set PM.BackgroundPicture = imlImage.ListImages(Index).Picture
End Sub

Private Sub mnuView_Toolbox_Click()
mnuView_Toolbox.Checked = Not mnuView_Toolbox.Checked
HideToolbox mnuView_Toolbox.Checked
On Error Resume Next
Select Case ActiveForm.SSTab1.Tab
    Case 0: ActiveForm.DHTML1.SetFocus
    Case 1: ActiveForm.rt1.SetFocus
    Case 2: ActiveForm.webPreview.SetFocus
End Select
End Sub

Private Sub mnuWindowCascade_Click()
  Me.Arrange vbCascade
End Sub


Private Sub mnuWindowTileHorizontal_Click()
  Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
  Me.Arrange vbTileVertical
End Sub


Private Sub SetLoadText(ByVal Text As String)
frmSplash.lbl.Caption = Text
End Sub

Private Sub MDIForm_Load()

DoEvents
SetLoadText "Start Loading Program..."

EditorCount = 0

'Prepare ToolBar
DoEvents
SetLoadText "Setting constants.."

cobSize.AddItem "1 (8 pt)"
cobSize.AddItem "2 (10 pt)"
cobSize.AddItem "3 (12 pt)"
cobSize.AddItem "4 (14 pt)"
cobSize.AddItem "5 (18 pt)"
cobSize.AddItem "6 (24 pt)"
cobSize.AddItem "7 (36 pt)"

cobSize.Refresh
Ready = True

fam(0).ZOrder 0
File1.Path = Tree1.Path

DoEvents
SetLoadText "Testing program..."

DE1.NewDocument

DoEvents
SetLoadText "Loading Databases..."

'Database work
DatCustom.DatabaseName = App.Path & "\Database\EffectLibraryDatabase.mdb"
DatLib.DatabaseName = App.Path & "\Database\EffectLibraryDatabase.mdb"
DatCustom.Refresh
DatLib.Refresh
lstCode.ReFill
mnuAbout.Caption = "About " & App.ProductName & "..."

DoEvents
SetLoadText "Loading User Interface..."


TmrStart.Enabled = True

DoEvents
SetLoadText "Building User Interface..."

FlatControls
cobFonts.MakeMeFlat
cobFonts.MakeText
cpBack.MakeMeFlat
cpFore.MakeMeFlat
Me.HideToolbox GetSet(App.ProductName, "option", "viewToolbox", True)
Me.HideProject GetSet(App.ProductName, "option", "viewProject", True)
Unload frmSplash

Me.Visible = True

End Sub

Private Sub mnuNew_Design_Click()
DoEvents
Me.NewBlankPage
End Sub

Public Sub NewBlankPage()
Dim Editor As New frmMain
EditorCount = EditorCount + 1
With Editor
.Caption = "New Page " & EditorCount
.txtFile.Text = App.Path & "\Tem" & EditorCount & ".htm"
.FileName = Mid(.txtFile.Text, 1, Len(.txtFile.Text) - 4) & "tmp.htm"
.cIsSave.Value = 0
.Visible = False
.DHTML1.ZOrder 0
.DHTML1.DocumentHTML = .NewPage.Text
.HTMLString = .NewPage.Text
.Flags = "OK"
.Visible = True
End With
End Sub

Sub SimpleNew()
Dim Editor As frmMain
Set Editor = New frmMain
EditorCount = EditorCount + 1
With Editor
.txtFile.Text = App.Path & "\Tem" & EditorCount & ".htm"
.FileName = Mid(.txtFile.Text, 1, Len(.txtFile.Text) - 4) & "tmp.htm"
.SSTab1.Tab = 0
.DHTML1.ZOrder 0
.DHTML1.DocumentHTML = .NewPage.Text
.rt1.Text = .NewPage.Text
.Caption = "New Page " & EditorCount
.Show
End With
End Sub

Private Sub pic1_Resize()
SB1.Top = Int((pic1.height - SB1.height) / 2)
End Sub

Private Sub pic2_Resize()
SB2.Top = Int((pic2.height - SB2.height) / 2)
End Sub



Private Sub picProjectMan_Resize()
OutBar.Move 0, 0, picProjectMan.width, picProjectMan.height
End Sub

'Private Sub outbarLeft_MenuItemClick(MenuNumber As Long, MenuItem As Long)

'Custom ID type
'Dim MenuID As String
'MenuID = Trim(Str(MenuNumber)) & "_" & Trim(Str(MenuItem))
'
'Select Case MenuID
'
'Case "1_1"
'If ActiveForm.SSTab1.Tab <> 0 Then Exit Sub
'Set frmProp.DHTML = ActiveForm.DHTML1
'frmProp.Show vbModal
'
'Case "1_2"
'InTable

'Case "1_3"
'If ActiveForm.SSTab1.Tab = 0 Then ActiveForm.DHTML1.execCommand DECMD_IMAGE
'
'Case "1_4"
'If ActiveForm.SSTab1.Tab = 0 Then frmMain.DHTML1.execCommand DECMD_HYPERLINK
'
'Case "1_5"
'MsgBox "Not Support Yet", vbMsgBoxSetForeground
''InFlash
''
'End Select
'1
'End Sub

'Sub InFlash()
'frmAddFlash.Show vbModal
'Dim inf As Object
'Set inf = ActiveForm.DHTML1.DOM.selection.createRange
'If frmAddFlash.FlashCode = "NO" Then Exit Sub
'inf.pasteHTML frmAddFlash.FlashCode
'End Sub


Private Sub picToolbox_Resize()

With SSTab1
.Left = 0
.Top = 0
.width = picToolbox.width
.height = picToolbox.height
End With

For i = 0 To 4
With fam(i)
.Left = 0
.Top = SSTab1.TabHeight
.width = picToolbox.width
On Error Resume Next
.height = picToolbox.height - SSTab1.TabHeight
End With
Next

With Tree1
.Top = 0
.Left = 0
.width = fam(0).width
On Error GoTo 1
.height = fam(0).height / 2
End With

With File1
.Top = Tree1.height
.Left = 0
.width = picToolbox.width
On Error GoTo 1
.height = fam(0).height - Tree1.height
End With

With lstTags
.Top = 0
.Left = 0
.width = fam(2).width
On Error GoTo 1
.height = fam(2).height
End With

With tbrForm
.Top = 0
.Left = 0
.width = fam(3).width
End With

With tv1
.Move 0, 0, fam(4).width, fam(4).height
End With

cmdAddCode.Move 0, 0, Int(fam(1).width / 2)
cmdEditCode.Move cmdAddCode.width, 0, Int(fam(1).width / 2)
lstCode.Move 0, cmdAddCode.height, fam(1).width, Int((fam(1).height - cmdAddCode.height) / 2)
rtc.Move 0, lstCode.Top + lstCode.height, fam(1).width, fam(1).height - cmdAddCode.height - lstCode.height

1
End Sub



Private Sub rtc_KeyPress(KeyAscii As Integer)
On Error Resume Next
syn.KeyPressEvent (KeyAscii)
End Sub



Private Sub SB1_Click()
mnuView_Toolbox_Click
End Sub


Private Sub SB2_Click()
mnuView_ProjectMan_Click
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
fam(SSTab1.Tab).ZOrder 0
FillHTMLTree
End Sub




Private Sub tbrEdit_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.ToolTipText
Case "Bold"
    ActiveForm.DHTML1.execCommand DECMD_BOLD
Case "Italic"
    ActiveForm.DHTML1.execCommand DECMD_ITALIC
Case "Underline"
    ActiveForm.DHTML1.execCommand DECMD_UNDERLINE
Case "Left"
    ActiveForm.DHTML1.execCommand DECMD_JUSTIFYLEFT
Case "Centre"
    ActiveForm.DHTML1.execCommand DECMD_JUSTIFYCENTER
Case "Right"
    ActiveForm.DHTML1.execCommand DECMD_JUSTIFYRIGHT
Case "Indent"
    ActiveForm.DHTML1.execCommand DECMD_INDENT
Case "Outdent"
    ActiveForm.DHTML1.execCommand DECMD_OUTDENT
Case "Bullet"
    ActiveForm.DHTML1.execCommand DECMD_UNORDERLIST
Case "Number"
    ActiveForm.DHTML1.execCommand DECMD_ORDERLIST
Case "Bring foreward"
    ActiveForm.DHTML1.execCommand DECMD_BRING_FORWARD
Case "Bring backward"
    ActiveForm.DHTML1.execCommand DECMD_SEND_BACKWARD
Case "Make Absolute"
    ActiveForm.DHTML1.execCommand DECMD_MAKE_ABSOLUTE
End Select
Timer2_Timer
1 End Sub


Private Sub tbrForm_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo 1

Select Case Button.Caption
Case "                                      Button"
InsertHTML HTMLButton
Case "Text Box"
InsertHTML HTMLTextbox
Case "Multiline Textbox"
InsertHTML HTMLMultiline
Case "Checkbox"
InsertHTML HTMLCheckBox
Case "Radio Button"
InsertHTML HTMLRadio
Case "Form"
InsertHTML HTMLForm
Case "Label"
InsertHTML HTMLLabel
Case "Combo Box"
InsertHTML HTMLCombo
Case "Horizontal Line"
InsertHTML HTMLLine
Case "Floating Text"
InsertHTML DIV

'Complicate Element
Case "Moving Text"

Case "Custom Script Block"

Case "Custom Control"

Case "ActiveX Control"

Case "Java Applet"

Case "Plug-ins"

Case "Effect Library ..."


End Select

1 End Sub

Private Sub tbrGeneral_ButtonClick(ByVal Button As MSComctlLib.Button)
Timer2_Timer
Select Case Button.ToolTipText
Case "New"
    mnuNew_Design_Click
Case "Open"
    mnuOpen_Click
Case "Save"
    mnuSave_Click
Case "Cut"
   On Error GoTo 1
    ActiveForm.DHTML1.execCommand DECMD_CUT
    Timer2_Timer
Case "Cut "
    On Error GoTo 1
    Clipboard.Clear: Clipboard.SetText ActiveForm.rt1.SelText: ActiveForm.rt1.SelText = ""
Case "Copy"
   On Error GoTo 1
    ActiveForm.DHTML1.execCommand DECMD_COPY
    Timer2_Timer
Case "Copy "
   On Error GoTo 1
    Clipboard.Clear: Clipboard.SetText ActiveForm.rt1.SelText
Case "Paste"
    On Error GoTo 1
    ActiveForm.DHTML1.execCommand DECMD_PASTE
    Timer2_Timer
Case "Paste "
    On Error GoTo 1
    ActiveForm.rt1.SelText = Clipboard.GetText
Case "Delete"
   On Error GoTo 1
    ActiveForm.DHTML1.execCommand DECMD_DELETE
    Timer2_Timer
Case "Delete "
    On Error GoTo 1
ActiveForm.rt1.SelText = ""
Case "Undo"
On Error GoTo 1
        Select Case ActiveForm.SSTab1.Tab
            Case 0: ActiveForm.DHTML1.execCommand DECMD_UNDO
            Case 1
            'doEvents
            'ActiveForm.rt1.SetFocus
            'SendKeys "{ctrl}{z}"
            'ActiveForm.rt1Undo
        End Select
    
Case "Redo"
    On Error GoTo 1
    ActiveForm.DHTML1.execCommand DECMD_REDO
Case "Play"
    On Error GoTo 1
    FileName = App.Path & "\tmp.htm"
    ActiveForm.DHTML1.SaveDocument FileName
    ActiveForm.webPreview.Offline = True
    ActiveForm.webPreview.navigate FileName

    ActiveForm.SSTab1.Tab = 2
Case "Stop"
   On Error GoTo 1
    ActiveForm.SSTab1.Tab = 0
End Select

1 End Sub

Private Sub tbrSimFunction_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.ToolTipText
Case "Find"
ActiveForm.DHTML1.execCommand DECMD_FINDTEXT
Case "Search"
frmMSearch.Show
Case "Image"
ActiveForm.DHTML1.execCommand DECMD_IMAGE
Case "Insert HTML"

Case "Link"
ActiveForm.DHTML1.execCommand DECMD_HYPERLINK

Case "Table"
InTable

Case "Document Properties"
Set frmProp.DHTML = ActiveForm.DHTML1
frmProp.Show vbModal

End Select
End Sub

Private Sub Timer1_Timer()
DoEvents
On Error GoTo 1
LoadUICore Me.ActiveForm.SSTab1.Tab
    With Me.SSTab1
    .TabEnabled(0) = True
    .TabEnabled(1) = True
    .TabEnabled(2) = True
    .TabEnabled(3) = True
    .TabEnabled(4) = True
    End With
Exit Sub

DoEvents
1 If Err.Number = 91 Then LoadUICore -1
With SSTab1
.TabEnabled(2) = False
.TabEnabled(3) = False
.TabEnabled(4) = False
End With
End Sub

Public Sub RefreshEditBar()
DoEvents
On Error GoTo 1
'Very boring...
'Dynamic General toolbar and menus
Select Case ActiveForm.SSTab1.Tab
Case 0
    With MfrmProgram
    .mnuCut = ButtonEnable(DECMD_CUT)
    .mnuCopy = ButtonEnable(DECMD_COPY)
    .mnuPaste = ButtonEnable(DECMD_PASTE)
    .mnuDelete = ButtonEnable(DECMD_DELETE)
    .mnuSelectAll = ButtonEnable(DECMD_SELECTALL)
    .mnuUndo = ButtonEnable(DECMD_UNDO)
    .mnuRedo = ButtonEnable(DECMD_REDO)
    End With

    With tbrGeneral
    '.Buttons(5).Value = buttonpress(DECMD_CUT))
    .Buttons(5).Enabled = ButtonEnable(DECMD_CUT)
    '.Buttons(7).Value = buttonpress(DECMD_COPY))
    .Buttons(7).Enabled = ButtonEnable(DECMD_COPY)
    '.Buttons(9).Value = buttonpress(DECMD_PASTE))
    .Buttons(9).Enabled = ButtonEnable(DECMD_PASTE)
    '.Buttons(11).Value = buttonpress(DECMD_DELETE))
    .Buttons(11).Enabled = ButtonEnable(DECMD_DELETE)
    .Buttons(14).Enabled = ButtonEnable(DECMD_UNDO)
    .Buttons(15).Enabled = ButtonEnable(DECMD_REDO)
    End With


    With MfrmProgram.tbrEdit
    On Error Resume Next
    MfrmProgram.cobFonts.Enabled = ButtonEnable(DECMD_SETFONTNAME)
    MfrmProgram.cobSize.Enabled = ButtonEnable(DECMD_SETFONTSIZE)
    MfrmProgram.cobFormat.Enabled = ButtonEnable(DECMD_SETBLOCKFMT)
        
    'Dynamic Edit Toolbar..

        'Just boring work !
    .Buttons(3).Value = ButtonPress(DECMD_BOLD)
    .Buttons(3).Enabled = ButtonEnable(DECMD_BOLD)
    .Buttons(4).Value = ButtonPress(DECMD_ITALIC)
    .Buttons(4).Enabled = ButtonEnable(DECMD_ITALIC)
    .Buttons(5).Value = ButtonPress(DECMD_UNDERLINE)
    .Buttons(5).Enabled = ButtonEnable(DECMD_UNDERLINE)
    .Buttons(7).Value = ButtonPress(DECMD_JUSTIFYLEFT)
    .Buttons(7).Enabled = ButtonEnable(DECMD_JUSTIFYLEFT)
    .Buttons(8).Value = ButtonPress(DECMD_JUSTIFYCENTER)
    .Buttons(8).Enabled = ButtonEnable(DECMD_JUSTIFYCENTER)
    .Buttons(9).Value = ButtonPress(DECMD_JUSTIFYRIGHT)
    .Buttons(9).Enabled = ButtonEnable(DECMD_JUSTIFYRIGHT)
    .Buttons(11).Value = ButtonPress(DECMD_INDENT)
    .Buttons(11).Enabled = ButtonEnable(DECMD_INDENT)
    .Buttons(12).Value = ButtonPress(DECMD_OUTDENT)
    .Buttons(12).Enabled = ButtonEnable(DECMD_OUTDENT)
    .Buttons(13).Value = ButtonPress(DECMD_UNORDERLIST)
    .Buttons(13).Enabled = ButtonEnable(DECMD_UNORDERLIST)
    .Buttons(14).Value = ButtonPress(DECMD_ORDERLIST)
    .Buttons(14).Enabled = ButtonEnable(DECMD_ORDERLIST)
    .Buttons(16).Value = ButtonPress(DECMD_BRING_FORWARD)
    .Buttons(16).Enabled = ButtonEnable(DECMD_BRING_FORWARD)
    .Buttons(17).Value = ButtonPress(DECMD_SEND_BACKWARD)
    .Buttons(17).Enabled = ButtonEnable(DECMD_SEND_BACKWARD)
    Me.cpFore.Enabled = ButtonEnable(DECMD_SETFORECOLOR)
    Me.cpBack.Enabled = ButtonEnable(DECMD_SETBACKCOLOR)
    .Buttons(22).Value = ButtonPress(DECMD_MAKE_ABSOLUTE)
    .Buttons(22).Enabled = ButtonEnable(DECMD_MAKE_ABSOLUTE)
End With

End Select
1
End Sub

Public Sub Timer2_Timer()

RefreshEditBar

End Sub

Private Sub Timer3_Timer()
DoEvents
With staMain

    On Error GoTo 1
    DoEvents
    On Error Resume Next
        .Panels(4).Text = "Title: " & ActiveForm.DHTML1.DocumentTitle
        .Panels(5).Text = "File on:  " & ActiveForm.DHTML1.CurrentDocumentPath
        .Panels(3).Text = "ROW " & ActiveForm.rt1.GetLineFromChar(ActiveForm.rt1.SelStart) + 1
    
    End With

1
End Sub


Private Sub Timer4_Timer()
DoEvents
On Error GoTo 1
Select Case ActiveForm.SSTab1.Tab
Case 0, 1
    For i = 0 To 4
    SSTab1.TabEnabled(i) = True
    Next
Case 2
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    SSTab1.TabEnabled(4) = False
End Select
1 End Sub

Private Sub Timer5_Timer()
DoEvents
staMain.Refresh
staProp.Refresh
If DatCustom.Recordset.RecordCount = 0 Then
rtc.Enabled = False
Else
rtc.Enabled = True
End If

End Sub

Private Sub tv1_DblClick()

On Error GoTo NoTabs

If ActiveForm.SSTab1.Tab = 1 Or ActiveForm.SSTab1.Tab = 2 Then Exit Sub

GoTo Start

NoTabs: Exit Sub

Start:

Dim c As Integer

On Error GoTo 2

    With ActiveForm.DHTML1.DOM
    
                Dim p As POINTAPI
                i = GetCursorPos(p)
                
                    'Set the startup position of the form
                    frmElementHTML.Top = p.Y * Screen.TwipsPerPixelY
                    frmElementHTML.Left = p.X * Screen.TwipsPerPixelX
                    
                c = CInt(Mid(tv1.SelectedItem.Key, 4)) - 1
    
        Select Case tv1.SelectedItem.Parent.Text
        
        
            Case "All Elements"
                frmElementHTML.Text1.Text = .All(c).innerHTML
                frmElementHTML.Show vbModal, Me
                
                If frmElementHTML.Value = "-1" Then Exit Sub
                .All(CInt(c) - 1).innerHTML = frmElementHTML.Value
                
            Case "Jave Applets"
                frmElementHTML.Text1.Text = .applets(c).innerHTML
                frmElementHTML.Show vbModal, Me
                
                If frmElementHTML.Value = "-1" Then Exit Sub
                .All(CInt(c) - 1).innerHTML = frmElementHTML.Value
                
            Case "Plug-ins"
                frmElementHTML.Text1.Text = .plugins(c).innerHTML
                frmElementHTML.Show vbModal, Me
                
                If frmElementHTML.Value = "-1" Then Exit Sub
                .All(CInt(c) - 1).innerHTML = frmElementHTML.Value
                
            Case "Images"
                frmElementHTML.Text1.Text = .images(c).innerHTML
                frmElementHTML.Show vbModal, Me
                
                If frmElementHTML.Value = "-1" Then Exit Sub
                .All(CInt(c) - 1).innerHTML = frmElementHTML.Value
                
            Case "Hyperlinks"
                frmElementHTML.Text1.Text = .All(c).innerHTML
                frmElementHTML.Show vbModal, Me
                
                If frmElementHTML.Value = "-1" Then Exit Sub
                .All(CInt(c) - 1).innerHTML = frmElementHTML.Value
                
            Case "Forms"
                frmElementHTML.Text1.Text = .Forms(c).innerHTML
                frmElementHTML.Show vbModal, Me
                
                If frmElementHTML.Value = "-1" Then Exit Sub
                .All(CInt(c) - 1).innerHTML = frmElementHTML.Value
                
            Case "Scripts"
                frmElementHTML.Text1.Text = .scripts(c).innerHTML
                frmElementHTML.Show vbModal, Me
                
                If frmElementHTML.Value = "-1" Then Exit Sub
                .All(CInt(c) - 1).innerHTML = frmElementHTML.Value
                
        End Select
        
    End With
2
End Sub

Sub FillHTMLTree()
    Dim a As IHTMLElement, i As Integer, strImg As String

            'Built select case list code:
            'Code in clipboard
            'Dim s As String
            'For i = 1 To imlHTML.ListImages.Count
            's = Clipboard.GetText
            'Clipboard.Clear
            'Clipboard.SetText s & " " & AP & imlHTML.ListImages(i).Key & AP & ","
            'Next
            
On Error Resume Next
With ActiveForm.DHTML1.DOM
tv1.Nodes.Clear
    tv1.Nodes.Add , , "all", "All Elements", "all"
    tv1.Nodes.Add 1, 2, "applet", "Jave Applets", "applet"
    tv1.Nodes.Add 1, 2, "embed", "Plug-ins", "embed"
    tv1.Nodes.Add 1, 2, "img", "Images", "img"
    tv1.Nodes.Add 1, 2, "a", "Hyperlinks", "a"
    tv1.Nodes.Add 1, 2, "form", "Forms", "form"
    tv1.Nodes.Add 1, 2, "script", "Scripts", "script"
                
            For i = 0 To .All.length - 1
                Set a = .All(i)
                On Error Resume Next
                strImg = ""
                
                    Select Case LCase(a.tagName)
                        Case "img", "a", "p", "span", "input", "script", "embed", "font", "object", "all", "hr", "html", "meta", "div", "label", "li", "img", "a", "p", "span", "input", "script", "embed", "font", "object", "all", "hr", "html", "meta", "div", "label", "li"
                        strImg = LCase(a.tagName)
                    End Select
                
                    If strImg = "" Then
                        tv1.Nodes.Add "all", 4, "all" & i, a.tagName
                    Else
                        tv1.Nodes.Add "all", 4, "all" & i, a.tagName, strImg
                    End If
            Next
    
    For i = 0 To .images.length - 1
    Set a = .images(i)
    tv1.Nodes.Add "img", 4, "img" & i, a.src, "img"
    Next
    
    For i = 0 To .Forms.length - 1
    Set a = .Forms(i)
    tv1.Nodes.Add "form", 4, "frm" & i, a.id, "form"
    Next
    
    For i = 0 To .All.length - 1
    Set a = .All(i)
    If UCase(a.tagName) = "A" Then tv1.Nodes.Add "a", 4, "ahr" & i, a.href, "a"
    Next
    
    For i = 0 To .applets.lenght - 1
    Set a = .applets(i)
    tv1.Nodes.Add "applet", 4, "apt" & i, a.id, "applet"
    Next
    
    For i = 0 To .embeds.length - 1
    Set a = .embeds(i)
    tv1.Nodes.Add "embed", 4, "ebd" & i, a.scr, "embed"
    Next
    
    For i = 0 To .scripts.length - 1
    Set a = .scripts(i)
    tv1.Nodes.Add "script", 4, "spt" & i, a.Text, "script"
    Next
    
End With
1
End Sub

Private Sub TmrStart_Timer()
    Static n As Integer
    n = n + 1
    
If WordCommand = "" Then
    
        If n = 5 Then
        DoEvents
        'mnuNew_Design_Click
        TmrStart.Enabled = False
        On Error Resume Next
        frmStart.Show vbModal
        Exit Sub
        End If

ElseIf LCase(Right(WordCommand, 3)) = "pcs" Then
        MsgBox "Opening Page Creator Web Site File", vbInformation
        TmrStart.Enabled = False
        Exit Sub

Else
    
    'MsgBox "LoadFile"
    If n = 2 Then
    TmrStart.Enabled = False
        
    On Error GoTo 2
    OpenFile WordCommand, WordCommand
    Exit Sub
2
    MsgBox Err.Number & " Error opening file123345!", vbExclamation
    mnuNew_Design_Click
    TmrStart.Enabled = False
    
    End If
    
    
End If


    
End Sub

Private Sub Tree1_OnDirChanged()
On Error GoTo 1
File1.Path = Tree1.Path
1 End Sub

