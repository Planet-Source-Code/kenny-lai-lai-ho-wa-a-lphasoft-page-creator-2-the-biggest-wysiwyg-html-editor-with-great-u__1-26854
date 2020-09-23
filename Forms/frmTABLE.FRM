VERSION 5.00
Begin VB.Form frmTable 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Insert Table"
   ClientHeight    =   3015
   ClientLeft      =   6630
   ClientTop       =   7020
   ClientWidth     =   5745
   Icon            =   "frmTABLE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.Frame Frame1 
      Caption         =   "Table"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton OkCmd 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton CancelCmd 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Rows 
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Cols 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox TableAttrs 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox CellAttrs 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox TableCaption 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label RowLabel 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Number of rows:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label ColLabel 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Number of columns:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label TableTagLabel 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Table Tag Attributes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label CellTagLabel 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Cell Tag Attributes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label CaptionLabel 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Caption:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1998 Microsoft Corporation.
' All rights reserved.
Public tableParam As DEInsertTableParam
Public Value As Boolean

Private Sub CancelCmd_Click()
Value = False
Me.Hide
End Sub

Private Sub Form_Load()
Value = False
    ' create the table parameter object
    Set tableParam = CreateObject("DEInsertTableParam.DEInsertTableParam.1")
    Rows = tableParam.NumRows
    Cols = tableParam.NumCols
    TableAttrs = tableParam.TableAttrs
    CellAttrs = tableParam.CellAttrs
    TableCaption = tableParam.Caption

End Sub

Private Sub OkCmd_Click()
    Value = True
    If Rows = "" Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    ElseIf IsNumeric(Rows) = False Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    ElseIf Rows <= 0 Then
        MsgBox "Please specify a positive integer for the number of table rows.", vbCritical
        Exit Sub
    End If
       
    If Cols = "" Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    ElseIf IsNumeric(Cols) = False Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    ElseIf Cols <= 0 Then
        MsgBox "Please specify a positive integer for the number of table columns.", vbCritical
        Exit Sub
    End If
    
    tableParam.NumRows = Rows
    tableParam.NumCols = Cols
    tableParam.CellAttrs = "Width=" & Int(100 / Cols) & "%"
    
    If Len(TableAttrs.Text) Then
        tableParam.TableAttrs = TableAttrs.Text
    End If
    If Len(CellAttrs.Text) Then
        tableParam.CellAttrs = CellAttrs.Text
    End If
    If Len(TableCaption.Text) Then
        tableParam.Caption = TableCaption.Text
    End If
    Value = True
Me.Hide
End Sub

Private Sub Rows_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Cols_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub
