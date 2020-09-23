VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
#Const DebugMode = 1


#If DebugMode Then
    Private mlClassDebugID As Long
#End If
Private mvarDescription As String
Private mvarIsGood As Boolean
Private mvarCommentName As String

Public Function GenerateCode() As String
GenerateCode = "CMT" & _
               IIf(Me.CommentName = "", "A Comment", Me.CommentName) & _
               "&!@#$1" & IIf(Me.IsGood = True, "True", "False") & _
               "&!@#$2" & Me.Description
End Function

Public Property Let CommentName(ByVal vData As String)
    mvarCommentName = vData
End Property


Public Property Get CommentName() As String
Attribute CommentName.VB_UserMemId = 0
    CommentName = mvarCommentName
End Property



Public Property Let IsGood(ByVal vData As Boolean)
    mvarIsGood = vData
End Property


Public Property Get IsGood() As Boolean
    IsGood = mvarIsGood
End Property



Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property


Public Property Get Description() As String
    Description = mvarDescription
End Property





Private Sub Class_Initialize()
    #If DebugMode Then
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub


Private Sub Class_Terminate()
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub


#If DebugMode Then
    Public Property Get ClassDebugID()
        ClassDebugID = mlClassDebugID
    End Property
#End If

