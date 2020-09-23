VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebPage"
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
Private mvarFilename As String
Private mvarResponsible As String

Public Function GenerateCode() As String

GenerateCode = "WPG" & _
             Me.Filename & "&!@#$1" & _
             Me.Responsible

End Function

Public Property Let Responsible(ByVal vData As String)
    mvarResponsible = vData
End Property


Public Property Get Responsible() As String
    Responsible = mvarResponsible
End Property



Public Property Let Filename(ByVal vData As String)
    mvarFilename = vData
End Property


Public Property Get Filename() As String
Attribute Filename.VB_UserMemId = 0
    Filename = mvarFilename
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

