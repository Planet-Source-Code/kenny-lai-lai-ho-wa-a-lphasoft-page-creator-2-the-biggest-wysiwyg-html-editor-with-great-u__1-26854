VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Task"
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
Private mvarTaskName As String
Private mvarDueDate As String
Private mvarStartDate As String
Private mvarTaskType As String
Private mvarDone As Boolean
Private mvarResponsible As String

Public Function GenerateCode() As String

GenerateCode = "TSK" & _
               IIf(Me.TaskName = "", "A Task", Me.TaskName) & _
               Me.TaskName & "&!@#$1" & _
               Me.TaskType & "&!@#$2" & _
               Me.Responsible & "&!@#$3" & _
               Me.StartDate & "&!@#$4" & _
               Me.DueDate & "&!@#%5" & _
               IIf(Me.Done = True, "True", "False") & _
               "&!@#%6" & Me.Description
               
End Function

Public Property Let Responsible(ByVal vData As String)
    mvarResponsible = vData
End Property


Public Property Get Responsible() As String
    Responsible = mvarResponsible
End Property



Public Property Let Done(ByVal vData As Boolean)
    mvarDone = vData
End Property


Public Property Get Done() As Boolean
    Done = mvarDone
End Property



Public Property Let TaskType(ByVal vData As String)
    mvarTaskType = vData
End Property


Public Property Get TaskType() As String
    TaskType = mvarTaskType
End Property



Public Property Let StartDate(ByVal vData As String)
    mvarStartDate = vData
End Property


Public Property Get StartDate() As String
    StartDate = mvarStartDate
End Property



Public Property Let DueDate(ByVal vData As String)
    mvarDueDate = vData
End Property


Public Property Get DueDate() As String
    DueDate = mvarDueDate
End Property



Public Property Let TaskName(ByVal vData As String)
    mvarTaskName = vData
End Property


Public Property Get TaskName() As String
Attribute TaskName.VB_UserMemId = 0
    TaskName = mvarTaskName
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

