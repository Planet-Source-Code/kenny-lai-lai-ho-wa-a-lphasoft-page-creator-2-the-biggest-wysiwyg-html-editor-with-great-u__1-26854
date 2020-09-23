VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Report"
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
Private mvarReportType As String
Private mvarReportName As String
Private mvarPerson As String
Private mvarDescription As String
Public Function GenerateCode() As String
GenerateCode = "RPT" & _
               IIf(Me.ReportName = "", "A Report", Me.ReportName) & "&!@#$1" & _
               Me.ReportType & "&!@#$2" & _
               Me.Description
End Function

Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property


Public Property Get Description() As String
    Description = mvarDescription
End Property



Public Property Let Person(ByVal vData As String)
    mvarPerson = vData
End Property


Public Property Get Person() As String
    Person = mvarPerson
End Property



Public Property Let ReportName(ByVal vData As String)
    mvarReportName = vData
End Property


Public Property Get ReportName() As String
Attribute ReportName.VB_UserMemId = 0
    ReportName = mvarReportName
End Property



Public Property Let ReportType(ByVal vData As String)
    mvarReportType = vData
End Property


Public Property Get ReportType() As String
    ReportType = mvarReportType
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

