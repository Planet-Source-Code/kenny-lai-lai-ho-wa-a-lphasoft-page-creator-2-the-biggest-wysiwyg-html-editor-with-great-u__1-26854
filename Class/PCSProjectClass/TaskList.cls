VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Task"
Attribute VB_Ext_KEY = "Member0" ,"Task"
Option Base 1

Private mCol As New Collection

Public Function Add(Optional ByVal TaskName As String, _
                                     Optional ByVal TaskType As String, _
                                     Optional ByVal StartDate As String, _
                                     Optional ByVal DueDate As String, _
                                     Optional ByVal Responsible As String, _
                                     Optional ByVal Done As Boolean, _
                                     Optional ByVal Description As String, _
                                     Optional sKey As String) _
                                                                                                             As Task
    Dim objNewMember As Task
    Set objNewMember = New Task

With objNewMember
    .TaskName = IIf(IsMissing(TaskName), "", TaskName)
    .TaskType = IIf(IsMissing(TaskType), "", TaskType)
    .StartDate = IIf(IsMissing(StartDate), "", StartDate)
    .DueDate = IIf(IsMissing(DueDate), "", DueDate)
    .Responsible = IIf(IsMissing(Responsible), "", Responsible)
    .Done = IIf(IsMissing(Done), False, Done)
    .Description = IIf(IsMissing(Description), "", Description)
End With

    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Task
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
   Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

