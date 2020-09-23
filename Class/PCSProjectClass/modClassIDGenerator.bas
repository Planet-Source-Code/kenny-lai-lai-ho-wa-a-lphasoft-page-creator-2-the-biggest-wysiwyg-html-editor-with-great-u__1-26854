Attribute VB_Name = "modClassIDGenerator"
Function GetNextClassDebugID() As Long
    '物件類別 ID 產生器
    Static lClassDebugID As Long
    lClassDebugID = lClassDebugID + 1
    GetNextClassDebugID = lClassDebugID
End Function

