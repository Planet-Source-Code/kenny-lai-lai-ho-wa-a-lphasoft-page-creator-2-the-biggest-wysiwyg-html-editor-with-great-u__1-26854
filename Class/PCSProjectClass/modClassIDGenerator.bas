Attribute VB_Name = "modClassIDGenerator"
Function GetNextClassDebugID() As Long
    'ª«¥óÃþ§O ID ²£¥Í¾¹
    Static lClassDebugID As Long
    lClassDebugID = lClassDebugID + 1
    GetNextClassDebugID = lClassDebugID
End Function

