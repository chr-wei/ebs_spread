VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaskSheetTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_FollowHyperlink(ByVal Target As hyperlink)
    Call Utils.RunTryCatchedCall(f:="TaskUtils.HandleFollowHyperlink", arg1:=Target, screenUpdating:=False)
End Sub
