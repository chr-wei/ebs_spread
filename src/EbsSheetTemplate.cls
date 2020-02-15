VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EbsSheetTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If StrComp(Target.Parent.name, Constants.EBS_SHEET_TEMPLATE_NAME) = 0 Then
        'Prohibit making changes in the template
        Exit Sub
    End If
    
    Call Utils.RunTryCatchedCall("EbsUtils.HandleEbsSheetChanges", , Target, screenUpdating:=False)
End Sub
