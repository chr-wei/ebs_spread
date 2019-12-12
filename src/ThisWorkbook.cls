VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call DeleteShortcuts
    
    'Move all tasks to virtual sheets prior to closing the workbook. This prevents high memory usage on next start.
    Call Planning.UI_VirtualizeTaskSheets
End Sub



Private Sub Workbook_Open()
    Call CreateShortcuts
End Sub



Function CreateShortcuts()
    Call Application.OnKey("+^{C}", "Planning.UI_CopyTask")                         'Ctrl+Shift+C
    Call Application.OnKey("+^{R}", "Planning.UI_DeleteSelectedTask")               'Ctrl+Shift+R
    Call Application.OnKey("+^{N}", "Planning.UI_AddNewTask")                       'Ctrl+Shift+N
    Call Application.OnKey("+^{S}", "Planning.UI_StartSelectedTask")                'Ctrl+Shift+S
    Call Application.OnKey("+^{X}", "Planning.UI_EndAllTasks")                      'Ctrl+Shift+X
    Call Application.OnKey("+^{F}", "Planning.UI_AddFiveteenMinutesToSelectedTask") 'Ctrl+Shift+F
    Call Application.OnKey("+^{+}", "Planning.UI_ShiftTaskUp")                      'Ctrl+Shift+'+'
    Call Application.OnKey("+^{-}", "Planning.UI_ShiftTaskDown")                    'Ctrl+Shift+'-'
    Call Application.OnKey("+^{M}", "Planning.UI_RunEbs")                           'Ctrl+Shift+'M'
    Call Application.OnKey("+^{T}", "Planning.UI_CollectTotalTimesSpent")           'Ctrl+Shift+'T'
    Call Application.OnKey("+^{G}", "Planning.UI_GatherTasks")                      'Ctrl+Shift+'T'
End Function



Function DeleteShortcuts()
    Call Application.OnKey("+^{C}", "") 'Ctrl+Shift+C
    Call Application.OnKey("+^{R}", "") 'Ctrl+Shift+R
    Call Application.OnKey("+^{N}", "") 'Ctrl+Shift+N
    Call Application.OnKey("+^{S}", "") 'Ctrl+Shift+S
    Call Application.OnKey("+^{X}", "") 'Ctrl+Shift+X
    Call Application.OnKey("+^{F}", "") 'Ctrl+Shift+F
    Call Application.OnKey("+^{+}", "") 'Ctrl+Shift+'+'
    Call Application.OnKey("+^{-}", "") 'Ctrl+Shift+'-'
    Call Application.OnKey("+^{M}", "") 'Ctrl+Shift+'M'
    Call Application.OnKey("+^{T}", "") 'Ctrl+Shift+'T'
    Call Application.OnKey("+^{G}", "") 'Ctrl+Shift+'G'
End Function
