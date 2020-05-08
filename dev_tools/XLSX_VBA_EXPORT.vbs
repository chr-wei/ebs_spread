' XLSX_VBA_EXPORT.vbs
'
' Extracts VBA objects from an Excel workbook.  Requires Microsoft Excel.
' Usage:
'  WScript XLSX_VBA_EXPORT.vbs <input file> <output folder without trailing backslash>

Option Explicit

' MsoAutomationSecurity
Const msoAutomationSecurityForceDisable = 3
' OpenTextFile iomode
Const ForReading = 1
Const ForAppending = 8

Dim App, AutoSec, Doc, FileSys
Set FileSys = CreateObject("Scripting.FileSystemObject")
If FileSys.FileExists(WScript.Arguments(1)) Then
	FileSys.DeleteFile WScript.Arguments(1)
End If
Set App = CreateObject("Excel.Application")

On Error Resume Next

App.DisplayAlerts = False
AutoSec = App.AutomationSecurity
App.AutomationSecurity = msoAutomationSecurityForceDisable
Err.Clear
Dim Component, componentCount, referenceCount, cmpIdx, refIdx, tmpIdx, Names(), Reference, TgtFilepath, TgtFile, TmpFile, TmpFilenames()
Set Doc = App.Workbooks.Open(WScript.Arguments(0), False, True)

If Err = 0 Then
	componentCount = Doc.VBProject.VBComponents.Count

	If componentCount > 0 Then
		ReDim Names(componentCount - 1)
		ReDim TmpFilenames(componentCount - 1)
		cmpIdx = 0
		For Each Component In Doc.VBProject.VBComponents
			Names(cmpIdx) = Component.Name
		 	cmpIdx = cmpIdx + 1
		Next
		Names = SortNames(Names)
		For cmpIdx = 0 To componentCount - 1
			TgtFilepath = FileSys.GetAbsolutePathName(WScript.Arguments(1) & "\" & Names(cmpIdx))
			Doc.VBProject.VBComponents(Names(cmpIdx)).Export TgtFilepath
			'WScript.Echo(TgtFilepath)
		Next
	End If
	referenceCount = Doc.VBProject.References.Count
	If referenceCount > 0 Then
	     Set TgtFile = FileSys.CreateTextFile(WScript.Arguments(1) & "\" & "REFERENCES", True)
		 TgtFile.WriteLine "'********REFERENCES********"
		ReDim Names(referenceCount - 1)
		refIdx = 0
		For Each Reference In Doc.VBProject.References
			Names(refIdx) = Reference.Name & Chr(9) & Reference.Description
			refIdx = refIdx + 1
		Next
		Names = SortNames(Names)
		For refIdx = 0 To referenceCount - 1
			If Names(refIdx) <> "" Then
				TgtFile.WriteLine "'" & Names(refIdx)
			End If
		Next
	End If
	TgtFile.Close
	Doc.Close False
End If

App.AutomationSecurity = AutoSec
App.Quit

For tmpIdx = 0 To UBound(TmpFilenames)
	If FileSys.FileExists(TmpFilenames(tmpIdx)) Then
		FileSys.DeleteFile TmpFilenames(tmpIdx)
	End If
Next

Sub SortNames(Names)
	Dim nameCount, J, T
	For nameCount = UBound(Names) - 1 To 0 Step -1
		For J = 0 To componentCount
			If Names(J) > Names(J + 1) Then
				T = Names(J + 1)
				Names(J + 1) = Names(J)
				Names(J) = T
			End If
		Next
	Next
	SortNames = Names
End Sub
