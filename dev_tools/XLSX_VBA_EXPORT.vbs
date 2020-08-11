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
Const ForWriting = 2
Const ForAppending = 8

Dim App, AutoSec, Doc, FileSys
Set FileSys = CreateObject("Scripting.FileSystemObject")
Set App = CreateObject("Excel.Application")

'On Error Resume Next

App.DisplayAlerts = False
AutoSec = App.AutomationSecurity
App.AutomationSecurity = msoAutomationSecurityForceDisable
Err.Clear
Dim Component, componentCount, referenceCount, cmpIdx, refIdx, tmpIdx, Names(), Reference, TgtFilepath, TgtFile
Set Doc = App.Workbooks.Open(WScript.Arguments(0), False, True)

If Err = 0 Then

    If FileSys.FolderExists(WScript.Arguments(1)) Then
	    FileSys.DeleteFolder WScript.Arguments(1)
    End If

	FileSys.CreateFolder WScript.Arguments(1)

	componentCount = Doc.VBProject.VBComponents.Count

	If componentCount > 0 Then
		ReDim Names(componentCount - 1)
		cmpIdx = 0
		For Each Component In Doc.VBProject.VBComponents
		    Names(cmpIdx) = Component.Name
		 	cmpIdx = cmpIdx + 1
		Next

		For cmpIdx = 0 To componentCount - 1
			TgtFilepath = FileSys.GetAbsolutePathName(WScript.Arguments(1) & "\" & Names(cmpIdx))
			Doc.VBProject.VBComponents(Names(cmpIdx)).Export TgtFilepath
			DeleteEmptyTrailingLines TgtFilepath
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
	
		For refIdx = 0 To referenceCount - 1
			If Names(refIdx) <> "" Then
				TgtFile.WriteLine "'" & Names(refIdx)
			End If
		Next
	End If
	TgtFile.Close
	Doc.Close False
Else
    WScript.Echo("Could not open workbook '" & WScript.Arguments(0) & "'")
End If

App.AutomationSecurity = AutoSec
App.Quit
WScript.Quit Err

Function DeleteEmptyTrailingLines(filepath)
	Dim tgtFile, strLine, strNewContents, strHeldBack
	Set tgtFile = Filesys.OpenTextFile(filepath, ForReading)
	strHeldBack = ""
	strLine = ""
	
	Do Until tgtFile.AtEndOfStream
    	strLine = TgtFile.Readline
		If Len(Trim(strLine)) > 0 Then
			strNewContents = strNewContents & strHeldBack & strLine
			strHeldBack = vbCrLf
    	Else
			'Hold back content with empty lines. These will only be added to content if non-empty lines follow
			'--> Prevents trailing lines at the end of the file.
			strHeldBack = strHeldBack & strLine & vbCrLf
		End If
	Loop

	TgtFile.Close

	Set tgtFile = Filesys.OpenTextFile(filepath, ForWriting)
	TgtFile.Write strNewContents
	TgtFile.Close
End Function