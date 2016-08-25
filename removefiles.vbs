'==========================================================================
'
' DESCRIPTION: Removes files.
'
' NAME: RemoveFiles.vbs
'
' AUTHOR: Russell Hill
' DATE  : 07/06/20016
'
' USAGE: Removes files in a given path. 
'
' PREREQ: Only use if there is no other method to remove the files.
'
' COMMENTS: Files do not have to exist.
'==========================================================================
Option Explicit

Dim strPathName, arrRemoveFiles

	'Enter path to the files here minus the C:\ at front and backslash  
	'at end (i.e. WINNT\System32\). Leave empty for files in C:\.
	
strPathName = ""

	'Enter files to be removed here. Will also except wildcards (i.e. *).
	
arrRemoveFiles = Array("", _
                       "", _
                       "")

RemoveStringsInFile strPathName, arrRemoveFiles

Function RemoveStringsInFile(strPath, arrFiles())
	Dim objShell, objFS, strSysDir, strFile, i
	Set objShell = CreateObject("Wscript.shell")
	Set objFS = CreateObject("Scripting.FileSystemObject")
	strSysDir = objShell.ExpandEnvironmentStrings("%SystemDrive%")
	If (Right(strSysDir, 1) <> "\") Then
		strSysDir = strSysDir & "\"
	End If
	On Error Resume Next
	For i = 0 To UBound(arrFiles)
		strFile = strSysDir & strPath & arrFiles(i)
		objFS.DeleteFile strFile, True
	Next
	Set objFS = Nothing
	Set objShell = Nothing		
End Function
