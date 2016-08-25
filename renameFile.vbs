'==========================================================================
'
' DESCRIPTION: Copies files.
'
' NAME: RenameFile.vbs
'
' AUTHOR: Russell Hill
' DATE  : 07/06/2016
'
' USAGE: Renames a file. 
'
' PREREQ: Only use if there is no other method to rename the files.
'
' COMMENTS:Can contain hardcoded paths.
'==========================================================================
Option Explicit

Dim strFromPathName, strFromFileName, strToFileName

	'Enter path from which the files are to be copied from. This 
	'should be in form %SystemDrive%\Folder  
	
strFromPathName = ""

	'Enter file to be renamed here.
	
strFromFileName = ""
                       
	'Enter new file name here. 
	
strToFileName = ""
                       
RenameFile strFromPathName, strFromFileName, strToFileName

Function RenameFile(strFromPath, strFromFile, strToFile)
	Dim objShell, objFS, strDir, strEnv, arrPath, arrFilPath, i
	Set objShell = CreateObject("WScript.Shell") 
	Set objFS = CreateObject("Scripting.FileSystemObject")
	If (Right(strFromPathName, 1) <> "\") Then
		strFromPathName = strFromPathName & "\"
	End If

	If (Left(strFromPathName, 1) = "%") Then
		arrPath = Split(strFromPathName, "\")
		strEnv = arrPath(0)
		arrFilPath = Filter(arrPath, strEnv, False, vbTextCompare)
		strFromPathName = Join(arrFilPath, "\")
		strDir = objShell.ExpandEnvironmentStrings(strEnv)
		If (Right(strDir, 1) <> "\") Then
			strDir = strDir & "\"
		End If
		strFromPathName = strDir & strFromPathName
	End If	
	
	On Error Resume Next
		strFromFile = strFromPathName & strFromFile
		strToFile = strFromPathName & strToFile
		objFS.MoveFile strFromFile, strToFile
		
	Set objFS = Nothing
	Set objShell = Nothing		
End Function
