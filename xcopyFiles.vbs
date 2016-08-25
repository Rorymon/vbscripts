'==========================================================================
'
' DESCRIPTION: Copies files.
'
' NAME: XcopyFiles.vbs
'
' AUTHOR: Russell Hill
' DATE  : 07/06/2016
'
' USAGE: Copies files in a given path using xcopy. 
'
' PREREQ: Only use if there is no other method to copy the files.
'
' COMMENTS: Files do not have to exist. Can contain hardcoded paths.
'==========================================================================
Option Explicit

Dim strFromPathName, strToPathName, arrFromFileNames, arrToFileNames
Dim blnFilesOnly

	'Enter path from which the files are to be copied from. This 
	'can be a UNC path or a environmental variable, hence to get  
	'the C drive use %SystemDrive% (i.e. %SystemDrive%\Folder or
	'\\UncPath\Folder).
	
strFromPathName = ""

	'Enter path to which the files are to be copied to. This can 
	'be a UNC path or a environmental variable, hence to get  
	'the C drive use %SystemDrive% (i.e. %SystemDrive%\Folder or
	'\\UncPath\Folder).
	
strToPathName = ""

	'Enter files to be copied here. Will also except wildcards (i.e. *).
	
arrFromFileNames = Array("*")
                       
	'Enter files to be renamed here. Leave empty if you do not wish to 
	'rename files. Size of array must be the same as arrFromFileNames
	'even if empty.
	
arrToFileNames = Array("")

	'Set to True to copy files only. Set to False to copy files and 
	'folders. To copy files and folders arrFromFileNames must be set
	'to wildcard (i.e. *)
	
blnFilesOnly = False
                      
CopyFiles strFromPathName, strToPathName, arrFromFileNames, arrToFileNames, blnFilesOnly

Function CopyFiles(strFromPath, strToPath, arrFromFiles(), arrToFiles(), blnFiles)
	Dim objShell, objFS, strDir, strFromFile, strToFile, strEnv
	Dim arrPath, arrFilPath, i, strCmd
	Set objShell = CreateObject("WScript.Shell") 
	Set objFS = CreateObject("Scripting.FileSystemObject")
	If (Right(strFromPath, 1) <> "\") Then
		strFromPath = strFromPath & "\"
	End If
	If (Right(strToPath, 1) <> "\") Then
		strToPath = strToPath & "\"
	End If
	If (Left(strFromPath, 1) = "%") Then
		arrPath = Split(strFromPath, "\")
		strEnv = arrPath(0)
		arrFilPath = Filter(arrPath, strEnv, False, vbTextCompare)
		strFromPath = Join(arrFilPath, "\")
		strDir = objShell.ExpandEnvironmentStrings(strEnv)
		If (Right(strDir, 1) <> "\") Then
			strDir = strDir & "\"
		End If
		strFromPath = strDir & strFromPath
	End If	
	If (Left(strToPath, 1) = "%") Then
		arrPath = Split(strToPath, "\")
		strEnv = arrPath(0)
		arrFilPath = Filter(arrPath, strEnv, False, vbTextCompare)
		strToPath = Join(arrFilPath, "\")
		strDir = objShell.ExpandEnvironmentStrings(strEnv)
		If (Right(strDir, 1) <> "\") Then
			strDir = strDir & "\"
		End If
		strToPath = strDir & strToPath
	End If	
	On Error Resume Next
	For i = 0 To UBound(arrFromFiles)
		strFromFile = strFromPath & arrFromFiles(i)
		strToFile = strToPath & arrToFiles(i)
		If (blnFiles) Then
			strCmd = "cmd.exe /c xcopy /y """ & strFromFile & """ /s """ & strToFile & """"
		Else
			strCmd = "cmd.exe /c xcopy /y """ & strFromFile & """ /s /e """ & strToFile & """"
		End If
		objShell.Run strCmd
	Next	
End Function
