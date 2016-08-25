'==========================================================================
'
' DESCRIPTION: Copies contents of one file to another.
'
' NAME: CopyFileContents.vbs
'
' AUTHOR: Russell Hill
' DATE  : 07/06/2016
'
' USAGE: Copies contents of one file to another without losing ACL. 
'
' PREREQ: Files to be copied do not have to exist. An empty file will be 
'         created. Folders must exist.
'
' COMMENTS: None.
'==========================================================================
Option Explicit

Dim strFromPathName, strToPathName, arrFromFileNames, arrToFileNames

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

	'Enter files to be copied here. Will not except wildcards.
	
arrFromFileNames = Array("", _
                         "", _
                         "")
                       
	'Enter files to be renamed here. Leave empty if you do not wish to 
	'rename files. Size of array must be the same as arrFromFileNames
	'even if empty.
	
arrToFileNames = Array("", _
                       "", _
                       "")
                       
CopyFileContents strFromPathName, strToPathName, arrFromFileNames, arrToFileNames

Function CopyFileContents(strFromPath, strToPath, arrFromFiles(), arrToFiles())
	Dim objShell, objFS, objFromFile, objToFile
	Dim strDir, strFromFile, strToFile, strEnv, strLine
	Dim arrPath, arrFilPath, i
	Const READ=1, WRITEFILE=2, APPEND=8
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
	If Not(objFS.FolderExists(strFromPath)) Then
		Exit Function
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
	If Not(objFS.FolderExists(strToPath)) Then
		Exit Function
	End If
	If (arrToFiles(0) = "") Then
		For i = 0 To UBound(arrFromFiles)
			arrToFiles(i) = arrFromFiles(i)
		Next
	End If	
	On Error Resume Next
	For i = 0 To UBound(arrFromFiles)
		strFromFile = strFromPathName & arrFromFiles(i)
		Set objFromFile = objFS.OpenTextFile(strFromFile, READ)
		strLine = objFromFile.ReadAll
		strToFile = strToPathName & arrToFiles(i)
		Set objToFile = objFS.OpenTextFile(strToFile, WRITEFILE, True) 
		objToFile.Write strLine
		objFromFile.Close
		objToFile.Close	
	Next	
End Function
