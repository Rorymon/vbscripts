'==========================================================================
'
' DESCRIPTION: Appends lines to a file.
'
' NAME: AppendToFile.vbs
'
' AUTHOR: Russell Hill
' DATE  : 07/06/2016
'
' USAGE: Appends lines to a file. 
'
' PREREQ: File must exist.
'
' COMMENTS: Amended to take UNC path or a environmental variable.
'==========================================================================
Option Explicit

Dim strFileName, arrAppendStrings

	'Enter the filename to which you wish to append to. This filename
	'can contain a UNC path or a environmental variable, hence to get  
	'the users profile use %UserProfile% (i.e. %UserProfile%\Folder\file.txt
	'or \\UncPath\Folder\file.txt).
	
strFileName = ""

	'Enter strings to be appended here. 
	
arrAppendStrings = Array("", _
				 		 "", _
						 "")

AppendStringsToFile strFileName, arrAppendStrings

Function AppendStringsToFile(strFile, arrStrings())
	Dim objShell, objFS, objFile, strDir, strLine, strEnv
	Dim arrPath, arrFilPath, i
	i = 0
	Const APPEND = 8
	Set objShell = CreateObject("Wscript.shell")
	Set objFS = CreateObject("Scripting.FileSystemObject")
	If (Left(strFile, 1) = "%") Then
		arrPath = Split(strFile, "\")
		strEnv = arrPath(0)
		arrFilPath = Filter(arrPath, strEnv, False, vbTextCompare)
		strFile = Join(arrFilPath, "\")
		strDir = objShell.ExpandEnvironmentStrings(strEnv)
		If (Right(strDir, 1) <> "\") Then
			strDir = strDir & "\"
		End If
		strFile = strDir & strFile
	End If	
	On Error Resume Next
	Set objFile = objFS.OpenTextFile(strFile, APPEND)
	If (Err.Number <> 0) Then Exit Function
	objFile.WriteBlankLines(1)
   	For i = 0 To UBound(arrStrings) 
		objFile.WriteLine arrStrings(i)
   	Next
	objFile.Close
End Function
