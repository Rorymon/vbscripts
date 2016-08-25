'==========================================================================
'
' DESCRIPTION: Removes folder.
'
' NAME: RemoveFolder.vbs
'
' AUTHOR: Russell Hill , HP
' DATE  : 08/07/2016
'
' USAGE: Removes folder, along with subfolders and files.
'
' PREREQ: Ensure that the folder to removed does not contain files and 
'         folders from another application from the same suite.
'
' COMMENTS: None
'
'==========================================================================

Option Explicit

Dim strRemoveFolder

	'Folder to be removed less C:\
	
strRemoveFolder = ""

RemoveFolder strRemoveFolder

Public Function RemoveFolder(strRemoveFolderName)
	Dim objFS, objShell, strSysDir
	Set objFS = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject( "WScript.Shell" )
	strSysDir = objShell.ExpandEnvironmentStrings("%SystemDrive%")
	If (Right(strSysDir, 1) <> "\") Then
		strSysDir = strSysDir & "\"
	End If
	strRemoveFolderName = strSysDir & strRemoveFolderName
	If (objFS.FolderExists(strRemoveFolderName)) Then
		objFS.DeleteFolder strRemoveFolderName, True
	End If
   	Set objShell = Nothing
   	Set objFS = Nothing
End Function
