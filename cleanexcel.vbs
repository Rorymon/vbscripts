'=================================================================================
'
' DESCRIPTION: Excel VBA to clean a spreadsheet. Removes all text BEFORE a
' certain character e.g. 1: Rory Monaghan will become Rory Monaghan
'
' NAME: cleanexcel.vbs
'
' AUTHOR: KuTools
' DATE  : August 23rd, 2016
'
' USAGE: Run as Excel VBA using Alt+F11 and navigating to Import > Module
' When prompted enter a range to apply the update and finally a character
'
' PREREQ: Excel 2013
'
' COMMENTS: All Credit to KuTools! This is just a dump of vbscripts I wish to keep
'=================================================================================

Sub RemoveAllButLastWord()

Dim Rng As Range
Dim WorkRng As Range
Dim xChar As String
On Error Resume Next
xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
xChar = Application.InputBox("String", xTitleId, "", Type:=2)
For Each Rng In WorkRng
    xValue = Rng.Value
    Rng.Value = VBA.Right(xValue, VBA.Len(xValue) - VBA.InStrRev(xValue, xChar))
Next
End Sub
