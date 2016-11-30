Option Explicit
'-----------------------------------------------------------------------------
'Code for calling in other files here, first thing our script does

Sub Include(otherFile)   
  Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")   
  Dim oFile : Set oFile = oFSO.OpenTextFile(otherFile, 1)   
  Dim sStr : sStr = oFile.ReadAll   
  oFile.Close   
  ExecuteGlobal 
  sStr
End Sub

'-----------------------------------------------------------------------------
'Call in all the files for the test, place this at the very end of a batch/controller vbs file (e.g. test-fil.vbs)

Include "test-file.vbs"
