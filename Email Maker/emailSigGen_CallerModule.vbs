' http://cyreath.blogspot.com/2014/02/vbscript-email-signature-generator.html
' http://www.youtube.com/subscription_center?add_user=cyreath

'-----------------------------------------------------------------------------------------
' This will prevent us from using a variable before declaring it
Option Explicit

' This is the Sub that opens external files and reads in the contents.
' In this way, you can have separate files for data and libraries of functions
Sub Include(yourFile)
  Dim oFSO, oFileBeingReadIn	' define Objects
  Dim sFileContents		' define Strings
  
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFileBeingReadIn = oFSO.OpenTextFile(yourFile & ".vbs", 1)
  sFileContents = oFileBeingReadIn.ReadAll
  oFileBeingReadIn.Close
  ExecuteGlobal sFileContents
End Sub

' Here we call the Include Sub, then pass it the name of the files we want items from
Include "emailSigGen_CollectingUserDetails"
Include "emailSigGen_HTMLFileCreator"

' Now we can call items from the files we read in, whenever we want to use them
GatherDetails
CreateHTMLFile
