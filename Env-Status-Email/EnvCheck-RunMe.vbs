' #####################################################################
' This file pulls in all files used to make the Environment Check email  
' See INCLUDE list below for the full set of files
' #####################################################################

' Make sure we declare all variables before use
Option Explicit

'-------------------------------------------------------------------------------
' Declare all the variables prior to using
Dim oMyApp, oMyItem, oFSO, oFile
Dim sStr, sEmailTitle

'-------------------------------------------------------------------------------
' Code to pull in additonal files containing emails list, email body copy and ask status questions
Sub Include(fileGroup)
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFile = oFSO.OpenTextFile(fileGroup & ".vbs", 1)
  
  sStr = oFile.ReadAll
  oFile.Close
  ExecuteGlobal sStr
End Sub

'-------------------------------------------------------------------------------
' List of files required for the program to run
Include "EnvCheck-EmailsList"			'(Contains the email addresses to use) 
Include "EnvCheck-StatusQuestions"		'(Asks the status of each test used to validate an environment is working)
Include "EnvCheck-EmailBodyCopy"		'(Provides the content of the email body to build the table in the email)

'-------------------------------------------------------------------------------
' Set-up Outlook as the application we want to use
Set oMyApp = CreateObject("Outlook.Application")
Set oMyItem = oMyApp.CreateItem(0)

'-------------------------------------------------------------------------------
' Set the subject title for the email
sEmailTitle = "Test Environment Status Report (" & Date & ")"

'-------------------------------------------------------------------------------
'Build the content of the email, matching fileds in outlook
With oMyItem
	.To =  sToMailSet
	.Cc =  sCCMailSet
	.Subject = sEmailTitle
	
     'Add the report table
	.HTMLBody = sEmailBodyCopy  '(String located in EnvCheck-EmailBodyCopy.vbs)
	
	.ReadReceiptRequested = False
End With

'-------------------------------------------------------------------------------
' Open the email for review, prior to sending
oMyItem.Display
