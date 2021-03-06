' String Comparison Script
' Compare two strings to see if the exactly match
' The script gives you the option to make all text lower case
'------------------------------------------------------------

'----------------- This script accepts 2 strings, then asks if you want to switch them both to lowercase 
'----------------- to get around the case sensitive check of StrComp

Option Explicit    

Dim sTitle, sPrompt, sStringFormat, sFirstString, sSecondString, sResult  

sTitle = "String Comparison"
sPrompt = "Enter the string you want to compare"  

'----------------- Get the strings to compare
sFirstString = InputBox(sPrompt, sTitle) 
sSecondString = InputBox(sPrompt, sTitle)  

'----------------- Check if the string's case should match or stay as they are
sStringFormat = MsgBox("Change case to match?", vbYesNoCancel+vbQuestion+vbDefaultButton2+vbSystemModal, sTitle)  

'----------------- Check the input and either leave the strings case as is or convert them both to lowercase
If sStringFormat = vbNo Then  
  sResult = StrComp(sFirstString,sSecondString)                  
      ElseIf  sStringFormat = vbYes  Then                                  
      sFirstString = LCase(sFirstString)                                  
      sSecondString = LCase(sSecondString)                                                  
        sResult = StrComp(sFirstString,sSecondString)                    
Else                                   
  MsgBox ("Looks like you hit Cancel") 
End If  

'----------------- Stating if the strings match
If sResult = 0 Then   
  '--- strings match then we're OK                  
  MsgBox ("The two strings DO match")                                  
Else ' --- Anything else is a fail                                                  
  MsgBox ("The strings DO NOT match")   
End If
