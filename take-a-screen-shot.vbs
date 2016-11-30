' ------- This script opens a web browser and takes a screen shot
' ------- You need to have Microsoft Word installed for this to work
' ------- BE AWARE - VBScript can be unreliable at making this work!

Option Explicit

' ------- Declare the variables  ----------------- 
Dim oIE, WshShell   

' ------- Wait until the webpage is loaded  -------------- 
Sub WaitForLoad   
  Do While oIE.Busy           
    WScript.Sleep 500   
  Loop 
End Sub 

' ------- Blocks of code for the test steps  ------------- 
Sub OpenPaint   
  Set WshShell = WScript.CreateObject("WScript.Shell")   
  WshShell.Run "mspaint"  
  WScript.Sleep 5000 
End Sub   

Sub OpenIEAndGoToGoogle   
  Set oIE = CreateObject("InternetExplorer.Application")   
  oIE.Visible = True   
  oIE.Navigate "https://www.google.co.uk"   
  Call WaitForLoad 
End Sub    

Sub ActivateIE   
  WShShell.AppActivate "Google"  
  WScript.Sleep 1000 
End Sub

Sub TakeScreenShot   
  Set Wshshell = CreateObject("Word.Basic")   
  WshShell.SendKeys "(%{1068})" 'Screenshots the currently active window, not the whole screen   
  WScript.Sleep 1000 
End Sub   

Sub ActivatePaintAndSaveTheImage   
  WshShell.AppActivate "Untitled - Paint"   
  WScript.Sleep 1500    
  WshShell.sendkeys "^(v)"   
  WScript.Sleep 1500    
  WshShell.sendkeys "^(s)"   
  WScript.Sleep 1500      
  WshShell.sendkeys "testing.jpg"  
  WScript.Sleep 1500      
  WshShell.sendkeys "%(s)"   
  WScript.Sleep 1500 
End Sub   

Sub ClosePaintAndIE  
  WshShell.AppClose "Paint"   
  WScript.Sleep 1500  
  WshShell.AppClose "Google"  
  WScript.Sleep 1500 
End Sub   

' ------- Call the Blocks of code  ---------------- 
Call OpenPaint 
Call OpenIEAndGoToGoogle 
Call ActivateIE
Call TakeScreenShot 
Call ActivatePaintAndSaveTheImage 
Call ClosePaintAndIE 
