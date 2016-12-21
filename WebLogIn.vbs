' ------- Set Up the VBScript environment
Option Explicit
Dim ie, ipf

Set ie = CreateObject("InternetExplorer.Application")

On Error Resume Next

Sub WaitForLoad
  Do While ie.Busy
   WScript.Sleep 500
   Loop
End Sub

Sub Find(x)
  set ipf = ie.Document.All.Item(x)
End Sub

' ------ Navigate to the intadmin environment
ie.Navigate "https://theWebsiteNameHere"

Call WaitForLoad
ie.Visible = True

' ------ Find and fill out the form fields
Call Find("username")
ipf.Value = "YOUR USER NAME HERE"

Call Find("Password")
ipf.Value = "YOUR PASSWORD HERE"

' ------ Then click submit to log-in
ipf.value = ie.document.forms(0).submit()
ipf.click

Call WaitForLoad

'You're now logged in
