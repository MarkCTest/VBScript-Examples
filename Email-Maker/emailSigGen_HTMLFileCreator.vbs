' http://cyreath.blogspot.com/2014/02/vbscript-email-signature-generator.html
' http://www.youtube.com/subscription_center?add_user=cyreath
'--------------------------------------------------------------------------------------------------

' This will prevent us from using a variable before declaring it
Option Explicit

' Declare all the variables, grouped by type for ease of organisation
Dim oFSO1, oOutputFile		' objects
Dim sTxtOutput			' strings

Function CreateHTMLFile
	' Ready the file we'll ouput our text to
	Set oFSO1 = CreateObject("Scripting.FileSystemObject")
	oOutputFile = "./" & Replace(sUserName, " ", "") & "_SigFile.html"
	Set sTxtOutput = oFSO1.OpenTextFile(oOutputFile, 2, True)

	' Create the HTML file with the details the user has entered

	sTxtOutput.writeline "<html>"

	sTxtOutput.writeline "<head>"
	sTxtOutput.writeline "<style>"
	sTxtOutput.writeline ".name{font-family:Calibri; font-size:12pt; font-weight:bold;}"
	sTxtOutput.writeline ".bodycopy{font-family:Calibri; font-size:10pt}"
	sTxtOutput.writeline "</style>"
	sTxtOutput.writeline "</head>"
			
	sTxtOutput.writeline "<body>"
	
	sTxtOutput.writeline "<div class='name'> " & sUsername & "</div>"
	sTxtOutput.writeline "<div class='bodycopy'> " & sOperationalTitle & "</div></br>"

	sTxtOutput.writeline "<div class='bodycopy'><b>Desk:</b> " & sUserPhone & "</div>"
	sTxtOutput.writeline "<div class='bodycopy'><b>Mobile:</b> " & sUserMobile & "</div>"
	sTxtOutput.writeline "<div class='bodycopy'><b>Email:</b> " & sUserEmail & "</div>"
	
	sTxtOutput.writeline "<p><img src='logo.jpg'></p>"
	
	sTxtOutput.writeline "</body>"

	sTxtOutput.writeline "</html>"

	'confirm file creation
	MsgBox "Thankyou, your HTML signature file has been created." & VbCrLf & "Please look in the current folder for a file named: " & VbCrLf & VbCrLf & Replace(sUserName, " ", "") & "_SigFile.html", vbOKOnly, "HTML Signature File Created"
End Function
