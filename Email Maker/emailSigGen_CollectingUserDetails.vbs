' http://cyreath.blogspot.com/2014/02/vbscript-email-signature-generator.html
' http://www.youtube.com/subscription_center?add_user=cyreath

'--------------------------------------------------------------------------------------------------
' This will prevent us from using a variable before declaring it
Option Explicit

'Declare our variables
Dim sUsername, sOperationalTitle, sUserPhone, sUserMobile, sUserEmail
Dim sTitle

'Set the title on the message boxes
sTitle = "Email Signature Generator"

'---------------------------------------------------------------------------------------------------
'The function that gets the data from the user
Function GatherDetails()
	sUsername = InputBox("Please enter your name, as you wish it to appear in your emails.", sTitle & " (1 of 5)")

	sOperationalTitle = InputBox("What is your Operational Role or Title?", sTitle & " (2 of 5)")
	 
	sUserPhone = InputBox("Enter your desk phone number, use the international format if required, e.g +44(0)1234 5678.", sTitle & " (3 of 5)")

	sUserMobile = InputBox("Enter your Mobile phone number, use the international format if required, e.g +44(0)1234 5678.", sTitle & " (4 of 5)")
	 
	sUserEmail = InputBox("Provide your company email address.", sTitle & " (5 of 5)", LCase(Replace(sUserName, " ", "") & "@yourCompanyName.com"))
End Function