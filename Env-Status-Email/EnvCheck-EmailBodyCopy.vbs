' ###############################################################
' This file provides the body copy for the Env Check status email   
' ###############################################################

' Make sure we declare all variables before use
Option Explicit

'---------------------------------------------------------------------------------------------------
'Declare any variables before use
Dim sEmailBodyCopy, sEmailStyleSheet, sLeaderComment

'-------------------------------------------------------------------------------
' Open the email for review, prior to sending
sLeaderComment = InputBox("Would you like to add a comment?", "Comment")

'-----------------------------------------------------------------------------------------------------
' This section creates the status table in the Environment Check email
' The set of variables inluded are located in EnvCheck-StatusQuestions.vbs

sEmailBodyCopy = "<html>" &_
"<head>" &_
"<style>" &_
"table {border-collapse:collapse; width:500px}" &_
"table,th,td{font-family:Calibri; font-size:11pt; border: 1px solid gray}" &_
"p{font-family:Calibri; font-size:11pt;}" &_
".header{font-size:14pt; font-weight:bold; background-color:#C8C8C8; text-align:center}" &_
"</style>" &_
"</head>" &_
"<body>" &_
"<p>All,</p>" &_
"<p>" & sLeaderComment & "</p>" &_
"<table>" &_
"<tr class='header'><td colspan='4'>RIT</td></tr>"&_
"<tr class='header'><td>System</td><td>Environment Name</td><td>Trade Message Flow</td><td>Status</td></tr>" &_
"<tr><td>Murex</td><td>MXDEV6</td><td>Lloyds FX</td><td>" & sLloydsFXMailStatus & "</td></tr>" &_
"<tr><td>MLC</td><td>MLC_UAT_07</td><td>FX / MM</td><td><b>TBC</b></td></tr>" &_
"<tr><td>Summit 5.5</td><td>OSUM5D4</td><td>Lloyds MM</td><td>" & sLloydsMMMailStatus & "</td></tr>" &_
"<tr><td>Summit 5.5</td><td>OSUM5D4</td><td>BOS FX</td><td>" & sHBOSFXMailStatus & "</td></tr>" &_
"<tr><td>Summit 5.5</td><td>OSUM5D4</td><td>BOS MM</td><td>" & sHBOSMMMailStatus & "</td></tr>" &_
"</table>" &_
"</body>" &_
"</html>"






