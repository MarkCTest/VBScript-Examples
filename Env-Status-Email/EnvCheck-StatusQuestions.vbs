' ###################################################
' This file asks the status of each Loop check trade, 
' which will be added to the status email   
' ###################################################

' Make sure we declare all variables before use
Option Explicit

'---------------------------------------------------------------------------------------------------
' Declare the set of variables we'll be using
Dim sLloydsFXStatus, sLloydsFXMailStatus
Dim sLloydsMMStatus, sLloydsMMMailStatus

Dim sHBOSFXStatus, sHBOSFXMailStatus
Dim sHBOSMMStatus, sHBOSMMMailStatus

' ------------------------------- Development -----------------------------------------------------
sLloydsFXStatus = MsgBox("Did the    Lloyds FX    trade complete the Loop?", vbYesNoCancel+vbQuestion+vbDefaultButton1, "Lloyds FX trade status")

	If sLloydsFXStatus = 6 Then 'Yes was clicked
		sLloydsFXMailStatus = "<font color='Green'>GREEN</font>"
	ElseIf sLloydsFXStatus = 7 Then ' No was clicked
		sLloydsFXMailStatus = "<font color='Red'>RED</font>"
	Else
		sLloydsFXMailStatus = "<font color='Orange'>AMBER</font>"
	End If

' ------------------------------- Integration -----------------------------------------------------
sLloydsMMStatus = MsgBox("Did the    Lloyds MM    trade complete the Loop?", vbYesNoCancel+vbQuestion+vbDefaultButton1, "Lloyds MM trade status")

	If sLloydsMMStatus = 6 Then 'Yes was clicked
		sLloydsMMMailStatus = "<font color='Green'>GREEN</font>"
	ElseIf sLloydsMMStatus = 7 Then ' No was clicked
		sLloydsMMMailStatus = "<font color='Red'>RED</font>"
	Else
		sLloydsMMMailStatus = "<font color='Orange'>AMBER</font>"
	End If
	
' ------------------------------- System Test -----------------------------------------------------
sHBOSFXStatus = MsgBox("Did the    HBOS FX    trade complete the Loop?", vbYesNoCancel+vbQuestion+vbDefaultButton1, "HBOS FX trade status")

	If sHBOSFXStatus = 6 Then 'Yes was clicked
		sHBOSFXMailStatus = "<font color='Green'>GREEN</font>"
	ElseIf sHBOSFXStatus = 7 Then ' No was clicked
		sHBOSFXMailStatus = "<font color='Red'>RED</font>"
	Else
		sHBOSFXMailStatus = "<font color='Orange'>AMBER</font>"
	End If

' ------------------------------- UAT -----------------------------------------------------
sHBOSMMStatus = MsgBox("Did the    HBOS MM    trade complete the Loop?", vbYesNoCancel+vbQuestion+vbDefaultButton1, "HBOS MM trade status")

	If sHBOSMMStatus = 6 Then 'Yes was clicked
		sHBOSMMMailStatus = "<font color='Green'>GREEN</font>"
	ElseIf sHBOSMMStatus = 7 Then ' No was clicked
		sHBOSMMMailStatus = "<font color='Red'>RED</font>"
	Else
		sHBOSMMMailStatus = "<font color='Orange'>AMBER</font>"
	End If

'--------------------------------------------------------------
' WRITE THE RESULTS TO FILE
