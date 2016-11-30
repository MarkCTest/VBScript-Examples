' ###################################################
' This file provides the email list to which the 
' Environment Check status email will be sent   
' ###################################################

' Make sure we declare all variables before use
Option Explicit

'---------------------------------------------------------------------------------------------------
'Declare variables for the Environment Check Status email group
Dim sToMailSet, sCCMailSet

'Declare Variables for the Troubleshooting Escalation email recipients
Dim sLloydsMMSummit, sLloydsFXMurex, sHBOSMMMLSASummit, sHBOSFXMLSASummit
Dim sLloydsMMandFX, sHBOSMMandFX, sLloydsMMFXandHBOSMMFX
Dim sLloydsLoopFailAll

' Declare the variables for the MONDAY morning check email recipients
Dim sMondaySummitDateRoll, sMondayMurexDateRoll, sMondayGDSNotUp

'------------------------------------------------------------------------------------------------------
' List of email addresses the Environment Check Status email will be sent TO and CC
sToMailSet = "Some Development Email Group; Some Test Team Email Group"
sCCMailSet = "Crowther, Mark"


'------------------------------------------------------------------------------------------------------
' List of email addresses for teams that will recieve the escalation emails to

sLloydsMMSummit = "ADM CB - Trading IT EBusiness Support"
sLloydsFXMurex = "ADM CB - Trading IT EBusiness Support; ADM CB - Trading Operations Murex Support"
sHBOSMMMLSASummit = "ADM CB - Trading IT EBusiness Support; ADM CB - Trading Risk - Change the Bank"
sHBOSFXMLSASummit = "ADM CB - Trading IT EBusiness Support; ADM CB - Trading Risk - Change the Bank"
sLloydsMMandFX = "ADM CB - Trading IT EBusiness Support"
sHBOSMMandFX = "ADM CB - Trading IT EBusiness Support; ADM CB - Trading Risk - Change the Bank"
sLloydsMMFXandHBOSMMFX = "ADM CB - Trading IT EBusiness Support"

sLloydsLoopFailAll ="Crowther, Mark"

'------------------------------------------------------------------------------------------------------
' List of email addresses for MONDAY MORNING failures

sMondaySummitDateRoll = ""
sMondayMurexDateRoll = ""
sMondayGDSNotUp = ""