'------------------ TEST SERVER STATUS CHECK -------------------------------
'----- This script will PING each server to see if it is up ----------------
'----You need to have a serverlist.txt file with the IP or www address in --
'----- it then writes the results to a .csv file for review ----------------
'---------------------------------------------------------------------------
dim strInputPath, strOutputPath, strStatusdim objFSO, objTextIn, objTextOut 

' ---- Location of the server list input file
strInputPath = "\\some\place\on\your\system\serverlist.txt"

' ---- The output file with server ping results
strOutputPath = "output.csv"

set objFSO = CreateObject("Scripting.FileSystemObject")
set objTextIn = objFSO.OpenTextFile( strInputPath,1 )
set objTextOut = objFSO.CreateTextFile( strOutputPath )
objTextOut.WriteLine("Target,Status")
      
      ' ---- Loop over the input file, testing each line (server name) until all are checked 
Do until objTextIn.AtEndOfStream = True    
        strComputer = objTextIn.ReadLine        
        if fPingTest( strComputer ) then             
          strStatus = "UP"        
        else             
          strStatus = "DOWN"        
        end if        
        objTextOut.WriteLine(strComputer & "," & strStatus)
loop
        
' ---- Ping the server by calling the ping service and adding the server name from the list 
  function fPingTest( strComputer )        
    dim objShell,objPing        
    dim strPingOut, flag        
    set objShell = CreateObject("Wscript.Shell")        
    set objPing = objShell.Exec("ping " & strComputer)    
        strPingOut = objPing.StdOut.ReadAll    
        if instr(LCase(strPingOut), "reply") then        
          flag = TRUE        
        else                
          flag = FALSE        
        end if        
        fPingTest = flag 
    end function
