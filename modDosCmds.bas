Attribute VB_Name = "modDosCmds"
Option Explicit

'==============================================================================
' FUNCTION
'   HostIsConnectible
'------------------------------------------------------------------------------
' DESCRIPTION
'   Uses a ping to determine whether a particular host address is available.
'------------------------------------------------------------------------------
' Parameters:
'   sHost:  The host address
'   iPings: The number of pings to test. Default is 2. Note often it takes 2
'       for slow connections.
'   iTO:    Time-out. Default is 1 second
'   iIPvX:  Force IPv4/IPv6 communications. 4 = IPv4, 6 = IPv6
'==============================================================================
Function HostIsConnectible(sHost As String, _
        Optional iPings As Long = 2, _
        Optional iTO As Long = 250, _
        Optional iIPvX As Long = 0)
    
    Dim nRes As Long
    Dim scCommand As String
    
    scCommand = "%comspec% /c ping.exe"
    Select Case iIPvX
        Case 4, 6:
            scCommand = scCommand & " -" & iIPvX
    End Select
    
    scCommand = scCommand & " -n " & iPings & " -w " & iTO
    scCommand = scCommand & " " & sHost & " | findstr ""time= time<"" > nul 2>&1"
    '============
    ' This command has the following components:
    ' - Runs ping with the supplied arguments
    ' - Pipes the result to the findstr DOS command which uses regular
    '   expressions.
    ' - The >nul syntax redirects the results to NUL
    ' - the command 2> redirects error output.
    ' NOTE: for some logic I don't understand, the combined redirection command
    '           >nul 2>&1
    '       redirects both normal and error outputs to NUL
    '============
    
    With CreateObject("WScript.Shell")
        nRes = .Run(scCommand, 0, True)
    End With
    HostIsConnectible = (nRes = 0)

End Function

