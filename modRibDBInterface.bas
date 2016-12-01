Attribute VB_Name = "modRibDBInterface"
Option Explicit

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   ddDBServer_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called by the Ribbon DropDown box that selects server to use.
'==============================================================================
Sub ddDBServer_Change(control As Office.IRibbonControl, id As String, index As Integer)

    Call SetDefaultDBServer(ldPowerAndUtilities, id)
    Call SetDefaultDBServer(ldMaintenance, id)
    Call SetDefaultDBServer(ldDocControl, id)
    Call SetDefaultDBServer(ldFinance, id)
    Call SetDefaultDBServer(ldPeople, id)
    Call SetDefaultDBServer(ldParts, id)
    
    '===========
    ' Close the connection in case it is open, so the new server is accessed on
    ' the next attempt.
    '===========
    Call CloseConnection(ldPowerAndUtilities)
    Call CloseConnection(ldMaintenance)
    Call CloseConnection(ldDocControl)
    Call CloseConnection(ldFinance)
    Call CloseConnection(ldPeople)
    Call CloseConnection(ldParts)
    
    '===========
    ' Save the entire workbook.
    '===========
    Call LihirAddIn.ThisWorkbook.Save
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   getDBServerItemID
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called by the Ribbon DropDown box to initialise itself
'==============================================================================
Sub getDBServerItemID(control As Office.IRibbonControl, ByRef itemID)
    Dim scServer As String
    
    scServer = wsParameters.Range("DBPUServerDefaultID")
    itemID = scServer

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnTestDBConn_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Callback for btnTestDBConn onAction
'==============================================================================
Sub btnTestDBConn_Click(control As Office.IRibbonControl)

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
#Else
    Dim cnn As Object
#End If


    If ConnectToDB(ldPowerAndUtilities, cnn) Then
        Call MsgBox("Connection established")
        Call CloseConnection(ldPowerAndUtilities)
    Else
        Call MsgBox("Failed to connect")
    End If
    
End Sub

