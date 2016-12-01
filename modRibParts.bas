Attribute VB_Name = "modRibParts"
Option Explicit

Public g_ufPartDisplay As ufPartDisplay

'
'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnPartsSearch_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Callback for btnPartsSearch onAction
'==============================================================================
Public Sub btnPartsSearch_Click(control As Office.IRibbonControl)
    Dim uf As ufPartsSearch
    
    Set uf = New ufPartsSearch
    
    uf.m_bInEventHandler = False
    Call uf.ResetForm
    
    Call uf.Show(vbModeless)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnDisplayPartDetail_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Callback for btnDisplayPartDetail onAction
'==============================================================================
Public Sub btnDisplayPartDetail_Click(control As Office.IRibbonControl)

    Dim uf As ufPartDisplay
    Dim iMaterial As Long
    
    '============
    ' Get the material number from the currently selected cell on the sheet
    '============
    If Selection.Count <> 1 Then
        Call MsgBox("Select a single cell only, that contains a material number")
        Exit Sub
    End If
    
On Error GoTo not_a_valid_material
    iMaterial = Val(Selection)
    
    GoTo display_material
    
not_a_valid_material:
    iMaterial = 0
 
display_material:
    
On Error GoTo cleanup_nicely
    Set g_ufPartDisplay = New ufPartDisplay
    
    If g_ufPartDisplay.Visible Then
        '=========
        ' If the user is clicking on the Parts Detail button, even though it
        ' exists, it may be because they can't see the window. Hence we toggle
        ' it's visibility.
        '=========
        Call g_ufPartDisplay.Hide
    Else
    
        Call g_ufPartDisplay.DisplayMaterial(iMaterial)

        Call g_ufPartDisplay.Show(vbModeless)
    End If
    
    
    Exit Sub
cleanup_nicely:

End Sub

'==============================================================================
' FUNCTION
'   CreateSimplePartNo
'------------------------------------------------------------------------------
' DESCRIPTION
'   Creates a simplified part number string by removing all characters that
' are not alphanumeric, and making all characters uppercase.
'==============================================================================
Public Function CreateSimplePartNo(scPartNo As String, Optional bLeaveInWildcard = False)

    Dim strTemp As String
    Dim i As Long
    
    If bLeaveInWildcard Then
        For i = 1 To Len(scPartNo)
            Select Case Asc(Mid(UCase(scPartNo), i, 1))
                Case 48 To 57, 65 To 90, 37
                    strTemp = strTemp & Mid(UCase(scPartNo), i, 1)
            End Select
        Next
    Else
        For i = 1 To Len(scPartNo)
            Select Case Asc(Mid(UCase(scPartNo), i, 1))
                Case 48 To 57, 65 To 90
                    strTemp = strTemp & Mid(UCase(scPartNo), i, 1)
            End Select
        Next
    End If
    CreateSimplePartNo = strTemp

End Function

