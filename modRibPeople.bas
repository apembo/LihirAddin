Attribute VB_Name = "modRibPeople"
Option Explicit

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnBrowseOrgChart2_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Callback for btnBrowseOrgChart onAction
'==============================================================================
Sub btnBrowseOrgChart_Click(control As Office.IRibbonControl)

    Dim ufOrgChart As ufOrgChartBrowser
    
    Set ufOrgChart = New ufOrgChartBrowser
    
    Call ufOrgChart.Show(vbModeless)
    
End Sub
