Attribute VB_Name = "modRibGeneral"
Option Explicit

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   GetVisible
'------------------------------------------------------------------------------
' DESCRIPTION
'   The event handler for all the custom ribbon controls. The event handler
' determined which control called the method, by the Tag property.
' Current tags defined are:
' - TagCustomTabLihir: The whole tab
' - TagGroupDBConfig: The DB Interface Group
' - TagDDDBServer: The DB Server Drop-down list
' - TagBtnTestDBConn: The Test DB Button
' - TagGroupLihirStaff: The Lihir Staff Group
' - TagBtnBrowseStaff: The Staff Browser Button
' - TagGroupLihirProjects: The Projects Group
' - TagBtnProjectBrowser: The Projects Browser Button
' - TagGroupMaint: The Maintenance Group
' - TagBtnUploadIW29: The IW29 Upload Button
' - TagBtnUploadIW39: The IW39 Upload Button
' - TagBtnUploadIW49: The IW49 Upload Button
'==============================================================================
Sub GetVisible(control As Office.IRibbonControl, ByRef returnedVal)
    Select Case control.Tag
    
        '===========
        ' No one - Not yet implemented
        '===========
        Case "TagGroupLihirProjects", _
            "TagBtnProjectBrowser", _
            "TagBtnReviewNotification", _
            "TagOpenRiskRegister", _
            "TagGroupRisk"
            returnedVal = False
            
        '===========
        ' Developers only
        '===========
        Case "TagGroupDBConfig", _
            "TagDDDBServer", _
            "TagBtnTestDBConn"
            Select Case Application.UserName
                Case "Adam Pemberton", "Adam", "Franz Hemetsberger"
                    returnedVal = True
                Case Else
                    returnedVal = False
            End Select
        
        '===========
        ' Developers and Admin Clerks
        '===========
        Case "TagMenuMaintUpload", _
            "TagMenuUploadFinance", _
            "TagLoadFinanceList", _
            "TagMenuUploadParts"
            Select Case Application.UserName
                Case "Adam Pemberton", "Adam", "Franz Hemetsberger", "Helen Bunbun", _
                    "Francisca Kamda", "Kesma Tapaua", "Desley Daimol", "Margaret Mota", _
                    "Alex Murdoch", "Tom Verevis", "Alistair Brian", "Donna Becher Lupalrea"
                    returnedVal = True
                Case Else
                    returnedVal = False
            End Select

        '===========
        ' Everyone
        '===========
        Case "TagGroupLihirStaff", _
            "TagBtnBrowseStaff", _
            "TagBtnBrowseOrgChart", _
            "TagGroupMaint", _
            "TagReviewWorkOrder", _
            "TagBtnFLOCBrowser", _
            "TagGroupFinance", _
            "TagGroupParts", _
            "TagPartsSearch"
            returnedVal = True

        '===========
        ' Everything Else (including TagCustomTabLihir) -> always
        '===========
        Case Else
            returnedVal = True

    End Select
End Sub


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnAboutLGL_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called by the Ribbon button
'==============================================================================
Public Sub btnAboutLGL_Click(control As Office.IRibbonControl)
    Dim scVer As String
    
    scVer = wsVer.Range("Version")
    
    Call MsgBox("Lihir Client Add-in Version " & scVer & ". Contact Adam Pemberton for more information.")
End Sub

