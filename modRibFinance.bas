Attribute VB_Name = "modRibFinance"
Option Explicit

Public g_ufGenFinanceList As ufGenFinanceList
Public g_ufPOSearch As ufPOSearch
Public g_ufPODisplay As ufPODisplay

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnLoadFinanceList_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called by the Ribbon "Import G/L List" button in the Finance Group.
'==============================================================================
Public Sub btnLoadFinanceList_Click(control As Office.IRibbonControl)

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    Dim oErr As typAppError
    Dim iRow As Long
    Dim dtDateFrom As Date
    Dim dtDateTo As Date
    Dim bSuccessful As Boolean

    '===========
    ' There is a strange error that occurs when the user clicks on the X to
    ' close a userform. The form seems to go to never-never-land but the user
    ' form pointer variable is not set to nothing, which means when I try and
    ' avoid reloading the form from scratch by checking for Nothing, it
    ' causes an error on the next invokation.
    ' The work-around is to put in a On Error code section.
    '===========
On Error GoTo handle_bad_userform_pointer

    If Not g_ufGenFinanceList Is Nothing Then
        Call Unload(g_ufGenFinanceList)
    End If

handle_bad_userform_pointer:
On Error GoTo 0 ' Clear error handler

    Set g_ufGenFinanceList = New ufGenFinanceList

    oErr.BeSilent = False

    Call g_ufGenFinanceList.Show

    If g_ufGenFinanceList Is Nothing Then
        GoTo cleanup
    ElseIf (g_ufGenFinanceList.Result <> vbOK) Then
        GoTo cleanup
    End If

    If g_ufGenFinanceList.Result = vbOK Then
        '=============
        ' Run a query and populate the FreeCanvas tab with the requested data.
        '=============

        '=============
        ' Clear the entire contents of the sheet
        '=============
        Call wsFreeCanvas.Cells.Clear

        '=============
        ' Disable event handling and calculations to speed things up
        '=============
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

        Select Case g_ufGenFinanceList.ItemRequired
            Case "PartnerObjects"

                wsFreeCanvas.Cells(1, 1) = "Partner Object"
                wsFreeCanvas.Cells(1, 2) = "Partner Object Type"

                With wsFreeCanvas.Range("A1:B1")
                    .Font.Bold = True
                    .Font.Italic = True
                End With

                '=============
                ' Run the query based on the supplied data
                '=============
                dtDateFrom = CDate(g_ufGenFinanceList.txtDateFrom)
                dtDateTo = CDate(g_ufGenFinanceList.txtDateTo)

                scSQLQuery = "SELECT DISTINCT fk_partner_obj, partner_obj_type from dbo.t_gen_ledger " & _
                                "WHERE fk_cost_center like '" & Replace(g_ufGenFinanceList.txtCCFilter, "*", "%") & "' AND " & _
                                "posting_date >= '" & Format(dtDateFrom, "yyyy-mm-dd") & "' AND " & _
                                "posting_date <= '" & Format(dtDateTo, "yyyy-mm-dd") & "' AND " & _
                                "NOT (fk_partner_obj is NULL) ORDER BY fk_partner_obj"

                Call GetDBRecordSet(ldFinance, cnn, scSQLQuery, rs)

                iRow = 2
                While Not (rs.EOF)
                    wsFreeCanvas.Cells(iRow, 1) = rs.Fields("fk_partner_obj")
                    wsFreeCanvas.Cells(iRow, 2) = rs.Fields("partner_obj_type")

                    iRow = iRow + 1
                    rs.MoveNext
                Wend

                bSuccessful = True

            Case "PurchaseDocs"

                wsFreeCanvas.Cells(1, 1) = "Purchase Doc"
                wsFreeCanvas.Cells(1, 2) = "Purchase Doc Item"
                wsFreeCanvas.Cells(1, 3) = "Cost Center"

                With wsFreeCanvas.Range("A1:C1")
                    .Font.Bold = True
                    .Font.Italic = True
                End With

                '=============
                ' Run the query based on the supplied data
                '=============
                dtDateFrom = CDate(g_ufGenFinanceList.txtDateFrom)
                dtDateTo = CDate(g_ufGenFinanceList.txtDateTo)

                '================
                ' Construct an SQL query that consists of the union of two separate
                ' sub-queries, one on the general ledger table and one on the
                ' order transaction table.
                ' The outer query then grabs a distinct list of the combined.
                '================
                scSQLQuery = "SELECT DISTINCT po.PO, po.POItem, po.CostCenter FROM " & _
                                "((SELECT DISTINCT gl.fk_purch_doc as PO, gl.fk_purch_doc_item as POItem, gl.fk_cost_center as CostCenter " & _
                                    "FROM finance.dbo.t_gen_ledger as gl " & _
                                    "WHERE gl.posting_date >= '" & Format(dtDateFrom, "yyyy-mm-dd") & "' " & _
                                    "AND gl.posting_date <= '" & Format(dtDateTo, "yyyy-mm-dd") & "' " & _
                                    "AND gl.fk_cost_center LIKE '" & Replace(g_ufGenFinanceList.txtCCFilter, "*", "%") & "' " & _
                                    "AND NOT (gl.fk_purch_doc is null)" & _
                                 ") UNION " & _
                                 "(SELECT DISTINCT wot.PO, wot.POItem, wot.CostCenter " & _
                                    "FROM finance.dbo.v_order_trans as wot " & _
                                    "WHERE wot.PostingDate >= '" & Format(dtDateFrom, "yyyy-mm-dd") & "' " & _
                                    "AND wot.PostingDate <= '" & Format(dtDateTo, "yyyy-mm-dd") & "' " & _
                                    "AND wot.CostCenter LIKE '" & Replace(g_ufGenFinanceList.txtCCFilter, "*", "%") & "' " & _
                                    "AND not (wot.PO is null)" & _
                                ")) as po ORDER by po.PO, po.POItem"


                Call GetDBRecordSet(ldFinance, cnn, scSQLQuery, rs)

                iRow = 2
                While Not (rs.EOF)
                    wsFreeCanvas.Cells(iRow, 1) = rs.Fields("PO")
                    wsFreeCanvas.Cells(iRow, 2) = rs.Fields("POItem")
                    wsFreeCanvas.Cells(iRow, 3) = rs.Fields("CostCenter")

                    iRow = iRow + 1
                    rs.MoveNext
                Wend

                bSuccessful = True

            Case "Materials"

                wsFreeCanvas.Cells(1, 1) = "Material"

                With wsFreeCanvas.Range("A1:A1")
                    .Font.Bold = True
                    .Font.Italic = True
                End With

                '=============
                ' Run the query based on the supplied data
                '=============
                dtDateFrom = CDate(g_ufGenFinanceList.txtDateFrom)
                dtDateTo = CDate(g_ufGenFinanceList.txtDateTo)

                scSQLQuery = "SELECT DISTINCT [fk_material] FROM dbo.t_gen_ledger " & _
                                "WHERE fk_cost_center like '" & Replace(g_ufGenFinanceList.txtCCFilter, "*", "%") & "' AND " & _
                                "posting_date >= '" & Format(dtDateFrom, "yyyy-mm-dd") & "' AND " & _
                                "posting_date <= '" & Format(dtDateTo, "yyyy-mm-dd") & "' AND " & _
                                "NOT (fk_material is NULL) ORDER BY fk_material"

                Call GetDBRecordSet(ldFinance, cnn, scSQLQuery, rs)

                iRow = 2
                While Not (rs.EOF)
                    wsFreeCanvas.Cells(iRow, 1) = rs.Fields("fk_material")

                    iRow = iRow + 1
                    rs.MoveNext
                Wend

                bSuccessful = True

        End Select
    End If

    '=============
    ' re-enable event handling and calculations
    '=============
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    If bSuccessful Then
        '============
        ' Make a copy of our FreeCanvas tab for the user
        '============
        Dim ws As Excel.Worksheet
        Dim wb As Excel.Workbook

        wsFreeCanvas.Copy

        Set ws = Application.ActiveSheet
        Set wb = Application.ActiveWorkbook

        ws.Name = g_ufGenFinanceList.ItemRequired
        ws.Cells(2, 1).Select
    End If

cleanup:
    Call Unload(g_ufGenFinanceList)

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnReviewPurchOrder_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called by the Ribbon "Review PO" button in the Finance Group.
'==============================================================================
Public Sub btnReviewPurchOrder_Click(control As Office.IRibbonControl)

    Dim uf As ufPODisplay
    Dim scPO As String

    '============
    ' Get the purchase order from the currently selected cell on the sheet
    '============
    If Selection.Count <> 1 Then
        Call MsgBox("Select a single cell only, that contains a valid Purchase Order number")
        Exit Sub
    End If

    scPO = Selection

On Error GoTo cleanup_nicely
    If g_ufPODisplay Is Nothing Then
        Set g_ufPODisplay = New ufPODisplay
    End If

    If g_ufPODisplay.Visible Then
        '=========
        ' If the user is clicking on the Parts Detail button, even though it
        ' exists, it may be because they can't see the window. Hence we toggle
        ' it's visibility.
        '=========
        Call g_ufPODisplay.Hide
    Else

        Call g_ufPODisplay.DisplayPO(scPO)

        Call g_ufPODisplay.Show(vbModeless)
    End If


    Exit Sub
cleanup_nicely:

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnPOSearch_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called by the Ribbon "PO Search" button in the Finance Group.
'==============================================================================
Public Sub btnPOSearch_Click(control As Office.IRibbonControl)

    '===========
    ' There is a strange error that occurs when the user clicks on the X to
    ' close a userform. The form seems to go to never-never-land but the user
    ' form pointer variable is not set to nothing, which means when I try and
    ' avoid reloading the form from scratch by checking for Nothing, it
    ' causes an error on the next invokation.
    ' The work-around is to put in a On Error code section.
    '===========
On Error GoTo handle_bad_userform_pointer

    If Not g_ufPOSearch Is Nothing Then
        Call Unload(g_ufPOSearch)
    End If

handle_bad_userform_pointer:
On Error GoTo 0 ' Clear error handler

    Set g_ufPOSearch = New ufPOSearch

    Call g_ufPOSearch.Show(vbModeless)
End Sub


