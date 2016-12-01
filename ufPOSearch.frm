VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPOSearch 
   Caption         =   "Search for Purchase Orders"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11655
   OleObjectBlob   =   "ufPOSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPOSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' Note on declarations in Object Modules (class module, userform code module or
' Workbook Code Module):
'---------------
'   Some code elements must be scoped as Private in an object module. These are
' constants, enums, and Declare statements. If you attempt to scope one of
' these items as Public, you will get a rather cryptic compiler error.
'==============================================================================

'==============================================================================
' PRIVATE MEMBER VARIABLES
'==============================================================================
Private Const c_iMaxPOsSearchCount As Long = 500

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public m_bInEventHandler As Boolean
Public m_ufPODisplay As ufPODisplay

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Call on userform creation
'==============================================================================
Private Sub UserForm_Initialize()
    Set m_ufPODisplay = New ufPODisplay
    Set m_ufPODisplay.m_oSearchFormParent = Me
    m_bInEventHandler = False
    
    txtMaterial.Enabled = False
    txtMaterial.BackColor = RGB(223, 223, 223)
    
#If DebugMode = 1 Then
    Height = 509
#Else
    Height = 345
#End If
End Sub

'==============================================================================
' SUBROUTINE
'   ResetForm
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clears the form
'==============================================================================
Public Sub ResetForm()
    txtDescriptionFilter.Text = "*"
    txtTrackingFilter.Text = "*"
    txtVendorFilter.Text = "*"
    Call btnClearVendorSearch_Click
    
    txtValueMinFilter.Text = "*"
    txtValueMaxFilter.Text = "*"
    
    cbIncludeClosedPOs.Value = False
    
    Call lbPurchOrders.Clear
    
    btnToggleDisplayPO.Caption = "Hide PO Detail"
    btnToggleDisplayPO.Enabled = False

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnSearch_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   OnChange event handler for the Search Button
'==============================================================================
Private Sub btnSearch_Click()
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim scEntry As String
    Dim iPOCount As Long
    Dim i As Long
    Dim iCount As Long
    Dim iColloquialCount As Long
    Dim scWhereClause As String
    Dim scTemp As String
    Dim scPO As String
    Dim scLastPO As String
    
    '==========
    ' Construct the WHERE clause of the query. It is basically a set of
    ' INTERSECT's.
    '==========
    
    '===========
    ' First ... descriptions.
    '===========
    If Me.txtDescriptionFilter.Text <> "" And Me.txtDescriptionFilter.Text <> "*" Then
        
        scTemp = Replace(Me.txtDescriptionFilter.Text, "'", "''")
        
        If InStr(scTemp, "*") > 0 Then
        
            scTemp = Replace(scTemp, "*", "%")
            scWhereClause = scWhereClause & _
                "(SELECT pk_purch_doc " & _
                            "FROM finance.dbo.t_purch_doc_item " & _
                            "WHERE short_text LIKE '" & scTemp & "')"
        Else
            scWhereClause = scWhereClause & _
                "(SELECT pk_purch_doc " & _
                            "FROM finance.dbo.t_purch_doc_item " & _
                            "WHERE short_text = '" & scTemp & "')"
        End If

    End If
    
    '===========
    ' next ... tracking field.
    '===========
    If Me.txtTrackingFilter.Text <> "" And Me.txtTrackingFilter.Text <> "*" Then
        
        scTemp = Replace(Me.txtTrackingFilter.Text, "'", "''")
        
        If (Len(scWhereClause) > 0) Then
            scWhereClause = scWhereClause & " INTERSECT "
        End If
        
        If InStr(scTemp, "*") > 0 Then
        
            scTemp = Replace(scTemp, "*", "%")
            scWhereClause = scWhereClause & _
                "(SELECT pk_purch_doc " & _
                            "FROM finance.dbo.t_purch_doc_item " & _
                            "WHERE tracking LIKE '" & scTemp & "')"
        Else
            scWhereClause = scWhereClause & _
                "(SELECT pk_purch_doc " & _
                            "FROM finance.dbo.t_purch_doc_item " & _
                            "WHERE tracking = '" & scTemp & "')"
        End If
    End If
    
    '===========
    ' Next ... Vendor's.
    '===========
    Dim scVendorIDList As String
    
    If Me.lbVendors.ListCount > 0 Then
    
        '========
        ' Get the vendor list string for the IN statement
        '========
        For i = 0 To (Me.lbVendors.ListCount - 1)
            If Len(scVendorIDList) > 0 Then scVendorIDList = scVendorIDList & ", "
            scVendorIDList = scVendorIDList & Me.lbVendors.List(i, 0)
        Next
        
    
        If (Len(scWhereClause) > 0) Then
            scWhereClause = scWhereClause & " INTERSECT "
        End If
        scWhereClause = scWhereClause & _
            "(SELECT DISTINCT pk_purch_doc " & _
                "FROM finance.dbo.t_purch_doc " & _
                            "WHERE fk_vendor IN (" & scVendorIDList & "))"
                
    End If
    
    '===========
    ' Next ... Value Range
    '===========
    Dim scValueRangeSubWhere As String
    Dim dVal As Double
    
    
    '===========
    ' Min Value
    '===========
    If ((txtValueMinFilter.Text <> "") And (txtValueMinFilter.Text <> "*")) Then
        
On Error GoTo bad_number
        dVal = Val(txtValueMinFilter.Text)
        scValueRangeSubWhere = "(total_order_value >= " & Format(dVal, "#,##0.00") & ")"
    End If
    
    '===========
    ' Max Value
    '===========
    If ((txtValueMaxFilter.Text <> "") And (txtValueMaxFilter.Text <> "*")) Then
        
        If Len(scValueRangeSubWhere) > 0 Then
            scValueRangeSubWhere = scValueRangeSubWhere & " AND "
        End If
        dVal = Val(txtValueMaxFilter.Text)
        scValueRangeSubWhere = "(total_order_value <= " & Format(dVal, "#,##0.00") & ")"
    End If
    
    '===========
    ' to be delivered.
    '===========
    If Not cbIncludeClosedPOs.Value Then
        If Len(scValueRangeSubWhere) > 0 Then
            scValueRangeSubWhere = scValueRangeSubWhere & " AND "
        End If
        scValueRangeSubWhere = "(total_to_be_delivered > 0.00)"
          
    End If
    
    '===========
    ' Do we have any value restrictions?
    '===========
    If Len(scValueRangeSubWhere) > 0 Then
    
        If (Len(scWhereClause) > 0) Then
            scWhereClause = scWhereClause & " INTERSECT "
        End If
        scWhereClause = scWhereClause & _
            "(SELECT pk_purch_doc " & _
                "FROM finance.dbo.v_purch_doc_totals " & _
                            "WHERE " & scValueRangeSubWhere & ")"
    End If
        
    GoTo purchasing_group_filter
    
bad_number:
    Call MsgBox("Min or Max values are not integers")
    Exit Sub
    
purchasing_group_filter:
    On Error GoTo 0 ' reset error handling
    
    '===========
    ' Next ... Purchasing Groups
    '===========
    Dim scPurchGroupSubGroup As String
    Dim scMaterialSubQuery As String
    
    If Not cbPOTypeCatalogued.Value Or _
        Not cbPOTypeFreeText.Value Or _
        Not cbPOTypeService.Value Or _
        Not cbPOTypeRotable.Value Or _
        Not cbPOTypeOther.Value Then
        
        Dim scPGCategoryIn As String
        
        If cbPOTypeFreeText.Value Then
            scPGCategoryIn = scPGCategoryIn & "'Free Text'"
        End If
        
        If cbPOTypeService.Value Then
            If Len(scPGCategoryIn) > 0 Then scPGCategoryIn = scPGCategoryIn & ", "
            scPGCategoryIn = scPGCategoryIn & "'Service'"
        End If
        
        If cbPOTypeRotable.Value Then
            If Len(scPGCategoryIn) > 0 Then scPGCategoryIn = scPGCategoryIn & ", "
            scPGCategoryIn = scPGCategoryIn & "'Rotable'"
        End If
        
        If cbPOTypeOther.Value Then
            If Len(scPGCategoryIn) > 0 Then scPGCategoryIn = scPGCategoryIn & ", "
            scPGCategoryIn = scPGCategoryIn & "'Other'"
        End If
        
        '============
        ' With the 'Catalogued' option, the user has the option of including
        ' a specific material.
        ' If they include the material, we have a separate OR of the form
        ' ((purch_group_category = 'Catalogued') AND (fk_material = 9123456))
        ' Else we just add it to the IN comma separated list.
        '============
        If cbPOTypeCatalogued.Value Then
        
            Dim bMaterialDefined As Boolean
            Dim iVal As Long
        
            If IsNumeric(txtMaterial.Text) Then
                
                iVal = CLng(txtMaterial.Text)
                If iVal > 9000000 And iVal < 9999999 Then
                    bMaterialDefined = True
                End If
            End If
            
            If bMaterialDefined Then
                scMaterialSubQuery = _
                    "(purch_group_category = 'Catalogued') AND (fk_material = " & iVal & ")"
            Else
                If Len(scPGCategoryIn) > 0 Then scPGCategoryIn = scPGCategoryIn & ", "
                scPGCategoryIn = scPGCategoryIn & "'Catalogued'"
            End If
        End If
        
        '==============
        ' Complete construction of the sub query
        '==============
        If Len(scPGCategoryIn) > 0 Then
            scPurchGroupSubGroup = "(purch_group_category IN (" & scPGCategoryIn & "))"
        End If
            
        If Len(scMaterialSubQuery) > 0 Then
            If Len(scPurchGroupSubGroup) > 0 Then scPurchGroupSubGroup = scPurchGroupSubGroup & " OR "
            scPurchGroupSubGroup = scPurchGroupSubGroup & "(" & scMaterialSubQuery & ")"
        End If
        
        If (Len(scWhereClause) > 0) Then
            scWhereClause = scWhereClause & " INTERSECT "
        End If
        scWhereClause = scWhereClause & _
            "(SELECT PO " & _
                "FROM finance.dbo.v_purch_orders WHERE (" & scPurchGroupSubGroup & "))"
        
    End If
    
    If Len(Trim(scWhereClause)) = 0 Then
        Call MsgBox("No search criteria specified")
        lblSearchResultsComment.Caption = "No search criteria specified"
        Exit Sub
    End If
    
    '=========
    ' Construct the full query
    '=========
    scWhereClause = "PO IN (" & scWhereClause & ")"
    
    
    scSQLQuery = "SELECT TOP " & c_iMaxPOsSearchCount & " PO, POItem, POItemDescription, CreationDate, NetPrice, ValueToBeDelivered " & _
            "FROM finance.dbo.v_purch_orders " & _
            "WHERE " & scWhereClause & " ORDER BY PO DESC, POItem"
            
    '============
    ' For debugging purposes only, display the query in the query textbox.
    '============
    txtQuery.Text = scSQLQuery
    
    '============
    ' Run the query.
    '============
    Call GetDBRecordSet(ldFinance, cnn, scSQLQuery, rs)
    
    '==========
    ' Clear the listbox
    '==========
    Call lbPurchOrders.Clear
        
    iPOCount = 0
    iCount = 0
    scLastPO = ""
    While Not rs.EOF
        iCount = iCount + 1
        scPO = rs.Fields("PO")
        
        If scPO <> scLastPO Then
            iPOCount = iPOCount + 1
            Call lbPurchOrders.AddItem(rs.Fields("PO"))
        Else
            Call lbPurchOrders.AddItem(" ")
        End If
        lbPurchOrders.List(lbPurchOrders.ListCount - 1, 1) = rs.Fields("POItem")
        lbPurchOrders.List(lbPurchOrders.ListCount - 1, 2) = rs.Fields("POItemDescription")
        lbPurchOrders.List(lbPurchOrders.ListCount - 1, 3) = Format(rs.Fields("CreationDate"), "dd/mm/yy")
        lbPurchOrders.List(lbPurchOrders.ListCount - 1, 4) = Format(rs.Fields("NetPrice"), "#,##0.00")
        lbPurchOrders.List(lbPurchOrders.ListCount - 1, 5) = Format(rs.Fields("ValueToBeDelivered"), "#,##0.00")
        lbPurchOrders.List(lbPurchOrders.ListCount - 1, 6) = scPO
        
        Call rs.MoveNext
        scLastPO = scPO
    Wend
    
    If iPOCount = 0 Then
        lblSearchResultsComment.Caption = "No PO's found"
    ElseIf iCount = c_iMaxPOsSearchCount Then
        lblSearchResultsComment.Caption = "Atleast " & iPOCount & " PO's found. " & c_iMaxPOsSearchCount & " listed."
    ElseIf iPOCount = 1 Then
        lblSearchResultsComment.Caption = "1 PO found."
    Else
        lblSearchResultsComment.Caption = iPOCount & " PO's found."
    End If
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnHelp_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Get's the relevant help range and invokes the ufHelp userform.
'==============================================================================
Private Sub btnHelp_Click()
    Dim uf As ufHelp
    
    Set uf = New ufHelp
    
    Call ufHelp.SetHelpFromRange(wsHelp.Range("Help_POSearch"))
    
    Call ufHelp.Show
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   cmdClose_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Hides the form
'==============================================================================
Private Sub cmdClose_Click()

    If Not Me.m_ufPODisplay Is Nothing Then
        If Me.m_ufPODisplay.Visible Then
           Call Me.m_ufPODisplay.Hide
        End If
    End If
    
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbPurchOrders_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the user selects a PO in the list
'==============================================================================
Private Sub lbPurchOrders_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim scPO As String
    
On Error GoTo exit_cleanly
    Application.EnableEvents = False
    
    Call Me.m_ufPODisplay.ClearDisplay(False, True)
    
    If lbPurchOrders.ListIndex < 0 Then
        Exit Sub
    End If
    
    '=============
    ' Get the PO and populate the display form
    '=============
    scPO = Val(lbPurchOrders.List(lbPurchOrders.ListIndex, 6))
    Call Me.m_ufPODisplay.DisplayPO(scPO)
    
    '=============
    ' Show the form
    '=============
    Me.btnToggleDisplayPO.Enabled = True
    
    If Not Me.m_ufPODisplay.Visible Then
        'Call Me.m_ufPODisplay.Show(vbModeless) ' Error will occur with "Can't show modeless form when Modal form is displayed"
        Call Me.m_ufPODisplay.Show(vbModeless)
        Me.btnToggleDisplayPO.Caption = "Hide PO Detail"
    End If
    
    Application.EnableEvents = True
    Exit Sub

exit_cleanly:
    Application.EnableEvents = True
    
    Call MsgBox("Error: " & Err.description)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnToggleDisplayPO_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   This is a cludge solution to handle when the PO Detail display
' disappears behind another window
'==============================================================================
Private Sub btnToggleDisplayPO_Click()
    If m_ufPODisplay Is Nothing Then
        btnToggleDisplayPO.Enabled = False
        Exit Sub
    End If
    
    If m_ufPODisplay.Visible Then
        Call m_ufPODisplay.Hide
        Me.btnToggleDisplayPO.Caption = "Show PO Detail"
    Else
        Call m_ufPODisplay.Show(vbModeless)
        Me.btnToggleDisplayPO.Caption = "Hide PO Detail"
    End If
        
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnListVendorsFromFilter_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the user does a vendor search
'==============================================================================
Private Sub btnListVendorsFromFilter_Click()
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    
    '==========
    ' Connect to the db
    '==========
    If Not ConnectToDB(ldFinance, cnn, True) Then
        Call MsgBox("Unable to connect to the database")
        Exit Sub
    End If
    
    If InStr(1, txtVendorFilter.Text, "*") > 0 Then
        '===========
        ' Then the user has included a wildcard, so we don't add extra's
        '===========
        scSQLQuery = "SELECT TOP 30 pk_vendor, vendor_name FROM finance.dbo.v_lihir_vendors WHERE " & _
            "vendor_name LIKE '" & Replace(txtVendorFilter.Text, "*", "%") & "'"
    Else
        scSQLQuery = "SELECT TOP 30 pk_vendor, vendor_name FROM finance.dbo.v_lihir_vendors WHERE " & _
            "vendor_name LIKE '%" & Replace(txtVendorFilter.Text, "*", "%") & "%'"
    End If
    
    Set rs = CreateObject("ADODB.Recordset")
    
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    Call Me.lbVendors.Clear
    
    While Not rs.EOF
        Call Me.lbVendors.AddItem(rs.Fields("pk_vendor"))
        Me.lbVendors.List(Me.lbVendors.ListCount - 1, 1) = rs.Fields("vendor_name")
    
        Call rs.MoveNext
    Wend
    Call rs.Close
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnClearVendorPartNoSearch_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clears and effectively disables searching by vendor part number
'==============================================================================
Private Sub btnClearVendorSearch_Click()
    txtVendorFilter.Text = "*"
    Call lbVendors.Clear
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnDownloadResults_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnDownloadResults_Click()

    Dim i As Long
    
    If lbPurchOrders.ListCount = 0 Then
        Exit Sub
    End If

    '=============
    ' Clear the entire contents of the sheet
    '=============
    Call wsFreeCanvas.Cells.Clear
    
    wsFreeCanvas.Cells(1, 1) = "Purch. Ord"
    wsFreeCanvas.Cells(1, 2) = "Item"
    wsFreeCanvas.Cells(1, 3) = "Description"
    wsFreeCanvas.Cells(1, 4) = "Creation Date"
    wsFreeCanvas.Cells(1, 5) = "Net Price"
    wsFreeCanvas.Cells(1, 6) = "Still to be Delivered"
    
    With wsFreeCanvas.Range("A1:F1")
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    For i = 0 To (lbPurchOrders.ListCount - 1)
        wsFreeCanvas.Cells(i + 2, 1) = lbPurchOrders.List(i, 6)
        wsFreeCanvas.Cells(i + 2, 2) = lbPurchOrders.List(i, 1)
        wsFreeCanvas.Cells(i + 2, 3) = lbPurchOrders.List(i, 2)
        wsFreeCanvas.Cells(i + 2, 4) = CDate(lbPurchOrders.List(i, 3))
        wsFreeCanvas.Cells(i + 2, 5) = Val(lbPurchOrders.List(i, 4))
        wsFreeCanvas.Cells(i + 2, 6) = Val(lbPurchOrders.List(i, 5))
    Next
    
    '============
    ' Make a copy of our FreeCanvas tab for the user
    '============
    Dim ws As Excel.Worksheet
    Dim wb As Excel.Workbook

    wsFreeCanvas.Copy
    
    Set ws = Application.ActiveSheet
    Set wb = Application.ActiveWorkbook
    
    '============
    ' Some basic formatting
    '============
    '============
    ' Autofit
    '============
    ws.Cells.Select
    ws.Cells.EntireColumn.AutoFit
    
    '============
    ' Dates and number formatting
    '============
    ws.Columns("D:D").Select
    Selection.NumberFormat = "d-mmm-yyyy"
    ws.Columns("E:F").Select
    Selection.NumberFormat = "#,##0.00"
    
    ws.Name = "SearchResults"
    ws.Cells(2, 1).Select
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   cbPOTypeCatalogued_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the user toggles the Catalogued (Materials) check box.
'==============================================================================
Private Sub cbPOTypeCatalogued_Click()

    If cbPOTypeCatalogued.Value Then
        txtMaterial.Enabled = True
        txtMaterial.BackColor = RGB(255, 255, 255)
    Else
        txtMaterial.Enabled = False
        txtMaterial.BackColor = RGB(223, 223, 223)
    End If
End Sub




