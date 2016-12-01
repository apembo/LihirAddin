VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPartsSearch 
   Caption         =   "Search for Parts"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11655
   OleObjectBlob   =   "ufPartsSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPartsSearch"
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
Private Const c_iMaxPartsSearchCount As Long = 500

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public m_bInEventHandler As Boolean
Public m_ufPartDisplay As ufPartDisplay


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Call on userform creation
'==============================================================================
Private Sub UserForm_Initialize()
    Set m_ufPartDisplay = New ufPartDisplay
    Set Me.m_ufPartDisplay.m_oSearchFormParent = Me
    m_bInEventHandler = False

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   ResetForm
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Public Sub ResetForm()
    Me.txtDescriptionFilter.Text = "*"
    cmbColloquial1.Text = ""
    Me.cmbColloquial2.Text = ""
    Call btnClearVendorPartNoSearch_Click
    
    Me.cbIncludeDeleted.Value = False
    
    Call Me.lbMaterials.Clear
    
    Me.btnToggleDisplayMaterial.Caption = "Hide Material Detail"
    btnToggleDisplayMaterial.Enabled = False

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
    Dim i As Long
    Dim scaColloquials(1 To 3) As String
    Dim iColloquialCount As Long
    Dim scWhereClause As String
    Dim sc
    
'On Error GoTo cleanup_nicely
    
    If Me.txtDescriptionFilter.Text = "" And Me.cmbColloquial1.Text = "" And Me.cmbColloquial2.Text = "" Then
        Call MsgBox("Need atleast a description filter or one colloquial filter for the search")
        Exit Sub
    End If
    
    '==========
    ' Connect to the db
    '==========
    If Not ConnectToDB(ldParts, cnn, True) Then
        Call MsgBox("Unable to connect to the database")
        Exit Sub
    End If
    
    '==========
    ' How many colloquial's?
    '==========
    iColloquialCount = 0
    
    If Me.cmbColloquial1.Text <> "" Then
        iColloquialCount = iColloquialCount + 1
        scaColloquials(iColloquialCount) = Replace(Me.cmbColloquial1.Text, "*", "%")
    End If

    If Me.cmbColloquial2.Text <> "" Then
        iColloquialCount = iColloquialCount + 1
        scaColloquials(iColloquialCount) = Replace(Me.cmbColloquial2.Text, "*", "%")
    End If

'    If Me.cmbColloquial3.Text <> "" Then
'        iColloquialCount = iColloquialCount + 1
'        scaColloquials(iColloquialCount) = Replace(Me.cmbColloquial3.Text, "*", "%")
'    End If
    
    '==========
    ' Construct the WHERE clause of the query. It is basically a set of
    ' INTERSECT's.
    '==========
    
    '===========
    ' First ... Colloquials
    '===========
    For i = 1 To iColloquialCount
        If i > 1 Then
            scWhereClause = scWhereClause & " INTERSECT "
        End If
        scWhereClause = scWhereClause & _
            "(SELECT DISTINCT (fk_sap_material) FROM parts.dbo.t_map_colloquial_material WHERE " & _
            "fk_colloquial IN (SELECT pk_colloquial FROM parts.dbo.t_colloquial WHERE colloquial LIKE '" & scaColloquials(i) & "'))"
    Next
    
    '===========
    ' Next ... descriptions. Haven't implemented custom descriptions as yet
    '===========
    If Me.txtDescriptionFilter.Text <> "" And Me.txtDescriptionFilter.Text <> "*" Then
        Dim scDescriptionFilter As String
        
        scDescriptionFilter = Replace(Me.txtDescriptionFilter.Text, "*", "%")
        If (Len(scWhereClause) > 0) Then
            scWhereClause = scWhereClause & " INTERSECT "
        End If
        scWhereClause = scWhereClause & _
            "(SELECT pk_sap_material " & _
                        "FROM parts.dbo.t_sap_material " & _
                        "WHERE description LIKE '" & scDescriptionFilter & "' OR long_description LIKE '" & scDescriptionFilter & "')"
                        
        '===========
        ' If the user selects to include user descriptions, we create a UNION
        ' query of the results of searching normal descriptions, plus user
        ' descriptions from the t_user_description table.
        '===========
        If Me.cbIncludeUserDescriptions.Value Then
            scWhereClause = "(" & scWhereClause & " UNION " & _
                "(SELECT fk_sap_material FROM parts.dbo.t_user_description " & _
                    "WHERE short_text LIKE '" & scDescriptionFilter & "' OR long_text LIKE '" & scDescriptionFilter & "'))"
        End If
    End If
    
    '===========
    ' Next ... Vendor Part Numbers.
    '===========
    Dim scPartNoFilter As String
    Dim scVendorIDList As String
'    Dim scVendor_Subquery
    Dim scVendor_Where As String
    Dim scPartNumber_SubQuery As String
    Dim scVendorPartNo_SubQuery As String
    
    '========
    ' Get the vendor filter if some vendors have been selected.
    '========
    If Me.lbVendors.ListCount > 0 Then
    
        '========
        ' Get the vendor list string for the IN statement
        '========
        For i = 0 To (Me.lbVendors.ListCount - 1)
            If Len(scVendorIDList) > 0 Then scVendorIDList = scVendorIDList & ", "
            scVendorIDList = scVendorIDList & Me.lbVendors.List(i, 0)
        Next
        
        scVendor_Where = "(mpn.fk_sap_supplier in (" & scVendorIDList & "))"
    End If
    
    '========
    ' Get the part number filter if there is one
    '========
    If Len(Trim(Me.txtPartNoFilter.Text)) = 0 Or Trim(Me.txtPartNoFilter.Text) = "*" Then
        scPartNumber_SubQuery = ""
    Else
        '==========
        ' Are
        '==========
        scPartNoFilter = CreateSimplePartNo(Replace(Me.txtPartNoFilter.Text, "*", "%"), True)
        scPartNumber_SubQuery = _
            "        (SELECT pn.pk_part_number " & vbCrLf & _
            "           FROM parts.dbo.t_part_number as pn " & vbCrLf & _
            "           WHERE pn.simple_part_number LIKE '" & scPartNoFilter & "')"
        
    End If
    
    '=============
    ' Do we include are part number or vendor filter?
    '=============
    If Len(scVendor_Where) > 0 Or Len(scPartNumber_SubQuery) > 0 Then
    
        If Len(scVendor_Where) > 0 Then
            scVendorPartNo_SubQuery = scVendor_Where
        End If
        
    
        If Len(scPartNumber_SubQuery) > 0 Then
            If Len(scVendorPartNo_SubQuery) > 0 Then
                scVendorPartNo_SubQuery = scVendorPartNo_SubQuery & " AND "
            End If
            scVendorPartNo_SubQuery = scVendorPartNo_SubQuery & " (mpn.fk_part_number IN " & scPartNumber_SubQuery & ")"
        End If
    

        If (Len(scWhereClause) > 0) Then
            scWhereClause = scWhereClause & " INTERSECT "
        End If
        scWhereClause = scWhereClause & _
            "(SELECT DISTINCT mpn.fk_sap_material FROM parts.dbo.t_sap_material_part_no as mpn WHERE " & scVendorPartNo_SubQuery & ")"
                    
    End If
    
    If Len(Trim(scWhereClause)) = 0 Then
        Call MsgBox("No search criteria specified")
        lblSearchResultsComment.Caption = "No search criteria specified"
        Exit Sub
    End If
    
    '=============
    ' Has any filter been applied?
    '=============
    scWhereClause = "pk_sap_material IN (" & scWhereClause & ")"
    
    '==========
    ' Include deleted materials?
    '==========
    If Not Me.cbIncludeDeleted.Value Then
        scWhereClause = scWhereClause & " AND deleted = 0"
    End If
    
    scSQLQuery = "SELECT TOP " & c_iMaxPartsSearchCount & " pk_sap_material, description, deleted, stock_level_total " & _
            "FROM parts.dbo.v_sap_material " & _
            "WHERE " & scWhereClause & " ORDER BY pk_sap_material"
            
    txtQuery.Text = scSQLQuery
    
    'Exit Sub ' DEBUG
    Set rs = CreateObject("ADODB.Recordset")
    
    '==========
    ' Has the user requested to suppress the actual query in order to allow
    ' debugging.
    '==========
    If cbSuppressActualQuery.Value Then
        Exit Sub
    End If
    
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    '==========
    ' Clear the listbox
    '==========
    Call Me.lbMaterials.Clear
    If Me.cbIncludeDeleted.Value Then
        Me.lbMaterials.ColumnCount = 4
        Me.lbMaterials.ColumnWidths = "50 pt;230 pt;30 pt;30 pt"
        lblListBoxCol3.Visible = True
        lblListBoxCol4.Caption = "Deleted"
    Else
        Me.lbMaterials.ColumnCount = 3
        Me.lbMaterials.ColumnWidths = "50 pt;260 pt;30 pt"
        lblListBoxCol3.Visible = False
        lblListBoxCol4.Caption = "Stock"
    End If
        
    i = 0
    While Not rs.EOF
        i = i + 1
        Me.lbMaterials.AddItem (rs.Fields("pk_sap_material"))
        Me.lbMaterials.List(Me.lbMaterials.ListCount - 1, 1) = rs.Fields("description")
        Me.lbMaterials.List(Me.lbMaterials.ListCount - 1, 2) = rs.Fields("stock_level_total")
        If Me.cbIncludeDeleted.Value Then
            Me.lbMaterials.List(Me.lbMaterials.ListCount - 1, 3) = rs.Fields("deleted")
        End If
        
        Call rs.MoveNext
    Wend
    
    If i = 0 Then
        lblSearchResultsComment.Caption = "No Materials found"
    ElseIf i = c_iMaxPartsSearchCount Then
        lblSearchResultsComment.Caption = "More than " & c_iMaxPartsSearchCount & " Materials found. " & c_iMaxPartsSearchCount & " listed."
    Else
        lblSearchResultsComment.Caption = i & " Materials found."
    End If
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   cmbColloquial1_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   OnChange event handler for the Colloquial Combobox
'==============================================================================
Private Sub cmbColloquial1_Change()
    Call HandleColloquialComboChange(Me.cmbColloquial1)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   cmbColloquial2_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   OnChange event handler for the Colloquial Combobox
'==============================================================================
Private Sub cmbColloquial2_Change()
    Call HandleColloquialComboChange(Me.cmbColloquial2, Replace(Trim(Me.cmbColloquial1.Text), "*", "%"))
End Sub
Private Sub cmbColloquial2_Click()
    Call HandleColloquialComboChange(Me.cmbColloquial2, Replace(Trim(Me.cmbColloquial1.Text), "*", "%"))
End Sub

'==============================================================================
' SUBROUTINE
'   HandleColloquialComboChange
'------------------------------------------------------------------------------
' DESCRIPTION
'   The common code event handler for the combo-box's
'==============================================================================
Private Sub HandleColloquialComboChange(oCombo As MSForms.ComboBox, Optional scPrevFilter As Variant = "%")

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    Dim scEntry As String
    Dim i As Long
    
    If (IsMissing(scPrevFilter) And Len(oCombo.Text) < 4) Or m_bInEventHandler Then
        Exit Sub
    End If
    
    m_bInEventHandler = True
On Error GoTo cleanup_nicely
    
    If Not ConnectToDB(ldParts, cnn, True) Then
        Call MsgBox("Unable to connect to the database")
        Exit Sub
    End If
    
    scEntry = Replace(Trim(oCombo.Text), "*", "%")
    If (InStr(1, scEntry, "%") < 1) Then
        scEntry = "%" & scEntry & "%"
    End If
    
    '============
    ' If this is the second combobox, then we need to reduce the
    ' set of available colloquials to those that share parts with
    ' the first combobox filter. This creates a complicated (but
    ' tested) query.
    '============
    If scPrevFilter <> "%" And scPrevFilter <> "" Then
        scSQLQuery = _
            "SELECT c2.colloquial FROM parts.dbo.t_colloquial as c2 WHERE c2.pk_colloquial IN " & _
                "(SELECT m2.fk_colloquial from parts.dbo.t_map_colloquial_material as m2 WHERE m2.fk_sap_material IN " & _
                    "(SELECT fk_sap_material FROM parts.dbo.t_map_colloquial_material WHERE fk_colloquial IN " & _
                        "(SELECT c1.pk_colloquial " & _
                            "FROM parts.dbo.t_colloquial AS c1 " & _
                            "WHERE c1.colloquial like '" & scPrevFilter & "'" & _
                        ")" & _
                    ")" & _
                ") AND c2.colloquial LIKE '" & scEntry & "'"
    Else
        scSQLQuery = "SELECT TOP 20 * FROM parts.dbo.t_colloquial WHERE colloquial LIKE '" & scEntry & "'"
    End If
    
    Set rs = CreateObject("ADODB.Recordset")
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    '===========
    ' Remove all items from the list
    '===========
    For i = 1 To oCombo.ListCount
        oCombo.RemoveItem (0)
    Next i
    
    While Not rs.EOF
        Call oCombo.AddItem(rs.Fields("colloquial"))
        
        Call rs.MoveNext
    Wend
    
cleanup_nicely:
    m_bInEventHandler = False
    
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
    
    Call ufHelp.SetHelpFromRange(wsHelp.Range("Help_PartsSearch"))
    
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

    If Not Me.m_ufPartDisplay Is Nothing Then
        If Me.m_ufPartDisplay.Visible Then
           Call Me.m_ufPartDisplay.Hide
        End If
    End If
    
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbMaterials_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the user selects a material in the list
'==============================================================================
Private Sub lbMaterials_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim iMaterial As Long
    
On Error GoTo exit_cleanly
    Application.EnableEvents = False
    
    Call Me.m_ufPartDisplay.ClearDisplay(False, True)
    
    If lbMaterials.ListIndex < 0 Then
        Exit Sub
    End If
    
    '=============
    ' Get the material and populate the display form
    '=============
    iMaterial = Val(lbMaterials.List(lbMaterials.ListIndex, 0))
    Call Me.m_ufPartDisplay.DisplayMaterial(Val(iMaterial))
    
    '=============
    ' Show the form
    '=============
    Me.btnToggleDisplayMaterial.Enabled = True
    
    If Not Me.m_ufPartDisplay.Visible Then
        Call Me.m_ufPartDisplay.Show(vbModeless)
        Me.btnToggleDisplayMaterial.Caption = "Hide Material Detail"
    End If
    
exit_cleanly:
    Application.EnableEvents = True
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnToggleDisplayMaterial_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   This is a cludge solution to handle when the Material Detail display
' disappears behind another window
'==============================================================================
Private Sub btnToggleDisplayMaterial_Click()
    If m_ufPartDisplay Is Nothing Then
        btnToggleDisplayMaterial.Enabled = False
        Exit Sub
    End If
    
    If m_ufPartDisplay.Visible Then
        Call m_ufPartDisplay.Hide
        Me.btnToggleDisplayMaterial.Caption = "Show Material Detail"
    Else
        Call m_ufPartDisplay.Show(vbModeless)
        Me.btnToggleDisplayMaterial.Caption = "Hide Material Detail"
    End If
        
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnListVendorsFromFilter_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the user selects a material in the list
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
Private Sub btnClearVendorPartNoSearch_Click()
    txtVendorFilter.Text = "*"
    txtPartNoFilter.Text = "*"
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
    
    If lbMaterials.ListCount = 0 Then
        Exit Sub
    End If

    '=============
    ' Clear the entire contents of the sheet
    '=============
    Call wsFreeCanvas.Cells.Clear
    
    wsFreeCanvas.Cells(1, 1) = "Material"
    wsFreeCanvas.Cells(1, 2) = "Description"
    wsFreeCanvas.Cells(1, 3) = "Stock Levels"
    
    With wsFreeCanvas.Range("A1:C1")
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    For i = 0 To (lbMaterials.ListCount - 1)
        wsFreeCanvas.Cells(i + 2, 1) = lbMaterials.List(i, 0)
        wsFreeCanvas.Cells(i + 2, 2) = lbMaterials.List(i, 1)
        wsFreeCanvas.Cells(i + 2, 3) = lbMaterials.List(i, 2)
    Next
    
    '============
    ' Make a copy of our FreeCanvas tab for the user
    '============
    Dim ws As Excel.Worksheet
    Dim wb As Excel.Workbook

    wsFreeCanvas.Copy
    
    Set ws = Application.ActiveSheet
    Set wb = Application.ActiveWorkbook
    
    ws.Name = "SearchResults"
    ws.Cells(2, 1).Select
    
End Sub
