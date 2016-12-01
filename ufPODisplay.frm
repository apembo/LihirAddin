VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPODisplay 
   Caption         =   "Review Purchase Order"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8715
   OleObjectBlob   =   "ufPODisplay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPODisplay"
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
Private Const g_cNullPO As String = "----------"
Private m_bUpdatingPO_DoNotFire As Boolean

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public m_scPurchOrd As String
Public m_oSearchFormParent As ufPOSearch
Public m_colPOItems As VBA.Collection



'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the form is created.
'==============================================================================
Private Sub UserForm_Initialize()
    
    m_bUpdatingPO_DoNotFire = False
    
    Call RemoveUserformCloseButton(Me)
    
End Sub

'==============================================================================
' SUBROUTINE
'   ClearDisplay
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Public Sub ClearDisplay(bEventsAlreadySuppressed As Boolean, bIncludingPO As Boolean)

    Dim i As Long

On Error GoTo cleanup_nicely
    '============
    ' Turn off event handling etc.
    '============
    If Not bEventsAlreadySuppressed Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        m_bUpdatingPO_DoNotFire = True
    End If
    
    If (bIncludingPO) Then
        txtDetailPO.Text = ""
        m_scPurchOrd = ""
    End If
    txtDetailDescription.Text = ""

    '============
    ' Set the Tag attribute for all the pages to 0 to indicate they haven't
    ' been populated.
    '============
    For i = 0 To (multiPageHeader.Pages.Count - 1)
        multiPageHeader.Pages.Item(i).Tag = "0"
    Next

    '============
    ' Header - Conditions
    '============
    txtNetValue.Text = ""
    txtCurrency.Text = ""
    
    txtPaymentTermsCode.Text = ""
    lblPaymentTermsDescription.Caption = "-"
    txtVendorAgreement.Text = ""
    
    '============
    ' Header - Vendor
    '============
    txtVendorID.Text = ""
    txtVendorDescription.Text = ""
    txtStreet.Text = ""
    txtPostCode.Text = ""
    txtCity.Text = ""
    txtCountryCode.Text = ""
    lblCountry.Caption = "-"
    
    txtTelephone.Text = ""

    '============
    ' Header - Vendor
    '============
    txtPurchasingOrg.Text = ""
    lblPurchasingOrgDescription.Caption = "-"

    txtPurchasingGroup.Text = ""
    lblPurchasingGroupDescription.Caption = "-"

    txtCompanyCode.Text = ""
    txtCompanyDescription.Caption = "-"

    '============
    ' Header - Status
    '============
    lblStatusOrdered.Caption = ""
    lblCurrency1.Caption = ""

    lblStatusDelivered.Caption = ""
    lblCurrency2.Caption = ""

    lblStatusStillToDeliver.Caption = ""
    lblCurrency3.Caption = ""

    lblStatusInvoiced.Caption = ""
    lblCurrency4.Caption = ""
    
    '============
    ' Clear the item list
    '============
    Call lbPOItems.Clear
    
    '============
    ' Item Display
    '============
    Call ClearItemDisplay(True)


cleanup_nicely:
    '============
    ' Turn event handling back on.
    '============
    If Not bEventsAlreadySuppressed Then
        m_bUpdatingPO_DoNotFire = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    
End Sub

'==============================================================================
' SUBROUTINE
'   DisplayPO
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Public Sub DisplayPO(Optional scPO As String = g_cNullPO)

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    
#If DebugBadType = 0 Then
   Dim oItem As clsPOItem
#Else
   Dim oItem As Object
#End If

    Dim scSQLQuery As String

    Dim dUpdateAge As Double
    Dim i As Long

On Error GoTo cleanup_nicely
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    m_bUpdatingPO_DoNotFire = True
    
    '==========
    ' Clear the display.
    '==========
    Call Me.ClearDisplay(True, False)
    If scPO <> g_cNullPO Then
        Me.txtDetailPO.Text = scPO
    End If
    
    '==========
    ' Valid PO?
    '==========
    If Not LooksLikeValidPO(txtDetailPO.Text) Then
        Call MsgBox("'" & txtDetailPO.Text & "' is not a valid Purchase Order.")
        GoTo cleanup_nicely
    End If
    
    '==========
    ' We first display the information outside the MultiPage
    '==========
    '==========
    ' Construct the query
    '==========
    scSQLQuery = "SELECT * FROM finance.dbo.v_purch_doc_detail WHERE PO = '" & txtDetailPO.Text & "'"
    Call GetDBRecordSet(ldFinance, cnn, scSQLQuery, rs)

    If rs.EOF Then
        txtDetailDescription.Text = "Purchase order not found"
        Call rs.Close
        GoTo cleanup_nicely
    End If

    '==========
    ' Display the PO
    '==========
    If Not IsNull(rs.Fields("PODescription")) Then
        txtDetailDescription.Text = rs.Fields("PODescription")
    End If
    
    '==========
    ' Multipage TAB
    '==========
    '==========
    ' Conditions
    '==========
    txtNetValue.Text = Format(rs.Fields("total_net_prices"), "#,##0.00")
    txtCurrency.Text = rs.Fields("currency")
    
    If Not IsNull(rs.Fields("pk_vendor_terms_of_pay")) Then
        txtPaymentTermsCode.Text = rs.Fields("pk_vendor_terms_of_pay")
        lblPaymentTermsDescription.Caption = rs.Fields("vendor_terms_of_pay")
    End If

    If Not IsNull(rs.Fields("outline_agreement")) Then
        txtVendorAgreement.Text = rs.Fields("outline_agreement")
    End If
    
    '==========
    ' Vendor
    '==========
On Error GoTo after_vendor_address_update

    If rs.Fields("pk_vendor") <> 0 Then
        txtVendorID.Text = rs.Fields("pk_vendor")
        txtVendorDescription.Text = rs.Fields("vendor_name")
        
        txtStreet.Text = rs.Fields("street")
        
        If Not IsNull(rs.Fields("postal_code")) Then
            txtPostCode.Text = rs.Fields("postal_code")
        End If
        txtCity.Text = rs.Fields("city")
        
        txtCountryCode.Text = rs.Fields("fk_country")
        lblCountry.Caption = rs.Fields("country")
        If Not IsNull(rs.Fields("telephone")) Then
            txtTelephone.Text = rs.Fields("telephone")
        End If
    End If
after_vendor_address_update:
On Error GoTo cleanup_nicely
    
    '==========
    ' Org. Data
    '==========
    If Not IsNull(rs.Fields("fk_purch_org")) Then
        txtPurchasingOrg.Text = rs.Fields("fk_purch_org")
        'lblPurchasingOrgDescription.Caption = rs.Fields("")
    Else
        txtPurchasingOrg.Text = "2351"
        lblPurchasingOrgDescription.Caption = "Lihir Operations"
    End If
    
    txtPurchasingGroup.Text = rs.Fields("pk_purch_group")
    lblPurchasingGroupDescription.Caption = rs.Fields("purch_group")
    
    If Not IsNull(rs.Fields("fk_company_code")) Then
        txtCompanyCode.Text = rs.Fields("fk_company_code")
        'txtCompanyDescription.Caption = rs.Fields("")
    Else
        txtCompanyCode.Text = "2351"
        txtCompanyDescription.Caption = "Lihir Gold Limited"
    End If
    
    '==========
    ' Status
    '==========
    lblStatusOrdered.Caption = Format(rs.Fields("total_order_value"), "#,##0.00")
    lblStatusDelivered.Caption = "?" 'rs.Fields("")
    lblStatusStillToDeliver.Caption = Format(rs.Fields("total_to_be_delivered"), "#,##0.00")
    lblStatusInvoiced.Caption = Format(rs.Fields("total_to_be_invoiced"), "#,##0.00")
    
    lblCurrency1.Caption = rs.Fields("currency")
    lblCurrency2.Caption = rs.Fields("currency")
    lblCurrency3.Caption = rs.Fields("currency")
    lblCurrency4.Caption = rs.Fields("currency")
    
    '==========
    ' The last update date string and colour
    '==========
    Me.txtLastUpdateFromSAP.Value = Format(rs.Fields("last_updated").Value, "ddd d-mmm-yy h:mm am/pm")

    dUpdateAge = (Date + Time) - rs.Fields("last_updated").Value
    Me.btnLastUpdateColour.BackColor = GetRGBFromAge(dUpdateAge, 1#, 7#, 28#)
    
    
    Call rs.Close
    
    '==========
    ' PO Items
    '==========
    '==========
    ' List Box
    '==========
    scSQLQuery = "SELECT * FROM finance.dbo.v_purch_doc_items " & _
        "WHERE pk_purch_doc = '" & txtDetailPO.Text & "'"
    Call GetDBRecordSet(ldFinance, cnn, scSQLQuery, rs)

    If rs.EOF Then
        txtDetailDescription.Text = "Purchase order items not found"
        Call rs.Close
        GoTo cleanup_nicely
    End If
    
    Set m_colPOItems = New VBA.Collection
    Call lbPOItems.Clear
    
    While Not rs.EOF
        
        '===========
        ' Create the POItem class instance, populate and add to the collection.
        '===========
        Set oItem = New clsPOItem
        Call oItem.Populate(rs.Fields)
        Call m_colPOItems.Add(oItem)
        
        '===========
        ' Add as a line to the list box.
        '===========
        Call lbPOItems.AddItem(oItem.pk_item)
        lbPOItems.List(lbPOItems.ListCount - 1, 1) = oItem.short_text
        lbPOItems.List(lbPOItems.ListCount - 1, 2) = oItem.order_qty
        lbPOItems.List(lbPOItems.ListCount - 1, 3) = oItem.order_unit
        lbPOItems.List(lbPOItems.ListCount - 1, 4) = Format(oItem.unit_price, "#,##0.00")
        lbPOItems.List(lbPOItems.ListCount - 1, 5) = Format(oItem.net_price, "#,##0.00")
    
        Call rs.MoveNext
    Wend
    

    Call rs.Close
        
cleanup_nicely:
    m_bUpdatingPO_DoNotFire = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

'==============================================================================
' FUNCTION
'   lbPOItems_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Populates the item specific details
'==============================================================================
Private Sub lbPOItems_Click()

#If DebugBadType = 0 Then
   Dim oItem As clsPOItem
#Else
   Dim oItem As Object
#End If
    
    Set oItem = m_colPOItems.Item(lbPOItems.ListIndex + 1)
    
    '===========
    ' Material Data
    '===========
    txtMaterialGroup.Text = oItem.fk_material_group
    lblMaterialGroupDescription.Caption = oItem.material_group
    
    '===========
    ' Account Assignment
    '===========
    txtRecipient.Text = oItem.goods_recipient
    txtGLAccount.Text = oItem.fk_cost_element
    lblCostElement.Caption = oItem.cost_element
    
    If oItem.WorkOrder = 0 Then
        txtWorkOrder.Text = ""
        lblWorkOrderDescription.Caption = ""
    Else
        txtWorkOrder.Text = oItem.WorkOrder
        lblWorkOrderDescription.Caption = oItem.WorkOrderDescription
    End If
    
    Call DisplayCostCenter(oItem)
    

    '===========
    ' Purchase Order History
    '===========


    '===========
    ' Enable the controls in case they have been disabled.
    '===========
    Call ConfigItemControls(False)
    
End Sub

'==============================================================================
' FUNCTION
'   ClearItemDisplay
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clears the item display (all tabs below the item list)
'==============================================================================
Private Sub ClearItemDisplay(Optional bDisableControls As Boolean = False)

    '============
    ' Material Display
    '============
    txtMaterialGroup.Text = ""
    lblMaterialGroupDescription.Caption = "-"
    
    '============
    ' Account Assignment
    '============
    txtUnloadingPoint.Text = ""
    txtRecipient.Text = ""
    txtGLAccount.Text = ""
    txtWorkOrder.Text = ""
    
    txtCostCenter.Text = ""
    lblCostCenterDescription.Caption = "-"
    lblCostCenterDepartment.Caption = "-"
    
    
    '============
    ' Optionally disable the controls to ensure avoid confusion about why they
    ' are not showing anything.
    '============
    Call ConfigItemControls(bDisableControls)
    
End Sub

'==============================================================================
' SUBROUTINE
'   ConfigItemControls
'------------------------------------------------------------------------------
' DESCRIPTION
'   Enables/Disable Item Controls and set background colour
'==============================================================================
Sub ConfigItemControls(bDisable As Boolean)

    If bDisable Then
        txtMaterialGroup.Enabled = False
        
        txtUnloadingPoint.Enabled = False
        txtRecipient.Enabled = False
        txtGLAccount.Enabled = False
        txtWorkOrder.Enabled = False
        
        txtCostCenter.Enabled = False
        lblCostCenterDescription.Enabled = False
        lblCostCenterDepartment.Enabled = False
    
        txtMaterialGroup.BackColor = RGB(223, 223, 223)
        txtUnloadingPoint.BackColor = RGB(223, 223, 223)
        txtRecipient.BackColor = RGB(223, 223, 223)
        txtGLAccount.BackColor = RGB(223, 223, 223)
        txtWorkOrder.BackColor = RGB(223, 223, 223)
        
        txtCostCenter.BackColor = RGB(223, 223, 223)

    Else
        txtMaterialGroup.Enabled = True
        
        txtUnloadingPoint.Enabled = True
        txtRecipient.Enabled = True
        txtGLAccount.Enabled = True
        txtWorkOrder.Enabled = True
        
        txtCostCenter.Enabled = True
        lblCostCenterDescription.Enabled = True
        lblCostCenterDepartment.Enabled = True
    
        txtMaterialGroup.BackColor = RGB(255, 255, 255)
        
        txtUnloadingPoint.BackColor = RGB(255, 255, 255)
        txtRecipient.BackColor = RGB(255, 255, 255)
        txtGLAccount.BackColor = RGB(255, 255, 255)
        txtWorkOrder.BackColor = RGB(255, 255, 255)

        txtCostCenter.BackColor = RGB(255, 255, 255)
    End If
End Sub

'==============================================================================
' FUNCTION
'   DisplayCostCenter
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
#If DebugBadType = 0 Then
Sub DisplayCostCenter(oItem As clsPOItem)
#Else
Sub DisplayCostCenter(oItem As Object)
#End If

    If oItem.CostCenter = "" Then
        txtCostCenter.Text = ""
        lblCostCenterDescription.Caption = "-"
        lblCostCenterDepartment.Caption = "-"
    Else
        txtCostCenter.Text = oItem.CostCenter
        lblCostCenterDescription.Caption = oItem.CCColloquialName
        lblCostCenterDepartment.Caption = oItem.Dept
    End If
End Sub


'==============================================================================
' FUNCTION
'   LooksLikeValidPO
'------------------------------------------------------------------------------
' DESCRIPTION
'   Determines whether the past in string is a valid PO ID.
' Basically it is, if:
' - it starts with 45 and is 10 digits long, or
' - Starts with 'LG...'
'==============================================================================
Public Function LooksLikeValidPO(scPO As String) As Boolean

    '===========
    ' Is it one of the old Ellipse PO's that start with "LG..."
    '===========
    If Left(scPO, 2) = "LG" Then
        LooksLikeValidPO = True
        Exit Function
    End If
    
On Error GoTo not_a_valid_PO

    If ((Left(scPO, 2) = "41") Or (Left(scPO, 2) = "45")) And _
        Len(scPO) = 10 Then
        
        LooksLikeValidPO = True
        Exit Function
    End If

not_a_valid_PO:
    LooksLikeValidPO = False
    
End Function

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtDetailPO_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Handles when the user enters a Purchase Order manually.
'==============================================================================
Private Sub txtDetailPO_Change()

    If m_bUpdatingPO_DoNotFire Then
        Exit Sub
    End If
    
    '=========='
    ' Does the PO entered look right?
    '==========
    If Not Me.LooksLikeValidPO(txtDetailPO.Text) Then
        Call ClearDisplay(False, False)
        Exit Sub
    End If
    
    Call DisplayPO
        
End Sub

'==============================================================================
' SUBROUTINE
'   DisplayPODetail
'------------------------------------------------------------------------------
' DESCRIPTION
'   Does the actual data collection and display, based on the contents of the
' PO text box.
'==============================================================================
Private Sub DisplayPODetail()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim dUpdateAge As Double
    Dim scStorageLocDescription As String
    Dim i As Long
    
    '==========
    ' Start with a blank slate to ensure we don't display old data
    '==========
    If Me.txtDetailPO.Text <> "" Then
        Call ClearDisplay(False, False)
    End If
    
    '=========='
    ' Does the PO entered look right?
    '==========
    If Not Me.LooksLikeValidPO(txtDetailPO.Text) Then
        Exit Sub
    End If
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Activate
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Private Sub UserForm_Activate()
    'Call ClearDisplay
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnClose_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Private Sub btnClose_Click()
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   DebugDisplaySQLQuery
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Public Sub DebugDisplaySQLQuery(scSQLQuery As String)
    If Not m_oSearchFormParent Is Nothing Then
        m_oSearchFormParent.txtQuery.Text = scSQLQuery
    End If
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnUpdatePODescription_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Event called on the custom description Update button click
'==============================================================================
Private Sub btnUpdatePODescription_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
#Else
    Dim cnn As Object
#End If
    Dim scSQLQuery As String

    Dim iRecordsAffected As Long
    
    If Not Me.LooksLikeValidPO(txtDetailPO.Text) Then
        Call MsgBox("The PO is not a valid PO")
        Exit Sub
    End If
        
    scSQLQuery = "UPDATE finance.dbo.t_purch_doc SET description = '" & _
        Left(Replace(txtDetailDescription.Text, "'", "''"), 100) & _
        "' WHERE pk_purch_doc = '" & txtDetailPO.Text & "'"
        
    If Not ConnectToDB(ldFinance, cnn) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Sub
    End If
    
On Error GoTo update_failed
    Call cnn.Execute(scSQLQuery, iRecordsAffected)
    
    If Abs(iRecordsAffected) = 0 Then
        Call MsgBox("No change in the database")
    Else
        Call MsgBox("Update successful")
    End If
    Exit Sub
    
update_failed:
    Call MsgBox("Error attempting to update the database: " & Err.description)
    
    On Error GoTo 0
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnDisplayWorkOrder_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Displays the work order if available
'==============================================================================
Private Sub btnDisplayWorkOrder_Click()
    
    Dim iWO As Long
    Dim uf As ufWorkOrderDisplay
    Dim oErr As typAppError
    
    If Not LooksLikeValidWorkOrder(txtWorkOrder.Text, iWO) Then
        Call MsgBox("Not a valid work order")
        Exit Sub
    End If
    
    Set uf = New ufWorkOrderDisplay
    
    If Not uf.SetWorkOrder(iWO, oErr, True) Then
        Call MsgBox("Could not display work order. Error: " & oErr.description)
        Exit Sub
    End If
    
    Call uf.Show
    
    Call Unload(uf)
    
End Sub


