VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPartDisplay 
   Caption         =   "Review Lihir Warehouse Material"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8715
   OleObjectBlob   =   "ufPartDisplay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPartDisplay"
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
Private Const c_iMaxColloquials As Long = 200
Private m_cbaColloquialStarsCheckBoxs(1 To 5) As Object
Private m_bIgnoreStarClicks As Boolean

Private m_colPicData As VBA.Collection
Private m_bUpdatingMaterialNo_DoNotFire As Boolean

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public m_oSearchFormParent As ufPartsSearch
Public m_iPictureIndex As Long
Public m_iMaterial As Long

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the form is created.
'==============================================================================
Private Sub UserForm_Initialize()

    Set m_cbaColloquialStarsCheckBoxs(1) = Me.CheckBox1
    Set m_cbaColloquialStarsCheckBoxs(2) = Me.CheckBox2
    Set m_cbaColloquialStarsCheckBoxs(3) = Me.CheckBox3
    Set m_cbaColloquialStarsCheckBoxs(4) = Me.CheckBox4
    Set m_cbaColloquialStarsCheckBoxs(5) = Me.CheckBox5
    
    m_bIgnoreStarClicks = False
    
    Set m_colPicData = New VBA.Collection
    
    Call SetStarRating(3)
    
    m_bUpdatingMaterialNo_DoNotFire = False
    
    Call RemoveUserformCloseButton(Me)
End Sub

'==============================================================================
' SUBROUTINE
'   ClearDisplay
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Public Sub ClearDisplay(bEventsAlreadySuppressed As Boolean, bIncludingMaterialNo As Boolean)

    Dim i As Long

On Error GoTo cleanup_nicely
    '============
    ' Turn off event handling etc.
    '============
    If Not bEventsAlreadySuppressed Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        m_bUpdatingMaterialNo_DoNotFire = True
    End If
    
    If (bIncludingMaterialNo) Then
        txtDetailMaterialNo.Text = ""
        Me.m_iMaterial = 0
    End If

    Me.txtBaseUnitOfMeas = ""
    Me.txtDetailDescription = ""
    Me.txtMaterialGroup = ""
    Me.txtMaterialLongText = ""
    Me.txtMaxStockLevel = ""
    Me.txtMRPType = ""
    Me.txtPlannedDeliveryTime = ""
    Me.txtPlantSpecificMaterialStatus = ""
    Me.txtReorderPoint = ""
    Me.txtStockLevel = ""
    Call Me.lbStoreLocations.Clear
    
    Me.txtLastUpdateFromSAP.Text = ""
    Me.btnLastUpdateColour.BackColor = RGB(255, 255, 255)
        
    '===========
    ' Colloquials
    '===========
    Call lbColloquials.Clear
    txtLihirColloquialsNew.Text = ""
    lblColloquialComments.Caption = ""
    '===========
    ' Pictures
    '===========
    Me.Image1.Picture = LoadPicture("")
    lblPictureInfo.Caption = ""
    lblPictureDate.Caption = ""
    
    '============
    ' Clear the picture collection.
    '============
    Call Me.ClearPicCollection
    
    '============
    ' Clear the descriptions
    '============
    Call Me.ClearDescriptions
    
    '============
    ' Clear the Material Part Numbers
    '============
    Call Me.lbMaterialPartNos.Clear
    
    '============
    ' Clear the usage list
    '============
    Call Me.lbUsageSearchResults.Clear
    Me.lblPartsUsageSearchResults.Caption = ""
    '============
    ' Set the Tag attribute for all the pages to 0 to indicate they haven't
    ' been populated.
    '============
    For i = 0 To (Me.MultiPage1.Pages.Count - 1)
        Me.MultiPage1.Pages.Item(i).Tag = "0"
    Next

cleanup_nicely:
    '============
    ' Turn event handling back on.
    '============
    If Not bEventsAlreadySuppressed Then
        m_bUpdatingMaterialNo_DoNotFire = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    
End Sub

'==============================================================================
' SUBROUTINE
'   DisplayMaterial
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Public Sub DisplayMaterial(iMaterial As Long)
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim dUpdateAge As Double
    Dim scStorageLocDescription As String
    Dim i As Long

On Error GoTo cleanup_nicely

    
    '==========
    ' Suppress event handling since we modify text boxes that have OnChange
    ' event handlers
    '==========
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    m_bUpdatingMaterialNo_DoNotFire = True
    
    '==========
    ' Reset various indices and flags.
    '==========
    Call Me.ClearDisplay(True, False)
    
    '==========
    ' Valid material?
    '==========
    If iMaterial >= 9000000 And iMaterial <= 9999999 Then
        Me.m_iMaterial = iMaterial
    
        '==========
        ' We first display the information outside the MultiPage
        '==========
        scSQLQuery = "SELECT * FROM parts.dbo.v_sap_material WHERE pk_sap_material = " & m_iMaterial

        Call GetDBRecordSet(ldParts, cnn, scSQLQuery, rs)
    
        If rs.EOF Then
            txtMaterialLongText.Text = "Material not found"
            Call rs.Close
            GoTo cleanup_nicely
        End If
    
        Me.txtDetailMaterialNo.Text = rs.Fields("pk_sap_material")
        Me.txtDetailDescription.Text = rs.Fields("description")
        Me.txtMaterialLongText.Text = rs.Fields("long_description")
    
        '==========
        ' The last update date string and colour
        '==========
        Me.txtLastUpdateFromSAP.Value = Format(rs.Fields("last_updated").Value, "ddd d-mmm-yy h:mm am/pm")
    
        dUpdateAge = (Date + Time) - rs.Fields("last_updated").Value
        Me.btnLastUpdateColour.BackColor = GetRGBFromAge(dUpdateAge, 1#, 7#, 28#)
    
        Call rs.Close
    End If

cleanup_nicely:
    Application.EnableEvents = True
    m_bUpdatingMaterialNo_DoNotFire = False
    Application.ScreenUpdating = True

    '===========
    ' Display whichever pages is showing.
    '===========
    Call MultiPage1_Change
    
End Sub

'==============================================================================
' SUBROUTINE
'   MultiPage1_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the Multipage page changes.
' The strategy is to only load the high demand pages if they are being
' displayed.
' At the moment, these are the Material Part number and Picture pages.
'==============================================================================
Private Sub MultiPage1_Change()

    If Me.MultiPage1.SelectedItem.Tag = "1" Then
        '==========
        ' Page already displayed
        '==========
        Exit Sub
    End If
    
    If m_iMaterial < 9000000 Or m_iMaterial > 9999999 Then
        Exit Sub
    End If
        
    '==========
    ' Which page has appeared
    '==========
    Select Case Me.MultiPage1.SelectedItem.Name
        Case "PgSapDetail"
            '==========
            ' Everything already displayed
            '==========
            Call LoadMaterialDetail
            
        Case "PgColloquials"
            '==========
            ' Everything already displayed
            '==========
            Call DisplayColloquials
            
        Case "PgDescriptions"
            '==========
            ' Everything already displayed
            '==========
            Call DisplayDescriptions

        Case "PgPictures"
            '==========
            ' Display Pictures
            '==========
            'Call DisplayPic(m_iPictureIndex)
            Call LoadAndDisplayPics
            
        Case "PgMaterialPartNos"
            '==========
            ' Display Pictures
            '==========
            Call LoadMaterialPartNoList
            
    End Select
    
    Me.MultiPage1.SelectedItem.Tag = "1"
End Sub

'==============================================================================
' SUBROUTINE
'   LoadMaterialDetail
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Private Sub LoadMaterialDetail()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim dUpdateAge As Double
    Dim scStorageLocDescription As String
    Dim i As Long
    Dim dtOldestUpdate As Date
    
    '==========
    ' Is it a valid material number?
    '==========
    If Me.m_iMaterial < 9000000 Then
        Exit Sub
    End If
    
On Error GoTo cleanup_nicely
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    m_bUpdatingMaterialNo_DoNotFire = True
    
    '==========
    ' First thing we do is set the material number. Note that the change event
    ' should be suppressed, either via the EnableEvents=false, or the
    ' m_bUpdatingMaterialNo_DoNotFire flag.
    '==========
    'Me.txtDetailMaterialNo.Text = rs.Fields("pk_sap_material")
    'Me.txtDetailDescription.Text = rs.Fields("description")
    'Me.txtMaterialLongText.Text = rs.Fields("long_description")
    
    '==========
    ' Pull the material detail info out of the DB
    '==========
    If Not ConnectToDB(ldParts, cnn, True) Then
        Call MsgBox("Unable to connect to the database")
        Exit Sub
    End If
    
    Set rs = CreateObject("ADODB.Recordset")
    
    scSQLQuery = "SELECT * FROM parts.dbo.v_sap_material WHERE pk_sap_material = " & m_iMaterial
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    If rs.EOF Then
        txtMaterialLongText.Text = "Material not found"
        Exit Sub
    End If
    
    '==========
    ' Populate the various fields.
    '==========
    Me.txtBaseUnitOfMeas.Text = rs.Fields("uom")
    
    Me.txtMaterialGroup.Text = rs.Fields("fk_material_group")
    Me.lblMaterialGroupDescription.Caption = rs.Fields("material_group_description")
    Me.txtMaxStockLevel.Text = rs.Fields("max_stock_level")
    Me.txtMRPType.Text = rs.Fields("fk_mrp_type")
    Me.lblMRPType.Caption = rs.Fields("mrp_type")
    Me.txtPlannedDeliveryTime.Text = rs.Fields("lead_time_days")
    If Not IsNull(rs.Fields("fk_plant_specific_mat_status")) Then
        Me.txtPlantSpecificMaterialStatus.Text = Format(rs.Fields("fk_plant_specific_mat_status"), "00")
        Me.lblPlantSpecificMaterialStatus.Caption = rs.Fields("psms_description")
        
        Select Case Me.txtPlantSpecificMaterialStatus.Text
            Case "01"
                Me.txtPlantSpecificMaterialStatus.BackColor = RGB(200, 255, 200) ' Light Green
            Case "09"
                Me.txtPlantSpecificMaterialStatus.BackColor = RGB(255, 0, 0) ' Red
            Case Else
                Me.txtPlantSpecificMaterialStatus.BackColor = RGB(255, 255, 0) ' Yellow
        End Select
    Else
        Me.txtPlantSpecificMaterialStatus.Text = ""
        Me.lblPlantSpecificMaterialStatus.Caption = ""
    End If
    Me.txtReorderPoint.Text = rs.Fields("reorder_point")
    Me.txtStockLevel.Text = Format(rs.Fields("stock_level_total"), "#,##0.0")
    
    '==========
    ' Deleted
    '==========
'    If rs.Fields("deleted") Then
'        Me.btnDeleted.Caption = "Yes"
'        Me.btnDeleted.BackColor = RGB(255, 0, 0)
'    Else
'        Me.btnDeleted.Caption = "No"
'        Me.btnDeleted.BackColor = RGB(200, 255, 200)
'    End If
    
    '==========
    ' The last update date string and colour
    '==========
'    Me.txtLastUpdateFromSAP.Value = Format(rs.Fields("last_updated").Value, "ddd d-mmm-yy h:mm am/pm")
'
'    dUpdateAge = (Date + Time) - rs.Fields("last_updated").Value
'    Me.btnLastUpdateColour.BackColor = GetRGBFromAge(dUpdateAge, 1#, 7#, 28#)
    
    Call rs.Close
    
    '=============
    ' Now the storage bins
    '=============
    Call Me.lbStoreLocations.Clear
    
    scSQLQuery = "SELECT bin, bin_stock_level, plant, storage_location, storage_loc_description, last_update " & _
            "FROM parts.dbo.v_materials_with_bins WHERE pk_sap_material = " & m_iMaterial & " AND material_storage_bin_deleted = 0 ORDER BY storage_location, bin"
            
    
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    If rs.EOF Then
        Call Me.lbStoreLocations.AddItem("Material not found")
        Exit Sub
    End If
    
    While Not rs.EOF
        
        Call Me.lbStoreLocations.AddItem(rs.Fields("plant"))
        Me.lbStoreLocations.List(Me.lbStoreLocations.ListCount - 1, 1) = Format(rs.Fields("storage_location"), "0000")
        Me.lbStoreLocations.List(Me.lbStoreLocations.ListCount - 1, 2) = rs.Fields("storage_loc_description")
        Me.lbStoreLocations.List(Me.lbStoreLocations.ListCount - 1, 3) = rs.Fields("bin")
        Me.lbStoreLocations.List(Me.lbStoreLocations.ListCount - 1, 4) = rs.Fields("bin_stock_level")
        Me.lbStoreLocations.List(Me.lbStoreLocations.ListCount - 1, 5) = Format(rs.Fields("last_update"), "d-mmm")
    
        Call rs.MoveNext
    Wend
    Call rs.Close
    
cleanup_nicely:
    '===========
    ' Restore event handling
    '===========
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    m_bUpdatingMaterialNo_DoNotFire = False

End Sub


'==============================================================================
' SUBROUTINE
'   LoadMaterialPartNoList
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Private Sub LoadMaterialPartNoList()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim scUsername As String
    Dim scWhere As String
    Dim i As Long
    
    '==========
    ' Clear the list
    '==========
    Me.lbMaterialPartNos.Clear
    
    '==========
    ' Is it a valid material number?
    '==========
    If Me.m_iMaterial < 9000000 Then
        Exit Sub
    End If

    '==========
    ' Connect to the db
    '==========
    If Not ConnectToDB(ldParts, cnn, True) Then
        Call MsgBox("Unable to connect to the database")
        Exit Sub
    End If
    
    scSQLQuery = "SELECT top 50 pk_sap_material_part_no " & vbCrLf & _
        ",fk_sap_material,fk_sap_supplier,vendor_name, part_number " & vbCrLf & _
        ",description, simple_part_number " & vbCrLf & _
        "FROM parts.dbo.v_material_part_numbers WHERE fk_sap_material = " & m_iMaterial
        
    Set rs = CreateObject("ADODB.Recordset")
    
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    While Not rs.EOF
                
        Call Me.lbMaterialPartNos.AddItem(rs.Fields("pk_sap_material_part_no"))
        Me.lbMaterialPartNos.List(Me.lbMaterialPartNos.ListCount - 1, 1) = rs.Fields("fk_sap_supplier")
        Me.lbMaterialPartNos.List(Me.lbMaterialPartNos.ListCount - 1, 2) = rs.Fields("vendor_name")
        Me.lbMaterialPartNos.List(Me.lbMaterialPartNos.ListCount - 1, 3) = rs.Fields("part_number")
        Me.lbMaterialPartNos.List(Me.lbMaterialPartNos.ListCount - 1, 4) = rs.Fields("description")
                
        Call rs.MoveNext
    Wend
    
    '===========
    ' Clear the fields at the bottom
    '===========
    Me.txtMaterialPartNo.Text = ""
    Me.txtMPNPartNoDescription.Text = ""
    Me.txtMPNPartNumber.Text = ""
    Me.txtMPNVendor.Text = ""
    Me.txtMPNVendorDescription.Text = ""
    
End Sub

Private Sub lbMaterialPartNos_Click()
    If Me.lbMaterialPartNos.ListIndex < 0 Then
        '===========
        ' Clear the fields at the bottom
        '===========
        Me.txtMaterialPartNo.Text = ""
        Me.txtMPNPartNoDescription.Text = ""
        Me.txtMPNPartNumber.Text = ""
        Me.txtMPNVendor.Text = ""
        Me.txtMPNVendorDescription.Text = ""
    Else
        Me.txtMaterialPartNo.Text = Me.lbMaterialPartNos.List(Me.lbMaterialPartNos.ListIndex, 0)
        Me.txtMPNPartNoDescription.Text = Me.lbMaterialPartNos.List(Me.lbMaterialPartNos.ListIndex, 4)
        Me.txtMPNPartNumber.Text = Me.lbMaterialPartNos.List(Me.lbMaterialPartNos.ListIndex, 3)
        Me.txtMPNVendor.Text = Me.lbMaterialPartNos.List(Me.lbMaterialPartNos.ListIndex, 1)
        Me.txtMPNVendorDescription.Text = Me.lbMaterialPartNos.List(Me.lbMaterialPartNos.ListIndex, 2)
    End If
End Sub



'==============================================================================
' SUBROUTINE
'   DisplayColloquials
'------------------------------------------------------------------------------
' DESCRIPTION
'   Looks after the specifics associated with displaying the colloquials.
'==============================================================================
Public Sub DisplayColloquials()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim scUsername As String
    Dim scWhere As String
    Dim i As Long
    
    
On Error GoTo exit_cleanly
    
    Call lbColloquials.Clear
    lblColloquialComments.Caption = ""
    
    '==========
    ' Is it a valid material number?
    '==========
    If Me.m_iMaterial < 9000000 Then
        Exit Sub
    End If
    
    '==========
    ' Get the users username which we'll need later.
    '==========
    If Not GetUserName(scUsername) Then
        scUsername = Application.UserName
    End If
    
    '==========
    ' Connect to the db
    '==========
    If Not ConnectToDB(ldParts, cnn, True) Then
        Call MsgBox("Unable to connect to the database")
        Exit Sub
    End If
    
    '===========
    ' Construct the WHERE clause based on the filter
    '===========
    If Me.cbColloquialFilterSAP.Value Then ' And Me.cbColloquialFilterUsers.Value) Then
        scWhere = "group_type = 'SYSTEM'"
    End If
    If cbColloquialFilterUsers.Value Then
        If (Len(scWhere) > 0) Then scWhere = scWhere & " OR "
        
        scWhere = scWhere & " (group_type = 'USER'"
        If rdColloquialFilterMeOnly.Value Then
            scWhere = scWhere & " AND owner_username = '" & Replace(scUsername, "'", "''") & "'"
        End If
        scWhere = scWhere & ")"
    End If
    If Len(scWhere) > 0 Then
        scWhere = "(" & scWhere & ") AND "
    End If
    scWhere = scWhere & "(pk_sap_material = " & m_iMaterial & ")"

    scSQLQuery = "SELECT pk_map_colloquial_material, colloquial, group_type, owner_username, Rating FROM parts.dbo.v_material_colloquials WHERE " & _
            scWhere & " ORDER BY colloquial"
            
    If Not m_oSearchFormParent Is Nothing Then
        m_oSearchFormParent.txtQuery.Text = scSQLQuery
    End If
            
    Set rs = CreateObject("ADODB.Recordset")
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    If rs.EOF Then
        Call Me.lbColloquials.AddItem("<None found>")
    End If
    
    i = 0
    While Not rs.EOF And i < c_iMaxColloquials
        
        Call lbColloquials.AddItem(rs.Fields("colloquial"))
        If rs.Fields("Rating") = 0 Then
            lbColloquials.List(lbColloquials.ListCount - 1, 1) = "-"
        Else
            lbColloquials.List(lbColloquials.ListCount - 1, 1) = Format(rs.Fields("Rating"), "0.0")
        End If
        lbColloquials.List(lbColloquials.ListCount - 1, 2) = rs.Fields("owner_username")
        
        '===========
        ' Store the colloquial-material map primary key in the 4th column which is hidden.
        '===========
        lbColloquials.List(lbColloquials.ListCount - 1, 3) = rs.Fields("pk_map_colloquial_material")
    
        Call rs.MoveNext
        i = i + 1
    Wend
    Call rs.Close
    lblColloquialComments.Caption = "Found " & i & " colloquials for material " & m_iMaterial
    
    Exit Sub

exit_cleanly:
    lblColloquialComments.Caption = "Error: " & Err.description
    
End Sub


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   cbColloquialFilterSAP_Click
'==============================================================================
Private Sub cbColloquialFilterSAP_Click()
    Call Me.DisplayColloquials
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   cbColloquialFilterUsers_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Enables or disables the user radio buttons
'==============================================================================
Private Sub cbColloquialFilterUsers_Click()
    Call Me.DisplayColloquials
    
    frmColloquialFilterUsers.Enabled = Me.cbColloquialFilterUsers.Value
    Me.rdColloquialFilterMeOnly.Enabled = Me.cbColloquialFilterUsers.Value
    Me.rdColloquialFilterEveryone.Enabled = Me.cbColloquialFilterUsers.Value
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   rdColloquialFilterEveryone_Click
'==============================================================================
Private Sub rdColloquialFilterEveryone_Click()
    Call Me.DisplayColloquials
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   rdColloquialFilterEveryone_Click
'==============================================================================
Private Sub rdColloquialFilterMeOnly_Click()
    Call Me.DisplayColloquials
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbColloquials_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Handler for the Click event on the lbColloquials Listbox.
' The handler displays the rating details of that colloquial.
'==============================================================================
Private Sub lbColloquials_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    
    Dim scColloquial As String
    Dim scRating As String
    Dim iPkMaterialColloquial As Long
    Dim scComment As String

    If lbColloquials.ListIndex < 0 Then Exit Sub
    
    scColloquial = lbColloquials.List(lbColloquials.ListIndex, 0)
    scRating = lbColloquials.List(lbColloquials.ListIndex, 1)
    iPkMaterialColloquial = Val(lbColloquials.List(lbColloquials.ListIndex, 3))
    
    Call lbMaterialColloquialRatings.Clear
    lblRatingAuthor.Caption = ""
    lblRatingScore.Caption = ""

    '==========
    ' Connect to the db
    '==========
    If Not ConnectToDB(ldParts, cnn, True) Then
        Call MsgBox("Unable to connect to the database")
        Exit Sub
    End If
    
    Set rs = CreateObject("ADODB.Recordset")
    
    scSQLQuery = "SELECT * FROM parts.dbo.t_rating " & vbCrLf & _
            "WHERE category = 'MATERIAL_COLLOQUIAL' AND fk_rated_item = " & iPkMaterialColloquial
    Me.DebugDisplaySQLQuery (scSQLQuery)
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)

    
    While Not rs.EOF
        Call lbMaterialColloquialRatings.AddItem(rs.Fields("author"))
        lbMaterialColloquialRatings.List(lbMaterialColloquialRatings.ListCount - 1, 1) = rs.Fields("rating_out_of_5")
        lbMaterialColloquialRatings.List(lbMaterialColloquialRatings.ListCount - 1, 2) = rs.Fields("pk_rating")
        scComment = rs.Fields("comment")
        If Len(scComment) = 0 Then
            scComment = "-"
        End If
        lbMaterialColloquialRatings.List(lbMaterialColloquialRatings.ListCount - 1, 3) = scComment
    
        Call rs.MoveNext
    Wend

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   CheckBoxX_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   The next 5 methods are the handlers for the 5 star rating checkboxes.
' The just call the SetStarRating method which displays appropriately to the
' desired rating.
'   It can get into a cyclic event loop though, so the member variable
' m_bIgnoreStarClicks is used to manage this.
'==============================================================================
Private Sub btnRatingsManager_Click()

    Dim ufMgr As ufRatingsManager
    Dim scColloquial As String
    Dim iPkColloquial As Long
    
    If lbColloquials.ListIndex < 0 Then
        Exit Sub
    End If
    
    scColloquial = lbColloquials.List(lbColloquials.ListIndex, 0)
    iPkColloquial = lbColloquials.List(lbColloquials.ListIndex, 3)
    
    Set ufMgr = New ufRatingsManager
    
    Call ufMgr.Initialise("Colloquial " & m_iMaterial & ": " & scColloquial, "MATERIAL_COLLOQUIAL", iPkColloquial, ldParts)
    
    Call ufMgr.Show
    
End Sub



'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   CheckBoxX_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   The next 5 methods are the handlers for the 5 star rating checkboxes.
' The just call the SetStarRating method which displays appropriately to the
' desired rating.
'   It can get into a cyclic event loop though, so the member variable
' m_bIgnoreStarClicks is used to manage this.
'==============================================================================
Private Sub CheckBox1_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(1)
End Sub

Private Sub CheckBox2_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(2)
End Sub

Private Sub CheckBox3_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(3)
End Sub

Private Sub CheckBox4_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(4)
End Sub

Private Sub CheckBox5_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(5)
End Sub

'==============================================================================
' FUNCTION
'   StarText
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the text associated with the star rating.
' Note that the method will also adjust the iStar integer if it's illegal.
'==============================================================================
Public Function StarText(ByRef iStar As Long)

    If iStar < 1 Then
        iStar = 1
    ElseIf iStar > 5 Then
        iStar = 5
    End If
    
    Select Case iStar
        Case 1
            StarText = "Terrible"
        Case 2
            StarText = "Poor"
        Case 3
            StarText = "Average"
        Case 4
            StarText = "Good"
        Case 5
            StarText = "Excellent"
    End Select
    
End Function

'==============================================================================
' SUBROUTINE
'   SetStarRating
'------------------------------------------------------------------------------
' DESCRIPTION
'   Sets the star rating by appropriately clearing and setting the checkbox's
' and displaying the relevant text
'==============================================================================
Private Sub SetStarRating(iStar As Long)
    Dim i As Long
    
On Error GoTo cleanup_nicely
    m_bIgnoreStarClicks = True
    
    i = 1
    For i = 1 To iStar
        m_cbaColloquialStarsCheckBoxs(i).Value = True
    Next
    For i = (iStar + 1) To 5
        m_cbaColloquialStarsCheckBoxs(i).Value = False
    Next
    
    Me.lblStarsRating.Caption = StarText(iStar)
    
cleanup_nicely:
    m_bIgnoreStarClicks = False
        
End Sub

'==============================================================================
' FUNCTION
'   GetStarRating
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the star rating based on the highest checked checkbox.
'==============================================================================
Private Function GetStarRating() As Long
    Dim i As Long

    For i = 5 To 1 Step -1
        If m_cbaColloquialStarsCheckBoxs(i).Value Then
            Exit For
        End If
    Next
    
    GetStarRating = i

End Function

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnRateColloquial_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Handler for the RATE button.
' It checks to see if the user has already put in a rating for the selected
' colloquial. If so, it gives the user the option to
'==============================================================================
Private Sub btnRateColloquial_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    
    Dim scColloquial As String
    Dim scUsername As String
    Dim iPkColloquialMaterialMap As Long
    Dim iPkRating As Long
    Dim iRating As Long
    Dim scRatingComment As String
    
    If Me.lbColloquials.ListIndex < 0 Then
        Call MsgBox("No colloquial selected")
        Exit Sub
    End If
    
    If Not GetUserName(scUsername) Then
        scUsername = Application.UserName
    End If
    
    scColloquial = Me.lbColloquials.List(Me.lbColloquials.ListIndex, 0)
    iPkColloquialMaterialMap = Val(Me.lbColloquials.List(Me.lbColloquials.ListIndex, 3))
    
    iRating = GetStarRating
    If Len(Trim(Replace(Me.txtMyRatingComments, "'", "''"))) > 0 Then
        scRatingComment = Replace(Trim(Me.txtMyRatingComments), "'", "''")
    Else
        scRatingComment = "-"
    End If
        
    
    '==========
    ' Connect to the db
    '==========
    If Not ConnectToDB(ldParts, cnn, True) Then
        Call MsgBox("Unable to connect to the database")
        Exit Sub
    End If
    
    Set rs = CreateObject("ADODB.Recordset")
    
    '==========
    ' Has this user already provided a rating?
    '==========
    scSQLQuery = "SELECT * FROM parts.dbo.t_rating WHERE fk_rated_item = " & m_iMaterial & " AND category = 'MATERIAL_COLLOQUIAL' AND author = '" & Replace(scUsername, "'", "''") & "'"
    Call Me.DebugDisplaySQLQuery(scSQLQuery)
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    If Not rs.EOF Then
    
        Dim iOldRating As Long
        Dim scOldComment As String
        
        '==========
        ' The user has already rated this colloquial.
        ' We tell them and give them the option to update the rating with
        ' the new entry
        '==========
        iOldRating = rs.Fields("rating_out_of_5")
        scOldComment = rs.Fields("comment")
        iPkRating = rs.Fields("pk_rating")
        Call rs.Close
        If (MsgBox("You already have set a rating of " & iOldRating & " for this colloquial. Do you want to update the rating?", vbYesNo) = vbYes) Then
            scSQLQuery = "UPDATE parts.dbo.t_rating SET rating_out_of_5 = " & iRating & _
                ", comment = '" & scRatingComment & "' WHERE pk_rating = " & iPkRating
            Call cnn.Execute(scSQLQuery)
        Else
            'Call SetStarRating(iOldRating)
            'txtMyRatingComments.Text = scOldComment
        End If
    Else
        Call rs.Close
        scSQLQuery = "INSERT INTO parts.dbo.t_rating " & _
            "(rating_out_of_5, fk_rated_item, category, comment, author) " & _
            "VALUES " & _
            "(" & iRating & ", " & iPkColloquialMaterialMap & ", 'MATERIAL_COLLOQUIAL', '" & _
                scRatingComment & "', '" & Replace(scUsername, "'", "''") & "')"
        Call Me.DebugDisplaySQLQuery(scSQLQuery)
        Call cnn.Execute(scSQLQuery)
        lblColloquialComments.Caption = "Material Colloquial '" & scColloquial & " successfully rated!"
    End If
    
    Call DisplayColloquials
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbMaterialColloquialRatings_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Handles for when the user clicks on the user ratings list box for the
' colloquials.
'==============================================================================
Private Sub lbMaterialColloquialRatings_Click()
    Dim scAuthor As String
    Dim iRating As Long
    Dim iPkRating As Long
    Dim scComment As String
    
    If lbMaterialColloquialRatings.ListIndex < 0 Then Exit Sub
    
    scAuthor = lbMaterialColloquialRatings.List(lbMaterialColloquialRatings.ListIndex, 0)
    iRating = Val(lbMaterialColloquialRatings.List(lbMaterialColloquialRatings.ListIndex, 1))
    iPkRating = Val(lbMaterialColloquialRatings.List(lbMaterialColloquialRatings.ListIndex, 2))
    
    scComment = lbMaterialColloquialRatings.List(lbMaterialColloquialRatings.ListIndex, 3)
    
    lblRatingAuthor.Caption = scAuthor
    lblRatingScore.Caption = iRating & " " & Me.StarText(iRating)
    txtMaterialColloquialRatingComment.Text = scComment
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtDetailMaterialNo_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Handles when the user enters a material number manually.
'==============================================================================
Private Sub txtDetailMaterialNo_Change()

    If m_bUpdatingMaterialNo_DoNotFire Then
        Exit Sub
    End If
    
    '==========
    ' The user has changed the material number. Hence clear display to
    ' avoid confusion, but only if it's not already clear.
    '==========
    If Me.txtDetailDescription.Text <> "" Then
        Call ClearDisplay(False, False)
    End If
    
    
On Error GoTo not_a_valid_material_no
    
    m_iMaterial = Val(Me.txtDetailMaterialNo.Text)
    
    GoTo process_material
    
not_a_valid_material_no:
    m_iMaterial = -1
    
process_material:
    If m_iMaterial < 9000000 Or m_iMaterial > 9999999 Then
        m_iMaterial = -1
        Exit Sub
    End If
    
    '============
    ' We appear to have a valid material number. So display it.
    '============
    Call DisplayMaterial(m_iMaterial)
    
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
'   cmdAddLihirColloquial_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Event Handler for the
'==============================================================================
Private Sub cmdAddLihirColloquial_Click()
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    
    Dim scColloquial As String
    Dim iPkColloquial As Long
    Dim bNewMaterialColloquial As Boolean
    Dim scUsername As String
    Dim scUserGroupName As String
    Dim iPkGroup As Long
    
    scColloquial = Trim(txtLihirColloquialsNew.Text)
    
    If Len(Trim(scColloquial)) < 4 Then
        Call MsgBox("All colloquials need to be atleast 4 characters to try and ensure some level of uniqueness")
        Exit Sub
    End If
    
    '==========
    ' Connect to the db
    '==========
    If Not ConnectToDB(ldParts, cnn, True) Then
        Call MsgBox("Unable to connect to the database")
        Exit Sub
    End If
    
    Set rs = CreateObject("ADODB.Recordset")
    
    '==========
    ' Does the colloquial already exist?
    '==========
    scSQLQuery = "SELECT pk_colloquial FROM [parts].[dbo].[t_colloquial] " & vbCrLf & _
        "WHERE colloquial = '" & Replace(txtLihirColloquialsNew.Text, "'", "''") & "'"
        
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    If rs.EOF Then
        Call rs.Close
        '==========
        ' Then it is a new colloquial. Hence it must be a new colloquial for
        ' the material.
        '==========
        bNewMaterialColloquial = True
        scSQLQuery = "SET NOCOUNT ON; INSERT INTO [parts].[dbo].[t_colloquial] " & vbCrLf & _
            "(colloquial) VALUES ('" & Replace(txtLihirColloquialsNew.Text, "'", "''") & "');" & _
            "SELECT SCOPE_IDENTITY() as pk_colloquial;"
            
        Set rs = cnn.Execute(scSQLQuery)
        iPkColloquial = rs.Fields("pk_colloquial")
        Call rs.Close
    Else
        iPkColloquial = rs.Fields("pk_colloquial")
        Call rs.Close
        '==========
        ' Does the Material Colloquial mapping exist?
        '==========
        scSQLQuery = "SELECT * FROM [parts].[dbo].[t_map_colloquial_material] " & vbCrLf & _
                "WHERE fk_sap_material = " & m_iMaterial & " AND fk_colloquial = " & iPkColloquial
                
        Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
        If rs.EOF Then
            bNewMaterialColloquial = True
        Else
            Call MsgBox("This colloquial is already assigned to material " & m_iMaterial & "!")
            bNewMaterialColloquial = False
        End If
        Call rs.Close
    End If
    
    If bNewMaterialColloquial Then
        '==========
        ' Create the new mapping.
        '==========
        '==========
        ' First we need to get the users group ID
        '==========
        If Not (GetUserName(scUsername)) Then
            scUsername = Application.UserName
        End If
        scUserGroupName = Application.UserName
        
        '==========
        ' Make the username string SQL query ready
        '==========
        scUsername = Replace(scUsername, "'", "''")
        scUserGroupName = Replace(scUserGroupName, "'", "''")
        '==========
        ' Construct and execute the query to find the users group ID
        '==========
        scSQLQuery = "SELECT * FROM parts.dbo.t_group " & vbCrLf & _
                "WHERE group_type = 'USER' AND owner_username = '" & scUsername & "'"
        Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
        If rs.EOF Then
            '==========
            ' This Users group doesn't exist, so we create it.
            '==========
            Call rs.Close
            scSQLQuery = "SET NOCOUNT ON; " & vbCrLf & _
                "INSERT INTO parts.dbo.t_group " & vbCrLf & _
                "(name, owner_username, group_type) VALUES " & vbCrLf & _
                "('" & scUserGroupName & "', '" & scUsername & "', 'USER');" & vbCrLf & _
                "SELECT SCOPE_IDENTITY() as pk_group"
                
            Set rs = cnn.Execute(scSQLQuery)
            iPkGroup = rs.Fields("pk_group")
            Call rs.Close
        Else
            iPkGroup = rs.Fields("pk_group")
            Call rs.Close
        End If
        
        '==========
        ' Now we create the new Colloquial-Material Map entry
        '==========
        scSQLQuery = "INSERT INTO [parts].[dbo].[t_map_colloquial_material] " & vbCrLf & _
                "(fk_sap_material, fk_colloquial, fk_group, reconcile_flag) " & vbCrLf & _
                "VALUES " & vbCrLf & _
                "(" & m_iMaterial & ", " & iPkColloquial & ", " & iPkGroup & ", 0)"
        Call cnn.Execute(scSQLQuery)
        Me.lblColloquialComments.Caption = "Successfully created colloquial-material mapping."
        
        Call Me.DisplayColloquials
    End If

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnAddPicture_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Private Sub btnAddPicture_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim oFileDialog As Office.FileDialog
    Dim ofile As Scripting.File
#Else
    Dim cnn As Object
    Dim rs As Object
    Dim oFileDialog As Object
    Dim ofile As Object
#End If

    Dim iResult As Long
    Dim scFromFileNameFull As String
    Dim iFileID As Long
    Dim scUser As String
    
    Dim scSQLQuery As String
    Dim scDescription As String
    
    Set oFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    Call GetUserName(scUser)

    With oFileDialog
        .AllowMultiSelect = False
        .Filters.Clear
        Call .Filters.Add("Jpeg Files", "*.jpg")
        Call .Filters.Add("Bitmap Files", "*.bmp")
        Call .Filters.Add("All Files", "*.*")
        iResult = .Show
    
        If iResult <> 0 Then
            scFromFileNameFull = .SelectedItems.Item(1)
        Else
            Exit Sub
        End If
            
        '==========
        ' Get a description
        '==========
        scDescription = InputBox("Please provide a brief description of the file contents (or click Cancel ... you can edit later).", "File Description")
        
        '==========
        ' Add the file to the database and get back the ID
        '==========
        If Not DBAddFile(scFromFileNameFull, "MATERIAL_PICTURES", scUser, scDescription, iFileID) Then
            Exit Sub
        End If
                
        Call DBGetFile(iFileID, ofile)
        
        Me.Image1.Picture = LoadPicture(ofile.Path)
        
        '========
        ' Connect to the DB
        '========
        If Not (ConnectToDB(ldMaintenance, cnn, True)) Then
            Call MsgBox("Unable to connect to the DB")
            Exit Sub
        End If
        
        '==========
        ' Add an entry to the mapping table
        '==========
        scSQLQuery = "INSERT into maint.dbo.t_file_map " & _
                "(fk_file, mapped_obj_category, fk_mapped_obj_int) VALUES " & _
                "(" & iFileID & ", 'MATERIAL_PICTURES', " & Me.txtDetailMaterialNo.Text & ")"
            
        Call cnn.Execute(scSQLQuery)
        
        '==========
        ' Finally force the form to redisplay it's data
        '==========
        Call Me.DisplayMaterial(Val(Me.txtDetailMaterialNo.Text))
    
    End With
End Sub

'==============================================================================
' SUBROUTINE
'   LoadAndDisplayPics
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Private Sub LoadAndDisplayPics()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

#If DebugBadType = 0 Then
    Dim cPicData As clsPictureData
#Else
    Dim cPicData As Object
#End If

    Dim scSQLQuery As String
    
On Error GoTo cleanup_nicely

    '============
    ' Clear the collection in case the user is calling this method to refresh
    ' the data
    '============
    Call Me.ClearPicCollection
    
    '============
    ' Any Pictures in the DB?
    '============
    
    scSQLQuery = "SELECT * FROM maint.dbo.v_mapped_files WHERE id = '" & m_iMaterial & "' AND base_path_category = 'MATERIAL_PICTURES'"
    
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)
  
    '========
    ' Cycle through the results, filling out the records in the pseudo
    ' collection implemented on this userform.
    '========
    While Not rs.EOF
        Set cPicData = New clsPictureData
        
        Call cPicData.PopulateFromRecordset(rs)
        
'        oFileInfo.FilePath = rs.Fields("base_path") & rs.Fields("relative_path")
'        oFileInfo.FileName = rs.Fields("relative_path")
'        If IsNull(rs.Fields("file_date")) Then
'            oFileInfo.filedate = 0
'        Else
'            oFileInfo.filedate = rs.Fields("file_date")
'        End If
'        If Not IsNull(rs.Fields("map_description")) Then
'            oFileInfo.description = rs.Fields("map_description")
'        Else
'            oFileInfo.description = ""
''        End If
        Call m_colPicData.Add(cPicData, cPicData.Key)
        
        Call rs.MoveNext
    Wend
    
    '============
    ' Enable/Disable the Previous/Next buttons as required.
    '============
    m_iPictureIndex = 1
    btnPrevPicture.Enabled = False
    
    If m_colPicData.Count > 1 Then
        btnNextPicture.Enabled = True
    Else
        btnNextPicture.Enabled = False
    End If
        
    '============
    ' Display the picture if there is one. Otherwise show blank.
    '============
    If m_colPicData.Count > 0 Then
        Call DisplayPic(1)
    Else
        Me.Image1.Picture = LoadPicture("")
        lblPictureCountPosition.Caption = "No Pics"
    End If
    
cleanup_nicely:
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnPrevPicture_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Private Sub btnPrevPicture_Click()
    
    If m_iPictureIndex > 1 Then
        m_iPictureIndex = m_iPictureIndex - 1
        
        btnNextPicture.Enabled = True
    
        If m_iPictureIndex = 1 Then
            btnPrevPicture.Enabled = False
        End If
        
        Call DisplayPic(m_iPictureIndex)
    Else
        '============
         ' This shouldn't occur as the button should be enabled if the index
         ' is 0. But we manage it anyway.
        '============
        btnPrevPicture.Enabled = False
        Call DisplayPic(-1)
    End If

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnNextPicture_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Private Sub btnNextPicture_Click()
    
    If m_iPictureIndex < m_colPicData.Count Then 'Me.PicCount > (m_iPictureIndex + 1) Then
    
        m_iPictureIndex = m_iPictureIndex + 1
        btnPrevPicture.Enabled = True
        
        If m_iPictureIndex = m_colPicData.Count Then
            btnNextPicture.Enabled = False
        End If
        
        Call DisplayPic(m_iPictureIndex)
        
    Else
        '============
         ' This shouldn't occur as the button should be enabled if the index
         ' at max. But we manage it anyway.
        '============
        btnNextPicture.Enabled = False
        Call DisplayPic(-1)
    End If
End Sub

'==============================================================================
' SUBROUTINE
'   DisplayPic
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Private Sub DisplayPic(iIndex As Long)
    
#If DebugBadType = 0 Then
    Dim cPicData As clsPictureData
#Else
    Dim cPicData As Object
#End If

    If (iIndex < 1) Or (iIndex > m_colPicData.Count) Then
        lblPictureCountPosition.Caption = ""
        lblPictureInfo.Caption = ""
        lblPictureDate.Caption = ""
        Me.Image1.Picture = LoadPicture("")
    Else
    
        Set cPicData = m_colPicData(iIndex)
        
        Me.Image1.Picture = LoadPicture(cPicData.FullPath)
            
        lblPictureCountPosition.Caption = iIndex & " of " & m_colPicData.Count
'        lblPictureInfo.Caption = Mid(cPicData.relative_path, 6, 200) & vbCrLf & _
'                "Date taken: " & Format(cPicData.file_date, "d-mmm-yy hh:mm")
        lblPictureInfo.Caption = cPicData.description
        lblPictureDate.Caption = Format(cPicData.file_date, "d-mmm-yyyy")
        lblPictureDate.ForeColor = RGB(180, 180, 0)
    
    End If
    
End Sub

Private Sub CopyFileAndOpen(iIndex As Long)

#If DebugBadType = 0 Then
    Dim cPicData As clsPictureData
#Else
    Dim cPicData As Object
#End If

    If (iIndex < 1) Or (iIndex > m_colPicData.Count) Then
        Exit Sub
    End If
    
    Set cPicData = m_colPicData(iIndex)

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
' SUBROUTINE
'   ClearPicCollection
'------------------------------------------------------------------------------
' DESCRIPTION
'   Resets the counter to indicate there are no pictures.
'==============================================================================
Public Sub ClearPicCollection()

    Set m_colPicData = Nothing
    Set m_colPicData = New VBA.Collection

    m_iPictureIndex = 1
End Sub

'==============================================================================
' SUBROUTINE
'   btnManagePictures_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Adds a custom user description
'==============================================================================
Private Sub btnManagePictures_Click()

    Dim ufPictMgr As ufPictureManager
    
    Set ufPictMgr = New ufPictureManager
    
    Call ufPictMgr.Configure("MATERIAL_PICTURES", CStr(Me.m_iMaterial))
    
    Call ufPictMgr.Show
    
    If Not (ufPictMgr Is Nothing) Then
        If ufPictMgr.DirtyFlag Then
            '===========
            ' Refresh the page. We do this by setting the Tag to "0" indicating
            ' it hasn't been loaded.
            '===========
            Me.MultiPage1.SelectedItem.Tag = "0"
            Call MultiPage1_Change
        End If
    End If
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnAddUpdateUserDescription_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Adds a custom user description
'==============================================================================
Private Sub btnAddUpdateUserDescription_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim scUsername As String
    Dim pk_user_description As Long

    '============
    ' Connect to the DB
    '============
    If Not (ConnectToDB(ldMaintenance, cnn, True)) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Sub
    End If

    Set rs = CreateObject("ADODB.Recordset")

    '============
    ' Does the user already have an entry.
    '============
    '==========
    ' Get the users username which we'll need later.
    '==========
    If Not GetUserName(scUsername) Then
        scUsername = Application.UserName
    End If

    scSQLQuery = "SELECT * FROM parts.dbo.t_user_description " & _
        "WHERE fk_sap_material = " & Me.m_iMaterial & " AND user_id = '" & Replace(scUsername, "'", "''") & "'"

    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    If Not rs.EOF Then

        If MsgBox("Update your existing entry?", vbYesNo) = vbNo Then
            Exit Sub
        End If

        pk_user_description = rs.Fields("pk_user_description")

        Call rs.Close
        scSQLQuery = "UPDATE parts.dbo.t_user_description SET short_text = '" & Replace(txtNewUserDescription.Text, "'", "''") & "', " & _
            "long_text = '" & Replace(Me.txtNewUserLongText.Text, "'", "''") & "' " & _
            "WHERE pk_user_description = " & pk_user_description

        Call cnn.Execute(scSQLQuery)
    Else
        Call rs.Close

        scSQLQuery = "INSERT INTO parts.dbo.t_user_description " & _
            "(fk_sap_material, user_id, short_text, long_text) VALUES (" & _
            Me.m_iMaterial & _
            ", '" & Replace(scUsername, "'", "''") & "', '" & _
            Replace(txtNewUserDescription.Text, "'", "''") & "', '" & _
            Replace(txtNewUserLongText.Text, "'", "''") & "')"

        Call cnn.Execute(scSQLQuery)
    End If
    
    '=============
    ' Display the new entry
    '=============
    Call DisplayDescriptions

End Sub

'==============================================================================
' SUBROUTINE
'   ClearDescriptions
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clears the whole MultPage page
'==============================================================================
Sub ClearDescriptions()
    Call lbUserDescriptions.Clear

    txtUserDescription.Text = ""
    txtNewUserDescription.Text = ""

    txtUserLongText.Text = ""
    txtNewUserLongText.Text = ""
    lblUserDescriptionAuthor.Caption = ""

End Sub


'==============================================================================
' SUBROUTINE
'   DisplayDescriptions
'------------------------------------------------------------------------------
' DESCRIPTION
'   Displays the contents of the Description Multipage
'==============================================================================
Sub DisplayDescriptions()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim scUsername As String
    Dim pk_user_description As Long
    
    '============
    ' Clear the list
    '============
    Call lbUserDescriptions.Clear
    
    '============
    ' Does the user already have an entry.
    '============
    '==========
    ' Get the users username which we'll need later.
    '==========
    If Not GetUserName(scUsername) Then
        scUsername = Application.UserName
    End If
    
    scSQLQuery = "SELECT * FROM parts.dbo.t_user_description " & _
        "WHERE fk_sap_material = " & Me.m_iMaterial

    If Not GetDBRecordSet(ldParts, cnn, scSQLQuery, rs) Then
        Exit Sub
    End If

    '==========
    ' Populate the list
    '==========
    While Not rs.EOF
        Call Me.lbUserDescriptions.AddItem(rs.Fields("short_text"))
        Me.lbUserDescriptions.List(Me.lbUserDescriptions.ListCount - 1, 1) = rs.Fields("user_id")
        Me.lbUserDescriptions.List(Me.lbUserDescriptions.ListCount - 1, 3) = rs.Fields("long_text")
        
        If scUsername = rs.Fields("user_id") Then
            Me.txtNewUserDescription.Text = rs.Fields("short_text")
            Me.txtNewUserLongText.Text = rs.Fields("long_text")
        End If
        
        Call rs.MoveNext
    Wend
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbUserDescriptions_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub lbUserDescriptions_Click()
    txtUserDescription.Text = lbUserDescriptions.List(lbUserDescriptions.ListIndex, 0)
    lblUserDescriptionAuthor.Caption = lbUserDescriptions.List(lbUserDescriptions.ListIndex, 1)
    txtUserLongText.Text = lbUserDescriptions.List(lbUserDescriptions.ListIndex, 3)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtNewUserDescription_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Displays the contents of the Description Multipage
'==============================================================================
Private Sub txtNewUserDescription_Change()
    If Len(txtNewUserDescription) = 0 Then
        lblNewDescriptionLength.Caption = "-"
    Else
        lblNewDescriptionLength.Caption = Len(txtNewUserDescription)
    End If
    
    If Len(txtNewUserDescription) > 40 Then
        txtNewUserDescription.BackColor = RGB(255, 0, 0)
    Else
        txtNewUserDescription.BackColor = RGB(255, 255, 255)
    End If
    

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnSearchPartsUsage_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnSearchPartsUsage_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim dSpend As Double
    Dim scCurrFormat As String
    
    '===========
    ' Blank the list display
    '===========
    Call lbUsageSearchResults.Clear
    
    '===========
    ' Construct the query
    '===========
    scSQLQuery = "SELECT * FROM dbo.v_materials_usage WHERE fk_material = " & Me.m_iMaterial
    If cbPartsUsageOnlyWorkOrders.Value Then
        scSQLQuery = scSQLQuery & " AND (NOT fk_order is NULL)"
    End If
    scSQLQuery = scSQLQuery & " ORDER BY posting_date DESC"
    
    '===========
    ' Get the list of records
    '===========
    Call GetDBRecordSet(ldFinance, cnn, scSQLQuery, rs)
    
    While Not rs.EOF
        Call lbUsageSearchResults.AddItem(rs.Fields("total_qty"))
        
        If Not IsNull(rs.Fields("posted_uom")) Then
            lbUsageSearchResults.List(lbUsageSearchResults.ListCount - 1, 1) = rs.Fields("posted_uom")
        End If
        
        '============
        ' Spend. Format without $ fraction if >= $10k
        '============
        lbUsageSearchResults.List(lbUsageSearchResults.ListCount - 1, 2) = rs.Fields("trans_currency")
        
        dSpend = rs.Fields("value_in_rep_currency")
        
        Select Case rs.Fields("trans_currency")
            Case "USD", "NZD", "AUD"
                scCurrFormat = "$#,##0"
            Case "PGK"
                scCurrFormat = "K#,##0"
            Case "EUR"
                scCurrFormat = CStr(Chr(128)) & "#,##0"
            Case Else
                scCurrFormat = "#,##0"
        End Select
        
        If Abs(dSpend) < 10000 Then
            scCurrFormat = scCurrFormat & ".00"
        End If
        lbUsageSearchResults.List(lbUsageSearchResults.ListCount - 1, 3) = Format(rs.Fields("value_in_trans_currency"), scCurrFormat)
        
        lbUsageSearchResults.List(lbUsageSearchResults.ListCount - 1, 4) = Format(rs.Fields("posting_date"), "d-mmm-yy")
        
        '============
        ' Work Order and functional location. Check for nulls.
        '============
        If Not IsNull(rs.Fields("fk_order")) Then
            lbUsageSearchResults.List(lbUsageSearchResults.ListCount - 1, 5) = rs.Fields("fk_order")
        Else
            lbUsageSearchResults.List(lbUsageSearchResults.ListCount - 1, 5) = "-"
        End If
        If Not IsNull(rs.Fields("fk_func_loc")) Then
            lbUsageSearchResults.List(lbUsageSearchResults.ListCount - 1, 6) = rs.Fields("fk_func_loc")
        Else
            lbUsageSearchResults.List(lbUsageSearchResults.ListCount - 1, 6) = "-"
        End If
    
        Call rs.MoveNext
    Wend
    
    If lbUsageSearchResults.ListCount > 0 Then
        lblPartsUsageSearchResults.Caption = "Found " & lbUsageSearchResults.ListCount & " transactions for this material."
    Else
        lblPartsUsageSearchResults.Caption = "No transactions for this material in the finance database."
    End If
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnSearchPartsUsage_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnOpenPicture_Click()

#If DevelopMode = 1 Then
    Dim oDestFile As Scripting.File
#Else
    Dim oDestFile As Object
#End If

#If DebugBadType = 0 Then
    Dim cPicData As clsPictureData
#Else
    Dim cPicData As Object
#End If

    Dim scTempFolderPath As String

    If m_colPicData.Count < 1 Then
        Exit Sub
    End If
    
    scTempFolderPath = GetSpecialFolderPath(sfTemp)
    
    Set cPicData = m_colPicData(m_iPictureIndex)
    
    If DBCopyFile(cPicData.pk_file, scTempFolderPath, oDestFile) Then
    
        '===========
        ' Open the file in its default application.
        '===========
        Dim iShell As Object
        Set iShell = CreateObject("Shell.Application")
        
        Call iShell.Open(oDestFile.Path)
    End If

End Sub


