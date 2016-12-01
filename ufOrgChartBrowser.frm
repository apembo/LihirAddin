VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufOrgChartBrowser 
   Caption         =   "Org Chart Browser"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12615
   OleObjectBlob   =   "ufOrgChartBrowser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufOrgChartBrowser"
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
' PRIVATE TYPE(S)
'==============================================================================

'==============================================================================
' PRIVATE MEMBER VARIABLES
'==============================================================================
Private m_iColPositionTable_pk_staff_position_sap As Long
Private m_iColPositionTable_job_title As Long
Private m_iColPositionTable_fk_parent_position_sap As Long
Private m_iColPositionTable_active As Long
Private m_iColPositionTable_level_no As Long

Private m_colPictures As VBA.Collection
Private m_bCanEditPictures As Boolean
Private m_iSelectedEncumbentIndex As Long
Private m_iSelectedRole As Long

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public m_colEncumbents As VBA.Collection

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub UserForm_Initialize()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim scUsername As String
    
    '=============
    ' Check the access level for this user.
    '=============
    Call GetUserName(scUsername)
    
    m_bCanEditPictures = HasPermissions(scUsername, PeoplePhotoEditing)
    btnManagePictures.Visible = m_bCanEditPictures
    
    m_iSelectedEncumbentIndex = -1
    
    Call InitializeTreeView
    
End Sub

'==============================================================================
' SUBROUTINE
'   InitializeTreeView
'------------------------------------------------------------------------------
' DESCRIPTION
'   Initializes the userform, adds the VBA treeview to the container frame on
' the userform and populates the treeview.
'==============================================================================
Private Sub InitializeTreeView()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim oNode As MSComctlLib.Node
#Else
    Dim cnn As Object
    Dim rs As Object
    Dim oNode As Object
#End If

    Dim scSQLQuery As String
    Dim iLevelNo As Long
    Dim bValuesReturned As Boolean
    
    Call TreeView1.Nodes.Clear
    
    Call ClearRoleDisplay
    Call ClearPositionDisplay
    Call ClearPersonDisplay
    
    '============
    ' Initialise some TreeView parameters
    '============
'    Me.TreeView1.Style = tvwTreelinesPlusMinusPictureText
'    Me.TreeView1.LineStyle = tvwRootLines
'    Me.TreeView1.Indentation = 10
'
'    Me.TreeView1.SingleSel = True
    
    '============
    ' Create the collection that will store node data.
    '============

    bValuesReturned = True
    iLevelNo = 1
    
    scSQLQuery = "SELECT pk_staff_role, role_name, pk_parent_role, pk_section, pk_department, level_no " & _
        "FROM people.dbo.v_staff_roles order by level_no"
    
    Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)
        
        
    '============
    ' Work through the returned entries.
    '============
    While Not rs.EOF
        Dim scKey As String
        Dim scParentKey As String
        Dim scRoleName As String
        
On Error GoTo report_error
        
        scKey = "N" & rs.Fields("pk_staff_role")
        scRoleName = rs.Fields("role_name")
        iLevelNo = rs.Fields("level_no")
        
        '=============
        ' DEBUG
        '=============
'        If rs.Fields("pk_staff_role") = 1242 Then
'            Call MsgBox("Here I am!")
'        End If
        '=============
        ' END DEBUG
        '=============
    
        If iLevelNo = 1 Then
            Set oNode = Me.TreeView1.Nodes.Add(, , scKey, scRoleName)
            oNode.Tag = rs.Fields("pk_staff_role")
        Else
            scParentKey = "N" & rs.Fields("pk_parent_role")
            Set oNode = Me.TreeView1.Nodes.Add(scParentKey, modCustomTypesAndEnums.TreeViewNodeType.tvwChild_, scKey, scRoleName)
            oNode.Tag = rs.Fields("pk_staff_role")
        End If
        
        GoTo skip_reporting_error
        
report_error:
        Call MsgBox("Error occurred with staff role ID " & scKey)


skip_reporting_error:
        Call rs.MoveNext
        
    Wend
    
    Call rs.Close
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   TreeView1_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub TreeView1_Click()

    'Dim iPkRole As Long


    If TreeView1.SelectedItem.index < 0 Then
        Exit Sub
    End If
    
    m_iSelectedRole = Val(Mid(TreeView1.SelectedItem.Key, 2))
    
    '===========
    ' Get the data for all members of this role, including vacant positions.
    '===========
    Call PopulateEncumbentsArray
    
    '===========
    ' Display the role
    '===========
    Call DisplayRole

End Sub

'==============================================================================
' SUBROUTINE
'   PopulateEncumbentsArray
'------------------------------------------------------------------------------
' DESCRIPTION
'   Fills out the encumbents collection with the
'==============================================================================
Private Sub PopulateEncumbentsArray()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

#If DebugBadType = 0 Then
   Dim oEncumbent As clsEncumbentData
#Else
   Dim oEncumbent As Object
#End If

    Dim scSQLQuery As String
    
    '==========
    ' Reset the collection by defining a new collection.
    '==========
    Set m_colEncumbents = New VBA.Collection
    
    scSQLQuery = "SELECT * from people.dbo.v_roles_positions_and_staff WHERE (PositionActive = 1) AND RoleID = " & m_iSelectedRole
    
    Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)
    
    While Not rs.EOF
        Set oEncumbent = New clsEncumbentData
        
        Call oEncumbent.Populate(rs.Fields)
        
        Call m_colEncumbents.Add(oEncumbent)
    
        Call rs.MoveNext
    Wend
End Sub

'==============================================================================
' SUBROUTINE
'   ClearRoleDisplay
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clears all role information and associated person and position info
'==============================================================================
Private Sub ClearRoleDisplay()

    lblRole.Caption = ""
    lblSection.Caption = ""
    lblRoleDept.Caption = ""
    lblRoleCount.Caption = ""
    lblRoleID.Caption = ""
    
    btnSelectRole.Enabled = False
    
    Call lbIncumbents.Clear

    
    Call ClearPositionDisplay
    
    Call ClearPersonDisplay
End Sub

'==============================================================================
' SUBROUTINE
'   DisplayRole
'------------------------------------------------------------------------------
' DESCRIPTION
'   Displays the role information and populates the incumbents listbox
'==============================================================================
Private Sub DisplayRole()
        
#If DebugBadType = 0 Then
   Dim oEncumbent As clsEncumbentData
#Else
   Dim oEncumbent As Object
#End If

    Dim i As Long
    
    
    Call ClearRoleDisplay
    
    m_iSelectedEncumbentIndex = -1
    
    lblRoleCount.Caption = Me.m_colEncumbents.Count
    
    If Me.m_colEncumbents.Count < 1 Then
        Exit Sub
    End If
    
    Set oEncumbent = Me.m_colEncumbents(1)
    
    '===========
    ' Fill out the basic role information at the top.
    '===========
    lblRole.Caption = oEncumbent.Role ' rs.Fields("role_name")
    lblSection.Caption = oEncumbent.Section ' rs.Fields("section_name")
    lblRoleDept.Caption = oEncumbent.Dept ' rs.Fields("sect_dept_name")
    lblRoleID.Caption = oEncumbent.RoleID
    
    '===========
    ' Loop through each encumbent, adding to the listbox.
    '===========
    For i = 1 To Me.m_colEncumbents.Count
        Set oEncumbent = m_colEncumbents(i)
        
        If oEncumbent.FullName = "" Then
            Call lbIncumbents.AddItem("Vacant")
        Else
            Call lbIncumbents.AddItem(oEncumbent.FullName)
        End If
        
'        lbIncumbents.List(lbIncumbents.ListCount - 1, 1) = oEncumbent.PosNo
'        lbIncumbents.List(lbIncumbents.ListCount - 1, 2) = oEncumbent.StaffID
        
'        Call rs.MoveNext
    Next

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   frmTreeView_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub frmTreeView_Click()
'    m_sWhatIsSelected = "Nothing"
'    m_sSelectedKey = ""
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbIncumbents_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub lbIncumbents_Click()
'    Dim iPosId As Long
'    Dim iStaffID As Long
     
    If lbIncumbents.ListIndex >= 0 Then
    
        m_iSelectedEncumbentIndex = lbIncumbents.ListIndex
'        m_iSelectedPosition = Val(lbIncumbents.List(lbIncumbents.ListIndex, 1))
'        m_iSelectedStaffID = Val(lbIncumbents.List(lbIncumbents.ListIndex, 2))
        
        Call DisplayPosition
        
'        m_sWhatIsSelected = "Encumbent"
'        m_sSelectedKey = lbIncumbents.Column(0)
        
        
        'm_iSelectedEncumbent = iStaffID
    Else
        Call ClearPositionDisplay
        Call ClearPersonDisplay
    End If
    
End Sub

'==============================================================================
' SUBROUTINE
'   ClearPositionDisplay
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub ClearPositionDisplay()

    Application.EnableEvents = False
On Error GoTo cleanup_nicely
    
    '========
    ' Position Tab
    '========
    txtSAPPositionNo.Text = ""
    lblJobTitle.Caption = ""
    lblOrgUnit.Caption = ""
    
    txtSAPPosReportsToNo.Text = ""
    lblSAPPosReportsToTitle.Caption = ""
    txtSAPPosReportsToNo.Locked = True
    
    btnUpdateSAPData.Enabled = False
    
cleanup_nicely:
    Application.EnableEvents = True

End Sub

'==============================================================================
' SUBROUTINE
'   DisplayPosition
'------------------------------------------------------------------------------
' DESCRIPTION
'   Displays the position and associated person (if the role is not vacant).
'==============================================================================
Private Sub DisplayPosition()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rsPic As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rsPic As Object
#End If

#If DebugBadType = 0 Then
    Dim oEncumbent As clsEncumbentData
    Dim cPicData As clsPictureData
#Else
    Dim oEncumbent As Object
    Dim cPicData As Object
#End If

    Dim scSQLQuery As String
    Dim bPersonFound As Boolean
    Dim iCycleDays As Long
    Dim dtSeedDate As Date
    Dim iPicCount As Long

    
    Call ClearPositionDisplay
    Call ClearPersonDisplay
    
'    scSQLQuery = "SELECT pk_staff_position_sap, job_title, fk_parent_position_sap, " & _
'        "parent_job_title, fk_org_unit," & _
'        "fk_staff_role, position_active, role_name, fk_section, full_name," & _
'        "SAP_id, person_active, org_description, pk_staff, fk_sap_name, " & _
'        "sap_work_schedule, sap_work_schedule_rule, nominal_seed_date, " & _
'        "roster_name, cycle_days, do_projects, do_actions " & _
'        "FROM people.dbo.v_position_detail " & _
'        "WHERE pk_staff_position_sap = " & m_iSelectedPosition
'
'    Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)

    Set oEncumbent = Me.m_colEncumbents(m_iSelectedEncumbentIndex + 1)
    
    bPersonFound = False
    
    '===========
    ' First the position information
    '===========
'    bPositionFound = True
    
    txtSAPPositionNo.Text = oEncumbent.PosNo ' rs.Fields("pk_staff_position_sap")
    lblJobTitle.Caption = oEncumbent.JobTitle ' rs.Fields("job_title")
    lblOrgUnit.Caption = oEncumbent.OrgUnit ' rs.Fields("org_description")
        
    '===========
    ' Parent Position
    '===========
    If oEncumbent.ParentPosNo = "" Then 'IsNull(rs.Fields("fk_parent_position_sap")) Then
        txtSAPPosReportsToNo = ""
        lblSAPPosReportsToTitle = "<Not Defined>"
    Else
        txtSAPPosReportsToNo.Text = oEncumbent.ParentPosNo ' rs.Fields("fk_parent_position_sap")
        lblSAPPosReportsToTitle.Caption = oEncumbent.ParentJobTitle ' rs.Fields("parent_job_title")
    End If
    txtSAPPosReportsToNo.Locked = False
    
    '===========
    ' Then the persons information, if the position isn't vacant.
    '===========
    If oEncumbent.FullName <> "" Then
        
        lblPkStaff.Caption = oEncumbent.StaffID
        
        '===========
        ' Person Tab
        '===========
        txtSAPID.Text = oEncumbent.SapID
        txtPersonsName.Text = oEncumbent.FullName
        txtSAPUsername.Text = oEncumbent.SapUserName
        
        '===========
        ' Additional Tab
        '===========
        txtSAPRosterID.Text = oEncumbent.RosterID
        lblSAPRosterDescription.Caption = oEncumbent.RosterName
        iCycleDays = oEncumbent.RosterCycleDays
        
        dtSeedDate = oEncumbent.RosterSeedDate
        
        dtSeedDate = Date - ((Date - dtSeedDate) Mod iCycleDays)
        If dtSeedDate <= (Date - CLng(iCycleDays * 0.6)) Then
            dtSeedDate = dtSeedDate + iCycleDays
        End If
        
        lblSAPRosterSeedDate.Caption = Format(dtSeedDate, "ddd, d-mmm-yyyy")
        
        cbDoesProjects.Enabled = True
        cbDoesProjects.Value = oEncumbent.DoesProjects
        
        cbDoesActions.Enabled = True
        cbDoesActions.Value = oEncumbent.DoesActions
        
        '===========
        ' Now pictures
        '===========
        Set m_colPictures = Nothing
        Set m_colPictures = New VBA.Collection
        
        iPicCount = 0
        
        '============
        ' Now read from the database
        '============
        scSQLQuery = "SELECT * FROM maint.dbo.v_mapped_files " & _
                "WHERE id = '" & lblPkStaff.Caption & "' AND base_path_category = 'PERSON_PICTURES' " & _
                "ORDER BY file_order, file_date"
                
        Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rsPic)
        
        While Not rsPic.EOF
            iPicCount = iPicCount + 1
            Set cPicData = New clsPictureData
            
            Call cPicData.PopulateFromRecordset(rsPic)
            
            Call m_colPictures.Add(cPicData, cPicData.Key)
            Call rsPic.MoveNext
        Wend
        
        Call rsPic.Close
        
        lblPicCount.Caption = iPicCount
        If (iPicCount > 0) Then
            lblPicIndex.Caption = "1"
        Else
            lblPicIndex.Caption = "0"
        End If
        
        bPersonFound = True
    End If
    
    '==========
    ' Display any pictures
    '==========
    If m_colPictures.Count > 0 Then
        '==========
        ' Get the first picture in the list and display it
        '==========
        lblPicIndex.Caption = 1
        Call DisplayCurrentPicture(1)
        
        If m_colPictures.Count > 1 Then
            btnNextPicture.Enabled = True
        End If
    End If

End Sub

'==============================================================================
' SUBROUTINE
'   ClearPersonDisplay
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub ClearPersonDisplay()

    Application.EnableEvents = False
On Error GoTo cleanup_nicely

    txtSAPID.Text = ""
    txtSAPID.Tag = "NotDirty"
    
    txtSAPUsername.Text = ""
    txtSAPUsername.Locked = True
    txtSAPUsername.Tag = "NotDirty"
    
    txtPersonsName.Text = ""
    txtPersonsName.Tag = "NotDirty"
    
    txtSAPRosterID.Text = ""
    txtSAPRosterID.Tag = "NotDirty"
    
    lblSAPRosterDescription.Caption = ""
    lblSAPRosterSeedDate.Caption = ""
    
    lblPkStaff.Caption = ""
    
    imgPerson.Picture = LoadPicture("")
    lblPictureDate.Caption = ""
    lblPictureInfo.Caption = ""
    
    Set m_colPictures = Nothing
    Set m_colPictures = New VBA.Collection
    
    cbDoesProjects.Value = False
    cbDoesProjects.Enabled = False
    
    cbDoesActions.Value = False
    cbDoesActions.Enabled = False
    
    btnUpdatePersonAdditional.Enabled = False
    
cleanup_nicely:
    Application.EnableEvents = True
    
End Sub

'==============================================================================
' SUBROUTINE
'   DisplayPerson
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub DisplayPerson()

End Sub

'==============================================================================
' SUBROUTINE
'   UpdatePositionDetailDisplay
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub UpdatePositionDetailDisplay(iListBoxIndex As Long)

#If DebugBadType = 0 Then
   Dim oEncumbent As clsEncumbentData
#Else
   Dim oEncumbent As Object
#End If

    Application.EnableEvents = False

On Error GoTo cleanup_nicely

    Set oEncumbent = Me.lbIncumbents.List(iListBoxIndex)
    
    '========
    ' Role Tab
    '========
    
    If oEncumbent.RoleID > 0 Then
        lblRole.Caption = oEncumbent.Role
        lblSection.Caption = oEncumbent.Section
        lblRoleDept.Caption = oEncumbent.Dept
        lblRoleID.Caption = oEncumbent.RoleID
    Else
        lblRole.Caption = ""
        lblSection.Caption = ""
        lblRoleDept.Caption = ""
        oEncumbent = ""
    End If
    
    '========
    ' SAP Tab.
    ' 1. Position details.
    '========
    Me.txtSAPPositionNo.Text = oEncumbent.PosNo
    Me.lblOrgUnit.Caption = oEncumbent.OrgUnit
    Me.lblJobTitle.Caption = oEncumbent.JobTitle
    
    '========
    ' 2. Reports to
    '========
    Me.txtSAPPosReportsToNo.Text = oEncumbent.ParentPosNo
    Me.lblSAPPosReportsToTitle.Caption = oEncumbent.ParentJobTitle
    
    Me.txtSAPPosReportsToNo.Locked = False ' Allow the user to change
    Me.txtSAPPosReportsToNo.Tag = "NotDirty" ' Not Dirty
    
    '========
    ' 3. Person details
    '========
    If oEncumbent.StaffID > 0 Then
        Me.txtSAPID.Text = oEncumbent.SapID
        Me.txtSAPRosterID.Text = oEncumbent.RosterName
        Me.lblSAPRosterDescription.Caption = oEncumbent.RosterDescription
    
        Me.txtSAPUsername.Text = oEncumbent.SapUserName
        Me.txtSAPUsername.Locked = False ' Allow the user to change if required
    Else
        Me.txtSAPID.Text = ""
        Me.txtSAPRosterID.Text = ""
        Me.lblSAPRosterDescription.Caption = ""
        
        Me.txtSAPUsername.Text = ""
        Me.txtSAPUsername.Locked = True ' Prevent changing ... confusing when there is no person
    End If
    Me.txtSAPUsername.Tag = "NotDirty" ' Not Dirty
    
    Me.btnUpdateSAPData.Enabled = False
    
    m_iSelectedEncumbentIndex = iListBoxIndex
    
cleanup_nicely:
    Application.EnableEvents = True
        
End Sub


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtSAPUsername_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clicked when the user changes the SAP Username for the person
'==============================================================================
Private Sub txtSAPUsername_Change()
    '=============
    ' Don't do anything if there is no longer an encumbent
    '=============
    If m_iSelectedEncumbentIndex < 0 Then
        Exit Sub
    End If
    
    If txtSAPID.Text <> "" Then
        txtSAPUsername.Tag = "Dirty" ' Dirty
        btnUpdateSAPData.Enabled = True
    End If
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtSAPPosReportsToNo_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clicked when the user changes the parent position number
'==============================================================================
Private Sub txtSAPPosReportsToNo_Change()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

#If DebugBadType = 0 Then
   Dim oEncumbent As clsEncumbentData
#Else
   Dim oEncumbent As Object
#End If

    Dim scSQLQuery As String


    '=============
    ' Don't do anything if there is no longer an encumbent
    '=============
    If m_iSelectedEncumbentIndex < 0 Then
        Exit Sub
    End If
    
    Set oEncumbent = m_colEncumbents(m_iSelectedEncumbentIndex + 1)
    
    '=============
    ' Check if this is a root position (i.e. GM) and the user is trying to make
    ' this position report to another (which will eliminate root nodes and
    ' potentially create a circular reporting structure).
    ' We only check if the parent was unassigned, the parent field is now
    ' non-blank and the text box field has not yet become dirty (to avoid
    ' repeatedly asking the question).
    '=============
    If (oEncumbent.ParentPosNo = "") And _
            (txtSAPPosReportsToNo.Text <> "") And _
            (txtSAPPosReportsToNo.Tag = "NotDirty") Then
        If MsgBox("Changing the parent position from no parent could create issues. Are you sure?", vbYesNo) = vbNo Then
            Application.EnableEvents = False
            txtSAPPosReportsToNo.Text = ""
            Application.EnableEvents = True
            Exit Sub
        End If
    End If
    
    txtSAPPosReportsToNo.Tag = "Dirty"
    btnUpdateSAPData.Enabled = True
    
    '===========
    ' If it looks like a valid position number, we attempt to update the title
    '===========
    If Len(txtSAPPosReportsToNo.Text) = 8 Then
        If Left(txtSAPPosReportsToNo.Text, 1) = "3" Then
            If IsNumeric(txtSAPPosReportsToNo.Text) Then
                scSQLQuery = "SELECT job_title FROM people.dbo.t_staff_position_sap WHERE pk_staff_position_sap = '" & Replace(txtSAPPosReportsToNo.Text, "'", "''") & "'"
                Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)
                
                If rs.EOF Then
                    lblSAPPosReportsToTitle.Caption = "<Undefined SAP Position>"
                Else
                    lblSAPPosReportsToTitle.Caption = rs.Fields("job_title")
                End If
            Else
                lblSAPPosReportsToTitle.Caption = "<Undefined SAP Position>"
            End If
        Else
            lblSAPPosReportsToTitle.Caption = "<Undefined SAP Position>"
        End If
    Else
        lblSAPPosReportsToTitle.Caption = ""
    End If
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnClose_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnClose_Click()
    Call Unload(Me)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnRefresh_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnRefresh_Click()
    Call InitializeTreeView
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtSAPPosReportsToNo_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clicked when the user changes the parent position number
'==============================================================================
Private Sub btnUpdateSAPData_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

#If DebugBadType = 0 Then
   Dim oEncumbent As clsEncumbentData
#Else
   Dim oEncumbent As Object
#End If

    Dim scSQLQuery As String

    Dim iRecordsAffected As Long
    Dim bUserSaysMoveAll As Boolean
    Dim scSQLIn As String
    
    Dim bParentPosNotFound As Boolean
    Dim scMsg As String
    
    Dim scPosNo As String
    Dim scOldParentPosID As String
    Dim scNewParentPosID As String
    Dim iRoleID As Long
    Dim iOldParentRoleID As Long
    Dim iNewParentRoleID As Long
    
    Dim scNewParentJobTitle As String
    Dim iNewParentSectionID As Long
    Dim iNewParentPosLevelNo As Long
    Dim iNewParentRoleLevelNo As Long
    Dim iPosInRoleCount As Long
    
    Dim iLevelNo As Long
    Dim iPositionsAffected As Long
    Dim iRolesAffected As Long
    
    '==========
    ' Has anything actually changed that warrant's an update?
    '==========
    If (txtSAPUsername.Tag <> "Dirty") And (txtSAPPosReportsToNo.Tag <> "Dirty") Then
        '===========
        ' Shouldn't occur
        '===========
        Exit Sub
    End If
    
    '========
    ' Connect to the DB
    '========
    If Not (ConnectToDB(ldPeople, cnn, True)) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Sub
    End If
    
    Set oEncumbent = m_colEncumbents(m_iSelectedEncumbentIndex + 1)
    scPosNo = oEncumbent.PosNo
    
    '========
    ' Has the username changed?
    '========
    If txtSAPUsername.Tag = "Dirty" Then
    
        scSQLQuery = "UPDATE people.dbo.t_staff SET fk_sap_name = '" & Replace(Trim(txtSAPUsername.Text), "'", "''") & "'" & _
            " WHERE pk_staff = " & oEncumbent.StaffID
            
        Call cnn.Execute(scSQLQuery, iRecordsAffected)
        If iRecordsAffected > 0 Then
            scMsg = "Successfully updated SAP Username for " & oEncumbent.FullName
            If Me.cbSupressPopups.Value Then
                lblGeneralInfo.Caption = scMsg
            Else
                Call MsgBox(scMsg)
            End If
            oEncumbent.SapUserName = Trim(txtSAPUsername.Text)
        End If
        txtSAPUsername.Tag = "NotDirty"
    End If
    
    '========
    ' Has the Parent Position Changed?
    '========
    If txtSAPPosReportsToNo.Tag = "Dirty" Then
    
        scNewParentPosID = Trim(Me.txtSAPPosReportsToNo.Text)
        '===========
        ' We check this parent position exists
        '===========
        scSQLQuery = "SELECT fk_staff_role, fk_section, job_title, pos_level_no, role_level_no FROM people.dbo.v_positions_and_roles WHERE pk_staff_position_sap = '" & _
                scNewParentPosID & "'"
                
        Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)
        
        If rs.EOF Then
            '===========
            ' Doesn't exist so tell the user
            '===========
            Call rs.Close
            Call MsgBox("Position number " & Trim(Me.txtSAPPosReportsToNo.Text) & _
                    " not found. Unable to update Parent Position")
        Else
            '===========
            ' Get some key data on the new parent position and role before we
            ' close the recordset.
            '===========
            scNewParentJobTitle = rs.Fields("job_title")
            iNewParentRoleID = rs.Fields("fk_staff_role")
            iNewParentSectionID = rs.Fields("fk_section")
            iNewParentPosLevelNo = rs.Fields("pos_level_no")
            iNewParentRoleLevelNo = rs.Fields("role_level_no")
            Call rs.Close
            
            iRoleID = Val(lblRoleID.Caption)
        
            '===========
            ' Parent position can't have the same role as this position.
            '===========
            If iRoleID = iNewParentRoleID Then
                Call MsgBox("You are trying to assign the parent position to a position with the same role as this position")
                Exit Sub
            End If
            
            '===========
            ' Next, we make sure we are not trying to assign the parent
            ' position to a child of this position, or another position with
            ' the same role. We do this by calling CheckPositionAndRolesToTop.
            '===========
            If Not CheckPositionAndRolesToTop(scNewParentPosID, scPosNo, iRoleID, iLevelNo, 20) Then
                Call MsgBox("Parent Position " & scNewParentPosID & " appears to be a descendant of this position.")
                Exit Sub
            End If
            
            '===========
            ' Next we set this position, and any other positions with the
            ' same role as this position, to the new parent position. This is
            ' because the alternative of splitting them up is too difficult.
            ' And remember, the reason why they were groupedin the first
            ' place is because they belonged to the same org group and had the
            ' same job title.
            '===========
            scSQLQuery = "UPDATE people.dbo.t_staff_position_sap SET fk_parent_position_sap = '" & scNewParentPosID & "' WHERE fk_staff_role = " & iRoleID
            Call cnn.Execute(scSQLQuery, iPositionsAffected)

            '===========
            ' Update this role to have the same parent role as the parent position.
            '===========
            scSQLQuery = "UPDATE people.dbo.t_staff_role SET fk_role_parent = " & iNewParentRoleID & " WHERE pk_staff_role = " & iRoleID
            Call cnn.Execute(scSQLQuery, iRolesAffected)
            
                        
            scMsg = "Successfully updated the parent position for position " & oEncumbent.PosNo & "-" & oEncumbent.JobTitle & " (" & iPositionsAffected & " position(s) and " & iRolesAffected & " role affected)"
            If Me.cbSupressPopups.Value Then
                lblGeneralInfo.Caption = scMsg
            Else
                Call MsgBox(scMsg)
            End If
        
            oEncumbent.ParentJobTitle = scNewParentJobTitle
            oEncumbent.ParentPosNo = Trim(Me.txtSAPPosReportsToNo.Text)
            oEncumbent.ParentRoleID = iNewParentRoleID
            
            '=============
            ' We also need to fix up the level numbers for the whole hierarchy
            '=============
            Call cnn.Execute("EXEC dbo.SetPositionLevelNumbers", iRecordsAffected)
            Call cnn.Execute("EXEC dbo.SetRoleLevelNumbers", iRecordsAffected)
            
        End If
end_of_parent_position_update:
        txtSAPPosReportsToNo.Tag = "NotDirty"
    End If

    Call UpdatePositionDetailDisplay(m_iSelectedEncumbentIndex)
    
    Me.btnUpdateSAPData.Enabled = False

End Sub

'==============================================================================
' SUBROUTINE
'   GetPositionLevelsToTop
'------------------------------------------------------------------------------
' DESCRIPTION
'   Through a recursive mechanism, this method determines how many levels
' there are between the supplied position and the head honcho (with a NULL
' parent).
'==============================================================================
Public Function CheckPositionAndRolesToTop( _
        scPos As String, _
        scIllegalPosNo As String, _
        iIllegalRoleID As Long, _
        ByRef iLevelCount As Long, _
        Optional maxRecursion As Long = 20) As Boolean
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    Dim iParentLevelCount As Long
    
    scSQLQuery = "SELECT fk_parent_position_sap, fk_role_parent FROM people.dbo.v_positions_and_roles WHERE pk_staff_position_sap = '" & scPos & "'"
    Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)
    
    If rs.EOF Then
        iLevelCount = 0
        CheckPositionAndRolesToTop = True
        Exit Function
    ElseIf maxRecursion < 1 Then
        iLevelCount = 0
        CheckPositionAndRolesToTop = False
        Exit Function
    ElseIf IsNull(rs.Fields("fk_parent_position_sap")) Then
        iLevelCount = 1
        CheckPositionAndRolesToTop = True
        Exit Function
    ElseIf scIllegalPosNo = rs.Fields("fk_parent_position_sap") Then
        CheckPositionAndRolesToTop = False
        Exit Function
    
    Else
        If Not IsNull(rs.Fields("fk_role_parent")) Then
            If iIllegalRoleID = rs.Fields("fk_role_parent") Then
                CheckPositionAndRolesToTop = False
                Exit Function
            End If
        End If
        
        CheckPositionAndRolesToTop = CheckPositionAndRolesToTop(rs.Fields("fk_parent_position_sap"), scIllegalPosNo, iIllegalRoleID, iParentLevelCount, maxRecursion - 1)
        iLevelCount = iParentLevelCount + 1
    End If
    
End Function


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnPrevPicture_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnPrevPicture_Click()
    
    Dim scPKStaff As String
    Dim iIndex As Long
    
    iIndex = CLng(lblPicIndex.Caption)
    
    scPKStaff = lblPkStaff.Caption

    If iIndex > 1 Then
        iIndex = iIndex - 1
        
        Call DisplayCurrentPicture(iIndex)

        btnNextPicture.Enabled = True

        If iIndex = 1 Then
            btnPrevPicture.Enabled = False
        End If
        
        lblPicIndex.Caption = iIndex

    Else
        btnPrevPicture.Enabled = False
    End If

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnNextPicture_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnNextPicture_Click()

    Dim scPKStaff As String
    Dim iIndex As Long
    Dim iPicCount As Long
    
    iIndex = CLng(lblPicIndex.Caption)
    iPicCount = CLng(lblPicCount.Caption)
    
    scPKStaff = lblPkStaff.Caption

    If iPicCount > iIndex Then
        iIndex = iIndex + 1
        
        Call DisplayCurrentPicture(iIndex)
        
        btnPrevPicture.Enabled = True
        If iIndex = iPicCount Then
            btnNextPicture.Enabled = False
        End If
        
        lblPicIndex.Caption = iIndex

    Else
        btnNextPicture.Enabled = False
    End If

End Sub

'==============================================================================
' SUBROUTINE
'   DisplayCurrentPicture
'------------------------------------------------------------------------------
' DESCRIPTION
'   Displays the current picture as defined by the supplied nodes PictureIndex
' parameter.
'==============================================================================
Private Sub DisplayCurrentPicture(iIndex As Long)

#If DebugBadType = 0 Then
    Dim cPicData As clsPictureData
#Else
    Dim cPicData As Object
#End If
        
    Set cPicData = m_colPictures(iIndex)
        
On Error GoTo could_not_find_picture
    imgPerson.Picture = LoadPicture(cPicData.FullPath)
    lblPictureInfo.Caption = cPicData.description
    lblPictureDate.Caption = Format(cPicData.file_date, "d-mmm-yyyy")
    lblPictureDate.ForeColor = RGB(180, 180, 0)
    
    lblPicIndex.Caption = iIndex
    
    Exit Sub
    
could_not_find_picture:
    lblPictureInfo.Caption = "Could not load picture"
    lblPictureDate.Caption = ""
    lblPicIndex.Caption = "-"
End Sub

'==============================================================================
' SUBROUTINE
'   btnOpenPicture_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Makes a copy of the picture in the temp folder, and opens using the default
' client picture application.
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
    Dim iIndex As Long
    
    '==========
    ' Is there a picture?
    '==========
    If m_colPictures.Count < 1 Then
        Exit Sub
    End If
    
    iIndex = CLng(lblPicIndex.Caption)
        
    Set cPicData = m_colPictures(iIndex)
        
    '==========
    ' Get the users Temp folder.
    '==========
    scTempFolderPath = GetSpecialFolderPath(sfTemp)
    
    If DBCopyFile(cPicData.pk_file, scTempFolderPath, oDestFile) Then
    
        '===========
        ' Open the file in its default application.
        '===========
        Dim oShell As Object
        Set oShell = CreateObject("Shell.Application")
        
        Call oShell.Open(oDestFile.Path)
    End If
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnManagePictures_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Creates and opens the picture manager form for managing pictures
' associated with this form.
'==============================================================================
Private Sub btnManagePictures_Click()
    Dim ufPictMgr As ufPictureManager
    
    Set ufPictMgr = New ufPictureManager
    
    Call ufPictMgr.Configure("PERSON_PICTURES", lblPkStaff.Caption)
    
    Call ufPictMgr.Show
    
    If Not (ufPictMgr Is Nothing) Then
        If ufPictMgr.DirtyFlag Then
            Call DisplayPosition
        End If
    End If
        
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   cbDoesProjects_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when clicking the 'Does Projects' Check box
'==============================================================================
Private Sub cbDoesProjects_Click()
    btnUpdatePersonAdditional.Enabled = True
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   cbDoesActions_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when clicking the 'Does Actions' Check box
'==============================================================================
Private Sub cbDoesActions_Click()
    btnUpdatePersonAdditional.Enabled = True
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnUpdatePersonAdditional_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Update the details on the Additional tab
'==============================================================================
Private Sub btnUpdatePersonAdditional_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    
    Dim iID As Long
    Dim iDoProjects As Long
    Dim iDoActions As Long
    
    iID = Val(lblPkStaff.Caption)
    
    If cbDoesProjects.Value Then
        iDoProjects = 1
    Else
        iDoProjects = 0
    End If
    
    If cbDoesActions.Value Then
        iDoActions = 1
    Else
        iDoActions = 0
    End If
    
    scSQLQuery = "UPDATE people.dbo.t_staff SET do_projects = " & _
            iDoProjects & ", do_actions = " & iDoActions & " WHERE pk_staff = " & iID
            
    '========
    ' Connect to the DB
    '========
    If Not (ConnectToDB(ldPeople, cnn, True)) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Sub
    End If
    
    Call cnn.Execute(scSQLQuery)
    
    btnUpdatePersonAdditional.Enabled = False
    
End Sub


