VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufFLOCBrowser 
   Caption         =   "Functional Location Browser"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14190
   OleObjectBlob   =   "ufFLOCBrowser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufFLOCBrowser"
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

Private Const g_cMaxWorkOrderCount As Long = 500
Private Const g_cMaxNotificationCount As Long = 500
Private Const g_cMaxARCount As Long = 500
Private Const g_cMaxShiftLogCount As Long = 250
Private Const g_cMaxPartsCount As Long = 500
Private Const g_cMaxFLOCSearch As Long = 20

'==============================================================================
' PRIVATE MEMBER VARIABLES
'==============================================================================
Private Const g_iDrillDownLevel As Long = 2

Private m_iColFlocTable_pk_func_loc As Long
Private m_iColFlocTable_Description As Long
Private m_iColFlocTable_floc_type As Long
Private m_iColFlocTable_fk_parent As Long
Private m_iColFlocTable_parent_floc_type As Long
Private m_iColFlocTable_fk_const_type As Long
Private m_iColFlocTable_level_no As Long
Private m_iColFlocTable_in_sap As Long

Private m_bNotificationOnDownloadIsRequiredEnd As Boolean

Private m_iFlocCount As Long
Private m_sWhatIsSelected As String
Private m_sSelectedKey As String

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public m_colNodeData As VBA.Collection

#If DevelopMode = 1 Then
    Public m_oSelectedNode As MSComctlLib.Node
#Else
    Public m_oSelectedNode As Object
#End If

Public m_ufPartDisplay As ufPartDisplay

Public m_wsFlocTable As Excel.Worksheet
Public m_wbFlocTable As Excel.Workbook

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub UserForm_Initialize()

    Dim wb As Excel.Workbook
    
    '===========
    ' Load a copy of the floc table, if more than 4 days out of date
    '===========
    
    Call Me.GetFlocTable
    
    Call InitialiseTreeView(m_wsFlocTable)
    
    Call m_wbFlocTable.Close
    Set m_wsFlocTable = Nothing
    
    m_sWhatIsSelected = "Nothing"
    Me.lblWOSearchResultsDescription.Caption = ""
    Me.lblNotiSearchResultsDescription.Caption = ""
    Me.lblSearchResultsDescription.Caption = ""
    Me.lblPictureCountPosition.Caption = ""
    
    Me.lblFLOCName.ForeColor = RGB(64, 64, 255)
    
    Call BlankFLOCData
End Sub

'==============================================================================
' SUBROUTINE
'   Reset
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Public Sub Reset()
    Call TreeView1.Refresh
    mpFLOCDetail.Value = 0
End Sub

'==============================================================================
' SUBROUTINE
'   InitializeTreeView
'------------------------------------------------------------------------------
' DESCRIPTION
'   Populates the treeview control with the contents of the worksheet.
'==============================================================================
Private Sub InitialiseTreeView(wsSource As Excel.Worksheet)

#If DevelopMode = 1 Then
    Dim oNode As MSComctlLib.Node
    Dim oParentNode As MSComctlLib.Node
#Else
    Dim oNode As Object
    Dim oParentNode As Object
#End If

#If DebugBadType = 0 Then
    Dim cNodeFLOC As clsNodeFLOC
#Else
    Dim cNodeFLOC As Object
#End If

    Dim iRow As Long
    Dim iImage As Long
    Dim iParentRow As Long
    Dim iErrorCount As Long
    Dim iDebug As Long
    Dim iParentNotFoundCount As Long
    Dim scNodeText As String

On Error GoTo exit_nicely
    
    '============
    ' Initialise some TreeView parameters
    '============
    Me.TreeView1.ImageList = ImageList1
    
    '============
    ' Create the collection that will store node data.
    '============
    Set m_colNodeData = New VBA.Collection

    '============
    ' First the NAVI table
    '============
    iRow = 2
    iParentNotFoundCount = 0
    While (wsSource.Cells(iRow, 1) <> "")
    
        Set cNodeData = New clsNodeFLOC
        
        '===========
        ' Get the row values
        '===========
        cNodeData.FuncLoc = wsSource.Cells(iRow, m_iColFlocTable_pk_func_loc)
        cNodeData.Name = wsSource.Cells(iRow, m_iColFlocTable_Description)
        cNodeData.FLOCType = Left(wsSource.Cells(iRow, m_iColFlocTable_floc_type), 4)
        cNodeData.FLOCParent = wsSource.Cells(iRow, m_iColFlocTable_fk_parent)
        cNodeData.FLOCParentType = Left(wsSource.Cells(iRow, m_iColFlocTable_parent_floc_type), 4)
  
        '===========
        ' Add the new node data object to the collection
        '===========
        Call m_colNodeData.Add(cNodeData, cNodeData.Key)
        
        '===========
        ' Create a shortened name that consists only of the extra component
        ' of the name not in the parent node.
        '===========
        scNodeText = ""
        If cNodeData.IsNAVI = cNodeData.ParentIsNAVI Then
            If Left(cNodeData.FuncLoc, Len(cNodeData.FLOCParent)) = cNodeData.FLOCParent Then
                scNodeText = Mid(cNodeData.FuncLoc, Len(cNodeData.FLOCParent) + 2) & "  " & cNodeData.Name
            End If
        End If
        If scNodeText = "" Then
            scNodeText = cNodeData.FuncLoc & "  " & cNodeData.Name
        End If
        
        '===========
        ' Determine the image
        '===========
        If cNodeData.IsNAVI Then
            iImage = 2
        Else
            iImage = 1
        End If

        '===========
        ' Add the node. Either a root node, or as the child of a parent.
        '===========
        If cNodeData.FLOCParent = "" Then
            Set oNode = Me.TreeView1.Nodes.Add(, , cNodeData.Key, scNodeText, iImage) 'iImage, iImage)
            oNode.Tag = "N"
            Set cNodeData.TreeViewNode = oNode
        ElseIf GetNode(cNodeData.ParentKey, oParentNode) Then
            Set oNode = Me.TreeView1.Nodes.Add(cNodeData.ParentKey, TreeViewNodeType.tvwChild_, cNodeData.Key, scNodeText, iImage)
            oNode.Tag = "N"
            Set cNodeData.TreeViewNode = oNode
        Else
            iParentNotFoundCount = iParentNotFoundCount + 1
        End If
        
        iRow = iRow + 1
        
    Wend
        
    '===========
    ' Display the FLOC detail tab
    '===========
    mpFLOCDetail.Value = 0
    
    Exit Sub

exit_nicely:
    Call MsgBox("Error on row " & iRow & ": " & Err.description & " (" & Err.Source & ")")
    
End Sub

'==============================================================================
' SUBROUTINE
'   GetNode
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the Node corresponding to the string Index, managing the
' error that is thrown if the index doesn't exist.
'==============================================================================
'Public Function GetNode(scIndex As String, ByRef oNode As MSComctlLib.Node) As Boolean
Public Function GetNode(scIndex As String, ByRef oNode As Object) As Boolean

On Error GoTo node_doesnt_exist
    
    Set oNode = Me.TreeView1.Nodes(scIndex)

    GetNode = True
    Exit Function

node_doesnt_exist:
    GetNode = False

End Function

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnClose_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnClose_Click()
    If Not Me.m_ufPartDisplay Is Nothing Then
        If Me.m_ufPartDisplay.Visible Then
            Call Me.m_ufPartDisplay.Hide
        End If
    End If
    
    If Me.cbExitMemory.Value Then
        Call Unload(Me)
    Else
        Call Hide
    End If
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   TreeView1_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   This click event is called whenever the node label is clicked on, or the
' expander is clicked on (to expand or collapse the node)
'==============================================================================
Private Sub TreeView1_Click()
    
#If DebugBadType = 0 Then
    Dim cNodeFLOC As clsNodeFLOC
#Else
    Dim cNodeFLOC As Object
#End If

    '========
    ' Blank the node view
    '========
    Call BlankFLOCData
    
    '========
    ' Is anything selected?
    '========
    If IsNull(TreeView1.SelectedItem) Then
        Exit Sub
    End If
    
    m_sSelectedKey = TreeView1.SelectedItem.Key
    
    Set m_oSelectedNode = TreeView1.SelectedItem
    
    Set cNodeFLOC = Me.m_colNodeData.Item(m_sSelectedKey)

    '========
    ' Display the information on the node
    '========
    If Not cNodeFLOC.DataPopulated Then
        Call cNodeFLOC.PopulateData
    End If

    If cNodeFLOC.DataPopulated Then
        Call DisplayFLOCData(cNodeFLOC)
    Else
        Call BlankFLOCData
    End If
    
End Sub

'==============================================================================
' SUBROUTINE
'   RefreshFlocView
'------------------------------------------------------------------------------
' DESCRIPTION
'   Call when-ever the entire FLOC detail display needs redoing.
'==============================================================================
Private Sub RefreshFlocView()
    
#If DebugBadType = 0 Then
    Dim cNodeFLOC As clsNodeFLOC
#Else
    Dim cNodeFLOC As Object
#End If

    If IsNull(TreeView1.SelectedItem) Then
        Call BlankFLOCData
        Exit Sub
    End If
    
    m_sSelectedKey = TreeView1.SelectedItem.Key
    
    Set m_oSelectedNode = TreeView1.SelectedItem
    
    Set cNodeFLOC = Me.m_colNodeData.Item(m_sSelectedKey)
    
    Call cNodeFLOC.PopulateData
    
    Call TreeView1_Click

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnSearchForParts_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Search for parts
'==============================================================================
Private Sub btnSearchForParts_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
        
    Dim iLbLineCount As Long
    Dim scSelectedNodeKey As String
    
#If DebugBadType = 0 Then
    Dim cNodeFLOC As clsNodeFLOC
#Else
    Dim cNodeFLOC As Object
#End If
    
    '========
    ' Clear the listbox and results label
    '========
    Call Me.lbPartsSearchResult.Clear
    Me.lblSearchResultsDescription.Caption = ""
    
    scSelectedNodeKey = Me.m_oSelectedNode.Key
    Set oNodeData = Me.m_colNodeData.Item(scSelectedNodeKey)
    
    '========
    ' Connect to the DB
    '========
    If Not (ConnectToDB(ldMaintenance, cnn, True)) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Sub
    End If
    
    '========
    ' Construct the query
    '========
    
    scSQLQuery = "SELECT DISTINCT TOP " & g_cMaxPartsCount & " t.Material, m.description,  count(*) as Usage " & _
            "FROM maint.dbo.v_work_orders as wo LEFT OUTER JOIN " & _
            "finance.dbo.v_order_trans as t ON wo.pk_work_order = t.WorkOrder LEFT OUTER JOIN " & _
            "parts.dbo.v_sap_material as m on t.Material = m.pk_sap_material " & _
            " WHERE wo.fk_func_loc "
            
    '========
    ' Use LIKE and a trailing % if we are looking at children as well.
    '========
    If Me.cbIncludeSubFlocs.Value Then
        scSQLQuery = scSQLQuery & "LIKE '" & oNodeData.FuncLoc & "%'"
    Else
        scSQLQuery = scSQLQuery & "= '" & oNodeData.FuncLoc & "'"
    End If
    scSQLQuery = scSQLQuery & " AND NOT t.Material is NULL " & _
            "GROUP BY t.Material, m.description ORDER BY Usage DESC, t.Material"

    '========
    ' Execute the query
    '========
    Set rs = CreateObject("ADODB.Recordset")
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)

    '========
    ' Display the results in the list box
    '========
    While Not rs.EOF
        Call Me.lbPartsSearchResult.AddItem(rs.Fields("Usage"))
        Me.lbPartsSearchResult.List(Me.lbPartsSearchResult.ListCount - 1, 1) = rs.Fields("Material")
        If Not IsNull(rs.Fields("description")) Then
            Me.lbPartsSearchResult.List(Me.lbPartsSearchResult.ListCount - 1, 2) = rs.Fields("description")
        End If

        Call rs.MoveNext
    Wend

    '========
    ' Complete the results label.
    '========
'    Const g_cMaxWorkOrderCount As Long = 500
'Const g_cMaxLogCount As Long = 500
    Dim scFeedback As String
    
    If rs.RecordCount = 0 Then
        scFeedback = "Found no results for FLOC '" & oNodeData.FuncLoc & "'"
    ElseIf rs.RecordCount = g_cMaxPartsCount Then
        scFeedback = "Results limited to first " & g_cMaxPartsCount & " for FLOC '" & oNodeData.FuncLoc & "'"
    Else
        scFeedback = "Found " & rs.RecordCount & " results for FLOC '" & oNodeData.FuncLoc & "'"
    End If
    
    If Me.cbIncludeSubFlocs.Value Then
        scFeedback = scFeedback & " and children."
    End If
    
    Me.lblSearchResultsDescription.Caption = scFeedback
End Sub

'==============================================================================
' SUBROUTINE
'   DisplayFLOCData
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
#If DebugBadType = 0 Then
Public Sub DisplayFLOCData(cNodeData As clsNodeFLOC)
#Else
Public Sub DisplayFLOCData(cNodeData As Object)
#End If

    lblFLOCName.Caption = cNodeData.m_sc_description
    
    txtCostCenter.Text = cNodeData.m_sc_fk_cost_centre
    If Len(cNodeData.m_sc_cc_colloquial) > 0 Then
        txtCostCenter.Text = txtCostCenter.Text & "  " & cNodeData.m_sc_cc_colloquial
    Else
        txtCostCenter.Text = txtCostCenter.Text & "  " & cNodeData.m_sc_CostCentreDescription
    End If
    
    Me.txtSortField.Text = cNodeData.m_sc_sort_field
    Me.txtSysStatus.Text = cNodeData.m_sc_system_status
    Me.txtUserStatus.Text = cNodeData.m_sc_user_status
    Me.txtFLOC.Text = cNodeData.m_sc_pk_func_loc
    
    Me.txtMWC.Text = cNodeData.m_sc_fk_main_work_centre
    Me.txtPlannerGroup.Text = cNodeData.m_sc_fk_planner_group
    
    
    btnNextPicture.Enabled = False
    btnPrevPicture.Enabled = False
    
    '==========
    ' Display any pictures
    '==========
    If cNodeData.Pictures.Count > 0 Then
        '==========
        ' Get the first picture in the list and display it
        '==========
        cNodeData.m_iPictureIndex = 1
        Call DisplayCurrentPicture(cNodeData)
        
        If cNodeData.Pictures.Count > 1 Then
            btnNextPicture.Enabled = True
        End If
        
    Else
        imgFLOCPic.Picture = LoadPicture("")
        lblPictureCountPosition.Caption = "No Pics"
    End If
End Sub

'==============================================================================
' SUBROUTINE
'   BlankFLOCData
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Public Sub BlankFLOCData()

    '=============
    ' Details Tab
    '=============
    lblFLOCName.Caption = ""
    txtCostCenter.Text = ""
    Me.txtSortField.Text = ""
    Me.txtFLOC.Text = ""
    
    Me.txtSysStatus.Text = ""
    Me.txtUserStatus.Text = ""
    Me.txtMWC.Text = ""
    Me.txtPlannerGroup.Text = ""
    
    lblGeneralInfo.Caption = ""
    
    '=============
    ' Pictures
    '=============
    imgFLOCPic.Picture = LoadPicture("")
    btnNextPicture.Enabled = False
    btnPrevPicture.Enabled = False
    lblPictureDate.Caption = ""
    lblPictureInfo.Caption = ""
    
    '=============
    ' Notification Tab
    '=============
    Call lbNotifications.Clear
    lblNotiSearchResultsDescription.Caption = ""
    
    '=============
    ' Parts Tab
    '=============
    Call lbPartsSearchResult.Clear
    lblSearchResultsDescription.Caption = ""
    
    '=============
    ' Work Orders Tab
    '=============
    Call lbWorkOrders.Clear
    lblWOSearchResultsDescription.Caption = ""
    
    '=============
    ' Logs Tab
    '=============
    Call lbLogs.Clear
    txtLogLongText.Text = ""
    txtLogWorkOrder.Text = ""
    txtLogReportedBy.Text = ""
    txtLogWorkTeam.Text = ""
    lblLogSearchResultsDescription.Caption = ""
    'txtLogWorkOrderType.Text = ""
    'txtLogWorkOrderShortText.Text = ""
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
    
    Call ufPictMgr.Configure("FLOC_PICTURES", Me.txtFLOC.Text)
    
    Call ufPictMgr.Show
    
    If Not (ufPictMgr Is Nothing) Then
        If ufPictMgr.DirtyFlag Then
            Call RefreshFlocView
        End If
    End If
        
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnPrevPicture_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnPrevPicture_Click()
    
#If DebugBadType = 0 Then
    Dim cNodeFLOC As clsNodeFLOC
#Else
    Dim cNodeFLOC As Object
#End If
    
    Dim scSelectedNodeKey As String
    
    scSelectedNodeKey = Me.m_oSelectedNode.Key
    Set oNodeData = Me.m_colNodeData.Item(scSelectedNodeKey)

    If oNodeData.m_iPictureIndex > 1 Then
        oNodeData.m_iPictureIndex = oNodeData.m_iPictureIndex - 1
        
        Call DisplayCurrentPicture(oNodeData)

        btnNextPicture.Enabled = True

        If oNodeData.m_iPictureIndex = 1 Then
            btnPrevPicture.Enabled = False
        End If

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

#If DebugBadType = 0 Then
    Dim cNodeFLOC As clsNodeFLOC
#Else
    Dim cNodeFLOC As Object
#End If

    Dim scSelectedNodeKey As String
    
    scSelectedNodeKey = Me.m_oSelectedNode.Key
    Set oNodeData = Me.m_colNodeData.Item(scSelectedNodeKey)

    If oNodeData.Pictures.Count > oNodeData.m_iPictureIndex Then
        oNodeData.m_iPictureIndex = oNodeData.m_iPictureIndex + 1
        
        Call DisplayCurrentPicture(oNodeData)
        'imgFLOCPic.Picture = LoadPicture(oNodeData.Pictures(oNodeData.m_iPictureIndex))
        'lblPictureCountPosition.Caption = oNodeData.m_iPictureIndex & " of " & oNodeData.Pictures.Count
        
        btnPrevPicture.Enabled = True
        If oNodeData.m_iPictureIndex = oNodeData.Pictures.Count Then
            btnNextPicture.Enabled = False
        End If

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
#If DebugBadType = 0 Then
Private Sub DisplayCurrentPicture(oNodeData As clsNodeFLOC)
#Else
Private Sub DisplayCurrentPicture(oNodeData As Object)
#End If

#If DebugBadType = 0 Then
    Dim cPicData As clsPictureData
#Else
    Dim cPicData As Object
#End If

    Dim iIndex As Long
    
    
    iIndex = oNodeData.m_iPictureIndex
        
    Set cPicData = oNodeData.Pictures(iIndex)
        
On Error GoTo could_not_find_picture
    imgFLOCPic.Picture = LoadPicture(cPicData.FullPath)
    lblPictureInfo.Caption = cPicData.description
    lblPictureDate.Caption = Format(cPicData.file_date, "d-mmm-yyyy")
    lblPictureDate.ForeColor = RGB(180, 180, 0)
    
    lblPictureCountPosition.Caption = iIndex & " of " & oNodeData.Pictures.Count
    
    Exit Sub
    
could_not_find_picture:
    lblPictureInfo.Caption = "Could not load picture"
    lblPictureDate.Caption = ""
    lblPictureCountPosition.Caption = "-"
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
    Dim cNodeFLOC As clsNodeFLOC
    Dim cPicData As clsPictureData
#Else
    Dim cNodeFLOC As Object
    Dim cPicData As Object
#End If

    Dim scTempFolderPath As String
    Dim scSelectedNodeKey As String
    
    '==========
    ' Get the Picture Data object
    '==========
    scSelectedNodeKey = m_oSelectedNode.Key
    Set oNodeData = m_colNodeData.Item(scSelectedNodeKey)

    If oNodeData.Pictures.Count < 1 Then
        Exit Sub
    End If
        
    Set cPicData = oNodeData.Pictures(oNodeData.m_iPictureIndex)
        
    '==========
    ' Get the users Temp folder.
    '==========
    scTempFolderPath = GetSpecialFolderPath(sfTemp)
    
    If DBCopyFile(cPicData.pk_file, scTempFolderPath, oDestFile) Then
    
        '===========
        ' Open the file in its default application.
        '===========
        Dim iShell As Object
        Set iShell = CreateObject("Shell.Application")
        
        Call iShell.Open(oDestFile.Path)
    End If
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbPartsSearchResult_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub lbPartsSearchResult_Click()
    Dim iMaterial As Long
    
On Error GoTo exit_cleanly
    Application.EnableEvents = False
    
    If m_ufPartDisplay Is Nothing Then
        Set m_ufPartDisplay = New ufPartDisplay
    End If
    
    Call Me.m_ufPartDisplay.ClearDisplay(False, True)
    
    If Me.lbPartsSearchResult.ListIndex < 0 Then
        Exit Sub
    End If
    
    '=============
    ' Get the material and populate the display form
    '=============
    iMaterial = Me.lbPartsSearchResult.List(Me.lbPartsSearchResult.ListIndex, 1)
    Call m_ufPartDisplay.DisplayMaterial(iMaterial)
    
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
' SUBROUTINE
'   GetFlocTable
'------------------------------------------------------------------------------
' DESCRIPTION
'   Extracts the complete FLOC table into a local table.
'==============================================================================
Public Sub GetFlocTable(Optional bForceUpdate As Boolean = False, Optional bIncludeRetired As Boolean = False)

#If DevelopMode = 1 Then
    Dim fso As Scripting.FileSystemObject
#Else
    Dim fso As Object
#End If

    Dim iMinRow As Long, iMaxRow As Long, iMinCol As Long, iMaxCol As Long, bEmpty As Boolean
    Dim scDBSource As String
    Dim scDBSource2 As String
    Dim scServer As String
    Dim scWorkstation As String
    Dim scDB As String
    Dim scSQLQuery As String
    Dim cnn As Excel.WorkbookConnection
    Dim scCnnName As String
    Dim qtTable As Excel.QueryTable
    Dim dtFLOCUpload As Date
    Dim bNeedDatabaseRefresh As Boolean
    Dim scFlocCacheFileName As String
    
    On Error GoTo exit_nicely
    
    '============
    ' Disable calculations etc
    '============
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
        
    '============
    ' When was the last FLOC upload? If recently, we don't reload.
    '============
    dtFLOCUpload = wsParameters.Range("LastFLOCUpload")
    
    bNeedDatabaseRefresh = True
    If ((Now() - dtFLOCUpload) * 24 < 8) And Not bForceUpdate Then
        bNeedDatabaseRefresh = False
    End If
    
    '===========
    ' Create a new instance of an Excel application to allow the
    ' cache instance to be made visible.
    '===========
    'Set appExcel = New Excel.Application
    'appExcel.Visible = False
    
    '============
    ' If we recently created a cache file, we open that cache file.
    '============
    If Not bNeedDatabaseRefresh Then
        scFlocCacheFileName = FlocCacheFilename
        
        '========
        ' Does the file exist?
        '========
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If Not fso.FileExists(scFlocCacheFileName) Then
            bNeedDatabaseRefresh = True
        Else
On Error GoTo corrupt_cache_file

            Set m_wbFlocTable = Workbooks.Open(scFlocCacheFileName, False, False)
            GoTo not_corrupt_cache_file
corrupt_cache_file:
            Set m_wbFlocTable = Workbooks.Add()
not_corrupt_cache_file:
On Error GoTo 0
        End If
    Else
        '============
        ' Create a new blank cache file.
        '============
        Set m_wbFlocTable = Workbooks.Add()
    End If
    
    Set m_wsFlocTable = m_wbFlocTable.Worksheets(1)
    
    '============
    ' We need to pull the complete FLOC data from the database.
    '------------
    ' Define the table columns. These need to be defined even if we don't
    ' actually do the load.
    '============
    m_iColFlocTable_pk_func_loc = 1
    m_iColFlocTable_Description = 2
    m_iColFlocTable_floc_type = 3
    m_iColFlocTable_fk_parent = 4
    m_iColFlocTable_parent_floc_type = 5
    m_iColFlocTable_fk_const_type = 6
    m_iColFlocTable_level_no = 7
    m_iColFlocTable_in_sap = 8
    
    '============
    ' If we appear to have a valid cache, we exit.
    '============
    If Not bNeedDatabaseRefresh Then
        Exit Sub
    End If
    
    '============
    ' Blank it.
    '============
    m_wsFlocTable.Cells.Clear
    
    '============
    ' Define the connection
    '============
    Call GetDBServer(ldMaintenance, scServer, scWorkstation)
    
    scDBSource = "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True" & _
            ";Data Source=" & scServer & ";Initial Catalog=maint"
    
    '============
    ' Define the query
    '============
    scSQLQuery = "SELECT pk_func_loc,description,floc_type,fk_parent,parent_floc_type,fk_const_type,level_no,in_sap " & _
            "FROM maint.dbo.v_floc " & _
            "WHERE is_archived = 0 " & _
            "ORDER BY floc_type, level_no, fk_parent, position, pk_func_loc"
            
    '============
    ' Create a list object to accept the results of the query, and then get it's QueryTable object.
    '============
    Dim oListObj
    
    Set qtTable = m_wsFlocTable.ListObjects.Add( _
        SourceType:=xlSrcQuery, _
        Source:=scDBSource, _
        Destination:=m_wsFlocTable.Range("$A$1")).QueryTable
    
    '============
    'Populate some major properties of the QueryTable.
    '============
    With qtTable
        .CommandText = scSQLQuery
        .CommandType = xlCmdSql
        
        '============
        'In order to see the output for the first time
        'we need to use the Refresh command.
        '============
        .Refresh
        .RefreshOnFileOpen = False
    End With
        
    '============
    ' Delete the connection so that we don't try to refresh again.
    '============
    Call m_wsFlocTable.Parent.Connections("Connection").Delete
    
    '============
    ' Save the updated Last Upload date
    '============
    wsParameters.Range("LastFLOCUpload") = Now()
    
    '============
    ' Save the whole add-in
    '============
    Dim wb As Excel.Workbook
    Set wb = wsParameters.Parent
    Call wb.Save
    
    Application.DisplayAlerts = False
    Call m_wbFlocTable.SaveAs(FlocCacheFilename)
    Application.DisplayAlerts = True
    'Call m_wbFlocTable.Close(True, FlocCacheFilename)
    
    Exit Sub
    
exit_nicely:

    Call MsgBox("Error occurred: " & Err.description & "(" & Err.Source & ")")
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnToggleDisplayMaterial_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Public Function FlocCacheFilename() As String

    Dim scFilename As String

    scFilename = GetSpecialFolderPath(sfTemp)
    If Not (Right(scFilename, 1) = "\") Then
        scFilename = scFilename & "\"
    End If
    scFilename = scFilename & "LihirAddin_V" & wsVer.Range("Ver_Maj") & "_FLOCCache.xlsx"
    
    FlocCacheFilename = scFilename

End Function

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnToggleDisplayMaterial_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
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
'   btnSearchNotifications_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnSearchNotifications_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

#If DebugBadType = 0 Then
    Dim cNodeFLOC As clsNodeFLOC
#Else
    Dim cNodeFLOC As Object
#End If

    Dim scSQLQuery As String
        
    Dim iLbLineCount As Long
    Dim scSelectedNodeKey As String
    
    '========
    ' Clear the listbox and results label
    '========
    Call Me.lbNotifications.Clear
    Me.lblSearchResultsDescription.Caption = ""
    
    scSelectedNodeKey = Me.m_oSelectedNode.Key
    Set oNodeData = Me.m_colNodeData.Item(scSelectedNodeKey)
    
    '========
    ' Construct the query
    '========
    scSQLQuery = "SELECT TOP " & g_cMaxNotificationCount & " pk_notification, short_text, notification_type, notification_date, required_end, noti_sys_status, primary_status " & _
            "FROM maint.dbo.v_notifications " & _
            " WHERE NOT (notification_type = 'AR') AND fk_func_loc "
            
    '========
    ' Use LIKE and a trailing % if we are looking at children as well.
    '========
    If Me.cbNotiIncludeSubFlocs.Value Then
        scSQLQuery = scSQLQuery & "LIKE '" & oNodeData.FuncLoc & "%'"
    Else
        scSQLQuery = scSQLQuery & "= '" & oNodeData.FuncLoc & "'"
    End If
    
    '========
    ' primary status
    '========
    Dim iCheckedCount As Long
    Dim scStatusIn As String
    
    iCheckedCount = 0
    If Me.cbNotiOutstanding.Value Then
        iCheckedCount = iCheckedCount + 1
        scStatusIn = "1"
    End If
    
    If Me.cbNotiInProcess.Value Then
        iCheckedCount = iCheckedCount + 1
        If Len(scStatusIn) > 0 Then scStatusIn = scStatusIn & ", "
        scStatusIn = scStatusIn & "2"
    End If
    
    If Me.cbNotiComplete.Value Then
        iCheckedCount = iCheckedCount + 1
        If Len(scStatusIn) > 0 Then scStatusIn = scStatusIn & ", "
        scStatusIn = scStatusIn & "3"
    End If
    
    If iCheckedCount < 3 Then
        scSQLQuery = scSQLQuery & " AND primary_status IN (" & scStatusIn & ") "
    End If
    
    If radDateIsRequiredEnd.Value Then
        lblNotiTableHeader_Date.Caption = "Required End Date"
        scSQLQuery = scSQLQuery & "ORDER BY required_end DESC"
    Else
        lblNotiTableHeader_Date.Caption = "Raised Date"
        scSQLQuery = scSQLQuery & "ORDER BY notification_date DESC"
    End If


    '========
    ' Execute the query
    '========
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)

    '========
    ' Display the results in the list box
    '========
    
    While Not rs.EOF
        Call Me.lbNotifications.AddItem(rs.Fields("pk_notification"))
        Me.lbNotifications.List(Me.lbNotifications.ListCount - 1, 1) = rs.Fields("short_text")
        Me.lbNotifications.List(Me.lbNotifications.ListCount - 1, 2) = rs.Fields("notification_type")
        
        Select Case rs.Fields("primary_status")
            Case 1
                Me.lbNotifications.List(Me.lbNotifications.ListCount - 1, 3) = "Outstanding"
            Case 2
                Me.lbNotifications.List(Me.lbNotifications.ListCount - 1, 3) = "In Process"
            Case Else
                Me.lbNotifications.List(Me.lbNotifications.ListCount - 1, 3) = "Complete"
        End Select
        
        If radDateIsRequiredEnd.Value Then
            If Not IsNull(rs.Fields("required_end")) Then
                Me.lbNotifications.List(Me.lbNotifications.ListCount - 1, 4) = Format(CDate(rs.Fields("required_end")), "d-mmm-yy")
            End If
        Else
            If Not IsNull(rs.Fields("notification_date")) Then
                Me.lbNotifications.List(Me.lbNotifications.ListCount - 1, 4) = Format(CDate(rs.Fields("notification_date")), "d-mmm-yy")
            End If
        End If
        Call rs.MoveNext
    Wend
    
    If radDateIsRequiredEnd.Value Then
        m_bNotificationOnDownloadIsRequiredEnd = True
    Else
        m_bNotificationOnDownloadIsRequiredEnd = False
    End If
    
    '========
    ' Complete the results label.
    '========
    Dim scFeedback As String
    
    If rs.RecordCount = 0 Then
        scFeedback = "Found no results for FLOC '" & oNodeData.FuncLoc & "'"
    ElseIf rs.RecordCount = g_cMaxNotificationCount Then
        scFeedback = "Results limited to first " & g_cMaxNotificationCount & " for FLOC '" & oNodeData.FuncLoc & "'"
    Else
        scFeedback = "Found " & rs.RecordCount & " results for FLOC '" & oNodeData.FuncLoc & "'"
    End If
    
    If Me.cbNotiIncludeSubFlocs.Value Then
        scFeedback = scFeedback & " and children."
    End If
    
    lblNotiSearchResultsDescription.Caption = scFeedback
    
    Call rs.Close
End Sub


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnSearchWorkOrders_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnDownloadNotifications_Click()
    Dim i As Long
    
    If lbNotifications.ListCount = 0 Then
        Exit Sub
    End If

    '=============
    ' Clear the entire contents of the sheet
    '=============
    Call wsFreeCanvas.Cells.Clear
    
    wsFreeCanvas.Cells(1, 1) = "Notification"
    wsFreeCanvas.Cells(1, 2) = "Description"
    wsFreeCanvas.Cells(1, 3) = "Type"
    wsFreeCanvas.Cells(1, 4) = "Primary Status"
    
    If m_bNotificationOnDownloadIsRequiredEnd Then
        wsFreeCanvas.Cells(1, 5) = "Required End Date"
    Else
        wsFreeCanvas.Cells(1, 5) = "Notification Date"
    End If
        
    With wsFreeCanvas.Range("A1:F1")
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    For i = 0 To (lbNotifications.ListCount - 1)
        wsFreeCanvas.Cells(i + 2, 1) = lbNotifications.List(i, 0)
        wsFreeCanvas.Cells(i + 2, 2) = lbNotifications.List(i, 1)
        wsFreeCanvas.Cells(i + 2, 3) = lbNotifications.List(i, 2)
        wsFreeCanvas.Cells(i + 2, 4) = lbNotifications.List(i, 3)
        wsFreeCanvas.Cells(i + 2, 5) = lbNotifications.List(i, 4)
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
    
    '============
    ' Autofit the columns
    '============
    For i = 1 To 6
        Call ws.Columns(i).EntireColumn.AutoFit
    Next

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnSearchWorkOrders_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnSearchWorkOrders_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

#If DebugBadType = 0 Then
    Dim cNodeFLOC As clsNodeFLOC
#Else
    Dim cNodeFLOC As Object
#End If

    Dim scSQLQuery As String
        
    Dim iLbLineCount As Long
    Dim scSelectedNodeKey As String
    
    '========
    ' Clear the listbox and results label
    '========
    Call Me.lbWorkOrders.Clear
    Me.lblWOSearchResultsDescription.Caption = ""
    
    scSelectedNodeKey = Me.m_oSelectedNode.Key
    Set oNodeData = Me.m_colNodeData.Item(scSelectedNodeKey)
    
    '========
    ' Construct the query
    '========
    scSQLQuery = "SELECT TOP " & g_cMaxWorkOrderCount & " pk_work_order, short_text, order_type, basic_start_date, sys_status, fk_func_loc, primary_status " & _
            "FROM maint.dbo.v_work_orders " & _
            " WHERE fk_func_loc "
            
    '========
    ' Use LIKE and a trailing % if we are looking at children as well.
    '========
    If Me.cbWkOrdIncludeSubFlocs.Value Then
        scSQLQuery = scSQLQuery & "LIKE '" & oNodeData.FuncLoc & "%'"
    Else
        scSQLQuery = scSQLQuery & "= '" & oNodeData.FuncLoc & "'"
    End If
    
    '========
    ' primary status
    '========
    Dim iCheckedCount As Long
    Dim scStatusIn As String
    
    iCheckedCount = 0
    If Me.cbWkOrdCreated Then
        iCheckedCount = iCheckedCount + 1
        scStatusIn = "1"
    End If
    
    If Me.cbWkOrdReleased Then
        iCheckedCount = iCheckedCount + 1
        If Len(scStatusIn) > 0 Then scStatusIn = scStatusIn & ", "
        scStatusIn = scStatusIn & "2"
    End If
    
    If Me.cbWkOrdTECO Then
        iCheckedCount = iCheckedCount + 1
        If Len(scStatusIn) > 0 Then scStatusIn = scStatusIn & ", "
        scStatusIn = scStatusIn & "3"
    End If
    
    If Me.cbWkOrdClosed Then
        iCheckedCount = iCheckedCount + 1
        If Len(scStatusIn) > 0 Then scStatusIn = scStatusIn & ", "
        scStatusIn = scStatusIn & "4"
    End If
    
    If iCheckedCount < 4 Then
        scSQLQuery = scSQLQuery & " AND primary_status IN (" & scStatusIn & ") "
    End If
    
    scSQLQuery = scSQLQuery & "ORDER BY basic_start_date DESC"

    '========
    ' Execute the query
    '========
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)

    '========
    ' Display the results in the list box
    '========
    While Not rs.EOF
        Call Me.lbWorkOrders.AddItem(rs.Fields("pk_work_order"))
        Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 1) = rs.Fields("short_text")
        Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 2) = rs.Fields("order_type")
        
        Select Case rs.Fields("primary_status")
            Case 1
                Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 3) = "Created"
            Case 2
                Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 3) = "Released"
            Case 3
                Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 3) = "TECO'd"
            Case Else
                Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 3) = "Closed"
        End Select
        
        If Not IsNull(rs.Fields("basic_start_date")) Then
            Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 4) = Format(CDate(rs.Fields("basic_start_date")), "d-mmm-yy")
        End If

        Call rs.MoveNext
    Wend
    
    
'Const g_cMaxLogCount As Long = 500

    '========
    ' Complete the results label.
    '========
    Dim scFeedback As String
    
    If rs.RecordCount = 0 Then
        scFeedback = "Found no results for FLOC '" & oNodeData.FuncLoc & "'"
    ElseIf rs.RecordCount = g_cMaxWorkOrderCount Then
        scFeedback = "Results limited to first " & g_cMaxWorkOrderCount & " for FLOC '" & oNodeData.FuncLoc & "'"
    Else
        scFeedback = "Found " & rs.RecordCount & " results for FLOC '" & oNodeData.FuncLoc & "'"
    End If
    
    If Me.cbWkOrdIncludeSubFlocs.Value Then
        scFeedback = scFeedback & " and children."
    End If
    
    lblWOSearchResultsDescription.Caption = scFeedback
    
    Call rs.Close
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbWorkOrders_DblClick
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub lbWorkOrders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim iWorkOrder As Long
    Dim oErr As typAppError
    
    If Me.lbWorkOrders.ListIndex >= 0 Then
        iWorkOrder = CLng(lbWorkOrders.List(lbWorkOrders.ListIndex, 0))
    
        If Not m_ufWorkOrder Is Nothing Then
            Call Unload(m_ufWorkOrder)
        End If
        
        Set m_ufWorkOrder = New ufWorkOrderDisplay
        
        oErr.BeSilent = False
On Error GoTo handle_error
        If Not m_ufWorkOrder.SetWorkOrder(iWorkOrder, oErr) Then
            Exit Sub
        End If
        
        GoTo show_form
        
handle_error:
        Call MsgBox("Order not valid or not found in the database")
        Exit Sub
        
show_form:
        Call m_ufWorkOrder.Show
    End If

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnSearchLogs_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnSearchLogs_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

#If DebugBadType = 0 Then
    Dim cNodeFLOC As clsNodeFLOC
#Else
    Dim cNodeFLOC As Object
#End If

    Dim scSQLAMQuery As String
    Dim scSQLLinkQuery As String
    Dim scSQLLocalQuery As String
    Dim scSQLQuery As String
    
    Dim scSelectedNodeKey As String
    Dim iShiftLogCount As Long
    Dim iARCount As Long
    
    '========
    ' Clear the list box
    '========
    Call Me.lbLogs.Clear
    
    '========
    ' Get the FLOC details
    '========
    scSelectedNodeKey = Me.m_oSelectedNode.Key
    Set oNodeData = Me.m_colNodeData.Item(scSelectedNodeKey)
    
    
    '========
    ' Construct the query.
    ' Note this is a complex query. It uses a Linked Server on our normal host,
    ' to the shiftlog server NMLLHRDB03. This was a little complex to setup.
    ' Review the document:
    ' Instructions for setting up a Linked Server between SQL Server 2012 and SQL Server 2000.docx
    ' for how to set it up.
    ' Once setup, you pull data from the linked server using the OPENQUERY T-SQL command.
    '========
    scSQLAMQuery = "SELECT TOP " & g_cMaxShiftLogCount & " log.dataID as ID, log.dataTeam as WorkTeam, log.dataDate as LoggedTime" & _
            ", log.dataLogger as LoggedBy, log.dataEquipment as Equip" & _
            ", log.dataOrder as WorkOrder, log.dataComments as Comments, equip.FLOC " & _
            "FROM [LGO-AM-RECORDS].dbo.ShiftLogData as log INNER JOIN " & _
            "[LGO-AM-RECORDS].dbo.EQUIPMENT as equip ON log.dataEquipment = equip.EquipNo WHERE equip.FLOC "
    scSQLLocalQuery = "SELECT TOP " & g_cMaxARCount & " pk_notification as ID, notification_date as LoggedTime, fk_main_work_centre as WorkTeam" & _
            ", reported_by as LoggedBy, sort_field as Equip, CAST(fk_order as varchar) as WorkOrder" & _
            ", short_text as Title, long_text as Comments, fk_func_loc as FLOC, 'AR' as DataSource " & _
            " FROM maint.dbo.v_notifications as ar WHERE ar.notification_type = 'AR' AND fk_func_loc "
    
    '========
    ' Use LIKE and a trailing % if we are looking at children as well.
    '========
    If Me.cbLogsIncludeSubFlocs.Value Then
        scSQLAMQuery = scSQLAMQuery & "LIKE ''" & oNodeData.FuncLoc & "%''"
        scSQLLocalQuery = scSQLLocalQuery & "LIKE '" & oNodeData.FuncLoc & "%'"
    Else
        scSQLAMQuery = scSQLAMQuery & "= ''" & oNodeData.FuncLoc & "''"
        scSQLLocalQuery = scSQLLocalQuery & "= '" & oNodeData.FuncLoc & "'"
    End If
    
        scSQLLinkQuery = "SELECT ID, LoggedTime, WorkTeam, LoggedBy, Equip, WorkOrder" & _
                ", left(Comments,50) as Title, Comments, FLOC, 'Log' as DataSource " & _
                " FROM OPENQUERY(NMLLHRDB03, ' " & scSQLAMQuery & " ' ) as ShiftLog "
    
    '========
    ' Construct the full Link Query
    '========
    If cbDisplayShiftLogs.Value And cbDisplayActivityReports.Value Then
        '========
        ' Do a UNION of the 2 queries
        '========
        scSQLQuery = scSQLLinkQuery & " UNION (" & scSQLLocalQuery & ") ORDER BY LoggedTime DESC"
    ElseIf cbDisplayShiftLogs.Value Then
        scSQLQuery = scSQLLinkQuery & " ORDER BY LoggedTime DESC"
    ElseIf cbDisplayActivityReports.Value Then
        scSQLQuery = scSQLLocalQuery & " ORDER BY LoggedTime DESC"
    Else
        Me.lblLogSearchResultsDescription.Caption = "No checkbox's selected"
        Call MsgBox("Need to select either 'ShiftLogs' or 'Act. Reports' (or both) to see entry's")
        Exit Sub
    End If
    
    
    '========
    ' DEBUG
    '========
'    txtLogLongText.Text = scSQLQuery
'    Exit Sub
    
    '========
    ' Execute the query
    '========
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)

    '========
    ' Display the results in the list box
    '========
    iShiftLogCount = 0
    iARCount = 0
    While Not rs.EOF
        Call Me.lbLogs.AddItem(rs.Fields("DataSource"))
        Me.lbLogs.List(Me.lbLogs.ListCount - 1, 1) = rs.Fields("ID")
        Me.lbLogs.List(Me.lbLogs.ListCount - 1, 2) = rs.Fields("Title")
        Me.lbLogs.List(Me.lbLogs.ListCount - 1, 3) = Format(TSQLDateStrToDate(rs.Fields("LoggedTime")), "d-mmm-yy")
        Me.lbLogs.List(Me.lbLogs.ListCount - 1, 4) = rs.Fields("WorkTeam")
        If Not IsNull(rs.Fields("Comments")) Then
            Me.lbLogs.List(Me.lbLogs.ListCount - 1, 5) = rs.Fields("Comments")
        Else
            Me.lbLogs.List(Me.lbLogs.ListCount - 1, 5) = ""
        End If
        If Not IsNull(rs.Fields("WorkOrder")) Then
            Me.lbLogs.List(Me.lbLogs.ListCount - 1, 6) = rs.Fields("WorkOrder")
            Me.lbLogs.List(Me.lbLogs.ListCount - 1, 7) = rs.Fields("LoggedBy")
            Me.lbLogs.List(Me.lbLogs.ListCount - 1, 8) = rs.Fields("WorkTeam")
        Else
            Me.lbLogs.List(Me.lbLogs.ListCount - 1, 6) = ""
            Me.lbLogs.List(Me.lbLogs.ListCount - 1, 7) = ""
            Me.lbLogs.List(Me.lbLogs.ListCount - 1, 8) = ""
        End If
        
        If rs.Fields("DataSource") = "AR" Then
            iARCount = iARCount + 1
        Else
            iShiftLogCount = iShiftLogCount + 1
        End If
        
        Call rs.MoveNext
    Wend
    
    '========
    ' Give user feedback.
    '========
    Dim scFeedback As String
    
    If rs.RecordCount = 0 Then
        scFeedback = "Found no results for FLOC '" & oNodeData.FuncLoc & "'"
    ElseIf (iARCount = g_cMaxARCount) Or (iShiftLogCount = g_cMaxShiftLogCount) Then
        If (iARCount = g_cMaxARCount) And (iShiftLogCount = g_cMaxShiftLogCount) Then
            scFeedback = "AR's and Shiftlog's were both limited to " & g_cMaxARCount & " and " & g_cMaxShiftLogCount & " respectively"
        ElseIf (iShiftLogCount = g_cMaxShiftLogCount) Then
            scFeedback = "Found " & iARCount & " AR's and Shiftlogs were limited to " & g_cMaxShiftLogCount
        Else
            scFeedback = "Found " & iShiftLogCount & " Shiftlogs and AR's were limited to " & g_cMaxARCount
        End If
    Else
        scFeedback = "Found " & iARCount & " AR's and " & iShiftLogCount & " Shiftlogs"
    End If
    scFeedback = scFeedback & " for FLOC '" & oNodeData.FuncLoc & "'"
    
    If Me.cbLogsIncludeSubFlocs.Value Then
        scFeedback = scFeedback & " and children."
    End If
    
    lblLogSearchResultsDescription.Caption = scFeedback

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbLogs_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub lbLogs_Click()
    txtLogLongText.Text = Me.lbLogs.List(Me.lbLogs.ListIndex, 5)
    txtLogWorkOrder.Text = Me.lbLogs.List(Me.lbLogs.ListIndex, 6)
    txtLogReportedBy.Text = Me.lbLogs.List(Me.lbLogs.ListIndex, 7)
    txtLogWorkTeam.Text = Me.lbLogs.List(Me.lbLogs.ListIndex, 8)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnLogShowWorkOrder_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnLogShowWorkOrder_Click()
    
    Dim iWorkOrder As Long
    Dim oErr As typAppError
    
    '===========
    ' Is the work order valid?
    '===========
On Error GoTo not_a_number
    iWorkOrder = Val(txtLogWorkOrder.Text)
    
    If (iWorkOrder > 10000000) And (iWorkOrder < 19999999) Then
        GoTo valid_work_order
    End If
    
not_a_number:
    Call MsgBox("This work order is not a valid work order!")
    Exit Sub
    
valid_work_order:

On Error GoTo 0 ' Clear error handling
On Error GoTo handle_error
    
    If Not m_ufWorkOrder Is Nothing Then
        Call Unload(m_ufWorkOrder)
    End If
    
    Set m_ufWorkOrder = New ufWorkOrderDisplay
    
    oErr.BeSilent = False

    If Not m_ufWorkOrder.SetWorkOrder(iWorkOrder, oErr, True) Then
        Exit Sub
    End If
    
    GoTo show_form
        
handle_error:
    Call MsgBox("Order not found in the database")
    Exit Sub
        
show_form:
    Call m_ufWorkOrder.Show

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnDownloadWorkOrders_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnDownloadWorkOrders_Click()
    Dim i As Long
    
    If lbWorkOrders.ListCount = 0 Then
        Exit Sub
    End If

    '=============
    ' Clear the entire contents of the sheet
    '=============
    Call wsFreeCanvas.Cells.Clear
    
    wsFreeCanvas.Cells(1, 1) = "Work Order"
    wsFreeCanvas.Cells(1, 2) = "Description"
    wsFreeCanvas.Cells(1, 3) = "Type"
    wsFreeCanvas.Cells(1, 4) = "Status"
    wsFreeCanvas.Cells(1, 5) = "Basic Start Date"
    
    With wsFreeCanvas.Range("A1:E1")
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    For i = 0 To (lbWorkOrders.ListCount - 1)
        wsFreeCanvas.Cells(i + 2, 1) = lbWorkOrders.List(i, 0)
        wsFreeCanvas.Cells(i + 2, 2) = lbWorkOrders.List(i, 1)
        wsFreeCanvas.Cells(i + 2, 3) = lbWorkOrders.List(i, 2)
        wsFreeCanvas.Cells(i + 2, 4) = lbWorkOrders.List(i, 3)
        wsFreeCanvas.Cells(i + 2, 5) = lbWorkOrders.List(i, 4)
    Next
    
    
'        Call Me.lbWorkOrders.AddItem(rs.Fields("pk_work_order"))
'        Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 1) = rs.Fields("short_text")
'        Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 2) = rs.Fields("order_type")
'
'        Select Case rs.Fields("primary_status")
'            Case 1
'                Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 3) = "Created"
'            Case 2
'                Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 3) = "Released"
'            Case 3
'                Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 3) = "TECO'd"
'            Case Else
'                Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 3) = "Closed"
'        End Select
'
'        If Not IsNull(rs.Fields("basic_start_date")) Then
'            Me.lbWorkOrders.List(Me.lbWorkOrders.ListCount - 1, 4) = Format(CDate(rs.Fields("basic_start_date")), "d-mmm-yy")
'        End If
    
    
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

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnDownloadFLOCParts_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnDownloadFLOCParts_Click()

    Dim i As Long

    If lbPartsSearchResult.ListCount = 0 Then
        Exit Sub
    End If

    '=============
    ' Clear the entire contents of the sheet
    '=============
    Call wsFreeCanvas.Cells.Clear
    
    wsFreeCanvas.Cells(1, 1) = "Material"
    wsFreeCanvas.Cells(1, 2) = "Description"
    wsFreeCanvas.Cells(1, 3) = "Usage"
    
    For i = 0 To (lbPartsSearchResult.ListCount - 1)
        wsFreeCanvas.Cells(i + 2, 1) = lbPartsSearchResult.List(i, 1)
        wsFreeCanvas.Cells(i + 2, 2) = lbPartsSearchResult.List(i, 2)
        wsFreeCanvas.Cells(i + 2, 3) = lbPartsSearchResult.List(i, 0)
    Next
    
    '============
    ' Make a copy of our FreeCanvas tab for the user
    '============
    Dim ws As Excel.Worksheet
    Dim wb As Excel.Workbook

    wsFreeCanvas.Copy
    
    Set ws = Application.ActiveSheet
    Set wb = Application.ActiveWorkbook
    
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    ws.Name = "SearchResults"
    ws.Cells(2, 1).Select

End Sub


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Terminate
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub UserForm_Terminate()

On Error GoTo cleanup_nicely
    If Not m_wbFlocTable Is Nothing Then
        Call m_wbFlocTable.Close(True, FlocCacheFilename)
    End If
cleanup_nicely:
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnFLOCSearch_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnFLOCSearch_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String
        
    
    Dim scSearch As String
    
    scSearch = Me.txtFLOCSearch.Text
    
    Call lbFLOCSearch.Clear
    
    If scSearch = "" Or scSearch = "*" Then
        Call MsgBox("You'll need to refine your search a little more than that!!")
        Exit Sub
    End If
    
    
    scSearch = Replace(scSearch, "*", "%")
    scSearch = Replace(scSearch, "'", "''")
    
    scSQLQuery = "SELECT TOP " & g_cMaxFLOCSearch & " pk_func_loc from maint.dbo.t_func_loc WHERE (pk_func_loc LIKE '" & _
        scSearch & "') OR ((sort_field <> '') AND (sort_field LIKE '" & scSearch & "'))"
        
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)
    
    If rs.RecordCount = 0 Then
        Call MsgBox("No items found. Broaden your search.")
        Exit Sub
    End If
    
    While Not rs.EOF
        Call lbFLOCSearch.AddItem(rs.Fields("pk_func_loc"))
    
        Call rs.MoveNext
    Wend
    
    If rs.RecordCount = g_cMaxFLOCSearch Then
        Call MsgBox("Limited the search to " & g_cMaxFLOCSearch & " entries.")
    End If
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbFLOCSearch_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub lbFLOCSearch_Click()

#If DevelopMode = 1 Then
    Dim oNode As MSComctlLib.Node
#Else
    Dim oNode As Object
#End If

    Dim scPk_func_loc As String
    
    scPk_func_loc = Me.lbFLOCSearch.List(Me.lbFLOCSearch.ListIndex)
    
On Error GoTo floc_not_in_tree
    Set oNode = TreeView1.Nodes.Item("N" & scPk_func_loc)
    
    oNode.Selected = True
    Call oNode.EnsureVisible
    
    Call TreeView1_Click
    Call TreeView1.SetFocus
    
    Exit Sub
    
floc_not_in_tree:
    Call MsgBox("The node was not found in the tree. It may because the node is retired and the tree only shows active functional locations")
End Sub

Private Sub CommandButton1_Click()
    Me.TreeView1.SingleSel = True
    Me.TreeView1.FullRowSelect = False
    TreeView1.Nodes.Item("N2301-1711-DG07").Selected = True
    Call TreeView1.Nodes.Item("N2301-1711-DG07").EnsureVisible
    Call TreeView1_Click
    Call TreeView1.SetFocus
    'Me.TreeView1.Nodes.Item("N2301-1711-DG07").Expanded = True
End Sub

