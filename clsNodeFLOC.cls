VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNodeFLOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
' CLASS
'   clsNodeFLOCData
'------------------------------------------------------------------------------
' DESCRIPTION
'   Simple data structure for the Adam's Treeview nodes. Just want to be
' create a new one each time.
'------------------------------------------------------------------------------
' VERSION
'   1.0     31-07-2015  Changes undocumented before this date
'==============================================================================

'==============================================================================
' PRIVATE MEMBER VARIABLES
'==============================================================================
Private m_bDataPopulated As Boolean
Private m_colPictures As VBA.Collection

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public m_sc_pk_func_loc As String
Public m_sc_description As String
Public m_sc_floc_type As String
Public m_sc_fk_parent As String
Public m_sc_parent_floc_type As String
Public m_sc_object_type As String
Public m_sc_sort_field As String
Public m_sc_fk_cost_centre As String
Public m_sc_CostCentreDescription As String
Public m_sc_CCArea As String
Public m_sc_CCDept As String
Public m_sc_fk_main_work_centre As String
Public m_sc_fk_planner_group As String
Public m_sc_system_status As String
Public m_sc_user_status As String
Public m_sc_position As String
Public m_i_fk_const_type As Long
Public m_sc_NAVIParent As String
Public m_i_level_no As Long
Public m_b_is_archived As Boolean
Public m_b_in_sap As Boolean
Public m_sc_cc_colloquial As String

Public m_iLevelsDrilledDown As Long
Public m_iPictureIndex As Long

#If DevelopMode = 1 Then
    Public TreeViewNode As MSComctlLib.Node
#Else
    Public TreeViewNode As Object
#End If


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   Class_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Initialize method
'==============================================================================
Private Sub Class_Initialize()
    m_iLevelsDrilledDown = 0
    
    m_bDataPopulated = False
    
    Set m_colPictures = New VBA.Collection

End Sub


'==============================================================================
' Property
'   Key
'------------------------------------------------------------------------------
' DESCRIPTION
'   This is a simplification of SetNode as, in this current implementation,
' the only actual requirement of this class is the Key
'==============================================================================
Property Get Key() As String
    Key = "N" & m_sc_pk_func_loc
End Property
Property Get ParentKey() As String
    ParentKey = "N" & m_sc_fk_parent
End Property

Property Let FuncLoc(scInput As String)
    m_sc_pk_func_loc = scInput
End Property
Property Get FuncLoc() As String
    FuncLoc = m_sc_pk_func_loc
End Property

Property Let Name(scInput As String)
     m_sc_description = scInput
End Property
Property Get Name() As String
    Name = m_sc_description
End Property

Property Let FLOCType(scInput As String)
    m_sc_floc_type = scInput
End Property
Property Get FLOCType() As String
    FLOCType = m_sc_floc_type
End Property

Property Let FLOCParent(scInput As String)
    m_sc_fk_parent = scInput
End Property
Property Get FLOCParent() As String
    FLOCParent = m_sc_fk_parent
End Property

Property Let FLOCParentType(scInput As String)
    m_sc_parent_floc_type = scInput
End Property
Property Get FLOCParentType() As String
    FLOCParentType = m_sc_parent_floc_type
End Property

Property Get IsNAVI() As Boolean
    If m_sc_floc_type = "NAVI" Then
        IsNAVI = True
    Else
        IsNAVI = False
    End If
End Property
Property Get ParentIsNAVI() As Boolean
    If m_sc_parent_floc_type = "NAVI" Then
        ParentIsNAVI = True
    Else
        ParentIsNAVI = False
    End If

End Property

'==============================================================================
' FUNCTION
'   PopulateData
'------------------------------------------------------------------------------
' DESCRIPTION
'   Gets the relevant data out of the database.
'==============================================================================
Public Function PopulateData() As Boolean
    
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
    Dim i As Long
    
    '========
    ' Construct the query
    '========
    scSQLQuery = "SELECT * FROM maint.dbo.v_floc_with_cc WHERE pk_func_loc = '" & Me.FuncLoc & "'"
    
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)
    
    If rs.EOF Then
        PopulateData = False
        Exit Function
    End If
    
    For i = 0 To rs.Fields.Count - 1
        Select Case rs.Fields(i).Name
            Case "pk_func_loc"
                m_sc_pk_func_loc = rs.Fields(i).Value
            Case "description"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_description = rs.Fields(i).Value
            Case "floc_type"
                m_sc_floc_type = rs.Fields(i).Value
            Case "fk_parent"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_fk_parent = rs.Fields(i).Value
            Case "object_type"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_object_type = rs.Fields(i).Value
            Case "sort_field"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_sort_field = rs.Fields(i).Value
            Case "fk_cost_centre"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_fk_cost_centre = rs.Fields(i).Value
            Case "CostCentreDescription"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_CostCentreDescription = rs.Fields(i).Value
            Case "CCArea"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_CCArea = rs.Fields(i).Value
            Case "CCDept"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_CCDept = rs.Fields(i).Value
            Case "fk_main_work_centre"
                m_sc_fk_main_work_centre = rs.Fields(i).Value
            Case "fk_planner_group"
                m_sc_fk_planner_group = rs.Fields(i).Value
            Case "system_status"
                m_sc_system_status = rs.Fields(i).Value
            Case "user_status"
                m_sc_user_status = rs.Fields(i).Value
            Case "position"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_position = rs.Fields(i).Value
            Case "fk_const_type"
                If Not IsNull(rs.Fields(i).Value) Then m_i_fk_const_type = rs.Fields(i).Value
            Case "fk_parent_navi"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_NAVIParent = rs.Fields(i).Value
            Case "level_no"
                If Not IsNull(rs.Fields(i).Value) Then m_i_level_no = rs.Fields(i).Value
            Case "is_archived"
                If Not IsNull(rs.Fields(i).Value) Then m_b_is_archived = rs.Fields(i).Value
            Case "in_sap"
                If Not IsNull(rs.Fields(i).Value) Then m_b_in_sap = rs.Fields(i).Value
            Case "parent_floc_type"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_parent_floc_type = rs.Fields(i).Value
            Case "cc_colloquial"
                If Not IsNull(rs.Fields(i).Value) Then m_sc_cc_colloquial = rs.Fields(i).Value
                
                
        End Select
    Next
    
    '============
    ' Any Pictures?
    '============

    '============
    ' Clear the collection in case the user is calling this method to refresh
    ' the data
    '============
'    While m_colPictures.Count > 0
'        Call m_colPictures.Remove(1)
'    Wend

    Set m_colPictures = Nothing
    Set m_colPictures = New VBA.Collection
    
    '============
    ' Now read from the database
    '============
    scSQLQuery = "SELECT * FROM maint.dbo.v_mapped_files " & _
            "WHERE id = '" & Replace(Me.FuncLoc, "'", "''") & "' AND base_path_category = 'FLOC_PICTURES' " & _
            "ORDER BY file_order, file_date"
            
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)
  
    While Not rs.EOF
        Set cPicData = New clsPictureData
        
        Call cPicData.PopulateFromRecordset(rs)
        
        Call m_colPictures.Add(cPicData, cPicData.Key)
        Call rs.MoveNext
    Wend
    
    m_bDataPopulated = True
    PopulateData = True
    
End Function

Public Property Get DataPopulated() As Boolean
    DataPopulated = m_bDataPopulated
End Property

Public Property Get Pictures() As VBA.Collection
    Set Pictures = m_colPictures

End Property

