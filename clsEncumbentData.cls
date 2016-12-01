VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEncumbentData"
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
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public Role As String
Public Section As String
Public Dept As String
Public ParentRole As String
Public RoleID As Long
Public SectionID As Long
Public DeptID As Long
Public ParentRoleID As Long

Public PosNo As String
Public JobTitle As String
Public ParentPosNo As String
Public ParentJobTitle As String
Public OrgUnit As String
Public OrgUnitID As Long

Public StaffID As Long
Public SapID As String
Public FullName As String
Public SapUserName As String
Public RosterID As String
Public RosterName As String
Public RosterDescription As String
Public RosterCycleDays As Long
Public RosterSeedDate As Date
Public IsContractor As Boolean
Public DoesProjects As Boolean
Public DoesActions As Boolean

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   Class_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the class is created.
'==============================================================================
Private Sub Class_Initialize()
    PosNo = -1
End Sub

'==============================================================================
' SUBROUTINE
'   Populate
'------------------------------------------------------------------------------
' DESCRIPTION
'   Takes an entry from the people.dbo.[v_roles_positions_and_staff] view
' and populates this class instance.

'==============================================================================
Public Sub Populate(oFields As Object) ' Fields)

    '=============
    ' Role
    '=============
    Role = oFields("RoleName")
    Section = oFields("Section")
    Dept = oFields("Dept")
    
    DeptID = oFields("DeptID")
    RoleID = oFields("RoleID")
    SectionID = oFields("SectionID")
    
    If Not IsNull(oFields("ParentRole")) Then
        ParentRole = oFields("ParentRole")
        ParentRoleID = oFields("ParentRoleID")
    End If
    
    '=============
    ' Position
    '=============
    PosNo = oFields("PosNo")
    JobTitle = oFields("PosJobTitle")
    OrgUnit = oFields("OrgUnit")
    OrgUnitID = oFields("OrgUnitID")
    
    If Not IsNull(oFields("ParentPosNo")) Then
        ParentPosNo = oFields("ParentPosNo")
        ParentJobTitle = oFields("PosParentJobTitle")
    Else
        ParentPosNo = -1
    End If
    
    '=============
    ' Staff/Person
    '=============
    If Not IsNull(oFields("StaffID")) Then
        StaffID = oFields("StaffID")
        SapID = oFields("PersonSapID")
        FullName = oFields("FullName")
        
        If Not IsNull(oFields("PersonSapName")) Then
            SapUserName = oFields("PersonSapName")
        End If
        
        If Not IsNull(oFields("RosterID")) Then
            RosterID = oFields("RosterID")
            RosterName = oFields("RosterName")
            RosterDescription = oFields("RosterDescription")
            RosterCycleDays = oFields("RosterCycleDays")
            RosterSeedDate = oFields("RosterSeedDate")
        Else
            RosterID = -1
        End If
        
        IsContractor = oFields("IsContractor")
        DoesProjects = oFields("do_projects")
        DoesActions = oFields("do_actions")
    Else
        StaffID = -1
    End If
End Sub
