Attribute VB_Name = "modPermissions"
Option Explicit

Public Enum eLIHPermissions
    PeoplePhotoEditing = 1
End Enum

'==============================================================================
' SUBROUTINE
'   HasPermissions
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns true if the specified username has the specified permission.
' Otherwise returns false.
'==============================================================================

Public Function HasPermissions(scUsername As String, ePermission As eLIHPermissions) As Boolean

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    Dim scPkPermissions As String
    
    '=============
    ' Check the access level for this user.
    '=============
    Call GetUserName(scUsername)
    If Not PermissionsStringFromPermission(ePermission, scPkPermissions) Then
        HasPermissions = False
        Exit Function
    End If
    
    scSQLQuery = "SELECT * FROM people.dbo.t_staff_permissions WHERE pk_user_name = '" & Replace(scUsername, "'", "''") & "' AND pk_permissions = '" & scPkPermissions & "'"
    Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)
    
    If Not rs.EOF Then
        If rs.Fields("status") = True Then
            HasPermissions = True
            Exit Function
        End If
    End If

    HasPermissions = False

End Function

'==============================================================================
' SUBROUTINE
'   HasPermissions
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns true if the specified username has the specified permission.
' Otherwise returns false.
'==============================================================================
Private Function PermissionsStringFromPermission(ePermission As eLIHPermissions, ByRef scPermission) As Boolean
    Select Case ePermission
        Case eLIHPermissions.PeoplePhotoEditing
            scPermission = "PEOPLE_PHOTO_EDITING"
        Case Else
            scPermission = "<undefined_permission>"
            PermissionsStringFromPermission = False
            Exit Function
    End Select

    PermissionsStringFromPermission = True
End Function
