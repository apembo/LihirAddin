VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSelectStandardRole 
   Caption         =   "Select a Standard Role"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6615
   OleObjectBlob   =   "ufSelectStandardRole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSelectStandardRole"
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
Private m_bSuppressEvents As Boolean

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the class is created
'==============================================================================
Private Sub UserForm_Initialize()
    m_bSuppressEvents = False
    Call RemoveUserformCloseButton(Me)
End Sub

'==============================================================================
' SUBROUTINE
'   InitialiseForm
'------------------------------------------------------------------------------
' DESCRIPTION
'   Initialises the form, including pulling in lists from the database.
'==============================================================================
Public Sub InitialiseForm(scJobTitle As String)
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    
On Error GoTo cleanup_nicely
    m_bSuppressEvents = True

    txtJobTitle.Text = scJobTitle
    
    '===============
    ' Populate the filter combobox
    '===============
    Call cmbCat1Filter.Clear
    
    scSQLQuery = "SELECT DISTINCT cat1 FROM [people].[dbo].[t_pos_category] order by cat1"
    
    Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)
    
    Call cmbCat1Filter.AddItem("<No Filter>")
    
    While Not rs.EOF
        Call cmbCat1Filter.AddItem(rs.Fields("cat1"))
        Call rs.MoveNext
    Wend
    cmbCat1Filter.ListIndex = 0
    
    '===============
    ' Populate the filter position category list
    '===============
    Call Me.lbPosCategory.Clear
    
    scSQLQuery = "SELECT * FROM [people].[dbo].[t_pos_category] order by pos_level, pk_pos_category"
    
    Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)
    
    While Not rs.EOF
        Call lbPosCategory.AddItem(rs.Fields("pk_pos_category"))
        lbPosCategory.List(lbPosCategory.ListCount - 1, 1) = rs.Fields("pos_level")
        
        Call rs.MoveNext
    Wend
    
cleanup_nicely:
    m_bSuppressEvents = False
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   cmbCat1Filter_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   ComboBox change event handler. Apply the filter based on the selected
' Cat1 value.
'==============================================================================
Private Sub cmbCat1Filter_Change()
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    
    If m_bSuppressEvents Then
        Exit Sub
    End If

    scSQLQuery = "SELECT * FROM [people].[dbo].[t_pos_category] " 'order by pos_level"
    
    If cmbCat1Filter.Text <> "<No Filter>" Then
        scSQLQuery = scSQLQuery & "WHERE cat1 = '" & cmbCat1Filter.Text & "'"
    End If
    scSQLQuery = scSQLQuery & " ORDER BY pos_level"
    
    Call Me.lbPosCategory.Clear
    
    Call GetDBRecordSet(ldPeople, cnn, scSQLQuery, rs)
    
    While Not rs.EOF
        Call lbPosCategory.AddItem(rs.Fields("pk_pos_category"))
        lbPosCategory.List(lbPosCategory.ListCount - 1, 1) = rs.Fields("pos_level")
        Call rs.MoveNext
    Wend
    
End Sub

'==============================================================================
' FUNCTION
'   GetChosenPosCategory
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the users chosen standard position category
'==============================================================================
Public Function GetChosenPosCategory()
    If lbPosCategory.ListIndex > 0 Then
        GetChosenPosCategory = lbPosCategory.List(lbPosCategory.ListIndex, 0)
    Else
        GetChosenPosCategory = ""
    End If
End Function

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnOK_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnOK_Click()
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnOK_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnCancel_Click()
    Call Hide
End Sub

'==============================================================================
' FUNCTION
'   GetStandardPositionCategoryForJobTitle
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the users chosen standard position category
'==============================================================================
Public Function GetStandardPositionCategoryForJobTitle(scJobTitle As String, ByRef scPosCategory As String)
    Call InitialiseForm(scJobTitle)
    
    Call Show
    
    scPosCategory = GetChosenPosCategory
    If Len(scPosCategory) > 0 Then
        GetStandardPositionCategoryForJobTitle = True
    Else
        GetStandardPositionCategoryForJobTitle = False
    End If
End Function


