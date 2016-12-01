VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSelectDept 
   Caption         =   "Select Department"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   OleObjectBlob   =   "ufSelectDept.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSelectDept"
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
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public m_bDeptSelected As Boolean
Public DeptID As Long
Public m_scDeptName As String

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnUploadOrgChart_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Callback for btnUploadOrgChart_Click onAction
'==============================================================================
Private Sub UserForm_Initialize()
    m_bDeptSelected = False
    DeptID = -1
    
    Call RemoveUserformCloseButton(Me)
End Sub

'==============================================================================
' SUBROUTINE
'   Initialise
'------------------------------------------------------------------------------
' DESCRIPTION
'   Set some basic parameters
'==============================================================================
Public Sub Initialise(scFormCaption As String, Optional scIntroMessage As String = "")

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
        
    Dim i As Long
    
    Me.Caption = scFormCaption
    Me.txtNewDept.Text = ""
    
    If scIntroMessage = "" Then
        lblExplanation.Visible = False
        lbDept.Top = 0
        lbDept.Height = 153
    Else
        lblExplanation.Visible = True
        lblExplanation.Caption = scIntroMessage
        lbDept.Top = 42
        lbDept.Height = 107
    End If
    
    '==========
    ' Connect to the database
    '==========
    If Not (ConnectToDB(eLihirDatabases.ldPeople, cnn, True)) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Sub
    End If
    
    scSQLQuery = "SELECT * FROM dbo.t_department WHERE fk_company = 1 order by department_name"
    
    Set rs = CreateObject("ADODB.Recordset")
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    i = 0
    While Not rs.EOF
        Call Me.lbDept.AddItem
        lbDept.Column(0, i) = rs.Fields("department_name")
        lbDept.Column(1, i) = rs.Fields("pk_department")
        
        i = i + 1
        Call rs.MoveNext
    Wend
    Call rs.Close
    
    Call RemoveUserformCloseButton(Me)

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnCancel_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnCancel_Click()
    m_bDeptSelected = False
    
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnOK_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnOK_Click()

    If (lbDept.ListIndex < 0) Then
        Call MsgBox("You have not selected a department.")
        Exit Sub
    ElseIf (Me.txtNewDept.Text <> "") Then
        Select Case MsgBox("You have entered text in the New Dept field. Did you mean to press New?", vbYesNo)
            Case vbNo
                ' Proceed through to assign the ID and hide.
            Case Else
                Exit Sub
        End Select
    End If
    
    'DeptID = lbDept.Column(1, lbDept.ListIndex)
    m_scDeptName = lbDept.List(lbDept.ListIndex, 0)
    DeptID = lbDept.List(lbDept.ListIndex, 1)
    
    m_bDeptSelected = True
    
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnNew_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnNew_Click()
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
        
    Dim i As Long
    Dim scNewName As String
    
    '==========
    ' Connect to the database
    '==========
    If Not (ConnectToDB(eLihirDatabases.ldPeople, cnn, True)) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Sub
    End If
    
    scNewName = Me.txtNewDept.Text
    
    scSQLQuery = "SET NOCOUNT ON; " & _
        "INSERT INTO people.dbo.t_department " & _
        "(department_name, name_short, fk_company) " & _
        "VALUES " & _
        "('" & Replace(scNewName, "'", "''") & "', '" & UCase(Left(scNewName, 3)) & "', 1); " & _
        "SELECT SCOPE_IDENTITY() as pk_department;"

On Error GoTo cleanup_nicely
    Set rs = cnn.Execute(scSQLQuery)
    
    '===========
    ' Get the newly created identity
    '===========
    If Not rs.EOF Then
        Me.DeptID = rs.Fields("pk_department").Value
        Me.m_bDeptSelected = True
    Else
        Call MsgBox("Unable to create new department. No explanation available")
        Me.m_bDeptSelected = False
    End If
    Call rs.Close
    Call Hide
    
    Exit Sub
    
cleanup_nicely:
    Call MsgBox("Create new dept failed with error: " & Err.description)
    Me.m_bDeptSelected = False
    
End Sub

'==============================================================================
' FUNCTION
'   GetUserSelectedDeptID
'------------------------------------------------------------------------------
' DESCRIPTION
'   This should be the main entry point for this form.
' Just create an instant of the form and call this method.
'==============================================================================
Public Function GetUserSelectedDeptID(scFormCaption As String, ByRef iDeptID As Long, ByRef scDept As String, Optional scExplanation As String = "") As Boolean

    Call Me.Initialise(scFormCaption, scExplanation)
    
    Call Me.Show
    
    If Me.m_bDeptSelected Then
        iDeptID = Me.DeptID
        scDept = Me.m_scDeptName
        GetUserSelectedDeptID = True
    Else
        GetUserSelectedDeptID = False
    End If
End Function


