VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufRating 
   Caption         =   "UserForm1"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2745
   OleObjectBlob   =   "ufRating.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufRating"
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
Private m_bIgnoreStarClicks As Boolean

Private m_cbaColloquialStarsCheckBoxs(1 To 5) As Object
Private m_bSuccess As Boolean

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   Initialise
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Public Sub Initialise(scTitle As String, Optional iRating As Long = 3, Optional scComment As String = "")
    m_bIgnoreStarClicks = False
    
    Call SetStarRating(iRating)
    
    Me.txtComments.Text = scComment
    Me.Caption = scTitle
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
Public Property Get StarRating() As Long
    Dim i As Long

    For i = 5 To 1 Step -1
        If m_cbaColloquialStarsCheckBoxs(i).Value Then
            Exit For
        End If
    Next
    
    StarRating = i

End Property

Public Property Get Success() As Boolean
    Success = m_bSuccess
End Property

Public Property Get Comment() As String
    Comment = Me.txtComments.Text
End Property

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the star rating based on the highest checked checkbox.
'==============================================================================
Private Sub UserForm_Initialize()

    Set m_cbaColloquialStarsCheckBoxs(1) = Me.CheckBox1
    Set m_cbaColloquialStarsCheckBoxs(2) = Me.CheckBox2
    Set m_cbaColloquialStarsCheckBoxs(3) = Me.CheckBox3
    Set m_cbaColloquialStarsCheckBoxs(4) = Me.CheckBox4
    Set m_cbaColloquialStarsCheckBoxs(5) = Me.CheckBox5
    
    m_bIgnoreStarClicks = False

    Call SetStarRating(3)
    
    Call RemoveUserformCloseButton(Me)
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnRate_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnRate_Click()
    m_bSuccess = True
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnCancel_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnCancel_Click()
    m_bSuccess = False
    Call Hide
End Sub

