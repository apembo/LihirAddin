VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufRatingsManager 
   Caption         =   "Ratings Manager"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5580
   OleObjectBlob   =   "ufRatingsManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufRatingsManager"
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
Private m_bIgnoreCommentChange As Boolean

Private m_eDatabase As eLihirDatabases
Private m_scSubjectCategory As String
Private m_iItemID As Long
Private m_scCurrentUser As String
Private m_scRatingAuthor As String
Private m_iPkRating As Long
Private m_bCurrentUserRatingFound As Boolean

Private m_cbaColloquialStarsCheckBoxs(1 To 5) As Object
Private m_bSuccess As Boolean

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the star rating based on the highest checked checkbox.
'==============================================================================
Private Sub UserForm_Initialize()

    Set m_cbaColloquialStarsCheckBoxs(1) = Me.cbStar1
    Set m_cbaColloquialStarsCheckBoxs(2) = Me.cbStar2
    Set m_cbaColloquialStarsCheckBoxs(3) = Me.cbStar3
    Set m_cbaColloquialStarsCheckBoxs(4) = Me.cbStar4
    Set m_cbaColloquialStarsCheckBoxs(5) = Me.cbStar5
    
    Call lbRatings.Clear
    btnRatingAction.Enabled = False
    
    lblTitle.Caption = ""
    
    m_bIgnoreStarClicks = False
    
'    m_bOperatingMode = opNoRatingSelected

    Call SetStarRating(0, False)
    
    '==========
    ' Remove the X button from the userform. Clicking the X button causes
    ' memory issues.
    '==========
    Call RemoveUserformCloseButton(Me)
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   Initialise
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Public Sub Initialise(scRatingsFor As String, scCategory As String, iID As Long, eDB As eLihirDatabases)

    m_bIgnoreStarClicks = False
    
    m_eDatabase = eDB
    m_scSubjectCategory = scCategory
    m_iItemID = iID
    lblTitle.Caption = scRatingsFor
    
    '=========
    ' Get the user
    '=========
    Call GetUserName(m_scCurrentUser)
    
    '=========
    ' Display all existing ratings.
    '=========
    Call Refresh
    
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbRatings_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub Refresh()
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String
    
    Dim scAuthor As String
    
    scSQLQuery = "SELECT * FROM dbo.t_rating " & _
        "WHERE fk_rated_item = " & m_iItemID & " AND category = '" & m_scSubjectCategory & "'"
    
    Call GetDBRecordSet(m_eDatabase, cnn, scSQLQuery, rs)
    
    Call lbRatings.Clear
    
'    bUserRatingFound = False
    
    Me.btnRatingAction.Caption = "Edit New"
    btnRatingAction.Enabled = True
    
    Call SetStarCheckBoxEnabled(False)
    Me.txtComments.Enabled = False
    
    m_bCurrentUserRatingFound = False
    
    While Not rs.EOF
        scAuthor = rs.Fields("author")
        Call lbRatings.AddItem(rs.Fields("author"))
        lbRatings.List(lbRatings.ListCount - 1, 1) = rs.Fields("rating_out_of_5")
        lbRatings.List(lbRatings.ListCount - 1, 2) = rs.Fields("comment")
        lbRatings.List(lbRatings.ListCount - 1, 3) = rs.Fields("pk_rating")
        
        '=========
        ' If we find the current user as an author, we populate the rating display.
        ' We'll also set this line as selected after finishing populating the list.
        '=========
        If scAuthor = m_scCurrentUser Then
            m_bCurrentUserRatingFound = True
            Me.btnRatingAction.Caption = "Update"
            Me.btnRatingAction.Enabled = False
'            Call SetStarCheckBoxEnabled(False)
'            Me.txtComments.Enabled = False
            
'            m_bIgnoreStarClicks = True
'            Call SetStarRating(rs.Fields("rating_out_of_5"))
'            txtComments.Text = rs.Fields("comment")
'            frmRating.Caption = "Rating by " & scAuthor
'            bUserRatingFound = True
'            iUserRatingLine = lbRatings.ListCount
'            m_bIgnoreStarClicks = False
        End If
        
        Call rs.MoveNext
    Wend
    
'    If bUserRatingFound Then
'        lbRatings.ListIndex = iUserRatingLine
'    End If
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbRatings_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub lbRatings_Click()

    m_bIgnoreStarClicks = True
    
    m_scRatingAuthor = lbRatings.List(lbRatings.ListIndex, 0)
    frmRating.Caption = "Rating of " & m_scRatingAuthor
    
    '=========
    ' Update the rating display to reflect the settings of this selected
    ' rating.
    '=========
    Call SetStarRating(lbRatings.List(lbRatings.ListIndex, 1), False)
    m_bIgnoreCommentChange = True
    txtComments.Text = lbRatings.List(lbRatings.ListIndex, 2)
    m_bIgnoreCommentChange = False
    
    m_iPkRating = lbRatings.List(lbRatings.ListIndex, 3)
    
    '=========
    ' Update the rating display to reflect the settings of this selected
    ' rating.
    '=========
    If m_scCurrentUser = m_scRatingAuthor Then
        '=========
        ' The user has selected his/her own rating.
        ' Because the Refresh method determined that this users rating existed,
        ' the button will have caption of "Update".
        ' We disable the Action button for now as it will act as the dirty
        ' flag. We enable the comment textbox and the star checkboxes.
        '=========
        txtComments.Enabled = True
        Call SetStarCheckBoxEnabled(True)
        btnRatingAction.Enabled = False
    Else
        '=========
        ' The user has selected another rating.
        ' If the user rating is listed elsewhere, the action button will have
        ' a caption of 'Update'. We don't want them to be able update another
        ' users entry so we disable the action button as well as the checkboxes
        ' and comment textbox.
        ' If the user rating is not listed elsewhere, the Action button will
        ' have a caption of "Edit New". In this case we enable it, to allow the
        ' user to begin editing a new entry.
        ' In either case, we initially disable the checkboxes and comment
        ' textbox.
        '=========
        txtComments.Enabled = False
        Call SetStarCheckBoxEnabled(False)
        
        If Me.btnRatingAction.Caption = "Edit New" Then
            Me.btnRatingAction.Enabled = True
        Else
            Me.btnRatingAction.Enabled = False
        End If
    End If
    
    m_bIgnoreStarClicks = False
End Sub

'==============================================================================
' SUBROUTINE
'   SetStarCheckBoxEnabled
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub SetStarCheckBoxEnabled(bEnable As Boolean)
    Dim i As Long
    
    For i = 1 To 5
        m_cbaColloquialStarsCheckBoxs(i).Enabled = bEnable
    Next
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
Private Sub cbStar1_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(1, True)
End Sub

Private Sub cbStar2_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(2, True)
End Sub

Private Sub cbStar3_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(3, True)
End Sub

Private Sub cbStar4_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(4, True)
End Sub

Private Sub cbStar5_Click()
    If m_bIgnoreStarClicks Then Exit Sub
    Call SetStarRating(5, True)
End Sub

'==============================================================================
' SUBROUTINE
'   SetStarRating
'------------------------------------------------------------------------------
' DESCRIPTION
'   Sets the star rating by appropriately clearing and setting the checkbox's
' and displaying the relevant text
'==============================================================================
Private Sub SetStarRating(iStar As Long, bUserAction As Boolean)
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
    
    If bUserAction Then
        btnRatingAction.Enabled = True
        If btnRatingAction.Caption = "Edit New" Then
            btnRatingAction.Caption = "Create New"
        End If
    End If
    
cleanup_nicely:
    m_bIgnoreStarClicks = False
        
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtComments_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub txtComments_Change()

    If m_bIgnoreCommentChange Then
        Exit Sub
    End If
    
    btnRatingAction.Enabled = True
    
    If btnRatingAction.Caption = "Edit New" Then
        btnRatingAction.Caption = "Create New"
    End If

End Sub


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnRatingAction_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnRatingAction_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
#Else
    Dim cnn As Object
#End If
    Dim scSQLQuery As String
    
    
    If btnRatingAction.Caption = "Edit New" Then
        Call SetStarCheckBoxEnabled(True)
        Call SetStarRating(0, False)
        
        m_bIgnoreCommentChange = True
        txtComments.Enabled = True
        txtComments.Text = ""
        m_bIgnoreCommentChange = False
        
        btnRatingAction.Caption = "Create"
        
        frmRating.Caption = "Rating of " & m_scCurrentUser
        
        lbRatings.ListIndex = -1
        
        Exit Sub
    End If
    
    '===========
    ' Connect to the dayabase
    '===========
    If Not ConnectToDB(m_eDatabase, cnn) Then
        Call MsgBox("Unable to connect to the database.")
        Exit Sub
    End If
    
    If btnRatingAction.Caption = "Update" Then
        If m_scCurrentUser = m_scRatingAuthor Then
    
            scSQLQuery = "UPDATE dbo.t_rating SET rating_out_of_5 = " & Me.StarRating & _
                ", comment = '" & Replace(txtComments.Text, "'", "''") & "' " & _
                " WHERE pk_rating = " & m_iPkRating
                
            Call cnn.Execute(scSQLQuery)
            btnRatingAction.Enabled = False
            
        End If
        
        Call Refresh
        
    Else
        
        scSQLQuery = "INSERT INTO dbo.t_rating " & _
            " (rating_out_of_5, fk_rated_item, category, comment, author) VALUES (" & _
            Me.StarRating & ", " & _
            m_iItemID & ", '" & _
            m_scSubjectCategory & "', '" & _
            Replace(txtComments.Text, "'", "''") & "', '" & _
            Replace(m_scCurrentUser, "'", "''") & "')"
            
        Call cnn.Execute(scSQLQuery)
        btnRatingAction.Enabled = False
        
        Call Refresh
    End If

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
        Case 0
            StarText = "-"
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
'   btnOK_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnOK_Click()

    If ((btnRatingAction.Caption = "Update") And (btnRatingAction.Enabled)) Or _
        btnRatingAction.Caption = "Create" Then
        If MsgBox("You have edited the rating. Discard?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If

    m_bSuccess = True
    Call Hide
End Sub

