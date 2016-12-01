VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPictureManager 
   Caption         =   "Picture Manager"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7470
   OleObjectBlob   =   "ufPictureManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPictureManager"
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
' USERFORM
'   ufPictureManager
'------------------------------------------------------------------------------
' DESCRIPTION
'   This userform manages pictures for the given host item.
' It looks after the following functions:
'   - Adding or deleting pictures
'   - Modifying picture attributes
'   - Changing the picture display order
'   - Adding, modifying or deleting ratings
'------------------------------------------------------------------------------
' IMPLEMENTATION DETAILS
'------------
' ADDING/DELETING PICTURES
'   This is fairly straight forward. Anyone can add a picture, while only
' the person who added a picture can delete it.
'------------------------------------------------------------------------------
' VERSION
'   1.0 - First release
'==============================================================================

'==============================================================================
' PRIVATE MEMBER VARIABLES
'==============================================================================
Private m_ufRating As ufRating
Private m_scKey As String
Private m_scCategory As String
Private m_bSuppressListboxClickEvent As Boolean

Private Const g_cDirtyFlagCount As Long = 5
Private m_baDirtyFlags(1 To g_cDirtyFlagCount) As Boolean
Private m_colPicData As VBA.Collection

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub UserForm_Initialize()
    Set m_colPicData = New VBA.Collection
    m_bSuppressListboxClickEvent = False
    
    Call RemoveUserformCloseButton(Me)
End Sub

'==============================================================================
' SUBROUTINE
'   SetDirtyFlag
'------------------------------------------------------------------------------
' DESCRIPTION
'   Set's one of the dirty flags and enables the Update button.
'==============================================================================
Private Sub SetDirtyFlag(eFlag As eDirtyFlags)

    m_baDirtyFlags(eFlag) = True

    Select Case eFlag
        Case OrderChanged, PictureDetailChanged
            Me.btnUpdate.Enabled = True
    End Select
    
End Sub

'==============================================================================
' SUBROUTINE
'   ClearDirtyFlag
'------------------------------------------------------------------------------
' DESCRIPTION
'   Resets one or more dirty flags and disables the update button where
' appropriate.
'==============================================================================
Private Sub ClearDirtyFlag(Optional eFlag As eDirtyFlags = 0)
    
    Dim i As Long

    If eFlag = 0 Then
        For i = 1 To g_cDirtyFlagCount
            m_baDirtyFlags(i) = False
        Next
    Else
        m_baDirtyFlags(eFlag) = False
    End If
    
    
    If (Not m_baDirtyFlags(OrderChanged)) And _
        (Not m_baDirtyFlags(PictureDetailChanged)) Then
        Me.btnUpdate.Enabled = False
    End If
    
    If Not m_baDirtyFlags(eDirtyFlags.PictureDetailChangedNotInList) Then
        btnUpdateDetail.Enabled = False
    End If

End Sub

'==============================================================================
' PROPERTY
'   DirtyFlag
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the state of the dirty flag corresponding to the supplied enum
'==============================================================================
Public Property Get DirtyFlag(Optional eFlag As eDirtyFlags = 0) As Boolean

    Dim bAtleastOneDirty As Boolean
    Dim i As Long
    
    If eFlag = 0 Then
        '==========
        ' Passing 0 or no value means the user is looking for any true dirty
        ' flag.
        '==========
        bAtleastOneDirty = False
        For i = 1 To g_cDirtyFlagCount
            If m_baDirtyFlags(i) Then
                bAtleastOneDirty = True
                Exit For
            End If
        Next
        DirtyFlag = bAtleastOneDirty
    Else
        DirtyFlag = m_baDirtyFlags(eFlag)
    End If
End Property

'==============================================================================
' SUBROUTINE
'   Configure
'------------------------------------------------------------------------------
' DESCRIPTION
'   Should be called by the form creator prior to showing.
'==============================================================================
Public Sub Configure(scCategory As String, scKey As String)

    m_scCategory = scCategory
    m_scKey = scKey
    
    Call ClearDirtyFlag
    
    Call Me.Refresh
    
End Sub

'==============================================================================
' SUBROUTINE
'   Refresh
'------------------------------------------------------------------------------
' DESCRIPTION
'   Call to refresh the display. Note that it will deselect the list.
'==============================================================================
Public Sub Refresh()

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
    
    Dim scName As String
    Dim iDataIndex As Long
    
    Call Me.lbPictures.Clear
    Call ClearPicDataCollection
    
    '========
    ' Construct the query
    '========
    scSQLQuery = "SELECT * " & _
            "FROM maint.dbo.v_mapped_files WHERE " & _
            "base_path_category = '" & m_scCategory & "' AND id = '" & m_scKey & "' " & _
            "ORDER BY file_order, file_date DESC"
    
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)

    While Not rs.EOF
        Set cPicData = New clsPictureData
        
        Call cPicData.PopulateFromRecordset(rs)
        Call m_colPicData.Add(cPicData, cPicData.Key)
    
        '----
        ' Name - modified relative_path
        '----
        Call Me.lbPictures.AddItem(cPicData.Name)
        
        '----
        ' file_date
        '----
        If IsNull(rs.Fields("file_date")) Then
            Me.lbPictures.List(Me.lbPictures.ListCount - 1, 1) = "-"
        Else
            Me.lbPictures.List(Me.lbPictures.ListCount - 1, 1) = cPicData.FileDateStr(DateShort)
        End If
                
        '----
        ' rating_avg
        '----
        If IsNull(rs.Fields("rating_avg")) Then
            Me.lbPictures.List(Me.lbPictures.ListCount - 1, 2) = "0.0"
        Else
            Me.lbPictures.List(Me.lbPictures.ListCount - 1, 2) = Format(cPicData.rating_avg, "0.0")
        End If
        
        '----
        ' pk_file_map
        '----
        Me.lbPictures.List(Me.lbPictures.ListCount - 1, 3) = cPicData.pk_file_map
        
        Call rs.MoveNext
    Wend
    Call rs.Close
    
    '========
    ' Display the title
    '========
    Select Case m_scCategory
        Case "FLOC_PICTURES"
            Me.lblTitle.Caption = "FLOC " & m_scKey
        Case "MATERIAL_PICTURES"
            Me.lblTitle.Caption = "Material " & m_scKey
        Case "PERSON_PICTURES"
            Me.lblTitle.Caption = "Person (ID: " & m_scKey & ")"
        Case Else
            Me.lblTitle.Caption = m_scCategory & " " & m_scKey
    End Select
    
    '========
    ' Clear various picture specific fields
    '========
    Call ClearFileDetailDisplay

End Sub

'==============================================================================
' SUBROUTINE
'   ClearFileDetailDisplay
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub ClearFileDetailDisplay()

    Me.lblFileOwner.Caption = ""
    Me.lblRating.Caption = ""
    Me.lblRatingCount.Caption = ""
    Me.dtpFileDate.Value = DateSerial(1970, 1, 1)
    Me.imgPreview.Picture = LoadPicture("")
    
    Me.btnOrderDown.Enabled = False
    Me.btnOrderUp.Enabled = False
    
    frmFileDetails.Caption = "File Details"
    
    Call ClearDirtyFlag(OrderChanged)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   dtpFileDate_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub dtpFileDate_Change()

    Call SetDirtyFlag(PictureDetailChangedNotInList)
    btnUpdateDetail.Enabled = True
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtDescription_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub txtDescription_Change()

    Call SetDirtyFlag(PictureDetailChangedNotInList)
    btnUpdateDetail.Enabled = True

End Sub

'==============================================================================
' FUNCTION
'   PictureDetailListNumber
'------------------------------------------------------------------------------
' DESCRIPTION
'   Gets the list index of the files who's details are displayed.
' Remember the list is 0 based.
'==============================================================================
Public Property Get PictureDetailListNumber() As Long
    If frmFileDetails.Caption = "File Details" Then
        PictureDetailListNumber = -1
    Else
        PictureDetailListNumber = CLng(Val(Mid(frmFileDetails.Caption, Len("File Details: ")))) - 1
    End If
End Property

'==============================================================================
' FUNCTION
'   HandleUnprocessedDetailChanges
'------------------------------------------------------------------------------
' DESCRIPTION
'   There are a number of situations where the user might try to do something
' after making some changes to file details, and forget to update the
' changes.
' This method should be called first and will ask the user if they want to save
' the changes.
' It will return true if it handles the situation, or false if the user selects
' cancel. If it returns False, the user should end the event handler.
'==============================================================================
Private Function HandleUnprocessedDetailChanges(Optional bIncludeCancel As Boolean = False) As Boolean
    Dim iPicNo As Long
    Dim oButtons As VBA.VbMsgBoxStyle
    
    If bIncludeCancel Then
        oButtons = vbYesNoCancel
    Else
        oButtons = vbYesNo
    End If
    
    '============
    ' Has the user made changes to picture details?
    '============
    If DirtyFlag(PictureDetailChangedNotInList) Then
        iPicNo = PictureDetailListNumber
        Select Case MsgBox("You have changed some details of picture " & (iPicNo + 1) & ". Update?", oButtons)
            Case vbYes
                Call btnUpdateDetail_Click
                
            Case vbNo
                Call ClearDirtyFlag(PictureDetailChangedNotInList)

            Case vbCancel

                HandleUnprocessedDetailChanges = False
                Exit Function
        End Select
    End If
    
    HandleUnprocessedDetailChanges = True
End Function

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   lbPictures_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub lbPictures_Click()

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
    
    Dim pk_file_map As Long
    Dim scPictureAuthor As String
    Dim scUser As String
    Dim iRatingCount As Long
    Dim iPicNo As Long
    Dim iPicIndex As Long
    Dim i As Long
    
    
    If lbPictures.ListIndex < 0 Then
        Call ClearFileDetailDisplay
        Exit Sub
    End If
    
    '==========
    ' Check if some details of the picture had been edited and give the user
    ' the option to save these changes before moving to the next picture.
    '==========
    Call HandleUnprocessedDetailChanges
    
    '==========
    ' Display the new picture info
    '==========
    frmFileDetails.Caption = "File Details: " & (lbPictures.ListIndex + 1)
    
    pk_file_map = Me.lbPictures.List(lbPictures.ListIndex, 3)
    Set cPicData = m_colPicData.Item(PicDataKey(pk_file_map))
    
    Me.imgPreview.Picture = LoadPicture(cPicData.FullPath)
    
    '----
    ' Description
    '----
    txtDescription.Text = cPicData.description
    
    '----
    ' file_date
    '----
    dtpFileDate.Value = cPicData.file_date
    
    '----
    ' owner
    '----
    scPictureAuthor = cPicData.owner
    Me.lblFileOwner.Caption = scPictureAuthor
    
    '----
    ' rating_count
    ' rating_avg
    '----
    Me.lblRatingCount.Caption = cPicData.rating_count
    Me.lblRating.Caption = Format(cPicData.rating_avg, "0.0")
    
    '========
    ' Enable or disable the Update and Delete button's depending on whether
    ' this user is the author of the picture/map.
    '========
    Call GetUserName(scUser)
    
    If LCase(scUser) = LCase(scPictureAuthor) Then
        Me.txtDescription.Enabled = True
        Me.dtpFileDate.Enabled = True
        'Me.btnUpdateDetail.Enabled = True
        Me.btnDelete.Enabled = True
    Else
        Me.txtDescription.Enabled = False
        Me.dtpFileDate.Enabled = False
        'Me.btnUpdateDetail.Enabled = False
        Me.btnDelete.Enabled = False
    End If
    
    Me.btnOrderDown.Enabled = True
    Me.btnOrderUp.Enabled = True
    
    '========
    ' Since we've just made changes to the text box etc which would have thrown
    ' change events which would have in turn set a dirty flag, we clear that
    ' dirty flag.
    '========
    Call ClearDirtyFlag(PictureDetailChangedNotInList)
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnAdd_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Add a picture
'==============================================================================
Private Sub btnAdd_Click()
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ofile As Scripting.File
    Dim oFileDialog As Office.FileDialog
#Else
    Dim cnn As Object
    Dim rs As Object
    Dim ofile As Object
    Dim oFileDialog As Object
#End If

    Dim scSQLQuery As String
        
    Dim iResult As Long
    Dim scFromFileNameFull As String
    Dim iFileID As Long
    Dim iMaxOrder As Long
    Dim i As Long
    Dim scUser As String
    Dim scDescription As String
    
    
    '==========
    ' Has the user forgotten to save details from a previous file edit?
    '==========
    If Not HandleUnprocessedDetailChanges Then
        Exit Sub
    End If

    '==========
    ' Use the file dialog object to get the file name.
    '==========
    Set oFileDialog = Application.FileDialog(msoFileDialogFilePicker)

    With oFileDialog
        .AllowMultiSelect = False
        .Filters.Clear
        Call .Filters.Add("JPEG Images", "*.jpg;*.jpeg")
        Call .Filters.Add("Bitmaps", "*.bmp;*.dib")
        Call .Filters.Add("Metafiles", "*.wmf;*.emf")
        Call .Filters.Add("GIF Images", "*.gif")
        Call .Filters.Add("Icons", "*.ico;*.cur")
        Call .Filters.Add("All Files", "*.*")
        iResult = .Show

        If iResult <> 0 Then
            scFromFileNameFull = .SelectedItems.Item(1)
        Else
            Exit Sub
        End If
    End With
    
    '==========
    ' Check the file type is supported.
    '==========
    Select Case LCase(FileExtensionFromPath(scFromFileNameFull))
        Case "jpg", "jpeg", "bmp", "dib", "gif", "wmf", "emf", "ico", "cur"
            ' All Good
        Case Else
            If MsgBox("Your file type appears to be an unsupported type. Proceed?", vbYesNo) = vbNo Then
                Exit Sub
            End If
    End Select
    
    Call GetUserName(scUser)

    '==========
    ' Get a description
    '==========
    scDescription = InputBox("Please provide a brief description of the file contents (or click Cancel ... you can edit later).", "File Description")
    
    '==========
    ' Add the file to the database and get back the ID
    '==========
    If Not DBAddFile(scFromFileNameFull, m_scCategory, scUser, scDescription, iFileID) Then
        Exit Sub
    End If

    If DBGetFile(iFileID, ofile) Then
        Me.imgPreview.Picture = LoadPicture(ofile.Path)
    End If

    '========
    ' Get the largest order number
    '========
    iMaxOrder = 1
    For i = 0 To Me.lbPictures.ListCount - 1
        If Me.lbPictures.List(i, 5) > iMaxOrder Then
            iMaxOrder = Me.lbPictures.List(i, 5)
        End If
    Next
    
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
            "(fk_file, mapped_obj_category, fk_mapped_obj_str, file_order) VALUES " & _
            "(" & iFileID & ", '" & m_scCategory & "', '" & m_scKey & "', " & (iMaxOrder + 1) & ")"

    Call cnn.Execute(scSQLQuery)

    '==========
    ' Finally refresh this picture manager.
    '==========
    Call Me.Refresh

    '==========
    ' Set the dirty flag to indicate the caller of this picture manager may
    ' need to refresh their display.
    '==========
    Call SetDirtyFlag(PicturesAddedOrDeleted)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnDelete_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Delete a picture
'==============================================================================
Private Sub btnDelete_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
#Else
    Dim cnn As Object
#End If

#If DebugBadType = 0 Then
    Dim cPicData As clsPictureData
#Else
    Dim cPicData As Object
#End If

    Dim scSQLQuery As String
        
    Dim pk_file_map As Long
    Dim iPicNo As Long
    Dim scUser As String
    Dim scPictureAuthor As String
    
    
    If lbPictures.ListIndex < 0 Then
        btnDelete.Enabled = False
        Exit Sub
    End If
    
    '=============
    ' What's the list index of this file?
    '=============
    iPicNo = PictureDetailListNumber
    If iPicNo < 0 Then
        Call ClearDirtyFlag(PictureDetailChangedNotInList)
        Exit Sub
    End If
    
    pk_file_map = Me.lbPictures.List(iPicNo, 3)
    Set cPicData = m_colPicData.Item(PicDataKey(pk_file_map))
    'iPicIndex = PicData_IndexFrom_pk_file_map(pk_file_map)
    
    '=============
    ' Double check that this user is the owner of the file.
    '=============
    scPictureAuthor = cPicData.owner
    Call GetUserName(scUser)
    
    If Not (LCase(scUser) = LCase(scPictureAuthor)) Then
        btnDelete.Enabled = False
        Exit Sub
    End If
    
    '========
    ' Connect to the DB
    '========
    If Not (ConnectToDB(ldMaintenance, cnn, True)) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Sub
    End If
    
    '==========
    ' Construct the query
    '==========
    scSQLQuery = "DELETE FROM maint.dbo.t_file_map WHERE pk_file_map = " & pk_file_map
    
    Call cnn.Execute(scSQLQuery)
    
    '==========
    ' Finally refresh this picture manager.
    '==========
    Call Me.Refresh

    '==========
    ' Set the dirty flag to indicate we may need
    '==========
    Call SetDirtyFlag(PicturesAddedOrDeleted)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnUpdateDetail_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnUpdateDetail_Click()

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
        
    Dim pk_file As Long
    Dim pk_file_map As Long
    Dim iPicNo As Long
    
    '=============
    ' We change the detail in the list item for now. The update will happen
    ' when the user hits the update button.
    '=============
    If Not DirtyFlag(PictureDetailChangedNotInList) Then
        btnUpdateDetail.Enabled = False
        Exit Sub
    End If
    
    '=============
    ' What's the list index of this file?
    '=============
    iPicNo = PictureDetailListNumber
    If iPicNo < 0 Then
        Call ClearDirtyFlag(PictureDetailChangedNotInList)
        Exit Sub
    End If
    
    pk_file_map = Me.lbPictures.List(iPicNo, 3)
    Set cPicData = m_colPicData.Item(PicDataKey(pk_file_map))
    'iPicIndex = PicData_IndexFrom_pk_file_map(pk_file_map)

    '=============
    ' We store the changes in the list for now
    '=============
    cPicData.description = Me.txtDescription.Text
    cPicData.file_date = Me.dtpFileDate.Value
    cPicData.changed = True
    
    '=============
    ' We also reflect the date change (the only thing displayed that the user
    ' can edit) in the listbox.
    '=============
    lbPictures.List(iPicNo, 1) = cPicData.FileDateStr(DateShort)
    
    '=============
    ' Indicate we have captured the changes (by storing them in the list) but
    ' have have not written them to the database.
    '=============
    Call SetDirtyFlag(PictureDetailChanged)
    Call ClearDirtyFlag(PictureDetailChangedNotInList)
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnOrderUp_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnOrderUp_Click()

    Dim iOldIndex As Long

    If lbPictures.ListIndex < 0 Then
        Exit Sub
    End If
    
    '==========
    ' Has the user forgotten to save details from a previous file edit?
    '==========
    If Not HandleUnprocessedDetailChanges Then
        Exit Sub
    End If

    
    If Me.lbPictures.ListIndex <= 0 Then
        Exit Sub
    End If
    
    iOldIndex = Me.lbPictures.ListIndex
    
    '==========
    ' Swap the list entries
    '==========
    Call MoveListboxItem(lbPictures, Me.lbPictures.ListIndex, Me.lbPictures.ListIndex - 1)
    Me.lbPictures.Selected(iOldIndex - 1) = True
    
    Call SetDirtyFlag(OrderChanged)
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnOrderDown_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnOrderDown_Click()

    Dim iOldIndex As Long

    If lbPictures.ListIndex < 0 Then
        Exit Sub
    End If
    
    '==========
    ' Has the user forgotten to save details from a previous file edit?
    '==========
    If Not HandleUnprocessedDetailChanges Then
        Exit Sub
    End If

    
    If Me.lbPictures.ListIndex >= (Me.lbPictures.ListCount - 1) Then
        Exit Sub
    End If
    
    iOldIndex = Me.lbPictures.ListIndex
       
    '==========
    ' Move the selected list entry down one position.
    '==========
    Call MoveListboxItem(lbPictures, Me.lbPictures.ListIndex, Me.lbPictures.ListIndex + 1)
    Me.lbPictures.Selected(iOldIndex + 1) = True
    
    Call SetDirtyFlag(OrderChanged)
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnRate_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnRate_Click()
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String
        
    Dim pk_file_map As Long
    Dim scAuthor As String
    Dim scComment As String
    Dim bPreviousRatingFound As Boolean
    Dim pk_rating As Long
    Dim iRecordsAffected As Long
    
    '==========
    ' Has the user forgotten to save details from a previous file edit?
    '==========
    If Not HandleUnprocessedDetailChanges Then
        Exit Sub
    End If

    '==========
    ' Is any picture selected?
    '==========
    If Me.lbPictures.ListCount < 0 Then
        Exit Sub
    End If
    
    '==========
    ' Get the file map ID and the user's name
    '==========
    pk_file_map = CLng(Me.lbPictures.List(Me.lbPictures.ListIndex, 3))

    If Not GetUserName(scAuthor) Then
        scAuthor = Application.UserName
    End If
    
    '==========
    ' See if the user has already entered a rating. We only allow one rating
    ' per user per item.
    '==========
    scSQLQuery = "SELECT * FROM dbo.t_rating " & _
        "WHERE category = 'FLOC_PICTURES' AND fk_rated_item = " & pk_file_map & _
            " AND author = '" & Replace(scAuthor, "'", "''") & "'"
    
    bPreviousRatingFound = False
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)
    
    '==========
    ' Get the Rating Userform.
    '==========
    If m_ufRating Is Nothing Then
        Set m_ufRating = New ufRating
    End If
    
    If Not rs.EOF Then
        bPreviousRatingFound = True
        pk_rating = rs.Fields("pk_rating")
        
        If IsNull(rs.Fields("comment")) Then
            scComment = ""
        Else
            scComment = rs.Fields("comment")
        End If
        Call m_ufRating.Initialise("Rating for Picture", rs.Fields("rating_out_of_5"), scComment)
    Else
        Call m_ufRating.Initialise("Rating for Picture")
    End If
    
    '==========
    ' Get the Rating
    '==========
    Call m_ufRating.Show
    
    If m_ufRating Is Nothing Then
        Exit Sub
    End If
    
    If m_ufRating.Success Then
        If bPreviousRatingFound Then
            scSQLQuery = "UPDATE dbo.t_rating SET rating_out_of_5 = " & m_ufRating.StarRating & _
                ", comment = '" & Replace(m_ufRating.Comment, "'", "''") & "' WHERE pk_rating = " & pk_rating
            Call cnn.Execute(scSQLQuery, iRecordsAffected)
        Else

            scSQLQuery = "INSERT INTO dbo.t_rating " & _
                "(rating_out_of_5,fk_rated_item,category,comment,author) VALUES (" & _
                m_ufRating.StarRating & ", " & _
                pk_file_map & ", " & _
                "'FLOC_PICTURES', '" & _
                Replace(m_ufRating.Comment, "'", "''") & "', '" & _
                Replace(scAuthor, "'", "''") & "')"
                
            Call cnn.Execute(scSQLQuery, iRecordsAffected)
            
        End If
    End If
        
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnCancel_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnCancel_Click()

    '===========
    ' Has the user updated one or more details of the currently selected file?
    '===========
    If Not HandleUnprocessedDetailChanges Then
        Exit Sub
    End If
    
    '===========
    ' Are there some changes that have not been reflected in the database?
    '===========
    If DirtyFlag(OrderChanged) Or DirtyFlag(PictureDetailChanged) Then
        If MsgBox("There are unsaved changes. So save them, click Cancel and click the Update button. To discard hit OK.", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    
    Call Hide
        
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnUpdate_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnUpdate_Click()

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
#Else
    Dim cnn As Object
#End If

#If DebugBadType = 0 Then
    Dim cPicData As clsPictureData
#Else
    Dim cPicData As Object
#End If

    Dim scSQLQuery As String
    Dim pk_file_map As Long
    Dim iItem As Long
    
    
    If Not DirtyFlag Then
        Exit Sub
    End If

    '===========
    ' Has the user changed the file order?
    '===========
    If Not (ConnectToDB(ldMaintenance, cnn, True)) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Sub
    End If
        
    '===========
    ' Has the user changed the file order?
    '===========
    If DirtyFlag(OrderChanged) Then
    
        '========
        ' If the order of the pictures has changed, we update
        '========
        For iItem = 0 To Me.lbPictures.ListCount - 1
            pk_file_map = CLng(Me.lbPictures.List(iItem, 3))
            Set cPicData = m_colPicData.Item(PicDataKey(pk_file_map))
            
            If Not cPicData.file_order = (iItem + 1) Then
                
                scSQLQuery = "UPDATE dbo.t_file_map SET file_order = " & (iItem + 1) & _
                    " WHERE pk_file_map = " & pk_file_map

                Call cnn.Execute(scSQLQuery)
            End If
            
        Next
    End If
    
    '===========
    ' Has the user changed any of the file details?
    '===========
    If DirtyFlag(PictureDetailChanged) Then
        For iItem = 0 To Me.lbPictures.ListCount - 1
            pk_file_map = CLng(Me.lbPictures.List(iItem, 3))
            Set cPicData = m_colPicData.Item(PicDataKey(pk_file_map))
        
            If cPicData.changed Then
                
                scSQLQuery = "UPDATE dbo.t_file SET description = '" & Replace(cPicData.description, "'", "''") & _
                "', file_date = '" & cPicData.FileDateStr(DateTimeTSQL) & _
                "' WHERE pk_file = " & cPicData.pk_file
                    
                Call cnn.Execute(scSQLQuery)
                cPicData.changed = False
            End If
        Next
    End If
    
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE
'   ClearPicDataCollection
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub ClearPicDataCollection()
    While m_colPicData.Count > 0
        Call m_colPicData.Remove(1)
    Wend
End Sub

'==============================================================================
' FUNCTION
'   PicDataKey
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Public Function PicDataKey(pk_file_map As Long) As String

#If DebugBadType = 0 Then
    Dim cPicData As clsPictureData
#Else
    Dim cPicData As Object
#End If
    
    Set cPicData = New clsPictureData
    
    PicDataKey = cPicData.Static_Key(pk_file_map)
End Function

'==============================================================================
'==============================================================================
' CODE SECTION
'   PicData storage methods.
'------------------------------------------------------------------------------
' DESCRIPTION
'   Trying to avoid using classes that add complexity
'==============================================================================
'==============================================================================



'Private Sub PicDataInitialise()
'    m_iPicDataCount = 0
'
'    ReDim m_taPicData(0 To 9)
'
'    tBadPicData.base_path = ""
'    tBadPicData.pk_file = -1
'    tBadPicData.pk_file_map = -1
'End Sub
'
'Private Property Get PicData(iIndex As Long) As typPictureData
'
'    If (iIndex < 0) Or (iIndex >= m_iPicDataCount) Then
'        PicData = tBadPicData
'    Else
'        PicData = m_taPicData(iIndex)
'    End If
'End Property
'
'Private Property Get PicDataCount() As Long
'    PicDataCount = m_iPicDataCount
'End Property
'
'Private Property Get PicDataArraySize() As Long
'    PicDataArraySize = UBound(m_taPicData) - LBound(m_taPicData) + 1
'End Property
'
'Private Function PicDataAdd(dat As typPictureData) As Long
'    If PicDataArraySize <= m_iPicDataCount Then
'        ReDim Preserve m_taPicData(0 To (m_iPicDataCount + 10))
'    End If
'    Call PicData_Copy(dat, m_taPicData(m_iPicDataCount))
'
'    PicDataAdd = m_iPicDataCount
'    m_iPicDataCount = m_iPicDataCount + 1
'End Function
'
'Private Sub PicDataDelete(iIndex As Long)
'
'    Dim i As Long
'
'    For i = (iIndex + 1) To (m_iPicDataCount - 1)
'        Call PicData_Copy(m_taPicData(i), m_taPicData(i - 1))
'    Next
'    m_iPicDataCount = m_iPicDataCount - 1
'
'End Sub
'
'Private Sub PicData_Copy(ByRef tFrom As typPictureData, ByRef tTo As typPictureData)
'    tTo.base_path = tFrom.base_path
'    tTo.changed = tFrom.changed
'    tTo.description = tFrom.description
'    tTo.file_date = tFrom.file_date
'    tTo.file_order = tFrom.file_order
'    tTo.owner = tFrom.owner
'    tTo.pk_file = tFrom.pk_file
'    tTo.pk_file_map = tFrom.pk_file_map
'    tTo.rating_avg = tFrom.rating_avg
'    tTo.rating_count = tFrom.rating_count
'    tTo.relative_path = tFrom.relative_path
'End Sub
'
'Private Sub PicDataClear()
'    m_iPicDataCount = 0
'End Sub
'
'Private Property Get PicData_Name(iIndex As Long) As String
'
'    Dim scRelPath As String
'
'    If (iIndex < 0) Or (iIndex >= m_iPicDataCount) Then
'        Exit Property
'    End If
'
'    scRelPath = m_taPicData(iIndex).relative_path
'
'    PicData_Name = Mid(scRelPath, InStr(5, scRelPath, "-") + 1)
'
'End Property
'
'Private Property Get PicData_FullPath(iIndex As Long) As String
'
'    Dim scBasePath As String
'    Dim scRelPath As String
'
'    If (iIndex < 0) Or (iIndex >= m_iPicDataCount) Then
'        Exit Property
'    End If
'
'    scBasePath = m_taPicData(iIndex).base_path
'    scRelPath = m_taPicData(iIndex).relative_path
'
'    If Right(scBasePath, 1) <> "\" Then
'        PicData_FullPath = scBasePath & "\" & scRelPath
'    Else
'        PicData_FullPath = scBasePath & scRelPath
'    End If
'End Property
'
'Private Sub PicData_CopyFromRecordset(rsFrom As ADODB.Recordset, tTo As typPictureData)
'
'    tTo.base_path = rsFrom.Fields("base_path")
'    tTo.file_order = rsFrom.Fields("file_order")
'    tTo.pk_file = rsFrom.Fields("pk_file")
'    tTo.pk_file_map = rsFrom.Fields("pk_file_map")
'    tTo.relative_path = rsFrom.Fields("relative_path")
'
'    tTo.changed = 0
'
'    '=============
'    ' Handle tidely those fields that could be null.
'    '=============
'    '----
'    ' file_date
'    '----
'    If Not IsNull(rsFrom.Fields("file_date")) Then
'        tTo.file_date = rsFrom.Fields("file_date")
'    Else
'        tTo.file_date = DateSerial(1970, 1, 1)
'    End If
'
'    '----
'    ' description
'    '----
'    If IsNull(rsFrom.Fields("description")) Then
'        tTo.description = ""
'    Else
'        tTo.description = rsFrom.Fields("description")
'    End If
'
'    '----
'    ' rating_count, rating_avg
'    '----
'    If IsNull(rsFrom.Fields("rating_count")) Then
'        tTo.rating_avg = 0#
'        tTo.rating_count = 0
'    Else
'        tTo.rating_avg = rsFrom.Fields("rating_avg")
'        tTo.rating_count = rsFrom.Fields("rating_count")
'    End If
'
'    '----
'    ' owner
'    '----
'    If IsNull(rsFrom.Fields("owner")) Then
'        tTo.owner = ""
'    Else
'        tTo.owner = rsFrom.Fields("owner")
'    End If
'
'End Sub
'
'Private Function PicData_IndexFrom_pk_file_map(pk_file_map As Long) As Long
'
'    Dim i As Long
'
'    For i = 0 To PicDataCount
'        If m_taPicData(i).pk_file_map = pk_file_map Then
'            PicData_IndexFrom_pk_file_map = i
'            Exit Function
'        End If
'    Next
'    PicData_IndexFrom_pk_file_map = -1
'End Function
