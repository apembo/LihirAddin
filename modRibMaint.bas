Attribute VB_Name = "modRibMaint"
Option Explicit

Public m_ufWorkOrder As ufWorkOrderDisplay
Public m_ufNotiDisplay As ufNotiDisplay

#If DebugBadType = 0 Then
    Public m_ufFLOCBrowser As ufFLOCBrowser
#Else
    Public m_ufFLOCBrowser As Object
#End If


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnReviewWorkOrder_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Callback for btnReviewWorkOrder onAction
'==============================================================================
Sub btnReviewWorkOrder_Click(control As Office.IRibbonControl)

    Dim oErr As typAppError

    If Selection.Count <> 1 Then
        Call MsgBox("Select a single cell only, that contains a work order number")
        Exit Sub
    End If

    '===========
    ' There is a strange error that occurs when the user clicks on the X to
    ' close a userform. The form seems to go to never-never-land but the user
    ' form pointer variable is not set to nothing, which means when I try and
    ' avoid reloading the form from scratch by checking for Nothing, it
    ' causes an error on the next invokation.
    ' The work-around is to put in a On Error code section.
    '===========
On Error GoTo handle_bad_userform_pointer

    If Not m_ufWorkOrder Is Nothing Then
        Call Unload(m_ufWorkOrder)
    End If

handle_bad_userform_pointer:
On Error GoTo 0
    Set m_ufWorkOrder = New ufWorkOrderDisplay

    oErr.BeSilent = False
On Error GoTo handle_error
    If Not m_ufWorkOrder.SetWorkOrder(Selection, oErr) Then
        Exit Sub
    End If

    GoTo show_form

handle_error:
    Call MsgBox("Order not valid or not found in the database")
    Exit Sub

show_form:
    Call m_ufWorkOrder.Show

End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnReviewNotification_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Callback for btnReviewNotification onAction
'==============================================================================
Sub btnReviewNotification_Click(control As Office.IRibbonControl)

    Dim oErr As typAppError

    If Selection.Count <> 1 Then
        Call MsgBox("Select a single cell only, that contains a notification number")
        Exit Sub
    End If

    '===========
    ' There is a strange error that occurs when the user clicks on the X to
    ' close a userform. The form seems to go to never-never-land but the user
    ' form pointer variable is not set to nothing, which means when I try and
    ' avoid reloading the form from scratch by checking for Nothing, it
    ' causes an error on the next invokation.
    ' The work-around is to put in a On Error code section.
    '===========
On Error GoTo handle_bad_userform_pointer

    If Not m_ufNotiDisplay Is Nothing Then
        Call Unload(m_ufNotiDisplay)
    End If

handle_bad_userform_pointer:
On Error GoTo 0
    Set m_ufNotiDisplay = New ufNotiDisplay

    oErr.BeSilent = False
On Error GoTo handle_error

    '==============
    ' Has the user selected an apparently valid notification?
    '==============
    Dim iNoti As Long
    
    iNoti = 0
    If Selection.Count = 1 Then
        Call LooksLikeValidNotification(Selection, iNoti)
    
'        Call MsgBox("Select a single cell only, that contains a notification number")
'        Exit Sub
    End If


    If Not m_ufNotiDisplay.SetNotification(iNoti, oErr) Then
        Exit Sub
    End If

    GoTo show_form

handle_error:
    Call MsgBox("Notification not valid or not found in the database")
    Exit Sub

show_form:
    Call m_ufNotiDisplay.Show(vbModeless)

End Sub

'==============================================================================
' SUBROUTINE
'   LooksLikeValidNotification
'------------------------------------------------------------------------------
' DESCRIPTION
'   Indicates whether or not the passed in string looks like a valid
' notification.
'==============================================================================
Public Function LooksLikeValidNotification(scNoti As String, ByRef iNotification As Long) As Boolean

    Dim iTemp As Long

    If Len(scNoti) >= 8 Then
        If IsNumeric(Left(scNoti, 8)) Then
            iTemp = Val(Left(scNoti, 8))
            
            If iTemp > 10000000 And iTemp < 19999999 Then
                LooksLikeValidNotification = True
                iNotification = iTemp
                Exit Function
            End If
        End If
    End If
    
    iNotification = 0
    LooksLikeValidNotification = False
    
End Function

'==============================================================================
' SUBROUTINE
'   LooksLikeValidWorkOrder
'------------------------------------------------------------------------------
' DESCRIPTION
'   Indicates whether or not the passed in string looks like a valid work
' order.
'==============================================================================
Public Function LooksLikeValidWorkOrder(scWO As String, ByRef iWorkOrder As Long) As Boolean

    Dim iTemp As Long

    If Len(scWO) >= 8 Then
        If IsNumeric(Left(scWO, 8)) Then
            iTemp = Val(Left(scWO, 8))
            
            If iTemp > 10000000 And iTemp < 19999999 Then
                LooksLikeValidWorkOrder = True
                iWorkOrder = iTemp
                Exit Function
            End If
        End If
    End If
    
    iWorkOrder = 0
    LooksLikeValidWorkOrder = False
    
End Function

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnFLOCBrowser_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Opens up the FLOC Browser
'==============================================================================
Sub btnFLOCBrowser_Click(control As Office.IRibbonControl)

    '===========
    ' There is a strange error that occurs when the user clicks on the X to
    ' close a userform. The form seems to go to never-never-land but the user
    ' form pointer variable is not set to nothing, which means when I try and
    ' avoid reloading the form from scratch by checking for Nothing, it
    ' causes an error on the next invokation.
    ' The work-around is to put in a On Error code section.
    '===========
On Error GoTo handle_bad_userform_pointer

    If m_ufFLOCBrowser Is Nothing Then
        Set m_ufFLOCBrowser = New ufFLOCBrowser
    Else
        Call m_ufFLOCBrowser.Reset
    End If

    GoTo show_form

handle_bad_userform_pointer:
    Set m_ufFLOCBrowser = New ufFLOCBrowser

On Error GoTo 0

show_form:
    Call m_ufFLOCBrowser.Show(vbModeless)

End Sub
