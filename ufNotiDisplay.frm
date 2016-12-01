VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufNotiDisplay 
   Caption         =   "Display PM Notification: Maintenance Request"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9150
   OleObjectBlob   =   "ufNotiDisplay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufNotiDisplay"
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
Private m_bSupressNotificationChangeHandling As Boolean

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Initialisation method for the form.
'==============================================================================
Private Sub UserForm_Initialize()
    m_bSupressNotificationChangeHandling = False
    Call RemoveUserformCloseButton(Me)
End Sub

'==============================================================================
' FUNCTION
'   SetNotification
'------------------------------------------------------------------------------
' DESCRIPTION
'   Populates the notification details from the database, based on the supplied
' number iNotification.
'==============================================================================
Public Function SetNotification(iNotification As Long, ByRef oErr As typAppError, Optional bSuppressErrorDisplay = False) As Boolean

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If

    Dim scSQLQuery As String
    
    Dim iErrorNo As Long
    Dim scError As String
    Dim iField As Long
    Dim dUpdateAge As Double
        
        
    oErr.number = eAppErrorCodes.aeNoError
    oErr.description = "<No Error>"
    
    Call Me.Blank(False)
    
    If iNotification = 0 Then
        m_bSupressNotificationChangeHandling = True
        Me.txtNotificationID.Text = ""
        m_bSupressNotificationChangeHandling = False
        
        SetNotification = True
        
        Exit Function
    End If
        
    If Not (ConnectToDB(ldMaintenance, cnn, True, iErrorNo, scError)) Then
        oErr.number = iErrorNo
        oErr.description = scError
        oErr.Source = "ufNotiDisplay.SetNotification"
        
        If (Not oErr.BeSilent) And (Not bSuppressErrorDisplay) Then
            Call MsgBox("Unable to connect to the 'maint' database. Error: " & oErr.number & " - " & oErr.description)
        End If
        SetNotification = False
        Exit Function
    End If
    
    '==============
    ' Get the relevant notification data
    '==============
    scSQLQuery = "SELECT * FROM dbo.v_notification_detail WHERE pk_notification = " & iNotification
    
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)
    
    If rs.EOF Then
        oErr.number = eAppErrorCodes.aeNotificationNotfound
        oErr.description = "Notification " & iNotification & " not found in the database"
        oErr.Source = "ufNotiDisplay.SetNotification"
        If Not bSuppressErrorDisplay Then
            Call DisplayError(oErr)
        End If
        SetNotification = False
        Exit Function
    End If
    
    '============
    ' Display the result
    '============
    m_bSupressNotificationChangeHandling = True
    Me.txtNotificationID.Value = iNotification
    m_bSupressNotificationChangeHandling = False
            
    For iField = 0 To (rs.Fields.Count - 1)
        If Not IsNull(rs.Fields(iField)) Then
            Select Case rs.Fields(iField).Name
                Case "pk_notification"
                Case "short_text"
                    txtNotiShortText.Text = rs.Fields(iField)
                    txtDescription.Text = rs.Fields(iField)
                    
                Case "fk_func_loc"
                    txtFuncLoc.Text = rs.Fields(iField)
                    
                Case "FlocDescription"
                    lblFuncLocDescription.Caption = rs.Fields(iField)
                
                Case "sort_field"
                    txtSortField.Text = rs.Fields(iField)
                
                Case "fk_main_work_centre"
                    txtWorkCenter.Text = rs.Fields(iField)
                
                Case "main_work_center"
                    lblMainWorkCenterDescription.Caption = rs.Fields(iField)
                
                Case "required_end"
                    txtRequiredEndDate.Text = Format(rs.Fields(iField).Value, "dd.mm.yyyy")
                
                Case "long_text"
                    txtNotiLongText.Text = rs.Fields(iField)
                
                Case "reported_by"
                    txtReportedBy.Text = rs.Fields(iField)
                
                Case "fk_order"
                    txtOrder.Text = rs.Fields(iField)
                
                Case "noti_user_status"
                    txtNotiUserStatus.Text = rs.Fields(iField)
                
                Case "noti_sys_status"
                    txtNotiSysStatus.Text = rs.Fields(iField)
                
                Case "notification_date"
                    txtNotiRaisedDate.Text = Format(rs.Fields(iField).Value, "dd.mm.yyyy")
                
                Case "notification_type"
                    txtNotiType.Text = rs.Fields(iField)
                
                Case "order_short_text"
                Case "order_type"
                Case "FlocInfo"
                Case "NotificationInfo"
                Case "ParentNavi"
                Case "ParentNaviDescription"
                Case "ParentNaviInfo"
                Case "FinYear"
                Case "primary_status"
                Case "fk_floc_cc"
                    txtCostCenter.Text = rs.Fields(iField)
                    
                    If Left(txtCostCenter.Text, 1) = "4" Then
                        txtCompanyCode.Text = "2351"
                        lblCompanyDescription.Caption = "Lihir Gold Limited               PNG"
                    Else
                        txtCompanyCode.Text = "????"
                        lblCompanyDescription.Caption = "Only currently designed for Lihir"
                    End If
                        
                Case "floc_cc_description"
                    lblCostCenterDescription.Caption = rs.Fields(iField)
                    
                Case "fk_wo_cc"
                    txtCostCenter.Text = rs.Fields(iField)
                    
                    If Left(txtCostCenter.Text, 1) = "4" Then
                        txtCompanyCode.Text = "2351"
                        lblCompanyDescription.Caption = "Lihir Gold Limited               PNG"
                    Else
                        txtCompanyCode.Text = "????"
                        lblCompanyDescription.Caption = "Only currently designed for Lihir"
                    End If
                    
                Case "wo_cc_description"
                    lblCostCenterDescription.Caption = rs.Fields(iField)
                    
                Case "last_updated_date"
                    
                    Me.txtLastUpdateFromSAP.Value = Format(rs.Fields(iField).Value, "ddd d-mmm-yy h:mm am/pm")
                    
                    dUpdateAge = (Date + Time) - rs.Fields(iField).Value
                    
                    btnLastUpdateColour.BackColor = GetRGBFromAge(dUpdateAge, 0#, 1#, 2#)
                    
                Case "maint_plant"
                    txtMaintPlant.Text = rs.Fields(iField)
                    txtMaintPlant2.Text = rs.Fields(iField)
                    txtMaintPlant3.Text = rs.Fields(iField)
                    
                Case "maint_plant_description"
                    lblMaintPlant3Description.Caption = rs.Fields(iField)
                    
                Case "priority_text"
                    txtPriority.Text = rs.Fields(iField)
            End Select
        End If
    Next

    SetNotification = True
    
End Function


'==============================================================================
' SUBROUTINE
'   Blank
'------------------------------------------------------------------------------
' DESCRIPTION
'   Blanks the entire notification display
'==============================================================================
Public Sub Blank(Optional bIncludeNoti As Boolean = True)

    '===========
    ' Header
    '===========
    If bIncludeNoti Then
        m_bSupressNotificationChangeHandling = True
        txtNotificationID.Text = ""
        m_bSupressNotificationChangeHandling = False
    End If
    
    txtNotiType.Text = ""
    txtNotiShortText.Text = ""
    txtNotiSysStatus.Text = ""
    txtNotiUserStatus.Text = ""
    txtOrder.Text = ""
    
    '===========
    ' Notification tab
    '===========
    txtFuncLoc.Text = ""
    lblFuncLocDescription.Caption = ""
    
    txtDescription.Text = ""
    txtNotiLongText.Text = ""
    
    txtPlannerGroupID.Text = ""
    txtMaintPlant.Text = ""
    lblPlannerGroupDescription.Caption = ""
    txtWorkCenter.Text = ""
    txtMaintPlant2.Text = ""
    lblMainWorkCenterDescription.Caption = ""
    txtReportedBy.Text = ""
    txtNotiRaisedDate.Text = ""
    
    txtRequiredEndDate.Text = ""
    txtPriority.Text = ""
    
    '===========
    ' Location Data tab
    '===========
    txtMaintPlant3.Text = ""
    lblMaintPlant3Description.Caption = ""
    txtSortField.Text = ""
    txtCompanyCode.Text = ""
    lblCompanyDescription.Caption = ""
    txtCostCenter.Text = ""
    lblCostCenterDescription.Caption = ""
    
    '===========
    ' Bottom Section
    '===========
    btnRefresh.Enabled = False
    Me.txtLastUpdateFromSAP.Value = ""
    Me.btnLastUpdateColour.BackColor = &H8000000F  ' Button Face
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnClose_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnClose_Click()
    Call Unload(Me)
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnRefresh_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Initialisation method for the form.
'==============================================================================
Private Sub btnRefresh_Click()
        
    Dim iNoti As Long
    Dim oErr As typAppError
    
On Error GoTo cleanup_nicely
    iNoti = Val(txtNotificationID.Text)
    
    Call Me.Blank
    
    Call Me.SetNotification(iNoti, oErr)
    Exit Sub
    
cleanup_nicely:
    Call MsgBox("Unable to refresh: #" & oErr.number & "-'" & oErr.description & "'")
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnDisplayOrder_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Opens the Display Work Order Userform
'==============================================================================
Private Sub btnDisplayOrder_Click()
    
    Dim ufWO As ufWorkOrderDisplay
    Dim iWO As Long
    Dim oErr As typAppError
    
    If txtOrder.Text = "" Then
        Call MsgBox("No work order specified for this notification")
        Exit Sub
    End If
    
    iWO = Val(txtOrder.Text)
    
    
    Set ufWO = New ufWorkOrderDisplay
    
    Call ufWO.SetWorkOrder(iWO, oErr, True)
    
    ufWO.Show
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtNotificationID_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Event handler for when the user changes the notification number.
'==============================================================================
Private Sub txtNotificationID_Change()

    Dim oErr As typAppError
    Dim iNoti As Long

    If m_bSupressNotificationChangeHandling Then
        Exit Sub
    End If
    
    Call Blank(False)
    
    If LooksLikeValidNotification(txtNotificationID.Text, iNoti) Then
        If Not SetNotification(iNoti, oErr) Then
            Call MsgBox("Error: " & oErr.description)
        End If
    End If
    
End Sub

