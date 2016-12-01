VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufWorkOrderDisplay 
   Caption         =   "Display Scheduled Maintenance Work Order 10204567: Central Header"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11805
   OleObjectBlob   =   "ufWorkOrderDisplay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufWorkOrderDisplay"
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
Private m_bSupressWorkOrderChangeHandling As Boolean

Private Type typOperationsFields
    operation_activity As String
    short_text As String
    fk_op_work_centre As String
    earliest_start As Date
    no_of_resource As Long
    normal_duration As Double
    planned_work As Double
    actual_work As Double
    unit_of_work As String
    sys_status As String
End Type

Private m_txtaOpFields(0 To 9, 0 To 9) As Object
Private m_taOpFields() As typOperationsFields
Private m_iOperationsCount As Long

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Initialisation method for the form.
'==============================================================================
Private Sub UserForm_Initialize()

    Set m_txtaOpFields(0, 0) = txtOpAc01
    Set m_txtaOpFields(1, 0) = txtOpAc02
    Set m_txtaOpFields(2, 0) = txtOpAc03
    Set m_txtaOpFields(3, 0) = txtOpAc04
    Set m_txtaOpFields(4, 0) = txtOpAc05
    Set m_txtaOpFields(5, 0) = txtOpAc06
    Set m_txtaOpFields(6, 0) = txtOpAc07
    Set m_txtaOpFields(7, 0) = txtOpAc08
    Set m_txtaOpFields(8, 0) = txtOpAc09
    Set m_txtaOpFields(9, 0) = txtOpAc10
    
    Set m_txtaOpFields(0, 1) = txtShortText01
    Set m_txtaOpFields(1, 1) = txtShortText02
    Set m_txtaOpFields(2, 1) = txtShortText03
    Set m_txtaOpFields(3, 1) = txtShortText04
    Set m_txtaOpFields(4, 1) = txtShortText05
    Set m_txtaOpFields(5, 1) = txtShortText06
    Set m_txtaOpFields(6, 1) = txtShortText07
    Set m_txtaOpFields(7, 1) = txtShortText08
    Set m_txtaOpFields(8, 1) = txtShortText09
    Set m_txtaOpFields(9, 1) = txtShortText10
    
    Set m_txtaOpFields(0, 2) = txtWorkCtr01
    Set m_txtaOpFields(1, 2) = txtWorkCtr02
    Set m_txtaOpFields(2, 2) = txtWorkCtr03
    Set m_txtaOpFields(3, 2) = txtWorkCtr04
    Set m_txtaOpFields(4, 2) = txtWorkCtr05
    Set m_txtaOpFields(5, 2) = txtWorkCtr06
    Set m_txtaOpFields(6, 2) = txtWorkCtr07
    Set m_txtaOpFields(7, 2) = txtWorkCtr08
    Set m_txtaOpFields(8, 2) = txtWorkCtr09
    Set m_txtaOpFields(9, 2) = txtWorkCtr10
    
    Set m_txtaOpFields(0, 3) = txtErlStart01
    Set m_txtaOpFields(1, 3) = txtErlStart02
    Set m_txtaOpFields(2, 3) = txtErlStart03
    Set m_txtaOpFields(3, 3) = txtErlStart04
    Set m_txtaOpFields(4, 3) = txtErlStart05
    Set m_txtaOpFields(5, 3) = txtErlStart06
    Set m_txtaOpFields(6, 3) = txtErlStart07
    Set m_txtaOpFields(7, 3) = txtErlStart08
    Set m_txtaOpFields(8, 3) = txtErlStart09
    Set m_txtaOpFields(9, 3) = txtErlStart10
    
    Set m_txtaOpFields(0, 4) = txtResourceNo01
    Set m_txtaOpFields(1, 4) = txtResourceNo02
    Set m_txtaOpFields(2, 4) = txtResourceNo03
    Set m_txtaOpFields(3, 4) = txtResourceNo04
    Set m_txtaOpFields(4, 4) = txtResourceNo05
    Set m_txtaOpFields(5, 4) = txtResourceNo06
    Set m_txtaOpFields(6, 4) = txtResourceNo07
    Set m_txtaOpFields(7, 4) = txtResourceNo08
    Set m_txtaOpFields(8, 4) = txtResourceNo09
    Set m_txtaOpFields(9, 4) = txtResourceNo10
    
    Set m_txtaOpFields(0, 5) = txtNorDur01
    Set m_txtaOpFields(1, 5) = txtNorDur02
    Set m_txtaOpFields(2, 5) = txtNorDur03
    Set m_txtaOpFields(3, 5) = txtNorDur04
    Set m_txtaOpFields(4, 5) = txtNorDur05
    Set m_txtaOpFields(5, 5) = txtNorDur06
    Set m_txtaOpFields(6, 5) = txtNorDur07
    Set m_txtaOpFields(7, 5) = txtNorDur08
    Set m_txtaOpFields(8, 5) = txtNorDur09
    Set m_txtaOpFields(9, 5) = txtNorDur10
    
    Set m_txtaOpFields(0, 6) = txtPlanWork01
    Set m_txtaOpFields(1, 6) = txtPlanWork02
    Set m_txtaOpFields(2, 6) = txtPlanWork03
    Set m_txtaOpFields(3, 6) = txtPlanWork04
    Set m_txtaOpFields(4, 6) = txtPlanWork05
    Set m_txtaOpFields(5, 6) = txtPlanWork06
    Set m_txtaOpFields(6, 6) = txtPlanWork07
    Set m_txtaOpFields(7, 6) = txtPlanWork08
    Set m_txtaOpFields(8, 6) = txtPlanWork09
    Set m_txtaOpFields(9, 6) = txtPlanWork10
    
    Set m_txtaOpFields(0, 7) = txtActWork01
    Set m_txtaOpFields(1, 7) = txtActWork02
    Set m_txtaOpFields(2, 7) = txtActWork03
    Set m_txtaOpFields(3, 7) = txtActWork04
    Set m_txtaOpFields(4, 7) = txtActWork05
    Set m_txtaOpFields(5, 7) = txtActWork06
    Set m_txtaOpFields(6, 7) = txtActWork07
    Set m_txtaOpFields(7, 7) = txtActWork08
    Set m_txtaOpFields(8, 7) = txtActWork09
    Set m_txtaOpFields(9, 7) = txtActWork10
    
    Set m_txtaOpFields(0, 8) = txtUnitOfWork01
    Set m_txtaOpFields(1, 8) = txtUnitOfWork02
    Set m_txtaOpFields(2, 8) = txtUnitOfWork03
    Set m_txtaOpFields(3, 8) = txtUnitOfWork04
    Set m_txtaOpFields(4, 8) = txtUnitOfWork05
    Set m_txtaOpFields(5, 8) = txtUnitOfWork06
    Set m_txtaOpFields(6, 8) = txtUnitOfWork07
    Set m_txtaOpFields(7, 8) = txtUnitOfWork08
    Set m_txtaOpFields(8, 8) = txtUnitOfWork09
    Set m_txtaOpFields(9, 8) = txtUnitOfWork10
    
    Set m_txtaOpFields(0, 9) = txtSysStatus01
    Set m_txtaOpFields(1, 9) = txtSysStatus02
    Set m_txtaOpFields(2, 9) = txtSysStatus03
    Set m_txtaOpFields(3, 9) = txtSysStatus04
    Set m_txtaOpFields(4, 9) = txtSysStatus05
    Set m_txtaOpFields(5, 9) = txtSysStatus06
    Set m_txtaOpFields(6, 9) = txtSysStatus07
    Set m_txtaOpFields(7, 9) = txtSysStatus08
    Set m_txtaOpFields(8, 9) = txtSysStatus09
    Set m_txtaOpFields(9, 9) = txtSysStatus10
    
    Me.sbOperations.SmallChange = 1
    Me.sbOperations.LargeChange = 5

End Sub


'==============================================================================
' FUNCTION
'   SetWorkOrder
'------------------------------------------------------------------------------
' DESCRIPTION
'   Populates the work order details from the database, based on the supplied
' work order number iWorkOrder.
'==============================================================================
Public Function SetWorkOrder(iWorkOrder As Long, ByRef oErr As typAppError, Optional bSuppressMissingNoti As Boolean = False) As Boolean

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
    Dim scaOperations(1 To 5, 1 To 4) As String
    Dim rng As Excel.Range
    Dim iNotification As Long
    Dim scVal As String
    Dim iVal As Long
    Dim dtVal As Date
    Dim dVal As Double
    
    oErr.number = eAppErrorCodes.aeNoError
    oErr.description = "<No Error>"

    Call Me.Blank(False)
    
    If iWorkOrder = 0 Then
        m_bSupressWorkOrderChangeHandling = True
        txtOrderNo.Text = ""
        m_bSupressWorkOrderChangeHandling = False
        
        SetWorkOrder = True
        
        Exit Function
    End If
    
    If Not (ConnectToDB(ldMaintenance, cnn, True, iErrorNo, scError)) Then
        oErr.number = iErrorNo
        oErr.description = scError
        oErr.Source = "ufWorkOrderDisplay.SetWorkOrder"
        
        If (Not oErr.BeSilent) Then
            Call MsgBox("Unable to connect to the 'maint' database. Error: " & oErr.number & " - " & oErr.description)
        End If
        SetWorkOrder = False
        Exit Function
    End If
    
    '==============
    ' Get the relevant work order data
    '==============
    scSQLQuery = "SELECT * FROM dbo.t_work_order WHERE pk_work_order = " & iWorkOrder
    
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)
    
    If rs.EOF Then
        oErr.number = eAppErrorCodes.aeWorkOrderNotFound
        oErr.description = "Work Order " & iWorkOrder & " not found in the database"
        oErr.Source = "ufWorkOrderDisplay.SetWorkOrder"
        Call DisplayError(oErr)
        SetWorkOrder = False
        Exit Function
    End If
    
    '============
    ' Display the result
    '============
    m_bSupressWorkOrderChangeHandling = True
    Me.txtOrderNo.Value = iWorkOrder
    m_bSupressWorkOrderChangeHandling = False
            
    For iField = 0 To (rs.Fields.Count - 1)
        Select Case rs.Fields(iField).Name
            Case "short_text"
                Me.txtOrderShortText.Value = GetStringFieldValueIdx(rs, iField)
            Case "order_type"
                Me.txtOrderType.Value = rs.Fields(iField).Value
            Case "fk_main_work_center"
                scVal = GetStringFieldValueIdx(rs, iField)
                'If Not IsNull(rs.Fields(iField).Value) Then scVal = GetStringFieldValueIdx(rs, iField)
                Me.txtMainWorkCenter.Value = scVal
                Me.txtMainWorkCenter2.Value = scVal
                
            Case "fk_planner_group"
                scVal = GetStringFieldValueIdx(rs, iField)
                'If Not IsNull(rs.Fields(iField).Value) Then scVal = rs.Fields(iField).Value
                Me.txtPlannerGroup.Value = scVal
                
            Case "fk_cost_center"
                scVal = GetStringFieldValueIdx(rs, iField)
                'If Not IsNull(rs.Fields(iField).Value) Then scVal = rs.Fields(iField).Value
                Me.txtCostCenter.Value = scVal
                Me.txtCostCenter2.Value = scVal
                
            Case "fk_notification"
                iVal = -1
                If Not IsNull(rs.Fields(iField).Value) Then
                    iVal = rs.Fields(iField).Value
                    Me.txtNotification.Value = iVal
                    iNotification = iVal
                End If
                
            Case "basic_start_date"
                dtVal = 0
                If Not IsNull(rs.Fields(iField).Value) Then
                    dtVal = rs.Fields(iField).Value
                    Me.txtBasicStartDate.Value = Format(dtVal, "DD.MM.YYYY")
                    Me.txtBasicStartTime.Value = Format(dtVal, "HH:MM")
                    Me.txtBasicStartDate2.Value = Me.txtBasicStartDate.Value
                    Me.txtBasicStartTime2.Value = Me.txtBasicStartTime.Value
                End If
                
            Case "fk_func_loc"
                Me.txtFuncLoc.Value = GetStringFieldValueIdx(rs, iField)
                Me.txtFuncLoc2.Value = GetStringFieldValueIdx(rs, iField)
            Case "revision"
                Me.txtRevision.Value = GetStringFieldValueIdx(rs, iField)
                Me.txtRevision2.Value = GetStringFieldValueIdx(rs, iField)
            Case "required_by_date"
                Me.txtRequiredByDate.Value = GetStringFieldValueIdx(rs, iField)
            Case "priority_text"
                Me.txtPriority.Value = GetStringFieldValueIdx(rs, iField)
                Me.txtPriority2.Value = GetStringFieldValueIdx(rs, iField)
            Case "sys_status"
                Me.txtSysStatus.Value = GetStringFieldValueIdx(rs, iField)
            Case "user_status"
                Me.txtUserStatus.Value = GetStringFieldValueIdx(rs, iField)
            Case "creation_on_date"
                Me.txtCreatedOn.Value = Format(rs.Fields(iField).Value, "DD.MM.YYYY")
            Case "entered_by"
                Me.txtEnteredBy.Value = GetStringFieldValueIdx(rs, iField)
            Case "last_changed_by"
                Me.txtChangedBy.Value = GetStringFieldValueIdx(rs, iField)
                
            Case "fk_sys_condition"
                '==================
                ' System Condition
                '------------------
                ' 0 - not in operation
                ' 1 - in operation
                '==================
                Me.txtSysCondition.Value = rs.Fields(iField).Value
                Select Case Me.txtSysCondition.Value
                    Case "0"
                        Me.lblSysConditionText.Caption = "not in operation"
                    Case "1"
                        Me.lblSysConditionText.Caption = "in operation"
                    Case Else
                End Select
                
            Case "fk_maint_act_type"
                '==================
                ' PM Activity type
                '------------------
                ' 100 - Corrective
                ' 120 - Accident Damage
                ' 130 - Rework
                ' 200 - Breakdown
                ' 210 - Improvement Project
                ' 330 - Condition Monitoring
                ' 340 - Safety
                '==================
                Me.txtPMActType.Value = rs.Fields(iField).Value
                Select Case Me.txtPMActType.Value
                    Case "100"
                        Me.lblPMActTypeText.Caption = "Corrective"
                    Case "120"
                        Me.lblPMActTypeText.Caption = "Accident Damage"
                    Case "130"
                        Me.lblPMActTypeText.Caption = "Rework"
                    Case "200"
                        Me.lblPMActTypeText.Caption = "Breakdown"
                    Case "210"
                        Me.lblPMActTypeText.Caption = "Improvement Project"
                    Case "330"
                        Me.lblPMActTypeText.Caption = "Condition Monitoring"
                    Case "340"
                        Me.lblPMActTypeText.Caption = "Safety"
                End Select
                
            Case "fk_maint_plan"
                Me.txtMaintenancePlan = rs.Fields(iField).Value
            Case "fk_maint_item"
                Me.txtMaintenanceItem = rs.Fields(iField).Value
            Case "last_updated_date"
                Me.txtLastUpdateFromSAP.Value = Format(rs.Fields(iField).Value, "ddd d-mmm-yy h:mm am/pm")
                
                dUpdateAge = (Date + Time) - rs.Fields(iField).Value
                
                Me.btnLastUpdateColour.BackColor = GetRGBFromAge(dUpdateAge, 0#, 1#, 2#)
'                If (dUpdateAge <= 0) Then
'                    ' less than 6 hours old -> Green
'                    Me.btnLastUpdateColour.BackColor = RGB(0, 255, 0) ' Green
'                ElseIf (dUpdateAge <= 1) Then
'                    ' less than 12 hours old -> slightly faded Green
'                    Me.btnLastUpdateColour.BackColor = RGB(155 * dUpdateAge, 255, 155 * dUpdateAge) ' Fading Green
'                ElseIf (dUpdateAge <= 2) Then
'                    ' less than 12 hours old -> very faded Green
'                    Me.btnLastUpdateColour.BackColor = RGB(255, 255 * (2 - dUpdateAge), 255 * (2 - dUpdateAge)) ' Increasing Red
'                Else
'                    Me.btnLastUpdateColour.BackColor = RGB(255, 0, 0) ' red
'                End If
                
            Case Else
                
        End Select
    Next
    
    Call rs.Close
    
    '==============
    ' Fill out some defaults
    '==============
    Me.txtMaintPlant = "2301"
    Me.txtCompanyCode = "2351"
    Me.lblCompanyCode = "Lihir Gold Ltd"
    Me.lblCompanyCode2 = "PNG"
    
    '==============
    ' Get the notification data if available
    '==============
    If iNotification > 0 Then
        
        scSQLQuery = "SELECT * FROM dbo.t_notification WHERE pk_notification = " & iNotification
        
        Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockOptimistic_)
        
        If rs.EOF Then
            oErr.number = eAppErrorCodes.aeWorkOrderNotFound
            oErr.description = "Notification " & iNotification & " not found in the database"
            oErr.Source = "ufWorkOrderDisplay.SetWorkOrder"
            If Not bSuppressMissingNoti Then
                Call DisplayError(oErr)
            End If
            Me.txtLongText.Value = ""
        Else
            Me.txtLongText.Value = rs.Fields("long_text")
        End If
        
        Call rs.Close
    Else
        Me.txtLongText.Value = ""
    End If
    
    '==============
    ' Get the functional location data if available
    '==============
    If Len(Me.txtFuncLoc.Value) > 0 Then
        
        scSQLQuery = "SELECT * FROM dbo.t_func_loc WHERE pk_func_loc = '" & Me.txtFuncLoc.Value & "'"
        
        Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockOptimistic_)
        
        If rs.EOF Then
            oErr.number = eAppErrorCodes.aeWorkOrderNotFound
            oErr.description = "Functional Location '" & Me.txtFuncLoc.Value & "' not found in the database"
            oErr.Source = "ufWorkOrderDisplay.SetWorkOrder"
            Call DisplayError(oErr)
        Else
            txtFuncLocDescription.Value = rs.Fields("description")
            Me.lblFuncLocText.Caption = rs.Fields("description")
            If Not IsNull(rs.Fields("sort_field")) Then txtSortField.Value = rs.Fields("sort_field")
        End If
        
        Call rs.Close
    End If
    
    '==============
    ' Get the associated operations
    '==============
    scSQLQuery = "SELECT * FROM dbo.t_operation WHERE fk_order = " & iWorkOrder & " ORDER BY operation_activity ASC"
    
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockOptimistic_)
    
    If rs.EOF Then
        oErr.number = eAppErrorCodes.aeNotificationNotfound
        oErr.description = "Notification for Work Order " & iWorkOrder & " not found in the database"
        oErr.Source = "ufWorkOrderDisplay.SetWorkOrder"
        Call DisplayError(oErr)
        SetWorkOrder = False
        Call modErrorHandler.DisplayError(oErr)
    Else
        '==============
        ' How many records
        '==============
        Dim iRowCount As Long
        Dim iRow As Long
    
        m_iOperationsCount = rs.RecordCount
        
        If (m_iOperationsCount < 1) Then
            Call MsgBox("Database says there are no operations found for work order " & Format(iWorkOrder, "0"))
            Exit Function
        End If
    
        '==============
        ' First we suck every operation out of the database into our array structure
        '==============
        ReDim m_taOpFields(m_iOperationsCount)
        
        iRowCount = 0
        While Not rs.EOF And (iRowCount < m_iOperationsCount)
            m_taOpFields(iRowCount).operation_activity = GetStringFieldValue(rs, "operation_activity")
            m_taOpFields(iRowCount).short_text = GetStringFieldValue(rs, "short_text")
            m_taOpFields(iRowCount).fk_op_work_centre = GetStringFieldValue(rs, "fk_op_work_centre")
            m_taOpFields(iRowCount).earliest_start = GetDateFieldValue(rs, "earliest_start")
            m_taOpFields(iRowCount).no_of_resource = GetLongFieldValue(rs, "no_of_resource")
            m_taOpFields(iRowCount).normal_duration = GetDoubleFieldValue(rs, "normal_duration")
            m_taOpFields(iRowCount).planned_work = GetDoubleFieldValue(rs, "planned_work")
            m_taOpFields(iRowCount).actual_work = GetDoubleFieldValue(rs, "actual_work")
            m_taOpFields(iRowCount).unit_of_work = GetStringFieldValue(rs, "unit_of_work")
            m_taOpFields(iRowCount).sys_status = GetStringFieldValue(rs, "sys_status")
            
            rs.MoveNext
            iRowCount = iRowCount + 1
        Wend
        
        '==============
        ' Set the scrollbar based on the number of records.
        '==============
        If m_iOperationsCount <= 10 Then
            sbOperations.Min = 0
            sbOperations.Max = 0
        Else
            sbOperations.Min = 0
            sbOperations.Max = m_iOperationsCount - 10
        End If
        
        '===========
        ' Display the operations
        '===========
        Call sbOperations_Change
    End If
    
    SetWorkOrder = True
    
End Function

Private Sub sbOperations_Change()
    Dim iOffset As Long
    Dim iRow As Long
    Dim iOpIndex As Long
    
    iOffset = sbOperations.Value
    
    iRow = 0
    iOpIndex = iOffset
    While iRow < 10 And iOpIndex < m_iOperationsCount
        
        m_txtaOpFields(iRow, 0) = m_taOpFields(iOpIndex).operation_activity
        m_txtaOpFields(iRow, 1) = m_taOpFields(iOpIndex).short_text
        m_txtaOpFields(iRow, 2) = m_taOpFields(iOpIndex).fk_op_work_centre
        m_txtaOpFields(iRow, 3) = Format(m_taOpFields(iOpIndex).earliest_start, "dd.mm.yyyy")
        m_txtaOpFields(iRow, 4) = m_taOpFields(iOpIndex).no_of_resource
        m_txtaOpFields(iRow, 5) = m_taOpFields(iOpIndex).normal_duration
        m_txtaOpFields(iRow, 6) = m_taOpFields(iOpIndex).planned_work
        m_txtaOpFields(iRow, 7) = m_taOpFields(iOpIndex).actual_work
        m_txtaOpFields(iRow, 8) = m_taOpFields(iOpIndex).unit_of_work
        m_txtaOpFields(iRow, 9) = m_taOpFields(iOpIndex).sys_status

        iRow = iRow + 1
        iOpIndex = iOpIndex + 1
    Wend
    
    While (iRow < 9)
        m_txtaOpFields(iRow, 0) = ""
        m_txtaOpFields(iRow, 1) = ""
        m_txtaOpFields(iRow, 2) = ""
        m_txtaOpFields(iRow, 3) = ""
        m_txtaOpFields(iRow, 4) = ""
        m_txtaOpFields(iRow, 5) = ""
        m_txtaOpFields(iRow, 6) = ""
        m_txtaOpFields(iRow, 7) = ""
        m_txtaOpFields(iRow, 8) = ""
        m_txtaOpFields(iRow, 9) = ""
        
        iRow = iRow + 1
    Wend

End Sub

'==============================================================================
' SUBROUTINE
'   Blank
'------------------------------------------------------------------------------
' DESCRIPTION
'   Blanks the entire work order display
'==============================================================================
Public Sub Blank(Optional bIncludeWO As Boolean = True)

    If bIncludeWO Then
        m_bSupressWorkOrderChangeHandling = True
        Me.txtOrderNo.Value = ""
        m_bSupressWorkOrderChangeHandling = False
    End If

    Me.txtActualStartDate.Value = ""
    Me.txtActualEndDate.Value = ""
    Me.txtAssembly.Value = ""
    Me.txtBasicStartDate.Value = ""
    Me.txtBasicStartDate2.Value = ""
    Me.txtBasicStartTime.Value = ""
    Me.txtBasicStartTime2.Value = ""
    Me.txtBasicFinishDate.Value = ""
    Me.txtBasicFinishTime.Value = ""
    Me.txtChangedBy.Value = ""
    Me.txtChangedOnDate = ""
    Me.txtCompanyCode = ""
    Me.txtCostCenter.Value = ""
    Me.txtCostCenter2.Value = ""
    Me.lblCostCenter.Caption = ""
    Me.lblCostCenter2.Caption = ""
    Me.txtCreatedOn.Value = ""
    Me.txtEnteredBy.Value = ""
    Me.txtFuncLoc.Value = ""
    Me.txtFuncLoc2.Value = ""
    Me.txtFuncLocDescription = ""
    Me.txtLongText.Value = ""
    Me.txtMaintPlant.Value = ""
    Me.txtMainWorkCenter.Value = ""
    Me.txtMainWorkCenter2.Value = ""
    Me.txtNotification.Value = ""
    Me.txtOrderShortText.Value = ""
    Me.txtOrderType.Value = ""
    Me.txtPlannedCosts.Value = ""
    Me.txtPlannerGroup.Value = ""
    Me.txtPMActType.Value = ""
    Me.txtPriority.Value = ""
    Me.txtPriority2.Value = ""
    Me.txtRevision.Value = ""
    Me.txtRevision2.Value = ""
    Me.txtRequiredByDate.Value = ""
    Me.txtSchedFinishDate.Value = ""
    Me.txtSchedFinishTime.Value = ""
    Me.txtSchedStartDate.Value = ""
    Me.txtSchedStartTime.Value = ""
    Me.txtSettlementOrder.Value = ""
    Me.txtSortField.Value = ""
    Me.txtSysCondition.Value = ""
    Me.txtSysStatus.Value = ""
    Me.txtUserStatus.Value = ""
    
    'Me.lblAssemblyText.Caption = ""
    Me.lblCompanyCode.Caption = "2351"
    Me.lblCurrency.Caption = "USD"
    Me.lblFuncLocText.Caption = ""
    Me.lblMainWorkCenter = ""
    Me.lblMainWorkCenter2 = ""
    Me.lblPlannerGroupText = ""
    Me.lblPMActTypeText.Caption = ""
    Me.lblSysConditionText.Caption = ""
    
    Me.txtLastUpdateFromSAP.Value = ""
    Me.btnLastUpdateColour.BackColor = &H8000000F  ' Button Face
    
    Call BlankOperations

End Sub

'==============================================================================
' SUBROUTINE
'   BlankOperations
'------------------------------------------------------------------------------
' DESCRIPTION
'   Puts the operations display into an initial state.
'==============================================================================
Sub BlankOperations()
    '==========
    ' Blank all text boxes
    '==========
    Dim iRow As Long
    Dim iCol As Long
    
    For iRow = 0 To 9
        For iCol = 0 To 9
            m_txtaOpFields(iRow, iCol).Value = ""
        Next
    Next
    
    Me.sbOperations.Min = 1
    Me.sbOperations.Max = 1
    Me.sbOperations.Value = 1
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
        
    Dim iWO As Long
    Dim oErr As typAppError
    
On Error GoTo cleanup_nicely
    iWO = Val(txtOrderNo.Text)
    
    Call Me.Blank
    Call Me.BlankOperations
    
    Call Me.SetWorkOrder(iWO, oErr)
    Exit Sub
    
cleanup_nicely:
    Call MsgBox("Unable to refresh: #" & oErr.number & "-'" & oErr.description & "'")
    
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   txtOrderNo_Change
'------------------------------------------------------------------------------
' DESCRIPTION
'   Event handler for when the user changes the work order number.
'==============================================================================
Private Sub txtOrderNo_Change()
    Dim oErr As typAppError
    Dim iWO As Long

    If m_bSupressWorkOrderChangeHandling Then
        Exit Sub
    End If
    
    Call Blank(False)
    
    If LooksLikeValidWorkOrder(Me.txtOrderNo.Text, iWO) Then
        If Not Me.SetWorkOrder(iWO, oErr, True) Then
            Call MsgBox("Error: " & oErr.description)
        End If
    End If
End Sub
