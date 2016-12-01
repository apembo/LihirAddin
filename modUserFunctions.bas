Attribute VB_Name = "modUserFunctions"
Option Explicit

Public Sub DescribeFunctions()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 3) As String
    
    FuncName = "LihirGetLogValue"
    FuncDesc = "Returns the log value corresponding to the supplied parameters"
    Category = 5 'Lookup and Reference
    ArgDesc(1) = "String that contains the Object Name (e.g. DG01, Linde)"
    ArgDesc(2) = "String that contains the object parameter (e.g. %, MW.h)"
    ArgDesc(3) = "Date/Time of the sample. Most will be midnight"

    Application.MacroOptions _
        Macro:=FuncName, _
        description:=FuncDesc, _
        Category:=Category, _
        ArgumentDescriptions:=ArgDesc
        
    'LihirGetMaintNotifParam(scParamName As String, scID As String)
    FuncName = "LihirGetMaintNotifParam"
    FuncDesc = "Returns the specified parameter belonging to the specified notification, or blank"
    Category = 5 'Lookup and Reference
    Dim ArgDesc2(1 To 2) As String
    ArgDesc2(1) = "Parameter name"
    ArgDesc2(2) = "Notification Number"


    Application.MacroOptions _
        Macro:=FuncName, _
        description:=FuncDesc, _
        Category:=Category, _
        ArgumentDescriptions:=ArgDesc2

    FuncName = "LihirGetMaintOrderParam"
    FuncDesc = "Returns the specified parameter belonging to the specified work order, or blank"
    Category = 5 'Lookup and Reference
    Dim ArgDesc3(1 To 2) As String
    ArgDesc3(1) = "Parameter name"
    ArgDesc3(2) = "Work order Number"


    Application.MacroOptions _
        Macro:=FuncName, _
        description:=FuncDesc, _
        Category:=Category, _
        ArgumentDescriptions:=ArgDesc3
        
    FuncName = "LihirCCEECostSummary"
    FuncDesc = "Returns the spend sum for the supplied cost center, financial year and optional GLAccount and Period."
    Category = 1 'Financial
    Dim ArgDesc4(1 To 4) As String
    ArgDesc4(1) = "Cost Center"
    ArgDesc4(2) = "Financial Year (e.g. 2016 = FY16)"
    ArgDesc4(3) = "Period (Optional). 1 = July, 12 = June."
    ArgDesc4(4) = "G/L Account (Optional)"


    Application.MacroOptions _
        Macro:=FuncName, _
        description:=FuncDesc, _
        Category:=Category, _
        ArgumentDescriptions:=ArgDesc4
        
    FuncName = "LihirCCEEQtySummary"
    FuncDesc = "Returns the quantity sum for the supplied cost center, financial year and optional GLAccount and Period."
    Category = 1 'Financial

    Application.MacroOptions _
        Macro:=FuncName, _
        description:=FuncDesc, _
        Category:=Category, _
        ArgumentDescriptions:=ArgDesc4
        
End Sub

'==============================================================================
' FUNCTION
'   LihirGetLogValue
'------------------------------------------------------------------------------
' DESCRIPTION
'   Get's a log value from the pu database for the given object name, parameter
' and Sample Date/time
'==============================================================================
Public Function LihirGetLogValue(scObjName As String, scObjParam As String, dtSampleDateTime As Date)
Attribute LihirGetLogValue.VB_Description = "Returns the log value corresponding to the supplied parameters"
Attribute LihirGetLogValue.VB_ProcData.VB_Invoke_Func = " \n5"

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    '==========
    ' Connect to the database
    '==========
'    If Not (ConnectToDB(ldPowerAndUtilities, cnn, True)) Then
'        Call MsgBox("Unable to connect to the DB")
'        LihirGetLogValue = xlErrNA
'        Exit Function
'    End If
    
    Set rs = CreateObject("ADODB.Recordset")
    
    scSQLQuery = "SELECT TOP 1 value_real FROM pu.dbo.v_log_values WHERE obj_name = '" & Replace(scObjName, "'", "''") & _
        "' AND param = '" & Replace(scObjParam, "'", "''") & _
        "' AND sample_date_time = '" & Format(dtSampleDateTime, "YYYY-MM-DD hh:mm:ss") & "'"
        
    Call GetDBRecordSet(ldPowerAndUtilities, cnn, scSQLQuery, rs)
        
'    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockOptimistic_)
    
    If Not rs.EOF Then
        LihirGetLogValue = rs.Fields("value_real")
    Else
        LihirGetLogValue = xlErrValue
    End If


End Function

'==============================================================================
' FUNCTION
'   LihirGetMaintNotifParam
'------------------------------------------------------------------------------
' DESCRIPTION
'   Get's the value of the supplied parameter in the supplied notification.
'==============================================================================
Public Function LihirGetMaintNotifParam(scParamName As String, scID As String) As Variant
Attribute LihirGetMaintNotifParam.VB_Description = "Returns the specified parameter belonging to the specified notification, or blank"
Attribute LihirGetMaintNotifParam.VB_ProcData.VB_Invoke_Func = " \n5"

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    Dim scParam As String

    '==========
    ' Connect to the database
    '==========
'    If Not (ConnectToDB(ldMaintenance, cnn, True)) Then
'        'Call MsgBox("Unable to connect to the DB")
'        LihirGetMaintNotifParam = "DB Connection Failure"
'        Exit Function
'    End If
    
    Select Case LCase(scParamName)
        Case "[pk_notification]", "[short_text]", "[fk_func_loc]", "[flocdescription]", _
            "[sort_field]", "[fk_main_work_centre]", "[required_end]", "[long_text]", _
            "[reported_by]", "[fk_order]", "[noti_user_status]", "[noti_sys_status]", _
            "[notification_date]", "[notification_type]", "[order_short_text]", _
            "[order_type]", "[flocfnfo]", "[notificationinfo]", "[parentnavi]", _
            "[parentnavidescription]", "[parentnaviinfo]", _
            "pk_notification", "short_text", "fk_func_loc", "flocdescription", _
            "sort_field", "fk_main_work_centre", "required_end", "long_text", _
            "reported_by", "fk_order", "noti_user_status", "noti_sys_status", _
            "notification_date", "notification_type", "order_short_text", _
            "order_type", "flocinfo", "notificationinfo", "parentnavi", _
            "parentnavidescription", "parentnaviinfo"
            
            scParam = scParamName
            
        Case "title", "shorttext"
            scParam = "short_text"
            
        Case "floc", "functional location", "functionallocation"
            scParam = "fk_func_loc"
            
        Case "flocdescription", "description", "floc description"
            scParam = "flocdescription"
                
        Case "sortfield", "sort field"
            scParam = "sort_field"
                
        Case "mainworkcentre", "mainworkcenter", "main work centre", "main work center", "mwc"
            scParam = "fk_main_work_centre"
                
        Case "required end", "required end date", "end date", "requiredend", "requiredenddate"
            scParam = "required_end"
                
        Case "detail", "longtext", "text detail"
            scParam = "long_text"
                
        Case "reportedby", "reported by"
            scParam = "reported_by"
                
        Case "work order", "order", "wo", "work order number", "workorder"
            scParam = "fk_order"
                
        Case "notification user status", "user status", "notificationuserstatus", "userstatus"
            scParam = "noti_user_status"
                
        Case "notification system status", "system status", "notificationsystemstatus", "systemstatus"
            scParam = "noti_sys_status"
                
        Case "raised date", "date", "raiseddate", "date raised", "dateraised"
            scParam = "notification_date"
                
        Case "type", "notification type", "notificationtype"
            scParam = "notification_type"
                
        Case "work order title", "workordertitle", "work order short text", "workordershorttext", "wo title", "wo short text"
            scParam = "order_short_text"
                
                
        Case "order type", "ordertype", "work order type", "workordertype", "wo type"
            scParam = "order_type"
                
                
        Case "floc info"
            scParam = "flocinfo"
                
                
        Case "notification info"
            scParam = "notificationinfo"
                
                
        Case "parent navi", "superior navi"
            scParam = "parentnavi"
                
                
        Case "parent navi description"
            scParam = "parentnavidescription"
                
                
        Case "parent navi info"
            scParam = "parentnaviinfo"
                
        Case Else
            LihirGetMaintNotifParam = "Parameter '" & scParamName & "' not recognised"
            Exit Function
                
    End Select
    
    '==========
    ' Remove extraneous spaces and strip leading and trailing square brackets
    '==========
    scParam = Trim(scParam)
    If Left(scParam, 1) = "[" Then
        scParam = Mid(scParam, 2, Len(scParam) - 2)
    End If
    
    '==========
    ' Extract from the database
    '==========
    scSQLQuery = "SELECT " & scParam & " FROM dbo.v_notifications WHERE pk_notification = " & Val(scID)
    
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)
    
    If Not rs.EOF Then
        If IsNull(rs.Fields(scParam)) Then
            LihirGetMaintNotifParam = ""
        Else
            LihirGetMaintNotifParam = rs.Fields(scParam)
        End If
    Else
        LihirGetMaintNotifParam = "Info not found. Check notification " & scID & " exists."
    End If

End Function

'==============================================================================
' FUNCTION
'   LihirGetMaintNotifParam
'------------------------------------------------------------------------------
' DESCRIPTION
'   Get's the value of the supplied parameter in the supplied notification.
'==============================================================================
Public Function LihirGetMaintOrderParam(scParamName As String, scID As String) As Variant
Attribute LihirGetMaintOrderParam.VB_Description = "Returns the specified parameter belonging to the specified work order, or blank"
Attribute LihirGetMaintOrderParam.VB_ProcData.VB_Invoke_Func = " \n5"

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    Dim scParam As String

    '==========
    ' Connect to the database
    '==========
'    If Not (ConnectToDB(ldMaintenance, cnn, True)) Then
'        'Call MsgBox("Unable to connect to the DB")
'        LihirGetMaintOrderParam = "DB Connection Failure"
'        Exit Function
'    End If
    
    Select Case LCase(scParamName)
            
        Case "pk_work_order", "short_text", "order_type", "fk_main_work_center", "fk_planner_group", "fk_cost_center", "fk_notification", "basic_start_date", _
            "fk_func_loc", "revision", "required_by_date", "priority_text", "sys_status", "user_status", "creation_on_date", "entered_by", "last_changed_by", _
            "fk_sys_condition", "fk_maint_act_type", "last_updated_date", "fk_custom_category_1", "fk_custom_category_2", "fk_maint_item", "fk_maint_plan", _
            "estimated_costs", "total_planned_costs", "total_actual_costs", "sched_start", "sched_finish", "floc_description", "NotificationLongText", "sort_field"
            
            scParam = scParamName
            
        Case "title", "shorttext"
            scParam = "short_text"
            
        Case "floc", "functional location", "functionallocation"
            scParam = "fk_func_loc"
            
        Case "flocdescription", "description", "floc description"
            scParam = "floc_description"
                
        Case "sortfield", "sort field"
            scParam = "sort_field"
                
        Case "mainworkcentre", "mainworkcenter", "main work centre", "main work center", "mwc"
            scParam = "fk_main_work_centre"
                
        Case "required by", "required date", "required by date", "requireddate", "required_date", "required_by_date"
            scParam = "required_by_date"
                
        Case "detail", "longtext", "text detail"
            scParam = "NotificationLongText"
                
        Case "work order title", "workordertitle", "work order short text", "workordershorttext", "wo title", "wo short text"
            scParam = "short_text"
                
                
        Case "order type", "ordertype", "work order type", "workordertype", "wo type"
            scParam = "order_type"
                
                
        Case "notification"
            scParam = "fk_notification"
                
                
        Case "parent navi", "superior navi"
            scParam = "fk_parent_navi"
                
        Case Else
            LihirGetMaintOrderParam = "Parameter '" & scParamName & "' not recognised"
            Exit Function
                
    End Select
    
    '==========
    ' Remove extraneous spaces and strip leading and trailing square brackets
    '==========
    scParam = Trim(scParam)
    If Left(scParam, 1) = "[" Then
        scParam = Mid(scParam, 2, Len(scParam) - 2)
    End If
    
    '==========
    ' Extract from the database
    '==========
    scSQLQuery = "SELECT " & scParam & " FROM dbo.v_work_orders WHERE pk_work_order = " & Val(scID)
    
    Call GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs)
    
    If Not rs.EOF Then
        If IsNull(rs.Fields(scParam)) Then
            LihirGetMaintOrderParam = ""
        Else
            LihirGetMaintOrderParam = rs.Fields(scParam)
        End If
    Else
        LihirGetMaintOrderParam = "Info not found. Check work order " & scID & " exists."
    End If

End Function

'==============================================================================
' FUNCTION
'   LihirCCEECostSummary
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the spend sum for the supplied cost center, financial year and
' optional GLAccount and Period.
'==============================================================================
Public Function LihirCCEECostSummary(scCostCenter As String, iFinYear As Long, Optional iPeriod As Long = -1, Optional iGLAccount As Long = -1) As Double
Attribute LihirCCEECostSummary.VB_Description = "Returns the spend sum for the supplied cost center, financial year and optional GLAccount and Period."
Attribute LihirCCEECostSummary.VB_ProcData.VB_Invoke_Func = " \n1"

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    Dim dtPeriodStart As Date
    
    '===========
    ' Do a quick sanity check on the supplied period. This allows us to return 0
    ' if it's in the future, significantly reducing unnecessary DB calls.
    '===========
    dtPeriodStart = GetStartDateFromFYPeriod(iFinYear, iPeriod)
    
    If dtPeriodStart > Date Then
        LihirCCEECostSummary = 0
        Exit Function
    End If
        
    
    scSQLQuery = "SELECT SUM(ValueUSD) as SummarySpend " & _
        "FROM [finance].[dbo].[v_gen_ledger] " & _
        "WHERE CostCenter = '" & scCostCenter & "' AND FinYear = " & iFinYear
        
    If iGLAccount > 0 Then
        scSQLQuery = scSQLQuery & " AND CostElement = " & iGLAccount
    End If
    If (iPeriod >= 1) And (iPeriod <= 12) Then
        scSQLQuery = scSQLQuery & " AND period = " & iPeriod
    End If
        
    Call GetDBRecordSet(ldFinance, cnn, scSQLQuery, rs)
    
    If rs.EOF Then
        LihirCCEECostSummary = 0
    ElseIf IsNull(rs.Fields("SummarySpend")) Then
        LihirCCEECostSummary = 0
    Else
        LihirCCEECostSummary = rs.Fields("SummarySpend")
    End If
End Function

'==============================================================================
' FUNCTION
'   LihirCCEEQtySummary
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the quantity sum for the supplied cost center, financial year and
' optional GLAccount and Period.
'==============================================================================
Public Function LihirCCEEQtySummary(scCostCenter As String, iFinYear As Long, Optional iPeriod As Long = -1, Optional iGLAccount As Long = -1) As Double
Attribute LihirCCEEQtySummary.VB_Description = "Returns the quantity sum for the supplied cost center, financial year and optional GLAccount and Period."
Attribute LihirCCEEQtySummary.VB_ProcData.VB_Invoke_Func = " \n1"

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
#Else
    Dim cnn As Object
    Dim rs As Object
#End If
    Dim scSQLQuery As String

    Dim dtPeriodStart As Date
    
    '===========
    ' Do a quick sanity check on the supplied period. This allows us to return 0
    ' if it's in the future, significantly reducing unnecessary DB calls.
    '===========
    dtPeriodStart = GetStartDateFromFYPeriod(iFinYear, iPeriod)
    
    If dtPeriodStart > Date Then
        LihirCCEEQtySummary = 0
        Exit Function
    End If
    
    
    scSQLQuery = "SELECT SUM(Qty) as SummaryQty " & _
        "FROM [finance].[dbo].[v_gen_ledger] " & _
        "WHERE CostCenter = '" & scCostCenter & "' AND FinYear = " & iFinYear
        
    If iGLAccount > 0 Then
        scSQLQuery = scSQLQuery & " AND CostElement = " & iGLAccount
    End If
    If (iPeriod >= 1) And (iPeriod <= 12) Then
        scSQLQuery = scSQLQuery & " AND period = " & iPeriod
    End If
        
    Call GetDBRecordSet(ldFinance, cnn, scSQLQuery, rs)
    
    
    If rs.EOF Then
        LihirCCEEQtySummary = 0
    ElseIf IsNull(rs.Fields("SummaryQty")) Then
        LihirCCEEQtySummary = 0
    Else
        LihirCCEEQtySummary = rs.Fields("SummaryQty")
    End If

End Function

Public Function GetStartDateFromFYPeriod(iFinYear As Long, iPeriod As Long)
    Dim dtTemp As Long
    
    If iPeriod <= 6 Then
        GetStartDateFromFYPeriod = DateSerial(iFinYear - 1, iPeriod + 6, 1)
    Else
        GetStartDateFromFYPeriod = DateSerial(iFinYear, iPeriod - 6, 1)
    End If
End Function

Public Function LihirGetUser() As String
    LihirGetUser = Application.UserName
End Function

Public Function LihirGetUserLogin() As String
    Dim scTemp As String
    
    Call GetUserName(scTemp)
    LihirGetUserLogin = scTemp
End Function
