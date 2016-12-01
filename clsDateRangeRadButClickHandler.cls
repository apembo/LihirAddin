VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDateRangeRadButClickHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
' CLASS
'   clsDateRangeRadButClickHandler
'------------------------------------------------------------------------------
' DESCRIPTION
'   Event handler class for Option Buttons
'==============================================================================

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public WithEvents btn As MSForms.OptionButton
Attribute btn.VB_VarHelpID = -1

Public OptionButtonType As eDateRangeTypes
Public DateRangeUserform As ufDateRangeSelection

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btn_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'   Event handler for Option Buttons
'==============================================================================
Private Sub btn_Click()
    Dim dtFirstDate As Date
    Dim dtLastDate As Date

    Select Case OptionButtonType
    
        Case eDateRangeTypes.drtFinancialYear
            Call DateRangeUserform.GetDateRange(btn.Caption, drtFinancialYear, dtFirstDate, dtLastDate)
            Call DateRangeUserform.SetMonthOptionButtons(dtFirstDate)
            DateRangeUserform.FirstDate = dtFirstDate
            DateRangeUserform.LastDate = dtLastDate
            
            Call DateRangeUserform.ClearCalYearButtons
            Call DateRangeUserform.ClearMonthButtons
            
        Case eDateRangeTypes.drtCalendarYear
            Call DateRangeUserform.GetDateRange(btn.Caption, drtCalendarYear, dtFirstDate, dtLastDate)
            Call DateRangeUserform.SetMonthOptionButtons(dtFirstDate)
            DateRangeUserform.FirstDate = dtFirstDate
            DateRangeUserform.LastDate = dtLastDate
        
            Call DateRangeUserform.ClearFinYearButtons
            Call DateRangeUserform.ClearMonthButtons
        
        Case eDateRangeTypes.drtMonth
            Call DateRangeUserform.GetDateRange(btn.Caption, drtMonth, dtFirstDate, dtLastDate)
            DateRangeUserform.FirstDate = dtFirstDate
            DateRangeUserform.LastDate = dtLastDate
            
            Call DateRangeUserform.ClearCalYearButtons
            Call DateRangeUserform.ClearFinYearButtons
    End Select
End Sub
