VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufDateRangeSelection 
   Caption         =   "Date Range Selection"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5475
   OleObjectBlob   =   "ufDateRangeSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufDateRangeSelection"
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
#If DebugBadType = 0 Then
    Private m_aradCalendarYears(1 To 4) As MSForms.OptionButton
    Private m_aradFinancialYears(1 To 4) As MSForms.OptionButton
    Private m_aradMonths(1 To 12) As MSForms.OptionButton
    Private m_oaOptionButtonHandlers(1 To 20) As clsDateRangeRadButClickHandler
#Else
    Private m_aradCalendarYears(1 To 4) As Object
    Private m_aradFinancialYears(1 To 4) As Object
    Private m_aradMonths(1 To 12) As Object
    Private m_oaOptionButtonHandlers(1 To 20) As Object
#End If

Private m_bFormResult As VBA.VbMsgBoxResult

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   butCancel_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub butCancel_Click()
    m_bFormResult = vbCancel
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   butOK_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub butOK_Click()
    m_bFormResult = vbOK
    Call Hide
End Sub

'==============================================================================
' PROPERTY - GET
'   Result
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Public Property Get Result() As VBA.VbMsgBoxResult
    Result = m_bFormResult
End Property


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the form is created.
'==============================================================================
Private Sub UserForm_Initialize()

#If DebugBadType = 0 Then
    Dim cOptionButtonClickHandler As clsDateRangeRadButClickHandler
#Else
    Dim cOptionButtonClickHandler As Object
#End If

    Dim i As Long
    
    Set m_aradCalendarYears(1) = Me.radioCalYr1
    Set m_aradCalendarYears(2) = Me.radioCalYr2
    Set m_aradCalendarYears(3) = Me.radioCalYr3
    Set m_aradCalendarYears(4) = Me.radioCalYr4
    
    Set m_aradFinancialYears(1) = Me.radioFinYr1
    Set m_aradFinancialYears(2) = Me.radioFinYr2
    Set m_aradFinancialYears(3) = Me.radioFinYr3
    Set m_aradFinancialYears(4) = Me.radioFinYr4
    
    Set m_aradMonths(1) = Me.radioMonth1
    Set m_aradMonths(2) = Me.radioMonth2
    Set m_aradMonths(3) = Me.radioMonth3
    Set m_aradMonths(4) = Me.radioMonth4
    Set m_aradMonths(5) = Me.radioMonth5
    Set m_aradMonths(6) = Me.radioMonth6
    Set m_aradMonths(7) = Me.radioMonth7
    Set m_aradMonths(8) = Me.radioMonth8
    Set m_aradMonths(9) = Me.radioMonth9
    Set m_aradMonths(10) = Me.radioMonth10
    Set m_aradMonths(11) = Me.radioMonth11
    Set m_aradMonths(12) = Me.radioMonth12
    
    For i = 1 To 4
        m_aradCalendarYears(i).Caption = CalYear(i - 3)
        Set cOptionButtonClickHandler = New clsDateRangeRadButClickHandler
        Set cOptionButtonClickHandler.btn = m_aradCalendarYears(i)
        Set cOptionButtonClickHandler.DateRangeUserform = Me
        cOptionButtonClickHandler.OptionButtonType = drtCalendarYear
        
        Set m_oaOptionButtonHandlers(i) = cOptionButtonClickHandler
        
        m_aradFinancialYears(i).Caption = FinYear(i - 3)
        Set cOptionButtonClickHandler = New clsDateRangeRadButClickHandler
        Set cOptionButtonClickHandler.btn = m_aradFinancialYears(i)
        Set cOptionButtonClickHandler.DateRangeUserform = Me
        cOptionButtonClickHandler.OptionButtonType = drtFinancialYear
    
        Set m_oaOptionButtonHandlers(i + 4) = cOptionButtonClickHandler
    Next
    
    For i = 1 To 12
        Set cOptionButtonClickHandler = New clsDateRangeRadButClickHandler
        Set cOptionButtonClickHandler.btn = m_aradMonths(i)
        Set cOptionButtonClickHandler.DateRangeUserform = Me
        cOptionButtonClickHandler.OptionButtonType = drtMonth
        Set m_oaOptionButtonHandlers(i + 8) = cOptionButtonClickHandler
    Next
    
    m_aradFinancialYears(3).Value = True
    
    Call RemoveUserformCloseButton(Me)

End Sub

'==============================================================================
' PROPERTY - GET/LET
'   FirstDate
'------------------------------------------------------------------------------
' DESCRIPTION
'   Get's or sets the FirstDate text box field
'==============================================================================
Public Property Get FirstDate() As Date
    FirstDate = CDate(txtFirstDate.Text)
End Property
Public Property Let FirstDate(rhs As Date)
    txtFirstDate.Text = Format(rhs, "d-mmm-yy")
End Property

'==============================================================================
' PROPERTY - GET/LET
'   LastDate
'------------------------------------------------------------------------------
' DESCRIPTION
'   Get's or sets the LastDate text box field
'==============================================================================
Public Property Get LastDate() As Date
    LastDate = CDate(txtLastDate.Text)
End Property
Public Property Let LastDate(rhs As Date)
    txtLastDate.Text = Format(rhs, "d-mmm-yy")
End Property

'==============================================================================
' SUBROUTINE
'   GetDateRange
'------------------------------------------------------------------------------
' DESCRIPTION
'   Static function (does not use dynamic values in this userform) to get the
' start and end dates of a particular range/range type.
' For example specifying FY14 and eDateRangeTypes.drtFinancialYear will return
' start and end dates of 1-Jul-13 and 30-Jun-14 respectively.
'==============================================================================
Public Sub GetDateRange(scRange As String, eRangeType As eDateRangeTypes, ByRef dtStart As Date, ByRef dtEnd As Date)

    Dim iYear As Long
    Dim dtTemp As Date

    Select Case eRangeType
        Case eDateRangeTypes.drtCalendarYear
            iYear = Val(scRange)
            dtStart = DateSerial(iYear, 1, 1)
            dtEnd = DateSerial(iYear, 12, 31)
            
        Case eDateRangeTypes.drtFinancialYear
            iYear = Year4DigitFrom2(Val(Right(scRange, 2)))
            dtStart = DateSerial(iYear - 1, 7, 1)
            dtEnd = DateSerial(iYear, 6, 30)
            
        Case eDateRangeTypes.drtMonth
            dtTemp = DateSerial(Val(Right(scRange, 2)) + 2000, Month(CDate(scRange)), 1)
            
            dtStart = DateSerial(Year(dtTemp), Month(dtTemp), 1)
            dtEnd = DateSerial(Year(dtTemp), Month(dtTemp) + 1, 1) - 1
            
        Case Else
            dtStart = DateSerial(1970, 1, 1)
            dtEnd = dtStart
    End Select
End Sub

'==============================================================================
' PROPERTY - GET
'   FinYear
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the financial year (string format 'FYxx') corresponding to the
' index where the index is relative to the current financial year.
' For example, 0 will produce a string representing the current financial year,
' -2 will produce a string representing 2 financial years ago etc.
'==============================================================================
Private Property Get FinYear(iIndex As Long) As String
    If Date < DateSerial(Year(Date), 7, 1) Then
        FinYear = "FY" & Right(Year(Date) + iIndex, 2)
    Else
        FinYear = "FY" & Right(Year(Date) + iIndex + 1, 2)
    End If
End Property

'==============================================================================
' PROPERTY - GET
'   CalYear
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the calendar year (4 digit integer) corresponding to the
' index where the index is relative to the current calendar year.
' For example, 0 will produce a the current calendar year, -2 will give the
' year from 2 calendar years ago etc.
'==============================================================================
Private Property Get CalYear(iIndex As Long) As Long
    CalYear = Year(Date) + iIndex
End Property

'==============================================================================
' PROPERTY - GET
'   FinYear
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the financial year (string format 'FYxx') corresponding to the
' index where the index is relative to the current financial year.
' For example, 0 will produce a string representing the current financial year,
' -2 will produce a string representing 2 financial years ago etc.
'==============================================================================
Public Function Year4DigitFrom2(i2DigitYear As Long)
    If (i2DigitYear >= 70) Then
        Year4DigitFrom2 = i2DigitYear + 1900
    Else
        Year4DigitFrom2 = i2DigitYear + 2000
    End If
End Function

'==============================================================================
' SUBROUTINE
'   SetMonthOptionButtons
'------------------------------------------------------------------------------
' DESCRIPTION
'   Sets the captions on the Month Radio buttons based on the start month
' supplied
'==============================================================================
Public Sub SetMonthOptionButtons(dtFirstMonth As Date)

    Dim i As Long
    
    For i = 1 To 12
    
        m_aradMonths(i).Caption = Format(DateSerial(Year(dtFirstMonth), Month(dtFirstMonth) + i - 1, 1), "mmm-yy")
        
        '============
        ' Clear the radiobutton
        '============
        m_aradMonths(i).Value = False
    Next
End Sub

'==============================================================================
' SUBROUTINE
'   ClearFinYearButtons
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clears all the financial year radio buttons
'==============================================================================
Public Sub ClearFinYearButtons()
    Dim i As Long
    For i = 1 To 4
        m_aradFinancialYears(i).Value = False
    Next
End Sub

'==============================================================================
' SUBROUTINE
'   ClearCalYearButtons
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clears all the financial year radio buttons
'==============================================================================
Public Sub ClearCalYearButtons()
    Dim i As Long
    For i = 1 To 4
        m_aradCalendarYears(i).Value = False
    Next
End Sub

'==============================================================================
' SUBROUTINE
'   ClearMonthButtons
'------------------------------------------------------------------------------
' DESCRIPTION
'   Clears all the month's radio buttons
'==============================================================================
Public Sub ClearMonthButtons()
    Dim i As Long
    For i = 1 To 12
        m_aradMonths(i).Value = False
    Next
End Sub





