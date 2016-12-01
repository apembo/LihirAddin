VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufGenFinanceList 
   Caption         =   "Get Item Listing"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6195
   OleObjectBlob   =   "ufGenFinanceList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufGenFinanceList"
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

Private m_eResult As VBA.VbMsgBoxResult
Private m_scItemRequired As String

Public frmSelectDateRange As ufDateRangeSelection

Public Property Get Result()
    Result = m_eResult
End Property

Public Property Get ItemRequired() As String
    ItemRequired = m_scItemRequired
End Property

Private Sub butOK_Click()
    m_eResult = vbOK
    
    '=============
    ' Specified the users requested list type
    '=============
    If (radioPurchDocs.Value) Then
        m_scItemRequired = "PurchaseDocs"
    ElseIf (RadioMaterials.Value) Then
        m_scItemRequired = "Materials"
    Else
        m_scItemRequired = "PartnerObjects"
    End If
    
    '=============
    ' Store the user selection for next time.
    '=============
    wsParameters.Range("GenLedgerListingCCFilter") = Me.txtCCFilter.Text
    wsParameters.Range("GenLedgerListingListItem") = m_scItemRequired
        
    Call Hide
End Sub

Private Sub butCancel_Click()
    m_eResult = vbCancel
    m_scItemRequired = ""
    Call Hide
End Sub

Private Sub butGetDateRange_Click()
    Dim oErr As typAppError

    If Not frmSelectDateRange Is Nothing Then
        Call Unload(frmSelectDateRange)
    End If
    
    Set frmSelectDateRange = New ufDateRangeSelection
    
    oErr.BeSilent = False

    Call frmSelectDateRange.Show
    
    If frmSelectDateRange.Result = vbOK Then
        Me.txtDateFrom = Format(frmSelectDateRange.FirstDate, "d-mmm-yy")
        Me.txtDateTo = Format(frmSelectDateRange.LastDate, "d-mmm-yy")
    End If
    
    Call Unload(frmSelectDateRange)
        
End Sub

Private Sub UserForm_Initialize()

    '============
    ' Default the date range to the last full month
    '============
    Me.txtDateFrom = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "d-mmm-yy")
    Me.txtDateTo = Format(DateSerial(Year(Date), Month(Date), 1) - 1, "d-mmm-yy")
    
    Me.txtCCFilter.Text = wsParameters.Range("GenLedgerListingCCFilter")
    
    Select Case wsParameters.Range("GenLedgerListingListItem")
        Case "PartnerObjects"
            radioPartnerObjects.Value = True
        Case "PurchaseDocs"
            Me.radioPurchDocs.Value = True
        Case "Materials"
            Me.RadioMaterials.Value = True
        Case Else
            radioPartnerObjects.Value = True
    End Select
    
    Call RemoveUserformCloseButton(Me)
            
End Sub
