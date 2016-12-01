Attribute VB_Name = "modErrorHandler"
Option Explicit

'==============================================================================
' USER DEFINED ENUMERATION
'   eAppErrorCodes
'------------------------------------------------------------------------------
' DESCRIPTION
'   Application specific custom error codes.
' The HRESULT standard says that if you set the 3rd highest order bit to 1, its
' a userdefined error.
'==============================================================================
Enum eAppErrorCodes
    aeNoError = &HA0000000
    aeWorkOrderNotFound = &HA0000001
    aeNotificationNotfound = &HA0000002
End Enum

'==============================================================================
' USER DEFINED TYPE
'   ACPError
'------------------------------------------------------------------------------
' DESCRIPTION
'   Simple error structure
'==============================================================================
Type typAppError
    number As Long
    description As String
    Source As String
    BeSilent As Boolean
End Type

'==============================================================================
' SUBROUTINE
'   DisplayError
'------------------------------------------------------------------------------
' DESCRIPTION
'   Generic error reporting
'==============================================================================
Public Sub DisplayError(oErr As typAppError, Optional scUserMessage As Variant)
    If oErr.BeSilent Then
        Exit Sub
    End If
    
    If (IsMissing(scUserMessage)) Then
        Call MsgBox("Error " & oErr.number & ": '" & oErr.description & "' (source:" & oErr.Source & ")")
    Else
        Call MsgBox("Error " & oErr.number & ": '" & scUserMessage & "'")
    End If
End Sub


