Attribute VB_Name = "modCustomTypesAndEnums"
Option Explicit

'==============================================================================
' ENUMERATION
'   eDateRangeTypes
'------------------------------------------------------------------------------
' DESCRIPTION
'   Used in the DateRangeSelection user form
'==============================================================================
Public Enum eDateRangeTypes
    drtFinancialYear
    drtCalendarYear
    drtMonth
End Enum

'==============================================================================
' ENUMERATION
'   eDateTimeFormat
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'==============================================================================
Public Enum eDateTimeFormat
    DateShort = 1
    DateLong = 2
    DateShortTime24Hr = 3
    DateShortTimeAmPm = 4
    DateLongTime24Hr = 5
    DateLongTimeAmPm = 6
    DateTSQL = 7
    DateTimeTSQL = 8
End Enum

'=============
' Dirty Flags
'=============
Public Enum eDirtyFlags
    OrderChanged = 1
    PicturesAddedOrDeleted = 2
    PictureDetailChanged = 3
    RatingChanged = 4
    PictureDetailChangedNotInList = 5
End Enum


'==============================================================================
' ENUMERATION
'   TreeViewNodeType
'------------------------------------------------------------------------------
' DESCRIPTION
'   Defined to allow the treeview control not to be referenced.
' The documentation says the following constants are defined
'   Constant    Value   Description
'   tvwLast     1   The Node is placed after all other nodes at the same level of the node named in relative.
'   tvwNext     2   The Node is placed after the node named in relative.
'   tvwPrevious 3   The Node is placed before the node named in relative.
'   tvwChild    4   The Node becomes a child node of the node named in relative.
'==============================================================================
Public Enum TreeViewNodeType
    tvwLast_ = 1
    tvwNext_ = 2
    tvwPrevious_ = 3
    tvwChild_ = 4
End Enum

'==============================================================================
' USER DEFINED TYPE
'   DBTableField
'------------------------------------------------------------------------------
' DESCRIPTION
'   Describes a parameter in a DB table
'==============================================================================
Public Type DBTableField
    FieldName As String
    FieldType As String
    FieldTypeSize As Long
End Type

'==============================================================================
' FUNCTION
'   DateRangeTypeString
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns a human friendly string representation of the different values of
' the eDateRangeTypes enumeration.
'==============================================================================
Public Function DateRangeTypeString(eType As eDateRangeTypes) As String
    Select Case eType
        Case eDateRangeTypes.drtCalendarYear
            DateRangeTypeString = "Calendar"
        Case eDateRangeTypes.drtFinancialYear
            DateRangeTypeString = "Financial"
        Case eDateRangeTypes.drtMonth
            DateRangeTypeString = "Month"
        Case Else
            DateRangeTypeString = "<Unknown>"
    End Select
End Function


