Attribute VB_Name = "modExcel_02"
Option Explicit

'==============================================================================
' VERSION
'   1.2
'------------------------------------------------------------------------------
' HISTORY
' 1.2 - Adam Pemberton
'   - Fixed a bug with the GetUsedRange function where it was returning a 0 as
' the minimum row.
'------------------
' 1.1 - Adam Pemberton
'    - Added GetColumnLetterString function
'    - Added GetA1AddressString function
'------------------
' 1.0 - Adam Pemberton - First Release - Used Range type functions.
'==============================================================================


'==============================================================================
' SUBROUTINE
'   GetUsedRange
'------------------------------------------------------------------------------
' DESCRIPTION
'   This function returns the true used range. Used is defined as has having a
' non-blank value. A formula that returns a blank results is considered an
' unused cell.
' If the sheet is completely empty, it sets the bEmpty flag and does not alter
' the 4 row/column parameters.
' Note:
' The traditional Worksheet.UsedRange function will include areas that have
' been used but have since had their content deleted.
'==============================================================================
Sub GetUsedRange(ByRef ws As Excel.Worksheet, _
                ByRef iRowFirst As Long, ByRef iRowLast As Long, ByRef iColFirst As Long, ByRef iColLast As Long, _
                ByRef bEmpty As Boolean)
                
    Dim bColUsed As Boolean
    Dim iRow As Long
    Dim iCol As Long
    Dim iR1 As Long, iR2 As Long, iC1 As Long, iC2 As Long
    Dim iColFirstActuallyFound As Long
    Dim iColLastActuallyFound As Long
    Dim iLastRowFound As Long
    
    '===========
    ' Start with the UsedRange function. It often over-reports which is why I
    ' have written this function.
    '===========
    'iRowFirst = ws.UsedRange.Row
    'iRowLast = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    'iColFirst = ws.UsedRange.Column
    'iColLast = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    iR1 = ws.UsedRange.Row
    iR2 = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    iC1 = ws.UsedRange.Column
    iC2 = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    
    '===========
    ' Is the sheet empty
    '===========
    If (iR1 = 1) And (iR2 = 1) And (iC1 = 1) And (iC2 = 1) Then
        bEmpty = True
        Exit Sub
    End If
    
    '===========
    ' Last Row. For each column, start at the last row and move to the first
    ' non
    '===========
    Dim iRowMax As Long
    
    iRowMax = 0
    For iCol = iC1 To iC2
        
        iLastRowFound = LastCellInColumn(ws, iCol)
        If (iLastRowFound = 1) And (ws.Cells(1, iCol) = "") Then
            iLastRowFound = 0
            If (iCol = iC1) Then
                iC1 = iCol + 1
            End If
        End If
        
        If (iLastRowFound > iRowMax) Then
            iRowMax = iLastRowFound
        End If
    Next
    
    If (iRowMax = 0) Then
        bEmpty = True
        Exit Sub
    End If
    
    iR2 = iRowMax
        
    '===========
    ' First Row
    '===========
    Dim iRowMin As Long
    
    iRowMin = iRowMax
    For iCol = iC1 To iC2
        Dim iFirstRowFound As Long
        
        iFirstRowFound = FirstCellInColumn(ws, iCol)
        If (iFirstRowFound = ws.Rows.Count) And (ws.Cells(ws.Rows.Count, iCol) = "") Then
            ' Then we didn't find anything in this column
            iFirstRowFound = iR1 + 1
        End If
        
        If (iFirstRowFound < iRowMin) Then
            iRowMin = iFirstRowFound
        End If
        If (iRowMin = 1) Then Exit For
    Next
    iR1 = iRowMin

    '===========
    ' Last Column
    '===========
    Dim bColEmpty As Boolean
    
    bColEmpty = True
    While bColEmpty And (iC2 > iC1)
        
        iLastRowFound = LastCellInColumn(ws, iC2)
        If (iLastRowFound = 1) And (ws.Cells(1, iC2) = "") Then
            iC2 = iC2 - 1
        Else
            bColEmpty = False
        End If
    Wend
            
    bEmpty = False
    
    iRowFirst = iR1
    iRowLast = iR2
    iColFirst = iC1
    iColLast = iC2

End Sub
            
'==============================================================================
' SUBROUTINE
'   LastCellInColumn
'------------------------------------------------------------------------------
' DESCRIPTION
'   Find the very last used cell in a Column.
' The original code from the net was:
'   Range("A" & Rows.Count).End(xlup).Select
'==============================================================================
Function LastCellInColumn(ws As Excel.Worksheet, iCol As Long) As Long
    If (ws.Cells(ws.Rows.Count, iCol) <> "") Then
        LastCellInColumn = ws.Rows.Count
    Else
        LastCellInColumn = ws.Cells(ws.Rows.Count, iCol).End(xlUp).Row
    End If
End Function

'==============================================================================
' SUBROUTINE
'   FirstCellInColumn
'------------------------------------------------------------------------------
' DESCRIPTION
'   Find the first used cell in a Column.
'==============================================================================
Function FirstCellInColumn(ws As Excel.Worksheet, iCol As Long) As Long
    If (ws.Cells(1, iCol) <> "") Then
        FirstCellInColumn = 1
    Else
        FirstCellInColumn = ws.Cells(1, iCol).End(xlDown).Row
    End If
End Function

'==============================================================================
' SUBROUTINE
'   LastCellBeforeBlankInColumn
'------------------------------------------------------------------------------
' DESCRIPTION
'   Find the last used cell, before a blank in a Column:
'==============================================================================
Sub LastCellBeforeBlankInColumn()
    Range("A1").End(xlDown).Select
End Sub

'==============================================================================
' SUBROUTINE
'   LastCellBeforeBlankInRow
'------------------------------------------------------------------------------
' DESCRIPTION
'   Find the last cell, before a blank in a Row:
'==============================================================================
Sub LastCellBeforeBlankInRow()
    Range("A1").End(xlToRight).Select
End Sub

'==============================================================================
' SUBROUTINE
'   LastCellInRow
'------------------------------------------------------------------------------
' DESCRIPTION
'   Find the very last used cell in a Row:
'==============================================================================
Sub LastCellInRow()
    Range("IV1").End(xlToLeft).Select
End Sub

'==============================================================================
' FUNCTION
'   GetColumnLetterString
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the column letter string corresponding to the supplied column
' number. For example:
'   9 -> 'I'
'   27 -> 'AA'
'==============================================================================
Function GetColumnLetterString(iColNumber As Long) As String
    Dim scTemp As String
    
    scTemp = Worksheets(1).Cells(1, iColNumber).Address(False, False)
    scTemp = Left(scTemp, Len(scTemp) - 1)
    
    GetColumnLetterString = scTemp
End Function


'==============================================================================
' FUNCTION
'   GetA1AddressString
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the column letter string corresponding to the supplied column
' number. For example:
'   9 -> 'I'
'   27 -> 'AA'
'==============================================================================
Function GetA1AddressString(iRow As Long, iCol As Long, _
        Optional bRowAbsolute As Boolean = False, _
        Optional bColAbsolute As Boolean = False) As String
        
    Dim scTemp As String
    
    GetA1AddressString = Worksheets(1).Cells(iRow, iCol).Address(bRowAbsolute, bColAbsolute)

End Function

