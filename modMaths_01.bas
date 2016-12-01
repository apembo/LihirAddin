Attribute VB_Name = "modMaths_01"
Option Explicit

'______________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' MODULE
'   modMaths
'______________________________________________________________________________
'
' DESCRIPTION
'   Generic math functions that VBA should provide ... but doesn't.
'______________________________________________________________________________
'
' VERSION
'
' Ver # | Date      | Description
'..............................................................................
' 1.1   | Jul 2015  |   Got rid of most of the Min Max functions.
'                       Added:
'                        - Modulus (remainder) function
'                        - Avg (average) function
'                        - Floor
'                        - Ceil
'..............................................................................
' 1.0   | Dec 2014  | Basically the Max Min functions.
'______________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'==============================================================================
' FUNCTION
'   Max
'------------------------------------------------------------------------------
' DESCRIPTION
'   Finds the maximum of 2 to 6 parameters.
'==============================================================================
Public Function Max(p1 As Variant, p2 As Variant, Optional p3 As Variant, Optional p4 As Variant, Optional p5 As Variant, Optional p6 As Variant) As Variant

    Dim dChosen As Double
    Dim dComparison As Double
    Dim pOut As Variant
    
    dChosen = p1
    pOut = p1 ' Preserve type
    
    '========
    ' p2
    '========
    dComparison = p2
    If (dComparison > dChosen) Then
        dChosen = dComparison
        pOut = p2
    End If
    
    '========
    ' p3
    '========
    If IsMissing(p3) Then
        Max = pOut
        Exit Function
    Else
        dComparison = p3
    End If
    
    If (dComparison > dChosen) Then
        dChosen = dComparison
        pOut = p3
    End If
    
    '========
    ' p4
    '========
    If IsMissing(p4) Then
        Max = pOut
        Exit Function
    Else
        dComparison = p4
    End If
    
    If (dComparison > dChosen) Then
        dChosen = dComparison
        pOut = p4
    End If
    
    '========
    ' p5
    '========
    If IsMissing(p5) Then
        Max = pOut
        Exit Function
    Else
        dComparison = p5
    End If
    
    If (dComparison > dChosen) Then
        dChosen = dComparison
        pOut = p5
    End If
    
    '========
    ' p6
    '========
    If IsMissing(p6) Then
        Max = pOut
        Exit Function
    Else
        dComparison = p6
    End If
    
    If (dComparison > dChosen) Then
        dChosen = dComparison
        pOut = p6
    End If
    
    Max = pOut
    
End Function

'==============================================================================
' FUNCTION
'   Min
'------------------------------------------------------------------------------
' DESCRIPTION
'   Finds the minimum of 2 to 6 parameters.
'==============================================================================
Public Function Min(p1 As Variant, p2 As Variant, Optional p3 As Variant, Optional p4 As Variant, Optional p5 As Variant, Optional p6 As Variant) As Variant

    Dim dChosen As Double
    Dim dComparison As Double
    Dim pOut As Variant
    
    dChosen = p1
    pOut = p1 ' Preserve type
    
    '========
    ' p2
    '========
    dComparison = p2
    If (dComparison < dChosen) Then
        dChosen = dComparison
        pOut = p2
    End If
    
    '========
    ' p3
    '========
    If IsMissing(p3) Then
        Min = pOut
        Exit Function
    Else
        dComparison = p3
    End If
    
    If (dComparison < dChosen) Then
        dChosen = dComparison
        pOut = p3
    End If
    
    '========
    ' p4
    '========
    If IsMissing(p4) Then
        Min = pOut
        Exit Function
    Else
        dComparison = p4
    End If
    
    If (dComparison < dChosen) Then
        dChosen = dComparison
        pOut = p4
    End If
    
    '========
    ' p5
    '========
    If IsMissing(p5) Then
        Min = pOut
        Exit Function
    Else
        dComparison = p5
    End If
    
    If (dComparison < dChosen) Then
        dChosen = dComparison
        pOut = p5
    End If
    
    '========
    ' p6
    '========
    If IsMissing(p6) Then
        Min = pOut
        Exit Function
    Else
        dComparison = p6
    End If
    
    If (dComparison < dChosen) Then
        dChosen = dComparison
        pOut = p6
    End If
    
    Min = pOut
    
End Function

'==============================================================================
' FUNCTION
'   Avg
'------------------------------------------------------------------------------
' DESCRIPTION
'   Finds the average of 2 to 6 parameters.
'==============================================================================
Public Function Avg(p1 As Variant, p2 As Variant, Optional p3 As Variant, Optional p4 As Variant, Optional p5 As Variant, Optional p6 As Variant) As Variant

    Dim dSum As Double
    Dim dComparison As Double
    Dim pOut As Variant
    
    dSum = p1 + p2
    
    '========
    ' p3
    '========
    If IsMissing(p3) Then
        Avg = dSum / 2#
        Exit Function
    Else
        dSum = dSum + p3
    End If
    
    '========
    ' p4
    '========
    If IsMissing(p4) Then
        Avg = dSum / 3#
        Exit Function
    Else
        dSum = dSum + p4
    End If
    
    '========
    ' p5
    '========
    If IsMissing(p5) Then
        Avg = dSum / 4#
        Exit Function
    Else
        dSum = dSum + p5
    End If
    
    '========
    ' p6
    '========
    If IsMissing(p6) Then
        Avg = dSum / 5#
        Exit Function
    Else
        dSum = dSum + p6
    End If
    
    Avg = dSum / 6#
    
End Function

'==============================================================================
' FUNCTION
'   Modulus
'------------------------------------------------------------------------------
' DESCRIPTION
'   The modulus, or remainder, operator divides number by divisor (rounding
' floating-point numbers to integers) and returns only the remainder as result.
' For example, in the following expression, A (result) equals 5.
'
'   Modulus(19,6.7) = 5
'
'   Usually, the data type of result is a Byte, Byte variant, Integer, Integer
' variant, Long, or Variant containing a Long, regardless of whether or not
' result is a whole number. Any fractional portion is truncated. However, if
' any expression is Null, result is Null. Any expression that is Empty is
' treated as 0.
'------------------------------------------------------------------------------
' REMARK
'   Obviously it is more efficient to use the Mod operator. This function is
' basically to remind one that it exists.
' The above description is straight out of documentation for the Mod operator.
'==============================================================================
Public Function Modulus(number As Variant, divisor As Variant) As Variant

    Modulus = number Mod divisor
    
End Function

'==============================================================================
' FUNCTION
'   Floor
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the whole number less than or equal to the passed in value.
'==============================================================================
Public Function Floor(dVal As Double) As Double
    Floor = CDbl(Fix(dVal))
End Function

'==============================================================================
' FUNCTION
'   Ceil
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the whole number greater than or equal to the passed in value.
'==============================================================================
Public Function Ceil(dVal As Double) As Double
    If dVal = CDbl(Fix(dVal)) Then
        Ceil = dVal
    Else
        Ceil = CDbl(Fix(dVal) + 1)
    End If
End Function
