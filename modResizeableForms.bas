Attribute VB_Name = "modResizeableForms"
Option Explicit

'Written: February 14, 2011
'Author:  Leith Ross
'
'NOTE:  This code should be executed within the UserForm_Activate() event.

Private Declare Function GetForegroundWindow Lib "User32.dll" () As Long

Private Declare Function GetWindowLong _
  Lib "User32.dll" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long) _
  As Long
               
Private Declare Function SetWindowLong _
  Lib "User32.dll" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) _
  As Long

Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

'==============================================================================
' SUBROUTINE
'   MakeFormResizable
'------------------------------------------------------------------------------
' DESCRIPTION
'   Use to make a Userform resizeable. This method should be executed within
' the UserForm_Activate() event.
'==============================================================================
Public Sub MakeFormResizable()

  Dim lStyle As Long
  Dim hWnd As Long
  Dim RetVal
  
    hWnd = GetForegroundWindow
  
    'Get the basic window style
     lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME

    'Set the basic window styles
     RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)

End Sub

