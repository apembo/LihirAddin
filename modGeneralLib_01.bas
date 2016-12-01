Attribute VB_Name = "modGeneralLib_01"
'==============================================================================
' MODULE
'   modGeneralLib
'------------------------------------------------------------------------------
' DESCRIPTION
'   Useful General methods
'------------------------------------------------------------------------------
' VERSION
'..............................................................................
'   1.0 | First release
'==============================================================================
Option Explicit

Private Const mcGWL_STYLE = (-16)
Private Const mcWS_SYSMENU = &H80000

'Windows API calls to handle windows
#If VBA7 Then
    Private Declare PtrSafe Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Const NoError = 0       'The Function call was successful


'==============================================================================
' Special folders enumeration associated with the GetSpecialFolderPath
' method.
'==============================================================================
Public Enum eSpecialFolders
    sfWindows = 0
    sfSystem = 1
    sfTemp = 2
    sfAllUsersDesktop
    sfAllUsersStartMenu
    sfAllUsersPrograms
    sfAllUsersStartup
    sfDesktop
    sfFavorites
    sfFonts
    sfMyDocuments
    sfNetHood
    sfPrintHood
    sfPrograms
    sfRecent
    sfSendTo
    sfStartMenu
    sfStartup
    sfTemplates
End Enum

'==============================================================================
' SUBROUTINE
'   GetRGBFromAge
'------------------------------------------------------------------------------
' DESCRIPTION
'   Function to return a colour based on an age, from green below a certain
' threhold (dAgeGood) through a transition, to red above a higher threshold
' (sAgeBad).
'   Returns a colour as follows:
' If Age <= AgeMaxGood                  -> Green
' if AgeMaxGood < Age < dAgeBorderline  -> A colour changing from Green to Orange
' if dAgeBorderline < Age < dAgeBad     -> A colour changing from Orange to Red
' If Age > dAgeBad                      -> Red
'==============================================================================
Public Function GetRGBFromAge(dAge As Double, dAgeGood As Double, dAgeBorderline As Double, dAgeBad As Double) As Long
    Dim dM As Double, dB As Double, dY As Double
    
    If dAge < dAgeGood Then
        GetRGBFromAge = RGB(0, 255, 0)
    ElseIf dAge <= dAgeBorderline Then
        dM = 255# / (dAgeBorderline - dAgeGood)
        dB = -1 * dM * dAgeGood
        
        dY = dAge * dM + dB
        GetRGBFromAge = RGB(CLng(dY), 255, CLng(dY))
    ElseIf dAge <= dAgeBad Then
        dM = -255# / (dAgeBorderline - dAgeGood)
        dB = -1 * dM * dAgeBad
        
        dY = dAge * dM + dB
        GetRGBFromAge = RGB(255, CLng(dY), CLng(dY))
    Else
        GetRGBFromAge = RGB(255, 0, 0)
    End If
    
End Function

'==============================================================================
' SUBROUTINE
'   GetRGBFromRange
'------------------------------------------------------------------------------
' DESCRIPTION
'   Function to return a colour based on a value in a range, from green below a certain
' threhold (dAgeGood) through a transition, to red above a higher threshold
' (sAgeBad).
'   Returns a colour as follows:
' If Age <= AgeMaxGood                  -> Green
' if AgeMaxGood < Age < dAgeBorderline  -> A colour changing from Green to Orange
' if dAgeBorderline < Age < dAgeBad     -> A colour changing from Orange to Red
' If Age > dAgeBad                      -> Red
'==============================================================================
Public Function GetRGBFromRange(dInput As Double, dMinimum As Double, dMiddle As Double, dMaximum As Double, Optional bSmallIsGood As Boolean = True) As Long
    
    Dim dM As Double, dB As Double, dY As Double
    
    If dInput < dMinimum Then
        GetRGBFromRange = RGB(0, 255, 0)
    ElseIf dInput <= dMiddle Then
        dM = 255# / (dMiddle - dMinimum)
        dB = -1 * dM * dMinimum
        
        dY = dInput * dM + dB
        GetRGBFromRange = RGB(CLng(dY), 255, CLng(dY))
    ElseIf dInput <= dMaximum Then
        dM = -255# / (dMiddle - dMinimum)
        dB = -1 * dM * dMaximum
        
        dY = dInput * dM + dB
        GetRGBFromRange = RGB(255, CLng(dY), CLng(dY))
    Else
        GetRGBFromRange = RGB(255, 0, 0)
    End If
    
End Function

'==============================================================================
' SUBROUTINE
'   GetUserName
'------------------------------------------------------------------------------
' DESCRIPTION
'   Get's the current users name
'==============================================================================
Public Function GetUserName(ByRef scUsername As String) As Boolean

    ' Buffer size for the return string.
    Const lpnLength As Integer = 255

    ' Get return buffer space.
    Dim status As Integer

    ' For getting user information.
    Dim lpName, lpUserName As String

    ' Assign the buffer size constant to lpUserName.
    lpUserName = Space$(lpnLength + 1)

    ' Get the log-on name of the person using product.
    status = WNetGetUser(lpName, lpUserName, lpnLength)

    ' See whether error occurred.
    If status = NoError Then
        ' This line removes the null character. Strings in C are null-
        ' terminated. Strings in Visual Basic are not null-terminated.
        ' The null character must be removed from the C strings to be used
        ' cleanly in Visual Basic.
        lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
        scUsername = lpUserName
        
    ElseIf Application.UserName <> "" Then
    
        scUsername = Application.UserName
        
    Else
       GetUserName = False
       Exit Function
       
    End If

    GetUserName = True

End Function

Sub ListLinks()
    'Updateby20140529
    Dim wb As Excel.Workbook
    Dim xIndex As Long
    Dim link
    
    Set wb = Application.ActiveWorkbook
    If Not IsEmpty(wb.LinkSources(xlExcelLinks)) Then
        wb.Sheets.Add
        xIndex = 1
        For Each link In wb.LinkSources(xlExcelLinks)
            Application.ActiveSheet.Cells(xIndex, 1).Value = link
            xIndex = xIndex + 1
        Next link
    End If
End Sub

'==============================================================================
' SUBROUTINE
'   MoveListboxItem
'------------------------------------------------------------------------------
' DESCRIPTION
'   Moves the item at From to the To index location.
'==============================================================================
Public Sub MoveListboxItem(lb As MSForms.ListBox, iFromIndex As Long, iToIndex As Long)
     '
     ' Swap listbox items
     '
    ReDim scaSubItems(lb.ColumnCount - 1) As String
    Dim iCol As Long
     
    For iCol = 0 To lb.ColumnCount - 1
        If IsNull(lb.List(iFromIndex, iCol)) Then
            scaSubItems(iCol) = ""
        Else
            scaSubItems(iCol) = lb.List(iFromIndex, iCol)
        End If
    Next
     
    Call lb.RemoveItem(iFromIndex)
    
    Call lb.AddItem(scaSubItems(0), iToIndex)
     
    For iCol = 1 To lb.ColumnCount - 1
        lb.List(iToIndex, iCol) = scaSubItems(iCol)
    Next
     
End Sub

'==============================================================================
' FUNCTION
'   GetFilenameFromPath
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'------------------------------------------------------------------------------
' Source
'
'==============================================================================
Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

'==============================================================================
' FUNCTION
'   SplitPath
'------------------------------------------------------------------------------
' DESCRIPTION
'   Splits the path into key subcomponents.
'------------------------------------------------------------------------------
' Source
'
'==============================================================================
Function SplitPath(scFullPath As String, ByRef scPath As String, ByRef scFilename As String, ByRef scFileNameNoExt As String, ByRef scExt As String) As Boolean
    
    'Dim scFilename As String
    
    '===========
    ' Functions ending in $ differs from there non-$ equivalents in that they
    ' operates directly on strings rather than variants which means they
    ' should be significantly faster.
    '===========
    scFilename = Right$(scFullPath, Len(scFullPath) - InStrRev(scFullPath, "\"))
    scExt = Right$(scFilename, InStrRev(scFullPath, ".") - 1)
    scFileNameNoExt = Left$(scFilename, Len(scFilename) - Len(scExt) - 1)
    
    scPath = Left$(scFullPath, InStrRev(scFullPath, "\"))
    
End Function

'==============================================================================
' FUNCTION
'   FileNameFromPath
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'------------------------------------------------------------------------------
' Source
' http://vba-tutorial.com/parsing-a-file-string-into-path-filename-and-extension/
'==============================================================================
Function FileNameFromPath(strFullPath As String) As String
 
     FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
 
End Function

'==============================================================================
' FUNCTION
'   FileNameNoExtensionFromPath
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'------------------------------------------------------------------------------
' Source
' http://vba-tutorial.com/parsing-a-file-string-into-path-filename-and-extension/
'==============================================================================
Function FileNameNoExtensionFromPath(strFullPath As String) As String
 
    Dim intStartLoc As Integer
    Dim intEndLoc As Integer
    Dim intLength As Integer
 
    intStartLoc = Len(strFullPath) - (Len(strFullPath) - InStrRev(strFullPath, "\") - 1)
    intEndLoc = Len(strFullPath) - (Len(strFullPath) - InStrRev(strFullPath, "."))
    intLength = intEndLoc - intStartLoc
 
    FileNameNoExtensionFromPath = Mid(strFullPath, intStartLoc, intLength)
 
End Function

'==============================================================================
' FUNCTION
'   FolderFromPath
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'------------------------------------------------------------------------------
' Source
' http://vba-tutorial.com/parsing-a-file-string-into-path-filename-and-extension/
'==============================================================================
Function FolderFromPath(ByRef strFullPath As String) As String
 
     FolderFromPath = Left(strFullPath, InStrRev(strFullPath, "\"))
 
End Function

'==============================================================================
' FUNCTION
'   FileExtensionFromPath
'------------------------------------------------------------------------------
' DESCRIPTION
'   Self explanatory
'------------------------------------------------------------------------------
' Source
' http://vba-tutorial.com/parsing-a-file-string-into-path-filename-and-extension/
'==============================================================================
Function FileExtensionFromPath(ByRef strFullPath As String) As String
 
     FileExtensionFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "."))
 
End Function

'==============================================================================
' FUNCTION
'   RemoveUserformCloseButton
'------------------------------------------------------------------------------
' DESCRIPTION
'   Removes the X close button on the top write corner
'------------------------------------------------------------------------------
' Source
' http://stackoverflow.com/questions/15153491/hide-close-x-button-on-excel-vba-userform-for-my-progress-bar
'==============================================================================
Public Sub RemoveUserformCloseButton(frm As Object)
    Dim lngStyle As Long
    Dim lngHWnd As Long

    lngHWnd = FindWindow(vbNullString, frm.Caption)
    lngStyle = GetWindowLong(lngHWnd, mcGWL_STYLE)

    If lngStyle And mcWS_SYSMENU > 0 Then
        Call SetWindowLong(hWnd:=lngHWnd, nIndex:=mcGWL_STYLE, dwNewLong:=(lngStyle And Not mcWS_SYSMENU))
    End If

End Sub

'==============================================================================
' FUNCTION
'   GetSpecialFolderPath
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the path to the special folder.
' Note:
'==============================================================================
Public Function GetSpecialFolderPath(enumFolder As eSpecialFolders) As String

    Dim scSpecialFolder As String
    Dim fld As Object ' Folder
 
    On Error GoTo ErrorHandler
    
    If enumFolder <= 3 Then
        Dim fso As Object, TmpFolder As Object
        
        Set fso = CreateObject("scripting.filesystemobject")
        '=============
        ' Use the GetSpecialFolder method of the fso object
        ' 0 = The Windows folder contains files installed by the Windows operating sys
        ' 1 = The System folder contains libraries, fonts, and device drivers
        ' 2 = The Temp folder is used to store temporary files. Its path is found in the TMP environment variable.
        Set fld = fso.GetSpecialFolder(enumFolder)
        
        GetSpecialFolderPath = fld.Path
        Exit Function
    End If
        
 
    Select Case enumFolder
        Case sfAllUsersDesktop
            scSpecialFolder = "AllUsersDesktop"
        Case sfAllUsersStartMenu
            scSpecialFolder = "AllUsersStartMenu"
        Case sfAllUsersPrograms
            scSpecialFolder = "AllUsersPrograms"
        Case sfAllUsersStartup
            scSpecialFolder = "AllUsersStartup"
        Case sfDesktop
            scSpecialFolder = "Desktop"
        Case sfFavorites
            scSpecialFolder = "Favorites"
        Case sfFonts
            scSpecialFolder = "Fonts"
        Case sfMyDocuments
            scSpecialFolder = "MyDocuments"
        Case sfNetHood
            scSpecialFolder = "NetHood"
        Case sfPrintHood
            scSpecialFolder = "PrintHood"
        Case sfPrograms
            scSpecialFolder = "Programs"
        Case sfRecent
            scSpecialFolder = "Recent"
        Case sfSendTo
            scSpecialFolder = "SendTo"
        Case sfStartMenu
            scSpecialFolder = "StartMenu"
        Case sfStartup
            scSpecialFolder = "Startup"
        Case sfTemplates
            scSpecialFolder = "Templates"
        
        Case Else
            GetSpecialFolderPath = ""
            Exit Function
    End Select
 
   '==========
   'Create a Windows Script Host Object
   '==========
   Dim objWSHShell As Object
   Set objWSHShell = CreateObject("WScript.Shell")
   
   '==========
   'Retrieve path
   '==========
   GetSpecialFolderPath = objWSHShell.SpecialFolders(scSpecialFolder)

   '==========
   ' Clean up
   '==========
   Set objWSHShell = Nothing
   Exit Function
 
ErrorHandler:
    GetSpecialFolderPath = ""
End Function

