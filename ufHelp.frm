VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufHelp 
   Caption         =   "Help"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
   OleObjectBlob   =   "ufHelp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufHelp"
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


'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   btnOK_Click
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub btnOK_Click()
    Call Hide
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   UserForm_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'
'==============================================================================
Private Sub UserForm_Initialize()
    Call WebBrowser1.Navigate("about:blank")
End Sub

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   RangetoHTML
'------------------------------------------------------------------------------
' DESCRIPTION
'   Brilliant Ron DeBruins method to convert a cell range to HTML.
' http://www.rondebruin.nl/win/s1/outlook/bmail1.htm
'==============================================================================
Public Sub SetHelpFromRange(rng As Excel.Range)
    '============
    ' Clear any existing page
    '============
    Call Me.WebBrowser1.Document.Close
    Call Me.WebBrowser1.Document.Open
    
    Call Me.WebBrowser1.Document.Write(RangetoHTML(rng))
End Sub


'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   RangetoHTML
'------------------------------------------------------------------------------
' DESCRIPTION
'   Brilliant Ron DeBruins method to convert a cell range to HTML.
' http://www.rondebruin.nl/win/s1/outlook/bmail1.htm
'==============================================================================
Function RangetoHTML(rng As Excel.Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2013
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Excel.Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        '===========
        ' PasteSpecial(Paste, Operation, SkipBlanks, Transpose)
        '===========
        .Cells(1).PasteSpecial Paste:=xlPasteColumnWidths
        .Cells(1).PasteSpecial Paste:=xlPasteValues ', , False, False
        .Cells(1).PasteSpecial Paste:=xlPasteFormats ', , False, False
        '.Cells(1).Paste
        .Cells(1).Select
'        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'        Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'        ActiveSheet.Pictures.Paste.Select
'        ActiveSheet.Shapes.Range(Array("Picture 1")).Select
        
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
