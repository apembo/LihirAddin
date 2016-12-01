Attribute VB_Name = "modImportExportVBA_2"
Option Explicit

'==============================================================================
' MODULE
'   modImportExportVBA
'------------------------------------------------------------------------------
' DESCRIPTION
'   Provides functionality for importing and exporting VBA modules.
' Written by Ron DeBruin:
' http://www.rondebruin.nl/win/s9/win002.htm
'------------------------------------------------------------------------------
' VERSION
'   1.01 | 22-Sep-2016   | Adam Pemberton
'   - Made late binding. Now does not require the reference.
'..............................................................................
'   1.00 | 26-Jul-2016   | Adam Pemberton
'   - Works
'==============================================================================

'------------------------------------------------------------------------------
' MEMBERS
'   SUB ExportModulesUL
'       Exports all code modules in a workbook
'
'   SUB ImportModulesUL
'       Deletes any existing code modules, then imports all modules from a
'       directory.
'
'   FUNCTION FolderWithVBAProjectFilesUL
'       Supporting function that provides the import/export folder.
'
'   SUB DeleteVBAModulesAndUserForms
'       Supporting method to delete all code modules in a workbook.
'==============================================================================

'==============================================================================
' ENUM
'   VBAComponentType
'------------------------------------------------------------------------------
' DESCRIPTION
'   The type values for the VBIDE.VBComponent.Type field.
'
' Microsoft documentation currently says these are the valid values:
'
'   vbext_ct_StdModule          1   Standard Module
'   vbext_ct_ClassModule        2   Class Module
'   vbext_ct_MSForm             3   Microsoft Form
'   vbext_ct_ActiveXDesigner    11  ActiveX Designer
'   vbext_ct_Document           100 Document Module
'
' https://msdn.microsoft.com/en-us/library/office/gg264162.aspx
'==============================================================================
Public Enum VBAComponentType
    ctStdModule = 1
    ctClassModule = 2
    ctMSForm = 3
    ctActiveXDesigner = 11
    ctDocument = 100
End Enum

'==============================================================================
' SUBROUTINE
'   ExportModulesUL
'------------------------------------------------------------------------------
' DESCRIPTION
'   The code modules will be exported in a folder named VBAProjectFiles in
' the Documents folder.
' The code below create this folder if it not exist or delete all files in the
' folder if it does exist.
'==============================================================================
Public Sub ExportModulesUL()

#If DevelopMode = 1 Then
    Dim cmpComponent As VBIDE.VBComponent
#Else
    Dim cmpComponent As Object
#End If

    Dim bExport As Boolean
    Dim wsSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String

    If FolderWithVBAProjectFilesUL = "Error" Then
        MsgBox "Export Folder does not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFilesUL & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wsSource = Application.Workbooks(szSourceWorkbook)
    
    If wsSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFilesUL & "\"
    
    For Each cmpComponent In wsSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case VBAComponentType.ctClassModule
                szFileName = szFileName & ".cls"
                
            Case VBAComponentType.ctMSForm
                szFileName = szFileName & ".frm"
                
            Case VBAComponentType.ctStdModule
                szFileName = szFileName & ".bas"
                
            Case VBAComponentType.ctDocument, VBAComponentType.ctActiveXDesigner
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wsSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub


'==============================================================================
' SUBROUTINE
'   ImportModulesUL
'------------------------------------------------------------------------------
' DESCRIPTION
'   Deletes all existing code modules, and then imports all the modules from
' the specified directory.
'==============================================================================
Public Sub ImportModulesUL()

#If DevelopMode = 1 Then
    Dim cmpComponents As VBIDE.VBComponents
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
#Else
    Dim cmpComponents As Object
    Dim objFSO As Object
    Dim objFile As Object
#End If

    Dim wbTarget As Excel.Workbook
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFilesUL = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFilesUL & "\"
        
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            Call cmpComponents.Import(objFile.Path)
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

'==============================================================================
' FUNCTION
'   FolderWithVBAProjectFilesUL
'------------------------------------------------------------------------------
' DESCRIPTION
'   Supporting method that supplies the relevant import/export folder.
'==============================================================================
Function FolderWithVBAProjectFilesUL() As String
    
#If DevelopMode = 1 Then
    Dim fso As Scripting.FileSystemObject
#Else
    Dim fso As Object
#End If
    
    Dim WshShell As Object
    Dim scBasePath As String
    Dim scVerPath As String
    Dim SpecialPath As String
    
    scBasePath = "C:\y\Dropbox\Code\Excel\09 Develop\LGO Addin\B-Client\vba\"
    scVerPath = wsVer.Range("Version")
    SpecialPath = scBasePath & scVerPath

    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    If fso.FolderExists(SpecialPath) = False Then
        On Error Resume Next
        MkDir SpecialPath
        On Error GoTo 0
    End If

    If fso.FolderExists(SpecialPath) = True Then
        FolderWithVBAProjectFilesUL = SpecialPath
    Else
        FolderWithVBAProjectFilesUL = "Error"
    End If
    
End Function

'==============================================================================
' FUNCTION
'   DeleteVBAModulesAndUserForms
'------------------------------------------------------------------------------
' DESCRIPTION
'   Supporting method that deletes all existing modules in the workbook.
'==============================================================================
Function DeleteVBAModulesAndUserForms()
    
#If DevelopMode = 1 Then
    Dim VBProj As VBIDE.VBProject
    Dim oVBComp As VBIDE.VBComponent
#Else
    Dim VBProj As Object
    Dim oVBComp As Object
#End If
    
    Set VBProj = ActiveWorkbook.VBProject
    
    For Each oVBComp In VBProj.VBComponents
        'If oVBComp.Type = vbext_ct_Document Then
        If oVBComp.Type = VBAComponentType.ctDocument Then
            'Thisworkbook or worksheet module
            'We do nothing
        Else
            VBProj.VBComponents.Remove oVBComp
        End If
    Next oVBComp
End Function

'==============================================================================
' FUNCTION
'   ExportVBACodeToExcel
'------------------------------------------------------------------------------
' DESCRIPTION
'   Exports code to a new multi-tab workbook, with the code on one tab call
' "AllCode" and a second tab created for all variable definitions.
'==============================================================================
Public Sub ExportVBACodeToExcel()

#If DevelopMode = 1 Then
    Dim oVBAComponent As VBIDE.VBComponent
#Else
    Dim oVBAComponent As Object
#End If

    Dim wbProj As Excel.Workbook
    Dim i As Long
    Dim scModuleName As String
    Dim wbDest As Excel.Workbook
    Dim wsCode As Excel.Worksheet
    Dim wsDefinitions As Excel.Worksheet
    Dim iRowCodeNext As Long
    Dim iRowDefNext As Long
    Dim scLine As String
    
    
    Set wbProj = ActiveWorkbook
    
    Set wbDest = Application.Workbooks.Add()
    If wbDest.Worksheets.Count < 2 Then
        wbDest.Worksheets.Add
    End If
    
    Set wsCode = wbDest.Worksheets(1)
    Set wsDefinitions = wbDest.Worksheets(2)
    
    wsCode.Name = "AllCode"
    wsDefinitions.Name = "Definitions"
    
    '=============
    ' A little bit of formatting
    '=============
    wsCode.Cells(1, 1) = "Module"
    wsCode.Cells(1, 2) = "Line #"
    wsCode.Cells(1, 3) = "Code"
    
    wsCode.Range("A1:C1").Font.Bold = True
    wsCode.Range("A1:C1").Font.Italic = True
    
    wsDefinitions.Cells(1, 1) = "Module"
    wsDefinitions.Cells(1, 2) = "Line #"
    wsDefinitions.Cells(1, 3) = "Code Line"
    wsDefinitions.Cells(1, 4) = "Type"
    
    wsDefinitions.Range("A1:D1").Font.Bold = True
    wsDefinitions.Range("A1:D1").Font.Italic = True
    
    iRowCodeNext = 2
    iRowDefNext = 2
    
    For Each oVBAComponent In wbProj.VBProject.VBComponents
        scModuleName = oVBAComponent.Name
        wsCode.Cells(iRowCodeNext, 1) = scModuleName
        iRowCodeNext = iRowCodeNext + 1
        For i = 1 To oVBAComponent.CodeModule.CountOfLines
            
            scLine = oVBAComponent.CodeModule.Lines(i, 1)
            
            wsCode.Cells(iRowCodeNext, 2) = i
            wsCode.Cells(iRowCodeNext, 3) = scLine
            iRowCodeNext = iRowCodeNext + 1
            
            If (InStr(1, scLine, " As ") > 0) Or (InStr(1, scLine, "Dim ") > 0) Then
                wsDefinitions.Cells(iRowDefNext, 1) = scModuleName
                wsDefinitions.Cells(iRowDefNext, 2) = i
                wsDefinitions.Cells(iRowDefNext, 3) = scLine
                
                Dim iIndex As Long
                Dim iEnd1, iEnd2, iEnd3, iEnd As Long
                Dim scType As String
                
                iIndex = InStr(1, scLine, " As ")
                While iIndex > 0
                    iEnd1 = InStr(iIndex + 4, scLine, ",")
                    If iEnd1 < 1 Then iEnd1 = 10000
                    
                    iEnd2 = InStr(iIndex + 4, scLine, ")")
                    If iEnd2 < 1 Then iEnd2 = 10000
                    
                    iEnd3 = InStr(iIndex + 4, scLine, " ")
                    If iEnd3 < 1 Then iEnd3 = 10000
                    
                    iEnd = Min(iEnd1, iEnd2, iEnd3)
                    
                    scType = Mid(scLine, iIndex + 4, iEnd - iIndex - 4)
                    wsDefinitions.Cells(iRowDefNext, 4) = scType
                    
                    iIndex = InStr(iEnd, scLine, " As ")
                    If iIndex > 0 Then
                        iRowDefNext = iRowDefNext + 1
                        wsDefinitions.Cells(iRowDefNext, 1) = scModuleName
                        wsDefinitions.Cells(iRowDefNext, 2) = i
                        wsDefinitions.Cells(iRowDefNext, 3) = scLine
                    End If
                Wend
                iRowDefNext = iRowDefNext + 1
                
            End If
            
        Next
next_module:
    Next
    
    '=============
    ' A little bit of formatting
    '=============
    wsCode.Columns("A:C").EntireColumn.AutoFit
    wsDefinitions.Columns("A:D").EntireColumn.AutoFit

End Sub

