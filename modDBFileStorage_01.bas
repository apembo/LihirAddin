Attribute VB_Name = "modDBFileStorage_01"
'==============================================================================
' MODULE
'   modDBStorage
'------------------------------------------------------------------------------
' DESCRIPTION
'   Handles copying the supplied file to the DB storage area, and updating
' the database table to reflect the new file.
'   The AddFile method tries to ensure multiple copies of the same file are not
' kept.
'------------------------------------------------------------------------------
' VERSION
'   1.1 | 03-Aug-2015   | Adam Pemberton
'   - Now captures and records original file name and size in the database.
'   - Makes sure to capture time as well as date for the file date/time
'   - Now does identical file checking and does not insert two identical files.
'       A file is treated as identical if it's, size, date and name are
'       identical.
'..............................................................................
'   1.0 | First release
'==============================================================================
Option Explicit

Public Type FileInfo
    FilePath As String
    FileName As String
    filedate As Date
    description As String
End Type

'==============================================================================
' SUBROUTINE
'   CopyFile
'------------------------------------------------------------------------------
' DESCRIPTION
'   Add's a file to the DB storage system and returns the database ID
'==============================================================================
Public Function DBAddFile(scFromFullPath As String, scBasePathCategory As String, _
        scOwner As String, scDescription As String, ByRef id As Long) As Boolean
    
#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim ofile As Scripting.File
    Dim fso As Scripting.FileSystemObject
#Else
    Dim cnn As Object
    Dim rs As Object
    Dim ofile As Object
    Dim fso As Object
#End If
    Dim scSQLQuery As String
    
    Dim scFilePath As String
    Dim scFilename As String
    Dim scOrigFileName As String
    Dim iOrigFileSize As Long
    Dim scFileDescription As String
    Dim scExt As String
    
    Dim dtFileDateTime As Date
        
    '========
    ' Does the file exist?
    '========
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(scFromFullPath) Then
        DBAddFile = False
        Exit Function
    End If
    
    Set ofile = fso.GetFile(scFromFullPath)
    dtFileDateTime = ofile.DateCreated
    scOrigFileName = ofile.Name
    iOrigFileSize = ofile.Size
    
    '========
    ' Connect to the DB
    '========
    If Not (ConnectToDB(ldMaintenance, cnn, True)) Then
        'Call MsgBox("Unable to connect to the DB")
        DBAddFile = False
        Exit Function
    End If
    
    '========
    ' Has someone already uploaded this exact file. We assume it is exact
    ' only if the original name, size and date are the same.
    '========
    scSQLQuery = "SELECT * FROM maint.dbo.t_file WHERE (file_date = '" & Format(dtFileDateTime, "YYYY-MM-DD hh:mm:ss") & _
        "') AND (orig_file_name = '" & Replace(scOrigFileName, "'", "''") & _
        "') AND (orig_size = " & iOrigFileSize & ")"
        
    If Not GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs) Then
        DBAddFile = False
        Exit Function
    End If
    
    If Not rs.EOF Then
        id = rs.Fields("pk_file")
        DBAddFile = True
        Exit Function
    End If
    
    '========
    ' Get the base path for this type of file
    '========
    Set rs = CreateObject("ADODB.Recordset")
    scSQLQuery = "SELECT TOP 1 * FROM maint.dbo.t_base_path WHERE path_category = '" & scBasePathCategory & "' ORDER BY priority ASC"
        
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    If rs.EOF Then
        DBAddFile = False
        Exit Function
    End If
    
    '========
    ' Extract the base path
    '========
    scFilePath = rs.Fields("path")
    If Right(scFilePath, 1) <> "\" Then
        scFilePath = scFilePath & "\"
    End If
    scFilename = FileNameFromPath(scFromFullPath)
    scExt = FileExtensionFromPath(scFromFullPath)
    
    Call rs.Close
    
    '========
    ' Insert the new information entry into the Database
    '========
    scSQLQuery = "SET NOCOUNT ON; " & _
            "INSERT INTO maint.dbo.t_file " & _
                        "(relative_path, base_path_category, type, owner, file_date, orig_size, orig_file_name, description) " & _
                        "VALUES " & _
                        "('" & Replace(scFilename, "'", "''") & "', '" & _
                            scBasePathCategory & "', '" & _
                            scExt & "', '" & _
                            Replace(scOwner, "'", "''") & "', '" & _
                            Format(dtFileDateTime, "YYYY-MM-DD hh:mm:ss") & "', " & _
                            iOrigFileSize & ", '" & _
                            Replace(Left(scOrigFileName, 250), "'", "''") & "', '" & _
                            Replace(Left(scDescription, 250), "'", "''") & "'); " & _
                        "SELECT SCOPE_IDENTITY() as pk_file;"

    Set rs = cnn.Execute(scSQLQuery)
    
    If rs.EOF Then
        DBAddFile = False
        Exit Function
    End If
        
    id = rs.Fields("pk_file")
    
    '============
    ' Adjust the filename to include the ID. This ensures it is unique.
    '============
    scFilename = Format(id, "0000") & "-" & scFilename
    
    Call ofile.Copy(scFilePath & scFilename, False)
    
    '============
    ' Update the filename in the database
    '============
    scSQLQuery = "UPDATE maint.dbo.t_file SET relative_path = '" & Replace(scFilename, "'", "''") & "' WHERE pk_file = " & id
    Call cnn.Execute(scSQLQuery)

    DBAddFile = True
End Function

'==============================================================================
' FUNCTION
'   DBGetFile
'------------------------------------------------------------------------------
' DESCRIPTION
'   Get's a reference to a File object from the DB file primary key.
' Returns false if unsuccessful.
'==============================================================================
Public Function DBGetFile(pk_file As Long, ByRef ofile As Object) As Boolean

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim fso As Scripting.FileSystemObject
#Else
    Dim cnn As Object
    Dim rs As Object
    Dim fso As Object
#End If
    Dim scSQLQuery As String
    

    
    Dim scFilePath As String
    Dim scBasePathCategory As String
    Dim scBasePath As String
    Dim scFileDescription As String
    
    '========
    ' Connect to the DB
    '========
    If Not (ConnectToDB(ldMaintenance, cnn, True)) Then
        Call MsgBox("Unable to connect to the DB")
        Exit Function
    End If
    
    scSQLQuery = "SELECT * from maint.dbo.t_file where pk_file = " & pk_file
    
    Set rs = CreateObject("ADODB.Recordset")
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    
    If rs.EOF Then
        DBGetFile = False
        Exit Function
    End If
    
    scBasePathCategory = rs.Fields("base_path_category")
    scFilePath = rs.Fields("relative_path")
    If Not IsNull(rs.Fields("description")) Then
        scFileDescription = rs.Fields("description")
    End If
    
    Call rs.Close
    scSQLQuery = "SELECT TOP 1 * FROM maint.dbo.t_base_path WHERE path_category = '" & scBasePathCategory & "' ORDER BY priority ASC"
        
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    If rs.EOF Then
        DBGetFile = False
        Exit Function
    End If
    
    '===========
    ' Get the base path and make sure it has a trailing slash.
    '===========
    scBasePath = rs.Fields("path")
    If Right(scBasePath, 1) <> "\" Then
        scBasePath = scBasePath & "\"
    End If
    scFilePath = scBasePath & scFilePath

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(scFilePath) Then
        DBGetFile = False
        Exit Function
    End If
    
    Set ofile = fso.GetFile(scFilePath)
    DBGetFile = True

End Function

'==============================================================================
' FUNCTION
'   DBCopyFile
'------------------------------------------------------------------------------
' DESCRIPTION
'   Copies the selected database file to the specified destination directory
' and returns a reference to the new file.
' Returns false if unsuccessful.
'==============================================================================
Public Function DBCopyFile(pk_file As Long, scDestPath As String, ByRef oDestFile As Object) As Boolean

#If DevelopMode = 1 Then
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim fso As Scripting.FileSystemObject
    Dim oDBFile As Scripting.File
#Else
    Dim cnn As Object
    Dim rs As Object
    Dim fso As Object
    Dim oDBFile As Object
#End If
    Dim scSQLQuery As String
    
    
    Dim scFilePath As String
    Dim scBasePathCategory As String
    Dim scFileDescription As String
    Dim scOriginalFileName As String
    
    '========
    ' Get the details of the file with the given primary key.
    '========
    scSQLQuery = "SELECT * from maint.dbo.t_file where pk_file = " & pk_file
    
    If Not GetDBRecordSet(ldMaintenance, cnn, scSQLQuery, rs) Then
        DBCopyFile = False
        Exit Function
    End If
    
    If rs.EOF Then
        DBCopyFile = False
        Exit Function
    End If
    
    scBasePathCategory = rs.Fields("base_path_category")
    scFilePath = rs.Fields("relative_path")
    If Not IsNull(rs.Fields("description")) Then
        scFileDescription = rs.Fields("description")
    End If
    
    If Not IsNull(rs.Fields("orig_file_name")) Then
        scOriginalFileName = rs.Fields("orig_file_name")
    End If
    
    If scOriginalFileName = "" Then
        scOriginalFileName = Mid(scFilePath, InStr(1, scFilePath, "-") + 1)
    End If
    Call rs.Close
    
    '========
    ' Determine the full path of the file
    '========
    
    scSQLQuery = "SELECT TOP 1 * FROM maint.dbo.t_base_path WHERE path_category = '" & scBasePathCategory & "' ORDER BY priority ASC"
        
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)
    If rs.EOF Then
        DBCopyFile = False
        Exit Function
    End If
    
    scFilePath = rs.Fields("path") & scFilePath

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(scFilePath) Then
        DBCopyFile = False
        Exit Function
    End If
    
    Set oDBFile = fso.GetFile(scFilePath)
    
    '============
    ' Construct the new full path
    '============
    Dim scDestFullPath As String
    Dim scNewFileName As String
    
    scDestFullPath = scDestPath
    If (Right(scDestFullPath, 1) <> "\") And (Right(scDestFullPath, 1) <> "/") Then
        scDestFullPath = scDestFullPath & "\"
    End If
    scDestFullPath = scDestFullPath & scOriginalFileName
    
    '============
    ' Copy the file and set the destination file pointer
    '============
    Call oDBFile.Copy(scDestFullPath, True)
    
    Set oDestFile = fso.GetFile(scDestFullPath)
    
    DBCopyFile = True

End Function

