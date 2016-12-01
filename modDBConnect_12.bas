Attribute VB_Name = "modDBConnect_12"
'==============================================================================
' MODULE
'   modDBConnect
'------------------------------------------------------------------------------
' DESCRIPTION
'   Handles connecting to a specific database
'------------------------------------------------------------------------------
' VERSION
'   1.12| 22-Sep-2016   | Adam Pemberton
'   - Changed the OneView Database to OneViewMaint to reflect it's a pointer
'     to one specific database.
'..............................................................................
'   1.11| 3-Aug-2016    | Adam Pemberton
'   - Added the Oneview Database
'..............................................................................
'   1.10| 26-Jul-2016   | Adam Pemberton
'   - Added the Contractor database
'..............................................................................
'   1.9 | 17-Jun-2016   | Adam Pemberton
'   - Added the Development Database Server NMLLHRDEVDB01
'..............................................................................
'   1.8 | 09-Jun-2016   | Adam Pemberton
'   - Added the parts database, and Franz asset management (including shift log)
'     database
'..............................................................................
'   1.7 | 28-Oct-2015   | Adam Pemberton
'   - Added the GetDatabaseEnumFromString method.
'   - Changed the database Enum name from eLihirDatabases to eLihirDatabases
'..............................................................................
'   1.6 | 26-Oct-2015   | Adam Pemberton
'   - Added the People database in a couple of places where it was missing.
'   - Added a CloseAllConnections method to close any open connections.
'..............................................................................
'   1.5 | 11-Sep-2015   | Adam Pemberton
'   - Converted it to late binding
'..............................................................................
'   1.4 | 01-Aug-2015   | Adam Pemberton
'   - Added the Doc Control DB
'..............................................................................
'   1.3 | 01-Aug-2015   | Adam Pemberton
'   - Added the Finance DB
'..............................................................................
'   1.2 | 08-Jun-2015   | Adam Pemberton
'   - General tidy-up.
'..............................................................................
'   1.1 | 23-Feb-2015   | Adam Pemberton
'   - Multiple database aware.
'..............................................................................
'   1.0 | First release
'==============================================================================
Option Explicit

'==============================================================================
' ENUMERATION's
'==============================================================================
Public Enum eLihirDatabases
    ldPowerAndUtilities = 0
    ldMaintenance = 1
    ldCCLAS = 2
    ldPPMSPlant = 3
    ldPPMSPower = 4
    ldFinance = 5
    ldDocControl = 6
    ldPeople = 7
    ldParts = 8
    ldAssetMgmt = 9
    ldContractors = 10
    ldOneViewMaint = 11
End Enum
Public Const g_cLihirDatabaseCount As Long = 12

'==============================================================================
' Replication of the ADODB CursorTypeEnum enumeration, so that we can remove
' references to the ActiveX Data Objects library
'==============================================================================
Public Enum ADODB_CursorTypeEnum
    adOpenUnspecified_ = -1 ' Does not specify the type of cursor.
    adOpenForwardOnly_ = 0 ' Default. Uses a forward-only cursor. Identical to
            ' a static cursor, except that you can only scroll forward through
            ' records. This improves performance when you need to make only one
            ' pass through a Recordset.
    adOpenKeyset_ = 1 ' Uses a keyset cursor. Like a dynamic cursor, except
            ' that you can't see records that other users add, although records
            ' that other users delete are inaccessible from your Recordset.
            ' Data changes by other users are still visible.
    adOpenDynamic_ = 2 ' Uses a dynamic cursor. Additions, changes, and
            ' deletions by other users are visible, and all types of movement
            ' through the Recordset are allowed, except for bookmarks, if the
            ' provider doesn't support them.
    adOpenStatic_ = 3 ' Uses a static cursor. A static copy of a set of records
            ' that you can use to find data or generate reports. Additions,
            ' changes, or deletions by other users are not visible.
End Enum

'==============================================================================
' Replication of the ADODB LockTypeEnum enumeration, so that we can remove
' references to the ActiveX Data Objects library
'==============================================================================
Public Enum ADODB_LockTypeEnum
    adLockUnspecified_ = -1 ' Unspecified type of lock. Clones inherits lock
            ' type from the original Recordset.
    adLockReadOnly_ = 1 ' Read-only records
    adLockPessimistic_ = 2 ' Pessimistic locking, record by record. The
            ' provider lock records immediately after editing
    adLockOptimistic_ = 3 ' Optimistic locking, record by record. The provider
            ' lock records only when calling update
    adLockBatchOptimistic_ = 4 ' Optimistic batch updates. Required for batch
            ' update mode
End Enum

'==============================================================================
' Replication of the ADODB ObjectStateEnum enumeration, so that we can remove
' references to the ActiveX Data Objects library
'==============================================================================
Public Enum ADODB_ObjectStateEnum
    adStateClosed_ = 0
    adStateOpen_ = 1
    adStateConnecting_ = 2
    adStateExecuting_ = 4
    adStateFetching_ = 8
End Enum

'==============================================================================
' PUBLIC/GLOBAL VARIABLES
'==============================================================================
Dim g_oaCnn(0 To (g_cLihirDatabaseCount - 1)) As Object ' ADODB.Connection

Dim g_scaConnString(0 To (g_cLihirDatabaseCount - 1)) As String

'==============================================================================
' FUNCTION
'   ConnectToDB
'------------------------------------------------------------------------------
' DESCRIPTION
'   Attempts to connect to the selected database.
'==============================================================================
Public Function ConnectToDB(eDB As eLihirDatabases, _
                            ByRef cnn As Object, _
                            Optional bLeaveOpen = False, _
                            Optional ByRef iErrNo As Long, _
                            Optional ByRef scError As String) As Boolean

    Dim scConnStr As String
    
    '============
    ' Initialise
    '============
    iErrNo = 0
    scError = ""
    
    '============
    ' Are we already connected? If so, return immediately.
    '============
    If Not (g_oaCnn(eDB) Is Nothing) Then
        Set cnn = g_oaCnn(eDB)
        If IsConnected(eDB) Then
            If (bLeaveOpen) Then
                ConnectToDB = True
                Exit Function
            Else
                Call g_oaCnn(eDB).Close
            End If
        End If
    Else
        Set g_oaCnn(eDB) = CreateObject("ADODB.Connection")
        Set cnn = g_oaCnn(eDB)
    End If
    
    '============
    ' What is the connection string? If it's blank, we use the default. If not
    ' then we use whatever it is defined as (may have been set by the user
    ' using the SetConnectionString method.
    '============
    If (g_scaConnString(eDB) = "") Then
        g_scaConnString(eDB) = GetConnectionString(eDB)
    End If
    
    '===========
    ' Open the connection
    '===========
On Error GoTo fail_handler
    Call g_oaCnn(eDB).Open(g_scaConnString(eDB))

    ConnectToDB = True
    Exit Function
    
fail_handler:
    iErrNo = Err.number
    scError = Err.description
    ConnectToDB = False
    
End Function

'==============================================================================
' FUNCTION
'   GetConnectionString
'------------------------------------------------------------------------------
' DESCRIPTION
'   If this method finds a connection string empty, it will call the method
' SetDBDefaults which musT be implemented externally to this module.
' The convention is to implement it on a module called modDBDefaults and it
' should set the connection string for all the connection strings that will
' be required.
'==============================================================================
Function GetConnectionString(eDB As eLihirDatabases) As String

    If (g_scaConnString(eDB) = "") Then
        Call SetDBDefaults
        'GetConnectionString = g_scaConnString(eDB)
        
        If (g_scaConnString(eDB) = "") Then
            Call MsgBox("Connection string undefined")
            GetConnectionString = ""
            Exit Function
        End If
    End If
    
    GetConnectionString = g_scaConnString(eDB)
        
End Function

'==============================================================================
' SUBROUTINE
'   SetConnectionString
'------------------------------------------------------------------------------
' DESCRIPTION
'   Changes the connection string for the particular database defined by eDB,
' to the string passed in. Note: this string must be complete. This module
' does not interpret or modify the string in anyway before attempting to
' connect.
'==============================================================================
Sub SetConnectionString(eDB As eLihirDatabases, scConnString As String)

    g_scaConnString(eDB) = scConnString
    
End Sub

'==============================================================================
' SUBROUTINE
'   GetConnection
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the connection object for the associated database.
'==============================================================================
Public Function GetConnection(eDB As eLihirDatabases) As Object 'ADODB.Connection
    Set GetConnection = g_oaCnn(eDB)
End Function

'==============================================================================
' SUBROUTINE
'   CloseConnection
'------------------------------------------------------------------------------
' DESCRIPTION
'   Closed the connection for the specified database.
'==============================================================================
Public Sub CloseConnection(eDB As eLihirDatabases)

    If Not (g_oaCnn(eDB) Is Nothing) Then
        If (g_oaCnn(eDB).State = adStateOpen_) Then
            Call g_oaCnn(eDB).Close
        End If
        Set g_oaCnn(eDB) = Nothing
    End If
    
End Sub

'==============================================================================
' SUBROUTINE
'   IsConnected
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns true if the specified connection is open, otherwise false.
'==============================================================================
Public Function IsConnected(eDB As eLihirDatabases) As Boolean
    If Not (g_oaCnn(eDB) Is Nothing) Then
        If ((g_oaCnn(eDB).State And adStateOpen_) = adStateOpen_) Then
            IsConnected = True
            Exit Function
        End If
    End If
    IsConnected = False
End Function

'==============================================================================
' SUBROUTINE
'   SetDefaultConnectionStrings
'------------------------------------------------------------------------------
' DESCRIPTION
'   Set's the default connection strings based on the supplied host.
'==============================================================================
Sub SetDefaultConnectionStrings(scHost As String)
    
    Dim scLocalServerString As String
    
    '===========
    ' Figure out the local database string
    '===========
    Select Case scHost
        Case "AdamMacbookAir"
            scLocalServerString = "Driver=SQL Server;Server=ADAMMACBOOKAIR;UID=db_configurator;PWD=get_it_right;Database="
        
        Case "NMLBC0XV32", "NML6Y4DM32"
            scLocalServerString = "Driver=SQL Server;Server=" & scHost & "\SQLEXPRESS;Trusted_Connection=True;Database="
            
        Case "AdamDellE7250"
            scLocalServerString = "Driver=SQL Server;Server=NML6Y4DM32\SQLEXPRESS;Trusted_Connection=True;Database="
        
        Case "DEVDB", "NMLLHRDEVDB01"
            scLocalServerString = "Driver=SQL Server;Server=NMLLHRDEVDB01;Trusted_Connection=True;Database="
        
        Case Else
            Call MsgBox("Unknown server host name '" & scHost & "'")
            Exit Sub
    End Select
    
    '===========
    ' pu database
    '===========
    g_scaConnString(eLihirDatabases.ldPowerAndUtilities) = scLocalServerString & "pu"
    
    '===========
    ' maint database
    '===========
    g_scaConnString(eLihirDatabases.ldMaintenance) = scLocalServerString & "maint"
        
    '===========
    ' CCLAS database
    '===========
    g_scaConnString(eLihirDatabases.ldCCLAS) = "Driver=SQL Server;Server=NMLLHRDB04;Database=LIHIR_CCLAS_CHD"
                
    '===========
    ' PPMS Plant database
    '===========
    g_scaConnString(eLihirDatabases.ldPPMSPlant) = "Driver=SQLServer;Server=NMLLHRPPMS02;UID=ppms_readonly;PWD=!PPMSUser2011;WSID=ASOILTH6HCZ1S;DATABASE=ProcessMoreProcessing"
                
    '===========
    ' PPMS Power database
    '===========
    g_scaConnString(eLihirDatabases.ldPPMSPower) = "Driver=SQLServer;Server=NMLLHRPPMS02;UID=ppms_readonlyPU;PWD=pplease;WSID=ASOILTH6HCZ1S;DATABASE=ProcessMorePower"
                    
    '===========
    ' finance database
    '===========
    g_scaConnString(eLihirDatabases.ldFinance) = scLocalServerString & "finance"
    
    '===========
    ' Doc Control database
    '===========
    g_scaConnString(eLihirDatabases.ldDocControl) = scLocalServerString & "controlled_docs"
    
    '===========
    ' people database
    '===========
    g_scaConnString(eLihirDatabases.ldPeople) = scLocalServerString & "people"

    '===========
    ' parts database
    '===========
    g_scaConnString(eLihirDatabases.ldParts) = scLocalServerString & "parts"

    '===========
    ' ShiftLog Database
    '===========
    g_scaConnString(eLihirDatabases.ldAssetMgmt) = "Driver=SQL Server;Server=NMLLHRDB03;UID=AMRDbUser;PWD=!pplease14;Database=LGO-AM-Records"

    '===========
    ' Contractor Database
    '===========
    g_scaConnString(eLihirDatabases.ldContractors) = scLocalServerString & "CC_LHR_Contractors"

    '===========
    ' OneView Database
    '===========
    g_scaConnString(eLihirDatabases.ldOneViewMaint) = g_connOneView
End Sub

'==============================================================================
' SUBROUTINE
'   GetStringFieldValue
'------------------------------------------------------------------------------
' DESCRIPTION
'   Gets the field value as a string, replacing NULL's with blanks.
'   rs must be an ADODB.recordset object
'==============================================================================
Public Function GetStringFieldValue(ByRef rs As Object, scFieldName As String) As String

On Error GoTo exit_cleanly

    If IsNull(rs.Fields(scFieldName)) Then
        GetStringFieldValue = ""
    Else
        GetStringFieldValue = rs.Fields(scFieldName)
    End If
    Exit Function
    
exit_cleanly:
    GetStringFieldValue = ""
End Function

'==============================================================================
' SUBROUTINE
'   GetDateFieldValue
'------------------------------------------------------------------------------
' DESCRIPTION
'   Gets the field value as a date, replacing NULL's with the date equivalent
' of 0.
'   rs must be an ADODB.recordset object
'==============================================================================
Public Function GetDateFieldValue(ByRef rs As Object, scFieldName As String) As Date

On Error GoTo exit_cleanly

    If IsNull(rs.Fields(scFieldName)) Then
        GetDateFieldValue = 0
    Else
        GetDateFieldValue = rs.Fields(scFieldName)
    End If
    Exit Function
    
exit_cleanly:
    GetDateFieldValue = 0
End Function

'==============================================================================
' SUBROUTINE
'   GetLongFieldValue
'------------------------------------------------------------------------------
' DESCRIPTION
'   Gets the field value as a long, replacing NULL's with 0's.
'   rs must be an ADODB.recordset object
'==============================================================================
Public Function GetLongFieldValue(ByRef rs As Object, scFieldName As String) As Long

On Error GoTo exit_cleanly

    If IsNull(rs.Fields(scFieldName)) Then
        GetLongFieldValue = 0
    Else
        GetLongFieldValue = rs.Fields(scFieldName)
    End If
    Exit Function
    
exit_cleanly:
    GetLongFieldValue = 0
End Function

'==============================================================================
' SUBROUTINE
'   GetDoubleFieldValue
'------------------------------------------------------------------------------
' DESCRIPTION
'   Gets the field value as a double, replacing NULL's with 0's.
'   rs must be an ADODB.recordset object
'==============================================================================
Public Function GetDoubleFieldValue(ByRef rs As Object, scFieldName As String) As Double

On Error GoTo exit_cleanly

    If IsNull(rs.Fields(scFieldName)) Then
        GetDoubleFieldValue = 0#
    Else
        GetDoubleFieldValue = rs.Fields(scFieldName)
    End If
    Exit Function
    
exit_cleanly:
    GetDoubleFieldValue = 0#
End Function


'==============================================================================
' SUBROUTINE
'   GetStringFieldValue
'------------------------------------------------------------------------------
' DESCRIPTION
'   Gets the field value as a string, replacing NULL's with blanks.
'   rs must be an ADODB.recordset object
'==============================================================================
Public Function GetStringFieldValueIdx(ByRef rs As Object, iFieldNo As Long) As String

On Error GoTo exit_cleanly

    If IsNull(rs.Fields(iFieldNo).Value) Then
        GetStringFieldValueIdx = ""
    Else
        GetStringFieldValueIdx = rs.Fields(iFieldNo).Value
    End If
    Exit Function
    
exit_cleanly:
    GetStringFieldValueIdx = ""
End Function

'==============================================================================
' SUBROUTINE
'   CloseAllConnections
'------------------------------------------------------------------------------
' DESCRIPTION
'   Closes all open connections.
'==============================================================================
Public Sub CloseAllConnections()
    Dim i As Long

    For i = 0 To (g_cLihirDatabaseCount - 1)
        If IsConnected(i) Then
            Call g_oaCnn(i).Close
        End If
    Next
End Sub

'==============================================================================
' SUBROUTINE
'   GetDatabaseEnumFromString
'------------------------------------------------------------------------------
' DESCRIPTION
'   Gets the field value as a string, replacing NULL's with blanks.
'   rs must be an ADODB.recordset object
'==============================================================================
Public Function GetDatabaseEnumFromString(scDB As String) As eLihirDatabases
    Select Case LCase(scDB)
        Case "pu"
            GetDatabaseEnumFromString = ldPowerAndUtilities
            
        Case "maint"
            GetDatabaseEnumFromString = ldMaintenance
            
        Case "LIHIR_CCLAS_CHD", "lihir_cclas_chd"
            GetDatabaseEnumFromString = ldCCLAS
            
        Case "ProcessMoreProcessing", "processmoreprocessing"
            GetDatabaseEnumFromString = ldPPMSPlant
            
        Case "ProcessMorePower", "processmorepower"
            GetDatabaseEnumFromString = ldPPMSPower
            
        Case "finance"
            GetDatabaseEnumFromString = ldFinance
            
        Case "controlled_docs"
            GetDatabaseEnumFromString = ldDocControl
            
        Case "people"
            GetDatabaseEnumFromString = ldPeople
            
        Case "parts"
            GetDatabaseEnumFromString = ldParts
            
        Case "AssetMgmt", "ShiftLog"
            GetDatabaseEnumFromString = ldAssetMgmt
            
        Case "cc_lhr_contractors"
            GetDatabaseEnumFromString = ldContractors
            
        Case "OneViewMaint"
            GetDatabaseEnumFromString = ldOneViewMaint
        
        Case Else
            GetDatabaseEnumFromString = -1
    End Select
End Function

'==============================================================================
' Function
'   GetDBRecordSet
'------------------------------------------------------------------------------
' DESCRIPTION
'   This function is just a wrapper for one of the most common ways we interact
' with databases.
'==============================================================================
Public Function GetDBRecordSet(eDB As eLihirDatabases, ByRef cnn As Object, scSQLQuery As String, ByRef rs As Object) As Boolean
    
    '========
    ' Connect to the DB
    '========
    If Not (ConnectToDB(eDB, cnn, True)) Then
        GetDBRecordSet = False
        Exit Function
    End If
    
    '========
    ' Construct the query
    '========

    If rs Is Nothing Then
        Set rs = CreateObject("ADODB.Recordset")
    Else
        If (rs.State And ADODB_ObjectStateEnum.adStateOpen_) = ADODB_ObjectStateEnum.adStateOpen_ Then
            Call rs.Close
        End If
    End If
    
    Call rs.Open(scSQLQuery, cnn, ADODB_CursorTypeEnum.adOpenStatic_, ADODB_LockTypeEnum.adLockReadOnly_)

    GetDBRecordSet = True

End Function

'==============================================================================
' Function
'   TSQLDateStrToDate
'------------------------------------------------------------------------------
' DESCRIPTION
'   Converts the TSQL format date string (YYYY-MM-DD) to a date
'==============================================================================
Public Function TSQLDateStrToDate(scTSQLDateString As String) As Date
    Dim iYear As Long
    Dim iMonth As Long
    Dim iDay As Long
    
    iYear = Val(Left(scTSQLDateString, 4))
    iMonth = Val(Mid(scTSQLDateString, 6, 2))
    iDay = Val(Mid(scTSQLDateString, 9, 2))
    
    TSQLDateStrToDate = DateSerial(iYear, iMonth, iDay)
End Function
