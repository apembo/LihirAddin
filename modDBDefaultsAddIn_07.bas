Attribute VB_Name = "modDBDefaultsAddIn_07"
Option Explicit

'==============================================================================
' MODULE
'   modDBDefaults
'------------------------------------------------------------------------------
' DESCRIPTION
'   Module that sets the default strings etc. This module supports the
' modDBConnect by providing the custom methods unique to the project.
'------------------------------------------------------------------------------
' VERSION
'   1.7 - Renamed GetServer to GetDBServer
'           Renamed SetDefaultServer to SetDefaultDBServer
'   1.6 - Includes the OneView DB
'   1.5 - Includes the Contractor Database
'   1.4 - Includes the DEV DB
'   1.3 - Handles the parts and Asset Mgmt databases
'   1.2 - Uses the renamed enumeration eLihirDatabases (formerly
'         LihirDatabase)
'   1.1 - Not sure
'   1.0 - First release
'==============================================================================

Public Const g_connPrefix_AdamMacbookAir = "Driver=SQL Server;Server=ADAMMACBOOKAIR;UID=db_configurator;PWD=get_it_right;Database="
Public Const g_connPrefix_NMLBC0XV32 = "DRIVER=SQL Server;Server=NMLBC0XV32\SQLEXPRESS;Trusted_Connection=True;Database="
Public Const g_connPrefix_NML6Y4DM32 = "DRIVER=SQL Server;Server=NML6Y4DM32\SQLEXPRESS;Trusted_Connection=True;Database="
Public Const g_connPrefix_NMLLHRDEVDB01 = "DRIVER=SQL Server;Server=NMLLHRDEVDB01;Trusted_Connection=True;Database="

Public Const g_connCCLAS = "Driver=SQL Server;Server=NMLLHRDB04;Database=LIHIR_CCLAS_CHD"
Public Const g_connDowntimePlant = "Driver=SQLServer;Server=NMLLHRPPMS02;UID=ppms_readonly;PWD=!PPMSUser2011;WSID=ASOILTH6HCZ1S;DATABASE=ProcessMoreProcessing"
Public Const g_connDowntimePower = "Driver=SQLServer;Server=NMLLHRPPMS02;UID=ppms_readonlyPU;PWD=pplease;WSID=ASOILTH6HCZ1S;DATABASE=ProcessMorePower"
Public Const g_connAssetMgmt = "Driver=SQL Server;Server=NMLLHRDB03;UID=AMRDbUser;PWD=!pplease14;Database=LGO-AM-Records"
Public Const g_connOneView = "Driver=SQL Server;Server=au04db501nml3\OneView;UID=OneView_Reports_LHR_HSS;PWD=LHR_R3@d_H$$;Database=DW_Newcrest"

'==============================================================================
' Email from Troy Ye of Corporate IT regarding the OneView Interface
'-------------------------------------
' Hi Adam,
'
' Yes - all the SAP work orders information is in the OneView database. Here is
' the connection detail:
'
' DB: au04db501nml3\OneView
' UserName: OneView_Reports_LHR_HSS
' Password: LHR_R3@d_H$$
'
' The notification information is in the table Report_IW38_EDGE_11. Please note
' the data in this table is not real-time, the frequency is:
'   CVO - Hourly Refresh,
'   other sites including Lihir's data are refreshed 4 times a day.
' Franz Hemetsberger was successfully connecting to OneView database for some
' self-service reporting, you may would like to catch up with him if you have
' any issues.
'==============================================================================
'==============================================================================
' Documentation on tables
'------------------------------------------------------------------------------
' There are many tables. Some useful ones appear to be:
' Report_IW38_EDGE_1
'   Basic work order listing including costing info. When tested it had 8446
'   Lihir work orders, spread across all system status's, and apparently only
'   for the last month.
' Report_IW38_EDGE_2:
'   Basic work order listing including costing info. When tested it had 2802
'   Lihir work orders, spread across all system status's, and apparently only
'   for the last month.
' Report_IW38_EDGE_3:
'   Basic work order listing including costing info. When tested it had about
'   5500 Lihir work orders, mostly TECO'd (2 also closed), and apparently only
'   for the last month.
' Report_IW38_EDGE_4:
'   Basic work order listing including costing info. When tested it had about
'   2500 Lihir work orders, nearly all TECO'd (6 Closed), and apparently only
'   for the last month.
' Report_IW38_EDGE_5:
'   Basic work order listing including costing info. When tested it had about
'   800 Lihir work orders, most closed and spread across 12 months.
' Report_IW38_EDGE_6:
'   Basic work order listing including costing info. When tested it had about
'   250 Lihir work orders mostly in the last month, but small numbers back up
'   to 5 months and forward 5 months
' Report_IW38_EDGE_7:
'   Seems to be a listing of work orders that have changed in the last 2
'   weeks. Does not include System Status. When tested, it had 661 Lihir Work
'   Orders spread over more than a year previous, and 4 months forward.
' Report_IW38_EDGE_9:
'   Basic work order listing including costing info. When tested it had about
'   18500 Lihir work orders, all created or released across 12 months before
'   and after.
' Report_IW38_EDGE_10:
'   Fewer fields, but otherwise identical to Report_IW38_EDGE_5
' Report_IW38_EDGE_11:
'   Small number of fields. Basis_Start_Date spread over more than 12 months
'   forward and back. Does not list System_Status. When tested, it had 1442
'   work orders.
' Report_IW38_EDGE_20:
'   Similar to Report_IW38_EDGE_88 but with less parameters.
' Report_IW38_EDGE_88:
'   Seems to have more detail than previous reports. When tested it had about
'   9000 Lihir work orders spread over 5 months before but no months after
'   the current month. All System Status's were represented.
' Report_IW38_EDGE_LIHIR_Adhoc:
'   Basic work order listing including costing info. When tested it had about
'   18500 Lihir work orders of all status's in the current and previous
'   month.
' REPORT_IW28_EDGE_CRTD_LIHIR:
'   Notification Listing for Lihir. At the time of testing, it returned about
'   200 notifications over a very wide range of creation dates, going back
'   2 years.
'==============================================================================



'==============================================================================
' SUBROUTINE
'   SetDBDefaults
'------------------------------------------------------------------------------
' DESCRIPTION
'   Method called by the modDBConnect module if it finds the connection strings
' are empty. The simplest implementation of this method is is to simply call
' the method SetDefaultConnectionStrings, passing in one of the default
' servers.
'   E.g.:
'       Call modDBConnect.SetDefaultConnectionStrings("NML6Y4DM32")
'==============================================================================
Sub SetDBDefaults()

    Dim scConnString As String
    
    '===========
    ' Power & Utilities database
    '===========
    Select Case wsParameters.Range("DBPUServerDefaultID")
        Case "NMLBC0XV32"
            scConnString = wsParameters.Range("NMLBC0XV32_ConnString") & "pu"
            
        Case "AdamMacbookAir"
            scConnString = wsParameters.Range("AdamMacbookAir_ConnString") & "pu"
            
        Case "AdamDellE7250", "NML6Y4DM32"
            scConnString = wsParameters.Range("AdamDellE7250_ConnString") & "pu"
            
        Case "DEVDB", "NMLLHRDEVDB01"
            scConnString = wsParameters.Range("NMLLHRDEVDB01_ConnString") & "pu"
        
        Case Else
            scConnString = ""
    End Select
            
    Call SetConnectionString(eLihirDatabases.ldPowerAndUtilities, scConnString)
        
    '===========
    ' Maintenance database
    '===========
    Select Case wsParameters.Range("DBMaintServerDefaultID")
        Case "NMLBC0XV32"
            scConnString = wsParameters.Range("NMLBC0XV32_ConnString") & "maint"
            
        Case "AdamMacbookAir"
            scConnString = wsParameters.Range("AdamMacbookAir_ConnString") & "maint"
            
        Case "AdamDellE7250", "NML6Y4DM32"
            scConnString = wsParameters.Range("AdamDellE7250_ConnString") & "maint"
            
        Case "DEVDB", "NMLLHRDEVDB01"
            scConnString = wsParameters.Range("NMLLHRDEVDB01_ConnString") & "maint"
        
        Case Else
            scConnString = ""
    End Select
        
    Call SetConnectionString(eLihirDatabases.ldMaintenance, scConnString)
        
    '===========
    ' CCLAS database
    '===========
    scConnString = g_connCCLAS
    Call SetConnectionString(eLihirDatabases.ldCCLAS, scConnString)
                    
    '===========
    ' PPMS Plant database
    '===========
    scConnString = g_connDowntimePlant
    Call SetConnectionString(eLihirDatabases.ldPPMSPlant, scConnString)
                    
    '===========
    ' PPMS Power database
    '===========
    scConnString = g_connDowntimePower
    Call SetConnectionString(eLihirDatabases.ldPPMSPower, scConnString)
                    
    '===========
    ' Finance database
    '===========
    Select Case wsParameters.Range("DBMaintServerDefaultID")
        Case "NMLBC0XV32"
            scConnString = wsParameters.Range("NMLBC0XV32_ConnString") & "finance"
            
        Case "AdamMacbookAir"
            scConnString = wsParameters.Range("AdamMacbookAir_ConnString") & "finance"
            
        Case "AdamDellE7250", "NML6Y4DM32"
            scConnString = wsParameters.Range("AdamDellE7250_ConnString") & "finance"
            
        Case "DEVDB", "NMLLHRDEVDB01"
            scConnString = wsParameters.Range("NMLLHRDEVDB01_ConnString") & "finance"
        
        Case Else
            scConnString = ""
    End Select
        
    Call SetConnectionString(eLihirDatabases.ldFinance, scConnString)
    
    '===========
    ' People database
    '===========
    Select Case wsParameters.Range("DBPeopleServerDefaultID")
        Case "NMLBC0XV32"
            scConnString = wsParameters.Range("NMLBC0XV32_ConnString") & "people"
            
        Case "AdamMacbookAir"
            scConnString = wsParameters.Range("AdamMacbookAir_ConnString") & "people"
            
        Case "AdamDellE7250", "NML6Y4DM32"
            scConnString = wsParameters.Range("AdamDellE7250_ConnString") & "people"
            
        Case "DEVDB", "NMLLHRDEVDB01"
            scConnString = wsParameters.Range("NMLLHRDEVDB01_ConnString") & "people"
        
        Case Else
            scConnString = ""
    End Select
        
    Call SetConnectionString(eLihirDatabases.ldPeople, scConnString)
    
    '===========
    ' Parts database
    '===========
    Select Case wsParameters.Range("DBPartsServerDefaultID")
        Case "NMLBC0XV32"
            scConnString = wsParameters.Range("NMLBC0XV32_ConnString") & "parts"
            
        Case "AdamMacbookAir"
            scConnString = wsParameters.Range("AdamMacbookAir_ConnString") & "parts"
            
        Case "AdamDellE7250", "NML6Y4DM32"
            scConnString = wsParameters.Range("AdamDellE7250_ConnString") & "parts"
            
        Case "DEVDB", "NMLLHRDEVDB01"
            scConnString = wsParameters.Range("NMLLHRDEVDB01_ConnString") & "parts"
        
        Case Else
            scConnString = ""
    End Select
        
    Call SetConnectionString(eLihirDatabases.ldParts, scConnString)
    
    '===========
    ' Lihir Asset Management (including Shift Log) Database
    '===========
    scConnString = g_connAssetMgmt

    Call SetConnectionString(eLihirDatabases.ldAssetMgmt, scConnString)

    '===========
    ' Contractors database
    '===========
    Select Case wsParameters.Range("DBContractorsServerDefaultID")
        Case "NMLBC0XV32"
            scConnString = wsParameters.Range("NMLBC0XV32_ConnString") & "CC_LHR_Contractors"
            
        Case "AdamMacbookAir"
            scConnString = wsParameters.Range("AdamMacbookAir_ConnString") & "CC_LHR_Contractors"
            
        Case "AdamDellE7250", "NML6Y4DM32"
            scConnString = wsParameters.Range("AdamDellE7250_ConnString") & "CC_LHR_Contractors"
            
        Case "DEVDB", "NMLLHRDEVDB01"
            scConnString = wsParameters.Range("NMLLHRDEVDB01_ConnString") & "CC_LHR_Contractors"
        
        Case Else
            scConnString = ""
    End Select

    Call SetConnectionString(eLihirDatabases.ldContractors, scConnString)
    
    '===========
    ' Corporate OneView database.
    '===========
    scConnString = g_connOneView

    Call SetConnectionString(eLihirDatabases.ldOneViewMaint, scConnString)

End Sub

'==============================================================================
' SUBROUTINE
'   SetDefaultDBServer
'------------------------------------------------------------------------------
' DESCRIPTION
'   Changes the default server for the particular database
'==============================================================================
Sub SetDefaultDBServer(eDB As eLihirDatabases, scServer As String)

    Select Case eDB
        Case eLihirDatabases.ldPowerAndUtilities
            wsParameters.Range("DBPUServerDefaultID") = scServer
        Case eLihirDatabases.ldMaintenance
            wsParameters.Range("DBMaintServerDefaultID") = scServer
        Case eLihirDatabases.ldCCLAS
            wsParameters.Range("DBCCLASServerDefaultID") = scServer
        Case eLihirDatabases.ldPPMSPlant, eLihirDatabases.ldPPMSPower
            wsParameters.Range("DBPPMSServerDefaultID") = scServer
        Case eLihirDatabases.ldFinance
            wsParameters.Range("DBFinanceServerDefaultID") = scServer
        Case eLihirDatabases.ldPeople
            wsParameters.Range("DBPeopleServerDefaultID") = scServer
        Case eLihirDatabases.ldParts
            wsParameters.Range("DBPartsServerDefaultID") = scServer
        Case eLihirDatabases.ldContractors
            wsParameters.Range("DBContractorsServerDefaultID") = scServer

        Case eLihirDatabases.ldCCLAS, _
            eLihirDatabases.ldPPMSPlant, _
            eLihirDatabases.ldPPMSPower, _
            eLihirDatabases.ldAssetMgmt, _
            eLihirDatabases.ldOneViewMaint
                ' Do nothing

    End Select
    
    Call SetDBDefaults
    
End Sub

'==============================================================================
' SUBROUTINE
'   GetDBServer
'------------------------------------------------------------------------------
' DESCRIPTION
'   Changes the default server for the particular database
'==============================================================================
Sub GetDBServer(eDB As eLihirDatabases, ByRef scServer As String, ByRef scWorkstation As String)
    
    Select Case eDB
        Case eLihirDatabases.ldPowerAndUtilities
            scWorkstation = wsParameters.Range("DBPUServerDefaultID")
        Case eLihirDatabases.ldMaintenance
            scWorkstation = wsParameters.Range("DBMaintServerDefaultID")
        Case eLihirDatabases.ldCCLAS
            scWorkstation = wsParameters.Range("DBCCLASServerDefaultID")
        Case eLihirDatabases.ldPPMSPlant, eLihirDatabases.ldPPMSPower
            scWorkstation = wsParameters.Range("DBPPMSServerDefaultID")
        Case eLihirDatabases.ldFinance
            scWorkstation = wsParameters.Range("DBFinanceServerDefaultID")
        Case eLihirDatabases.ldPeople
            scWorkstation = wsParameters.Range("DBPeopleServerDefaultID")
        Case eLihirDatabases.ldParts
            scWorkstation = wsParameters.Range("DBPartsServerDefaultID")
        Case eLihirDatabases.ldContractors
            scWorkstation = wsParameters.Range("DBContractorsServerDefaultID")
        Case eLihirDatabases.ldOneViewMaint
            scWorkstation = "au04db501nml3"
    End Select
    
    If scWorkstation = "AdamDellE7250" Then
        scWorkstation = "NML6Y4DM32"
    End If
    
    Select Case scWorkstation
        Case "NMLBC0XV32", "NML6Y4DM32", "AdamMacbookAir", "AdamDellE7250"
            scServer = scWorkstation & "\SQLEXPRESS"
        Case "au04db501nml3"
            scServer = "au04db501nml3\OneView"
        Case Else
            scServer = scWorkstation
    End Select

End Sub

