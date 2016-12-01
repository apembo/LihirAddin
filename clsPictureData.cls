VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPictureData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
Private m_i_pk_file As Long
Private m_i_pk_file_map As Long
Private m_sc_relative_path As String
Private m_sc_base_path As String
Private m_dt_file_date As Date
Private m_sc_description As String
Private m_i_file_order As Long
Private m_i_rating_count As Long
Private m_d_rating_avg As Double
Private m_sc_owner As String
Private m_b_changed As Boolean

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   Class_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called on class creation
'==============================================================================
Private Sub Class_Initialize()
    m_i_pk_file = -1
    m_i_pk_file_map = -1
    m_sc_base_path = ""
    m_sc_relative_path = ""
End Sub

'==============================================================================
' PROPERTY - GET
'   pk_file
'==============================================================================
Public Property Get pk_file() As Long
    pk_file = m_i_pk_file
End Property

'==============================================================================
' PROPERTY - GET
'   pk_file_map
'==============================================================================
Public Property Get pk_file_map() As Long
    pk_file_map = m_i_pk_file_map
End Property
Public Property Get Key() As String
    Key = Static_Key(m_i_pk_file_map)
End Property
'==============================================================================
' FUNCTION
'   Static_Key
'------------------------------------------------------------------------------
' DESCRIPTION
'   Although static methods are not available in vba, this method uses this
' concept. This method allows the one method to define how a key is defined
' for both internal class methods and external methods.
' Unlike true static methods, external functions will need to create a
' a specific instance of this class in order to call this method. But they
' won't need to populate the class with anything so the overheads shouldn't
' be too bad.
'==============================================================================
Public Function Static_Key(pk_file_map As Long) As String
    Static_Key = "pk:" & Format(pk_file_map, "00000")
End Function

'==============================================================================
' PROPERTY - GET
'   relative_path
'==============================================================================
Public Property Get relative_path() As String
    relative_path = m_sc_relative_path
End Property
'==============================================================================
' PROPERTY - GET
'   base_path
'==============================================================================
Public Property Get base_path() As String
    base_path = m_sc_base_path
End Property

'==============================================================================
' PROPERTY - GET, SET
'   file_date
'==============================================================================
Public Property Get file_date() As Date
    file_date = m_dt_file_date
End Property
Public Property Let file_date(dtRhs As Date)
    m_dt_file_date = dtRhs
End Property

Public Property Get FileDateStr(eFormat As eDateTimeFormat) As String

    Dim scFormat As String
    
    Select Case eFormat
        Case eDateTimeFormat.DateShort
            scFormat = "D-MMM-YY"
            
        Case eDateTimeFormat.DateLong
            scFormat = "DDD D-MMM-YYYY"
            
        Case eDateTimeFormat.DateShortTime24Hr
            scFormat = "D-MMM-YY hh:mm"
        
        Case eDateTimeFormat.DateShortTimeAmPm
            scFormat = "D-MMM-YY h:mm am/pm"
        
        Case eDateTimeFormat.DateLongTime24Hr
            scFormat = "DDD D-MMM-YYYY hh:mm:ss"
        
        Case eDateTimeFormat.DateLongTimeAmPm
            scFormat = "DDD D-MMM-YYYY h:mm:ss am/pm"
            
        Case eDateTimeFormat.DateTSQL
            scFormat = "YYYY-MM-DD"
        
        Case eDateTimeFormat.DateTimeTSQL
            scFormat = "YYYY-MM-DD hh:mm:ss"
        
        Case Else
            scFormat = "D-MMM-YY"
    End Select
    
    FileDateStr = Format(m_dt_file_date, scFormat)
            
End Property

'==============================================================================
' PROPERTY - GET, SET
'   description
'==============================================================================
Public Property Get description() As String
    description = m_sc_description
End Property
Public Property Let description(scRhs As String)
    m_sc_description = scRhs
End Property

'==============================================================================
' PROPERTY - GET, SET
'   file_order
'==============================================================================
Public Property Get file_order() As Long
    file_order = m_i_file_order
End Property
Public Property Let file_order(iRhs As Long)
    m_i_file_order = iRhs
End Property

'==============================================================================
' PROPERTY - GET, SET
'   rating_count
'==============================================================================
Public Property Get rating_count() As Long
    rating_count = m_i_rating_count
End Property
Public Property Let rating_count(iRhs As Long)
    m_i_rating_count = iRhs
End Property

'==============================================================================
' PROPERTY - GET, SET
'   rating_avg
'==============================================================================
Public Property Get rating_avg() As Double
    rating_avg = m_d_rating_avg
End Property
Public Property Let rating_avg(dRhs As Double)
    m_d_rating_avg = dRhs
End Property

'==============================================================================
' PROPERTY - GET
'   owner
'==============================================================================
Public Property Get owner() As String
    owner = m_sc_owner
End Property

'==============================================================================
' PROPERTY - GET, SET
'   changed
'==============================================================================
Public Property Get changed() As Boolean
    changed = m_b_changed
End Property
Public Property Let changed(bRhs As Boolean)
    m_b_changed = bRhs
End Property

'==============================================================================
' PROPERTY - GET
'   Name - Composite
'==============================================================================
Public Property Get Name() As String
    Name = Mid(m_sc_relative_path, InStr(5, m_sc_relative_path, "-") + 1)
End Property

'==============================================================================
' PROPERTY - GET
'   FullPath - Composite
'==============================================================================
Public Property Get FullPath() As String
    If Right(m_sc_base_path, 1) <> "\" Then
        FullPath = m_sc_base_path & "\" & m_sc_relative_path
    Else
        FullPath = m_sc_base_path & m_sc_relative_path
    End If
End Property

'==============================================================================
' FUNCTION
'   CreateFromRecordset
'------------------------------------------------------------------------------
' DESCRIPTION
'   Creates a new instance of this class from the data in the provided
' RecordSet.
'==============================================================================
Public Sub PopulateFromRecordset(rsFrom As Object) 'ADODB.Recordset)

    m_sc_base_path = rsFrom.Fields("base_path")
    m_i_file_order = rsFrom.Fields("file_order")
    m_i_pk_file = rsFrom.Fields("pk_file")
    m_i_pk_file_map = rsFrom.Fields("pk_file_map")
    m_sc_relative_path = rsFrom.Fields("relative_path")
    
    m_b_changed = 0

    '=============
    ' Handle tidely those fields that could be null.
    '=============
    '----
    ' file_date
    '----
    If Not IsNull(rsFrom.Fields("file_date")) Then
        m_dt_file_date = rsFrom.Fields("file_date")
    Else
        m_dt_file_date = DateSerial(1970, 1, 1)
    End If
    
    '----
    ' description
    '----
    If IsNull(rsFrom.Fields("description")) Then
        m_sc_description = ""
    Else
        m_sc_description = rsFrom.Fields("description")
    End If
    
    '----
    ' rating_count, rating_avg
    '----
    If IsNull(rsFrom.Fields("rating_count")) Then
        m_d_rating_avg = 0#
        m_i_rating_count = 0
    Else
        m_d_rating_avg = rsFrom.Fields("rating_avg")
        m_i_rating_count = rsFrom.Fields("rating_count")
    End If
    
    '----
    ' owner
    '----
    If IsNull(rsFrom.Fields("owner")) Then
        m_sc_owner = ""
    Else
        m_sc_owner = rsFrom.Fields("owner")
    End If
    
End Sub
