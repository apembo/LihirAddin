VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPOItem"
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
Private m_bErrorsPopulating As Boolean

'==============================================================================
' PUBLIC MEMBER VARIABLES
'==============================================================================
Public pk_purch_doc As String
Public pk_item As String
Public short_text As String
Public fk_material As Long
Public fk_material_group As Long
Public material_group As String
Public goods_recipient As String
Public release_indicator As String
Public Tracking As String
Public order_qty As Double
Public order_unit As String
Public unit_price As Double
Public net_price As Double
Public curr As String
Public fk_cost_element As Long
Public cost_element As String
Public qty_to_be_delivered As Double
Public value_to_be_delivered As Double
Public qty_to_be_invoiced As Double
Public value_to_be_invoiced As Double
Public DeliveryStatus As String
Public CostCenter As String
Public Dept As String
Public Area As String
Public CCColloquialName As String
Public WorkOrder As Long
Public WorkOrderDescription As String
Public CostingType As String
Public Costing As String

'==============================================================================
' SUBROUTINE - EVENT HANDLER
'   Class_Initialize
'------------------------------------------------------------------------------
' DESCRIPTION
'   Called when the class is created.
'==============================================================================
Private Sub Class_Initialize()
    pk_purch_doc = ""
End Sub

'==============================================================================
' SUBROUTINE
'   Populate
'------------------------------------------------------------------------------
' DESCRIPTION
'   Takes an entry from the finance.dbo.[v_purch_doc_items] view
' and populates this class instance.
'==============================================================================
Public Sub Populate(oFields As Object) ' Fields)

On Error GoTo end_with_errors

    pk_purch_doc = oFields("pk_purch_doc")
    pk_item = oFields("pk_item")
    short_text = oFields("short_text")
    
    If Not IsNull(oFields("fk_material")) Then
        fk_material = oFields("fk_material")
    End If
    fk_material_group = oFields("fk_material_group")
    material_group = oFields("material_group")
    
    If Not IsNull(oFields("goods_recipient")) Then
        goods_recipient = oFields("goods_recipient")
    End If
    
    If Not IsNull(oFields("release_indicator")) Then
        release_indicator = oFields("release_indicator")
    End If
    
    If Not IsNull(oFields("Tracking")) Then
        Tracking = oFields("Tracking")
    End If
    
    
    order_qty = oFields("order_qty")
    order_unit = oFields("order_unit")
    unit_price = oFields("unit_price")
    net_price = oFields("net_price")
    curr = oFields("currency")
    fk_cost_element = oFields("fk_cost_element")
    cost_element = oFields("cost_element")
    
    qty_to_be_delivered = oFields("qty_to_be_delivered")
    value_to_be_delivered = oFields("value_to_be_delivered")
    qty_to_be_invoiced = oFields("qty_to_be_invoiced")
    value_to_be_invoiced = oFields("value_to_be_invoiced")
    
    DeliveryStatus = oFields("DeliveryStatus")
    
    If Not IsNull(oFields("CostCenter")) Then
        CostCenter = oFields("CostCenter")
        Dept = oFields("Dept")
        If Not IsNull(oFields("Area")) Then
            Area = oFields("Area")
        End If
        CCColloquialName = oFields("CCColloquialName")
    End If
    
    If Not IsNull(oFields("WorkOrder")) Then
        WorkOrder = oFields("WorkOrder")
        WorkOrderDescription = oFields("WorkOrderDescription")
    End If
    
    CostingType = oFields("CostingType")
    Costing = oFields("Costing")
    
    m_bErrorsPopulating = False
    Exit Sub
    
end_with_errors:
    m_bErrorsPopulating = True
    
End Sub

'==============================================================================
' PROPERTY
'   ErrorPopulating
'------------------------------------------------------------------------------
' DESCRIPTION
'   Returns the status of the m_bErrorsPopulating flag
'==============================================================================
Public Property Get ErrorPopulating() As Boolean
    ErrorPopulating = m_bErrorsPopulating
End Property

