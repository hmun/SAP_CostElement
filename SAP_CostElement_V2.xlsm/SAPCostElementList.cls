VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCostElementList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public COST_ELEM As String
Public VALID_FROM As String
Public VALID_TO As String
Public CELEM_CATEGORY As String
Public CELEM_ATTRIBUTE As String
Public RECORD_QUANTITY As String
Public UNIT_OF_MEASURE As String
Public DEFAULT_COSTCENTER As String
Public DEFAULT_ORDER As String
Public JV_REC_IND As String
Public NAME As String
Public DESCRIPT As String
Public FUNC_AREA As String

Public Function create(p_COST_ELEM, p_VALID_FROM, p_VALID_TO, p_CELEM_CATEGORY, p_CELEM_ATTRIBUTE, _
    p_RECORD_QUANTITY, p_UNIT_OF_MEASURE, p_DEFAULT_COSTCENTER, _
    p_DEFAULT_ORDER, p_JV_REC_IND, p_NAME, p_DESCRIPT, p_FUNC_AREA)
    
    COST_ELEM = p_COST_ELEM
    VALID_FROM = p_VALID_FROM
    VALID_TO = p_VALID_TO
    CELEM_CATEGORY = p_CELEM_CATEGORY
    CELEM_ATTRIBUTE = p_CELEM_ATTRIBUTE
    RECORD_QUANTITY = p_RECORD_QUANTITY
    UNIT_OF_MEASURE = p_UNIT_OF_MEASURE
    DEFAULT_COSTCENTER = p_DEFAULT_COSTCENTER
    DEFAULT_ORDER = p_DEFAULT_ORDER
    JV_REC_IND = p_JV_REC_IND
    NAME = p_NAME
    DESCRIPT = p_DESCRIPT
    FUNC_AREA = p_FUNC_AREA
End Function

