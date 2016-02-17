VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAPCostType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function createMultiple(pControllingArea As String, pCostElemClass As String, pTestRun As String, pData As Collection)
    On Error GoTo createMultiple_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCostType As Object
    Dim oReturn As Object
    Dim oCostElementList As Object
    Dim aControllingArea As String
    Dim aTestRun As String

    Set boCostType = MySAPCon.SAPBapiControl.GetSAPObject("CostType")
    Set oReturn = MySAPCon.SAPBapiControl.DimAs(boCostType, "CreateMultiple", "Return")
    Set oCostElementList = MySAPCon.SAPBapiControl.DimAs(boCostType, "CreateMultiple", "CostElementList")

    Dim aCELRow As Object
    Dim lRow As SAPCostElementList
    For Each lRow In pData
        Set aCELRow = oCostElementList.Rows.Add
        aCELRow.Value("COST_ELEM") = lSAPFormat.unpack(lRow.COST_ELEM, 10)
        aCELRow.Value("VALID_FROM") = lRow.VALID_FROM
        aCELRow.Value("VALID_TO") = lRow.VALID_TO
        aCELRow.Value("CELEM_CATEGORY") = lRow.CELEM_CATEGORY
        aCELRow.Value("CELEM_ATTRIBUTE") = lRow.CELEM_ATTRIBUTE
        aCELRow.Value("UNIT_OF_MEASURE") = lRow.UNIT_OF_MEASURE
        aCELRow.Value("DEFAULT_COSTCENTER") = lSAPFormat.unpack(lRow.DEFAULT_COSTCENTER, 10)
        aCELRow.Value("DEFAULT_ORDER") = lSAPFormat.unpack(lRow.DEFAULT_ORDER, 12)
        aCELRow.Value("JV_REC_IND") = lRow.JV_REC_IND
        aCELRow.Value("NAME") = lRow.NAME
        aCELRow.Value("DESCRIPT") = lRow.DESCRIPT
        aCELRow.Value("FUNC_AREA") = lRow.FUNC_AREA
    Next lRow
    Application.Cursor = xlWait
    If pTestRun = "X" Then
        boCostType.createMultiple CoArea:=pControllingArea, _
        TestRun:=pTestRun, _
        CostElemClass:=pCostElemClass, _
        CostElementList:=oCostElementList, _
        Return:=oReturn
    Else
        boCostType.createMultiple CoArea:=pControllingArea, _
        CostElementList:=oCostElementList, _
        CostElemClass:=pCostElemClass, _
        Return:=oReturn
    End If

    Application.Cursor = xlDefault

    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit

    If oReturn.RowCount = 0 Then
        createMultiple = "Success ???"
        aSAPBapiTranctionCommit.commit
    Else
        Dim retRow As Object
        Set retRow = oReturn.Rows(1)
        If retRow("TYPE") = "S" Then
            createMultiple = retRow("MESSAGE")
            aSAPBapiTranctionCommit.commit
        Else
            For Each retRow In oReturn.Rows
                createMultiple = createMultiple & ";" & retRow("MESSAGE")
            Next retRow
        End If
    End If
    Set boCostType = Nothing
    Exit Function

createMultiple_Error:
    Application.Cursor = xlDefault
    createMultiple = "Error"
    MySAPErr.MSGProt "boCostType", "createMultiple", "", Err.Number, Err.Description
    Exit Function
End Function

Function changeMultiple(pControllingArea As String, pLanguageKey As String, pTestRun As String, pData As Collection)
    On Error GoTo changeMultiple_Error
    Dim lSAPFormat As New SAPFormat
    Dim boCostType As Object
    Dim oReturn As Object
    Dim oCostElementList As Object
    Dim aControllingArea As String
    Dim aTestRun As String
    Dim oLanguage As Object

    Set boCostType = MySAPCon.SAPBapiControl.GetSAPObject("CostType")
    Set oReturn = MySAPCon.SAPBapiControl.DimAs(boCostType, "ChangeMultiple", "Return")
    Set oCostElementList = MySAPCon.SAPBapiControl.DimAs(boCostType, "ChangeMultiple", "CostElementList")
    Set oLanguage = MySAPCon.SAPBapiControl.DimAs(boCostType, "ChangeMultiple", "Language")
    oLanguage("LANGU") = aLanguageKey

    Dim aCELRow As Object
    Dim lRow As SAPCostElementList
    For Each lRow In pData
        Set aCELRow = oCostElementList.Rows.Add
        aCELRow.Value("COST_ELEM") = lSAPFormat.unpack(lRow.COST_ELEM, 10)
        aCELRow.Value("VALID_FROM") = lRow.VALID_FROM
        aCELRow.Value("VALID_TO") = lRow.VALID_TO
        aCELRow.Value("CELEM_CATEGORY") = lRow.CELEM_CATEGORY
        aCELRow.Value("CELEM_ATTRIBUTE") = lRow.CELEM_ATTRIBUTE
        aCELRow.Value("UNIT_OF_MEASURE") = lRow.UNIT_OF_MEASURE
        aCELRow.Value("DEFAULT_COSTCENTER") = lSAPFormat.unpack(lRow.DEFAULT_COSTCENTER, 10)
        aCELRow.Value("DEFAULT_ORDER") = lSAPFormat.unpack(lRow.DEFAULT_ORDER, 12)
        aCELRow.Value("JV_REC_IND") = lRow.JV_REC_IND
        aCELRow.Value("NAME") = lRow.NAME
        aCELRow.Value("DESCRIPT") = lRow.DESCRIPT
        aCELRow.Value("FUNC_AREA") = lRow.FUNC_AREA
    Next lRow
    Application.Cursor = xlWait
    If pTestRun = "X" Then
        boCostType.changeMultiple CoArea:=pControllingArea, _
        TestRun:=pTestRun, _
        Language:=oLanguage, _
        CostElementList:=oCostCenterList, _
        Return:=oReturn
    Else
        boCostType.changeMultiple CoArea:=pControllingArea, _
        CostElementList:=oCostCenterList, _
        Language:=oLanguage, _
        Return:=oReturn
    End If

    Application.Cursor = xlDefault

    Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit

    If oReturn.RowCount = 0 Then
        changeMultiple = "Success ???"
        aSAPBapiTranctionCommit.commit
    Else
        Dim retRow As Object
        Set retRow = oReturn.Rows(1)
        If retRow("TYPE") = "S" Then
            changeMultiple = retRow("MESSAGE")
            aSAPBapiTranctionCommit.commit
        Else
            For Each retRow In oReturn.Rows
                changeMultiple = changeMultiple & ";" & retRow("MESSAGE")
            Next retRow
        End If
    End If
    Set boCostType = Nothing
    Exit Function

changeMultiple_Error:
    Application.Cursor = xlDefault
    changeMultiple = "Error"
    MySAPErr.MSGProt "boCostType", "changeMultiple", "", Err.Number, Err.Description
    Exit Function
End Function
