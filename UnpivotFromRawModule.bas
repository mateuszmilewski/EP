Attribute VB_Name = "UnpivotFromRawModule"
Option Explicit

'The MIT License (MIT)
'
'Copyright (c) 2020 FORREST
' Mateusz Milewski mateusz.milewski@mpsa.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.


Public Sub parseAndTryToUnpivot(ictrl As IRibbonControl)
    prepareForUnPivot
End Sub





Public Sub prepareForUnPivot()


    Dim shData As Worksheet
    Dim matrixDic As New Dictionary
    
    Dim raw As Worksheet
    Set raw = ThisWorkbook.Sheets("RAW_DATA")
    
    Dim tmpM2 As IMatrix
    
    
    Dim strArr As Variant
    strArr = Array("P1", "P2", "R1", "R2", "R3", "M1", "M2", "M3", "T1", "T2", "T3", "S1")
    
    Dim c As New Common
    Dim data As New Collection
    
    Dim ii As Variant
    ' strArr = Array("P1", "P2", "R1", "R2", "M1", "M2")
    For ii = LBound(strArr) To UBound(strArr)
        'Set tmpM2 = Nothing
        'Set tmpM2 = c.collectMatrix(raw, "" & CStr(strArr(ii)))
        matrixDic.Add "" & CStr(strArr(ii)) & "M", c.collectMatrix(raw, "" & CStr(strArr(ii)))
    Next ii
    

    
    Dim p1 As ISourceProxy, p2 As ISourceProxy
    Dim r1 As ISourceProxy, r2 As ISourceProxy, r3 As ISourceProxy
    Dim m1 As ISourceProxy, m2 As ISourceProxy, m3 As ISourceProxy
    Dim t1 As ISourceProxy, t2 As ISourceProxy, t3 As ISourceProxy
    Dim s1 As ISourceProxy
    
    
    
    Set p1 = New PerturbationHandler
    Set p1.m = matrixDic("P1M")
    p1.amplify Nothing
    
    ' p1.AmpedData.testReturnData ' OK
    On Error Resume Next
    Set data = p1.AmpedData.dataset
    
    
    ' test on count
    'Debug.Print data.Count
    
    Set p2 = New PerturbationHandler
    Set p2.m = matrixDic("P2M")
    ' amped data is nothing this below will not work
    ' Set p2.AmpedData.dataset = data
    p2.amplify data
    
    On Error Resume Next
    Set data = p2.AmpedData.dataset
    
    
    
    
    ' retard portion
    ' ===================================================
    Set r1 = New RetardHandler
    Set r1.m = matrixDic("R1M")
    r1.amplify data
    
    On Error Resume Next
    Set data = r1.AmpedData.dataset
    
    
    Set r2 = New RetardHandler
    Set r2.m = matrixDic("R2M")
    r2.amplify data
    
    On Error Resume Next
    Set data = r2.AmpedData.dataset
    
    Set r3 = New RetardHandler
    Set r3.m = matrixDic("R3M")
    r3.amplify data
    
    On Error Resume Next
    Set data = r3.AmpedData.dataset
    ' ===================================================
    
    
    Set m1 = New ManquantsHandler
    Set m1.m = matrixDic("M1M")
    m1.amplify data
    
    Set data = m1.AmpedData.dataset
    
    
    Set m2 = New ManquantsHandler
    Set m2.m = matrixDic("M2M")
    m2.amplify data
    
    On Error Resume Next
    Set data = m2.AmpedData.dataset
    
    Set m3 = New ManquantsHandler
    Set m3.m = matrixDic("M3M")
    m3.amplify data
    
    On Error Resume Next
    Set data = m3.AmpedData.dataset
    
    
    ' --------------------------------------
    
    Set t1 = New TransferHandler
    Set t1.m = matrixDic("T1M")
    t1.amplify data
    
    Set data = t1.AmpedData.dataset
    
    Set t2 = New TransferHandler
    Set t2.m = matrixDic("T2M")
    t2.amplify data
    
    Set data = t2.AmpedData.dataset
    
    Set t3 = New TransferHandler
    Set t3.m = matrixDic("T3M")
    t3.amplify data
    
    On Error Resume Next
    Set data = t2.AmpedData.dataset
    
    ' --------------------------------------
    
    
    ' STOCK
    ' --------------------------------------
    
    Set s1 = New StockHandler
    Set s1.m = matrixDic("S1M")
    s1.amplify data
    
    On Error Resume Next
    Set data = t1.AmpedData.dataset
    
    ' --------------------------------------
    
    
    ' Debug.Print data.Count ' OK!
    
    Set shData = ThisWorkbook.Sheets.Add
    shData.name = c.tryToRenameIt(shData, "DATA_", "")
    
    Dim rData As Range
    Set rData = shData.Range("A2")
    
    With shData
        .Cells(1, EP.E_DATA.E_DATA_PLT).Value = "PLT"
        .Cells(1, EP.E_DATA.E_DATA_T).Value = "T"
        .Cells(1, EP.E_DATA.E_DATA_Y).Value = "YEAR"
        .Cells(1, EP.E_DATA.E_DATA_M).Value = "MONTH"
        .Cells(1, EP.E_DATA.E_DATA_STR_MONTH).Value = "STR_MONTH"
        .Cells(1, EP.E_DATA.E_DATA_REF_DATE) = "DATE_REF"
        ' new 3 dates
        ' -----------------------------------------------------------------------
        .Cells(1, EP.E_DATA.E_DATA_INPUT_DATE) = "INPUT_DATE"
        .Cells(1, EP.E_DATA.E_DATA_RESULT_DATE) = "RESULT_DATE"
        .Cells(1, EP.E_DATA.E_DATA_VISIBILITY_DATE) = "VISIBILITY_DATE"
        ' -----------------------------------------------------------------------
        .Cells(1, EP.E_DATA.E_DATA_COMMON_DATATYPE_CODE).Value = "DTC"
        .Cells(1, EP.E_DATA.E_DATA_SPECIFIC_DATATYPE_CODE).Value = "DATATYPE"
        .Cells(1, EP.E_DATA.E_DATA_STR_DATATYPE_CODE).Value = "STR_DT"
        .Cells(1, EP.E_DATA.E_DATA_V).Value = "V"
        .Cells(1, EP.E_DATA.E_DATA_YELLOW_TRIGGER).Value = "YELLOW"
        .Cells(1, EP.E_DATA.E_DATA_GREEN_TRIGGER).Value = "GREEN"
        
    End With
    
    Dim i As AmpedItem
    
    Dim pltC As Integer
    Dim tC As Integer, yC As Integer, mC As Integer, drC As Integer
    Dim dtC1 As Integer, dtC2 As Integer, dtC3 As Integer, strMonthInt As Integer
    Dim vC As Integer
    
    pltC = EP.E_DATA.E_DATA_PLT - 1
    
    tC = EP.E_DATA.E_DATA_T - 1
    yC = EP.E_DATA.E_DATA_Y - 1
    mC = EP.E_DATA.E_DATA_M - 1
    drC = EP.E_DATA.E_DATA_REF_DATE - 1
    
    
    
    ' new 3 dates
    '===================================================
    'E_DATA_INPUT_DATE
    'E_DATA_RESULT_DATE
    'E_DATA_VISIBILITY_DATE
    '===================================================
    dtC1 = EP.E_DATA.E_DATA_COMMON_DATATYPE_CODE - 1
    dtC2 = EP.E_DATA.E_DATA_SPECIFIC_DATATYPE_CODE - 1
    dtC3 = EP.E_DATA_STR_DATATYPE_CODE - 1
    
    
    strMonthInt = EP.E_DATA.E_DATA_STR_MONTH - 1
    
    vC = EP.E_DATA_V - 1
    
    
    For Each i In data
        
        rData.offset(0, pltC).Value = CStr(i.stdPlt)
        rData.offset(0, tC).Value = CStr(i.stdTiming)
        rData.offset(0, yC).Value = CStr(i.stdYear)
        On Error Resume Next
        rData.offset(0, mC).Value = CDbl(i.stdMonth)
        
        
        rData.offset(0, drC).Value = i.referenceDate
        
        ' new 3 dates
        ' E_DATA_INPUT_DATE
        rData.offset(0, drC + 1).Value = i.inputDate
        ' E_DATA_RESULT_DATE
        rData.offset(0, drC + 2).Value = i.resultDate
        ' E_DATA_VISIBILITY_DATE
        rData.offset(0, drC + 3).Value = i.visibleDate
        ' -----------------------------------------------
        
        ' str month
        ' -----------------------------------------------
        rData.offset(0, strMonthInt).Value = i.strMonth
        ' -----------------------------------------------
        
        
        
        rData.offset(0, dtC1).Value = Left(i.stdDataType, 1)
        rData.offset(0, dtC2).Value = CStr(i.stdDataType)
        rData.offset(0, dtC3).Value = CStr(i.stdImprovedDataType)
        
        ' Debug.Assert Not (i.v Like "*541*001")
        If i.v = "" Then
        ElseIf IsError(i.v) Then
        ElseIf IsNumeric(i.v) Then
            rData.offset(0, vC).Value = CDbl(i.v)
        Else
        End If
        
        ' nre
        rData.offset(0, E_DATA.E_DATA_YELLOW_TRIGGER - 1).Value = i.yellowTrigger
        rData.offset(0, E_DATA.E_DATA_GREEN_TRIGGER - 1).Value = i.greenTrigger
        
        Set rData = rData.offset(1, 0)
        
        If rData.row Mod 20 = 0 Then
            rData.Select
        End If
    Next i
    
    
    
End Sub
