Attribute VB_Name = "TP04Module"
Option Explicit

'The MIT License (MIT)
'
'Copyright (c) 2021 FORREST
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

' THIS GOING FIRST
' ============================================
' ============================================
Public Sub getPricesFromTP04()
    
    Dim sapH As New SAP_Handler, outFromTp04 As Worksheet
    ' ActiveSheet sheet as input for logic inside
    sapH.runMainLogicForSQ01__forEP ActiveSheet, "A2", outFromTp04, "TP04_377"
    
    ' E_FROM_SQ01_QUASI_TP04
    outFromTp04FillLabels outFromTp04
    removeEmptyRows outFromTp04
    extend_and_std_tp04_data outFromTp04
    
    
    Set sapH = Nothing
End Sub
' ============================================
' ============================================


' THEN THIS!
' ============================================
' ============================================
Public Sub putPricesInVisibleTransferItems()
    
    Dim ti As Worksheet
    Set ti = ThisWorkbook.Sheets("TransferItems")
    
    Dim outTp04 As Worksheet
    Set outTp04 = ThisWorkbook.Sheets("OUT_TP04")
    
    Dim pltCodeRef As Range
    Set pltCodeRef = ThisWorkbook.Sheets("register").Range("G200").offset(1, 0)
    
    
    Dim initialPrice As Range, preSerialPrice As Range
    
    Dim divStr As String, coforStr As String
    
    Dim tir As Range
    Set tir = ti.Cells(2, 1)
    Do
        
        If Not tir.EntireRow.Hidden Then
            
            divStr = getDivision(ti.Cells(tir.row, EP.E_OUT_TR_PLT), pltCodeRef)
            divStr = "5" & divStr & "0"
            
            coforStr = Split(CStr(ti.Cells(tir.row, EP.E_OUT_TR_COFOR).Value), "-")(0)
            
            
            Set initialPrice = findArticleWithoutIndex(coforStr, divStr, ti.Cells(tir.row, EP.E_OUT_TR_ARTICLE), outTp04)
            Set preSerialPrice = findArticleWithIndex(coforStr, divStr, ti.Cells(tir.row, EP.E_OUT_TR_ARTICLE), ti.Cells(tir.row, EP.E_OUT_TR_INDICE), outTp04)
            
            
            If Not initialPrice Is Nothing Then ti.Cells(tir.row, E_OUT_TR_SYNTHESIS + 1).Value = initialPrice.Value
            If Not preSerialPrice Is Nothing Then ti.Cells(tir.row, E_OUT_TR_SYNTHESIS + 2).Value = initialPrice.Value
            
        End If
        
        Set tir = tir.offset(1, 0)
    Loop Until Trim(tir.Value) = ""
    
End Sub
' ============================================
' ============================================

Private Function findArticleWithoutIndex(cofor1 As String, divString As String, a1 As Range, lookHere As Worksheet) As Range
    Set findArticleWithoutIndex = Nothing
    
    
    Dim ra As Range, rdiv As Range, rc As Range
    Set ra = lookHere.Cells(2, EP.E_FROM_SQ01_QUASI_TP04_ARTICLE)
    Set rdiv = lookHere.Cells(2, EP.E_FROM_SQ01_QUASI_TP04_DIV)
    Set rc = lookHere.Cells(2, EP.E_FROM_SQ01_QUASI_TP04_FOUR)
    
    Do
    
    
        If Trim(ra.Value) = Trim(a1.Value) Then
            If Trim(rdiv.Value) = Trim(divString) Then
                
                Set findArticleWithoutIndex = lookHere.Cells(ra.row, EP.E_FROM_SQ01_QUASI_TP04_II_RATE_CALCED_SUM)
                
                
                If Trim(rc.Value) = Trim(cofor1) Then
                    Set findArticleWithoutIndex = lookHere.Cells(ra.row, EP.E_FROM_SQ01_QUASI_TP04_II_RATE_CALCED_SUM)
                    Exit Do
                End If
            End If
        End If
    
    
        Set ra = ra.offset(1, 0)
        Set rdiv = rdiv.offset(1, 0)
        Set rc = rc.offset(1, 0)
    
    Loop Until Trim(ra.Value) = ""
    
End Function

Private Function findArticleWithIndex(cofor1 As String, divString As String, a1 As Range, i1 As Range, lookHere As Worksheet) As Range
    Set findArticleWithIndex = Nothing
    
    Dim ra As Range, rdiv As Range, rc As Range
    Set ra = lookHere.Cells(2, EP.E_FROM_SQ01_QUASI_TP04_ARTICLE)
    Set rdiv = lookHere.Cells(2, EP.E_FROM_SQ01_QUASI_TP04_DIV)
    Set rc = lookHere.Cells(2, EP.E_FROM_SQ01_QUASI_TP04_FOUR)
    
    Dim ai As String
    ai = Trim(a1.Value) & "-0" & Trim(i1.Value)
    
    Do
    
    
        If Trim(ra.Value) = Trim(ai) Then
            If Trim(rdiv.Value) = Trim(divString) Then
                
                Set findArticleWithIndex = lookHere.Cells(ra.row, EP.E_FROM_SQ01_QUASI_TP04_II_RATE_CALCED_SUM)
                
                If Trim(rc.Value) = Trim(cofor1) Then
                    Set findArticleWithIndex = lookHere.Cells(ra.row, EP.E_FROM_SQ01_QUASI_TP04_II_RATE_CALCED_SUM)
                    Exit Do
                End If
            End If
        End If
    
    
        Set ra = ra.offset(1, 0)
        Set rdiv = rdiv.offset(1, 0)
        Set rc = rc.offset(1, 0)
    
    Loop Until Trim(ra.Value) = ""
End Function

Private Function getDivision(plt As Range, rf As Range) As String
    
    getDivision = ""
    
    Set rf = ThisWorkbook.Sheets("register").Range("G200").offset(1, 0)
    
    Do
        If UCase(plt.Value) = UCase(rf.offset(0, 1).Value) Then
            getDivision = CStr(rf.Value)
            Exit Do
        End If
        
        Set rf = rf.offset(1, 0)
    Loop Until Trim(rf.Value) = ""
End Function





 
 ' ============================================
Private Sub removeEmptyRows(sh As Worksheet)
    
    Dim r As Range
    Set r = sh.Cells(2, 2)
    
    recurForRemoveRow sh, r.End(xlDown)
    
    sh.Cells(2, 2).Select
End Sub
' ============================================


Private Sub recurForRemoveRow(sh As Worksheet, ir As Range)

    If ir.End(xlDown).row < 1048576 Then
    
        Set ir = ir.offset(1, 0)
        ir.Select
    
        If Trim(ir.Value) = "" And Trim(ir.offset(0, 1).Value) = "" Then
            ir.EntireRow.Delete xlShiftUp
            
            Set ir = sh.Cells(2, 2)
            recurForRemoveRow sh, ir.End(xlDown)
        End If
    End If
End Sub

Public Sub testForEmptyRows() ' OK test success
    
    'outFromTp04FillLabels ActiveSheet ' OK
    'removeEmptyRows ActiveSheet ' OK
    ' extend_and_std_tp04_data ActiveSheet ' OK!
    
End Sub

Private Sub extend_and_std_tp04_data(sh As Worksheet)


    ' rate based on string currency
    Dim r As Range, wrkRng As Range
    Set r = sh.Range("B2")
    
    Dim currFomurla As String, currString As String
    Dim unitFormula As String, unitString As String
    
    
    Set wrkRng = sh.Range(r, r.End(xlDown))
    Set wrkRng = wrkRng.offset(0, E_FROM_SQ01_QUASI_TP04_II_RATE_TO_EUR - 2)
    
    wrkRng.FormulaR1C1Local = ThisWorkbook.Sheets("register").Range("D399").FormulaR1C1Local
    wrkRng.Value = wrkRng.Value
    
    
    Set wrkRng = sh.Range(r, r.End(xlDown))
    Set wrkRng = wrkRng.offset(0, EP.E_FROM_SQ01_QUASI_TP04_II_RATE_UNIT_VALUE - 2)
    
    wrkRng.FormulaR1C1Local = ThisWorkbook.Sheets("register").Range("J398").FormulaR1C1Local
    wrkRng.Value = wrkRng.Value
    
    
    ' final sum
    Set wrkRng = sh.Range(r, r.End(xlDown))
    Set wrkRng = wrkRng.offset(0, EP.E_FROM_SQ01_QUASI_TP04_II_RATE_CALCED_SUM - 2)
    
    ' HERE SUPER STATIC!
    ' =============================================
    wrkRng.FormulaR1C1 = "=RC[-5]/RC[-1]/RC[-2]"
    wrkRng.Value = wrkRng.Value
    ' =============================================
    
    ' this is too long
    'Do
    '    currFomurla = ThisWorkbook.Sheets("register").Range("D399").Formula
    '    currString = Chr(34) & CStr(sh.Cells(r.row, EP.E_FROM_SQ01_QUASI_TP04_CURRENCY).Value) & Chr(34)
    '    Set wrkRng = sh.Cells(r.row, EP.E_FROM_SQ01_QUASI_TP04_II_RATE_TO_EUR)
    '    wrkRng.Formula = Replace(currFomurla, "X", currString)
    '    wrkRng.Value = wrkRng.Value
    '
    '
    '    Set r = r.offset(1, 0)
    'Loop Until Trim(r.Value) = ""
    
    ' from ep.replaceDotWDecPntrMod
    ' EP.replaceDotWithDecimalPointer (str)
    
    MsgBox "std on output from sq01 ready!"
End Sub



Private Sub outFromTp04FillLabels(sh As Worksheet)
    With sh
        .Cells(1, E_FROM_SQ01_QUASI_TP04_DOMAIN).Value = "DOMAIN"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_ARTICLE).Value = "ARTICLE"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_RU).Value = "RU"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_DOC_ACHAT).Value = "DOC_ACHAT"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_POSTE).Value = "POSTE"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_TYPE).Value = "TYPE"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_DIV).Value = "DIV"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_FOUR).Value = "FOUR"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_DATE_DEBUT).Value = "DATE_DEBUT"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_DATE_FIN).Value = "DATE_FIN"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_UNITE).Value = "UNITE"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_SUM).Value = "SUM"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_UNITE2_EMPTY).Value = "EMPTY"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_CURRENCY).Value = "CURRENCY"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_II_RATE_UNIT_VALUE).Value = "UNIT"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_II_RATE_TO_EUR).Value = "RATE_TO_EUR"
        .Cells(1, E_FROM_SQ01_QUASI_TP04_II_RATE_CALCED_SUM).Value = "__SUM2"
    End With
End Sub




