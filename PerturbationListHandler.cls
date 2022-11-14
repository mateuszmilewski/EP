VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerturbationListHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


' PerturbationListHandler
' this class going for normalisation for original list from previous KPI file

Private orgPertuSh As Worksheet
Private e As E_PERTU_LIST

Private item As PerturbationListItem
Private collectionOfItems As Collection

Private normPertuSh As Worksheet




Private Sub Class_Initialize()
    Set orgPertuSh = Nothing
    Set normPertuSh = Nothing
    
    
    tryToCatchWorksheets
End Sub

Private Sub tryToCatchWorksheets()
    
    On Error Resume Next
    Set orgPertuSh = ThisWorkbook.Sheets(EP.G_PERTURBATION_SH_NM_PERTURBATION)
    
    On Error Resume Next
    Set normPertuSh = ThisWorkbook.Sheets("NormPertuList_backupCopy")
    
End Sub

Private Sub Class_Terminate()
    Set orgPertuSh = Nothing
    Set normPertuSh = Nothing
End Sub


Private Function parseStringToDate(dateLike As String) As Date
    
    If IsDate(dateLike) Then
        parseStringToDate = CDate(dateLike)
    Else
        
        ' this example is a little bit more sophi... well...
        ' for now treat as error
        parseStringToDate = CDate(0)
        
    End If
End Function


Private Function parseAccounting(oui_q As String) As Boolean

    parseAccounting = False

    If UCase("oui") = UCase(oui_q) Then parseAccounting = True
    If UCase("yes") = UCase(oui_q) Then parseAccounting = True
    If UCase("tak") = UCase(oui_q) Then parseAccounting = True
    If UCase("da") = UCase(oui_q) Then parseAccounting = True
    If UCase("true") = UCase(oui_q) Then parseAccounting = True
    If UCase("vrai") = UCase(oui_q) Then parseAccounting = True
    If UCase("vraie") = UCase(oui_q) Then parseAccounting = True
    
End Function


Public Sub takeDataOneByOneAndParse()

    Dim listItem As PerturbationListItem
    Set collectionOfItems = Nothing
    Set collectionOfItems = New Collection
    

    Dim ir As Range, lr As Range, startingPointInOrgPerturbationList As Integer
    startingPointInOrgPerturbationList = 4
    Set ir = orgPertuSh.Cells(startingPointInOrgPerturbationList, EP.E_PERTU_LIST_SITE)
    
    ' B1048576
    Set lr = orgPertuSh.Cells(1048576, EP.E_PERTU_LIST_SITE).End(xlUp)
    
    Debug.Print "ir: " & ir.Address & ", lr: " & lr.Address
    
    Dim i As Range
    For Each i In orgPertuSh.Range(ir, lr)
    
        ' Debug.Print i.Address & " " & i.Value & " " & i.MergeArea.Address & " " & i.MergeArea.Count
        
        ' test
        'If i.MergeArea.Count > 1 Then
        '    If i.Value = "" Then
        '        Debug.Print "i: " & i.Address & " i.MergeArea: " & i.MergeArea.Address & " " & _
        '           i.MergeArea.Cells(1, 1).Value & " " & i.MergeArea.Cells(1, 1).Address
        '    End If
        'End If
        
        Set listItem = Nothing
        Set listItem = New PerturbationListItem
        
        'Debug.Print "current adr: " & i.Address
        'Debug.Assert i.Address <> "$B$266"
        
        With listItem
            .site = generalParseForListItem(i, EP.E_PERTU_LIST_SITE)
            
            ' date
            ' -------------------------------------------------------------------------
            ' as date
            .strDate = generalParseForListItem(i, EP.E_PERTU_LIST.E_PERTU_LIST_DATE)
            .d = CDate(parseStringToDate(.strDate))
            .strMonth = makeStrMonth(.d)
            ' -------------------------------------------------------------------------
            
            .veh = generalParseForListItem(i, EP.E_PERTU_LIST.E_PERTU_LIST_VEH)
            .dan = generalParseForListItem(i, EP.E_PERTU_LIST.E_PERTU_LIST_DAN)
            .perturbation = generalParseForListItem(i, E_PERTU_LIST_PERTURBATION)
            
            ' acc ---------------------------------------------------------------------
            .accounting = generalParseForListItem(i, E_PERTU_LIST_ACCOUNTING)
            ' as boolean
            .bacc = CBool(parseAccounting(.accounting))
            ' -------------------------------------------------------------------------
            
            .causePerturbation = generalParseForListItem(i, E_PERTU_LIST_CAUSE)
            .responsibility = generalParseForListItem(i, E_PERTU_LIST_RESP)
            .Disturbance_origin_BOM = generalParseForListItem(i, E_PERTU_LIST_ORIGIN_BOM)
            .str_Nr_of_affected_cars = generalParseForListItem(i, E_PERTU_LIST_NUM_OF_AFFECTED_CARS)
            .int_Nr_of_affected_cars = tryToParseToNum(.str_Nr_of_affected_cars)
            
            ' 5 WHY
            ' -------------------------------------------------------------------------
            .fivewhy = generalParseForListItem(i, E_PERTU_LIST_5WHY)
            .fivewhyDescription = generalParseForListItem(i, E_PERTU_LIST_5WHY)
            
            .b5w = CBool(.fivewhy <> "")
            ' -------------------------------------------------------------------------
            
            
            
            .comment = ""
            
            
            
            
            
        End With
        
        listItem.calculateQuality
        collectionOfItems.Add listItem
        
        
    
    Next i
    
End Sub

Public Sub replaceDataInNormPertuList()


    ' clear the norm list
    normPertuSh.UsedRange.Clear
    
    remakeLabels normPertuSh
    
    If Not collectionOfItems Is Nothing Then
        
        If collectionOfItems.Count > 0 Then
        
        
            Dim orng As Range, eo As EP.E_PERTU_OUT_LIST, wiersz As Long
            Set orng = normPertuSh.Cells(2, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_SITE)
            wiersz = 2
            
            For Each item In collectionOfItems
            
                'E_PERTU_OUT_LIST_SITE = 1
                'E_PERTU_OUT_LIST_STR_DATE
                'E_PERTU_OUT_LIST_DATE
                'E_PERTU_OUT_LIST_VEH
                'E_PERTU_OUT_LIST_DAN
                'E_PERTU_OUT_LIST_PERTURBATION
                'E_PERTU_OUT_LIST_ACCOUNTING
                'E_PERTU_OUT_LIST_BACC
                'E_PERTU_OUT_LIST_CAUSE
                'E_PERTU_OUT_LIST_RESP
                'E_PERTU_OUT_LIST_ORIGIN_BOM
                'E_PERTU_OUT_LIST_NUM_OF_AFFECTED_CARS
                'E_PERTU_OUT_LIST_5WHY
                'E_PERTU_OUT_LIST_B_5WHY
                'E_PERTU_OUT_LIST_Q
                ' str month
            
                With normPertuSh
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_SITE).Value = item.site
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_STR_DATE).Value = item.strDate
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_DATE).Value = item.d
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_VEH).Value = item.veh
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_DAN).Value = item.dan
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_PERTURBATION).Value = item.perturbation
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_ACCOUNTING).Value = item.accounting
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_BACC).Value = (-1 * Int(item.bacc)) ' idk why but int(true) == -1
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_CAUSE).Value = item.causePerturbation
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_RESP).Value = item.responsibility
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_ORIGIN_BOM).Value = item.Disturbance_origin_BOM
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_NUM_OF_AFFECTED_CARS).Value = item.int_Nr_of_affected_cars
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_5WHY).Value = item.fivewhy
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_Q).Value = item.quality
                    .Cells(wiersz, EP.E_PERTU_OUT_LIST.E_PERTU_OUT_LIST_STR_MONTH).Value = CStr(item.strMonth)
                End With
                
                wiersz = wiersz + 1
                
            Next item
            
            With columnS("A:O")
                .HorizontalAlignment = xlGeneral
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            
            MsgBox "Ready!", vbInformation
        Else
            MsgBox "the collection is empty, no perturbations found!", vbInformation
        End If
    Else
        MsgBox "there is not data!", vbCritical
        
    End If
End Sub

Private Sub remakeLabels(osh As Worksheet)
    'E_PERTU_OUT_LIST_SITE = 1
    'E_PERTU_OUT_LIST_STR_DATE
    'E_PERTU_OUT_LIST_DATE
    'E_PERTU_OUT_LIST_VEH
    'E_PERTU_OUT_LIST_DAN
    'E_PERTU_OUT_LIST_PERTURBATION
    'E_PERTU_OUT_LIST_ACCOUNTING
    'E_PERTU_OUT_LIST_BACC
    'E_PERTU_OUT_LIST_CAUSE
    'E_PERTU_OUT_LIST_RESP
    'E_PERTU_OUT_LIST_ORIGIN_BOM
    'E_PERTU_OUT_LIST_NUM_OF_AFFECTED_CARS
    'E_PERTU_OUT_LIST_5WHY
    'E_PERTU_OUT_LIST_B_5WHY
    'E_PERTU_OUT_LIST_Q
    osh.Cells(1, E_PERTU_OUT_LIST_SITE).Value = "PLT"
    osh.Cells(1, E_PERTU_OUT_LIST_STR_DATE).Value = "STR_DATE"
    osh.Cells(1, E_PERTU_OUT_LIST_DATE).Value = "D"
    osh.Cells(1, E_PERTU_OUT_LIST_VEH).Value = "VEH"
    osh.Cells(1, E_PERTU_OUT_LIST_DAN).Value = "DAN"
    osh.Cells(1, E_PERTU_OUT_LIST_PERTURBATION).Value = "PERTU"
    osh.Cells(1, E_PERTU_OUT_LIST_ACCOUNTING).Value = "ACCOUNTING"
    osh.Cells(1, E_PERTU_OUT_LIST_BACC).Value = "BOOL_ACC"
    osh.Cells(1, E_PERTU_OUT_LIST_CAUSE).Value = "CAUSE"
    osh.Cells(1, E_PERTU_OUT_LIST_RESP).Value = "RESP"
    osh.Cells(1, E_PERTU_OUT_LIST_ORIGIN_BOM).Value = "Disturbance_origin_BOM"
    osh.Cells(1, E_PERTU_OUT_LIST_NUM_OF_AFFECTED_CARS).Value = "NUM_OF_CARS"
    osh.Cells(1, E_PERTU_OUT_LIST_5WHY).Value = "FIVE-WHY"
    osh.Cells(1, E_PERTU_OUT_LIST_B_5WHY).Value = " BOOL_FIVE-WHY"
    osh.Cells(1, E_PERTU_OUT_LIST_Q).Value = "Q"
    osh.Cells(1, EP.E_PERTU_OUT_LIST_STR_MONTH).Value = "STR_MONTH"
End Sub



Private Function generalParseForListItem(i As Range, e1 As EP.E_PERTU_LIST) As Variant
    generalParseForListItem = ""
    
    
    If i.offset(0, e1 - 2).Value <> "" Then
        generalParseForListItem = CStr(i.offset(0, e1 - 2).Value)
    Else
        If i.MergeArea.Count > 1 Then
            generalParseForListItem = CStr(i.MergeArea.Cells(1, 1).offset(0, e1 - 2).Value)
            
        Else
            generalParseForListItem = ""
        End If
    End If

End Function
    
    
    
Private Function tryToParseToNum(str1 As String) As Long
    tryToParseToNum = 0
    
    If IsNumeric(str1) Then
        tryToParseToNum = CLng(str1)
    Else
        
        Dim arr As Variant, x As Variant
        arr = Split(str1, " ")
        
        
        ' risky, but why not!
        For x = LBound(arr) To UBound(arr)
            If IsNumeric(arr(x)) Then
            
                tryToParseToNum = 0
                On Error Resume Next
                tryToParseToNum = CLng(arr(x))
                Exit For
            End If
        Next x
    End If
End Function




Private Function makeStrMonth(d1 As Date) As String

    Dim yyyy As String, yy As String, mm As String
    yyyy = CStr(Year(d1))
    yy = Right(yyyy, 2)
    mm = CStr(Month(d1))
    If Len(mm) = 1 Then mm = "0" & CStr(mm)
    
    

    makeStrMonth = CStr(yy) & "-" & CStr(mm) & " (" & ThisWorkbook.Sheets("register").Range("G1").offset(Int(Month(d1)), 0).Value & ")"
End Function

    
    
