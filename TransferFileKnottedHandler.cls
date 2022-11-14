VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferFileKnottedHandler"
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


' this data taking only info from table
Private dic As Dictionary

' uber Dictionary
'Private uD As Dictionary

' but first I will try with the collection
Private ik As TransferFileKnottedItemHandler
Private collectionOfKnottedItems As Collection



Public Sub showCollection()
    
    For Each ik In collectionOfKnottedItems
        
        Debug.Print ik.data.row & " " & ik.data.pn & " " & ik.data.des & " " & ik.tb.MDL
    Next ik
    
End Sub

Public Sub putCollectionOfDataIntoWorksheet(ByRef rRef As Range)


    ' becuase ik as TransferFileKnottedItemHandler
    ' have 2 layers of data and that is why we are having this extra if statemenets
    ' with checkings on Nothing elemenets to prevent running failure
    
    Dim sh As Worksheet, wiersz As Long, dataCounter As Long
    Set sh = ThisWorkbook.Sheets("TransferItems")
    
    wiersz = 2
    wiersz = findFirstEmpty(sh)
    
    dataCounter = 0 - wiersz
    
    Debug.Print "collectionOfKnottedItems.Count: " & collectionOfKnottedItems.Count
    
    For Each ik In collectionOfKnottedItems
        With sh
        
        
            If Not ik Is Nothing Then
        
                If UCase(ik.data.fltr) = "X" Then
                    
                    ' e out tr plant name
                    'E_OUT_TR_PHOTO_DATE
                    'E_OUT_TR_PROJECT
                    'E_OUT_TR_SOUS_PROJET
                    'E_OUT_TR_MDL
                    'E_OUT_TR_AMC
                    'E_OUT_TR_DIFF1
                    'E_OUT_TR_DIFF2
                    'E_OUT_TR_ROW
                    ' col s ' new - auto date from sigapp
                    ' col s ' cw; new
                    ' ool s month ' new
                    'E_OUT_TR_ARTICLE
                    'E_OUT_TR_INDICE
                    'E_OUT_TR_SGR_FAB
                    'E_OUT_TR_DESC
                    'E_OUT_TR_DOMAIN
                    'E_OUT_TR_DECISION
                    'E_OUT_TR_SOUS_PROJET_2
                    'E_OUT_TR_SYNTHESIS
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_PLT).Value = CStr(rRef.offset(0, -1).Value)
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_PHOTO_DATE).Value = CDate(ik.photoDate)
                    
                    If Not ik.tb Is Nothing Then
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_PROJECT).Value = ik.tb.str
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_SOUS_PROJET).Value = ik.tb.sousProjet
                    
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_MDL).Value = CDate(ik.tb.MDL)
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_AMC).Value = CDate(ik.tb.AMC)
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_DIFF1).Value = CLng(ik.tb.MDL - ik.photoDate)
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_DIFF2).Value = CLng(ik.tb.AMC - ik.photoDate)
                    
                        If CLng(ik.tb.MDL - ik.photoDate) > 10 And CLng(ik.tb.MDL - ik.photoDate) < 30 Then
                            .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_SCOPE).Value = 1
                        Else
                            .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_SCOPE).Value = 0
                        End If
                        
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_CW).Formula = assignIsoCWFormula(EP.E_OUT_TR.E_OUT_TR_CW)
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_MONTH).Value = assignMonthFromFriday(CDate(ik.tb.AMC))
                    End If
                    
                    
                    ' column S section
                    ' ----------------------------------------------------------------------------------------------------------
                    If IsDate(ik.data.columnS) Then
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_COL_S_DATE).Value = CDate(ik.data.columnS)
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_COL_S_MONTH).Value = assignMonthFromFriday(CDate(ik.data.columnS))
                    End If
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_COL_S_CW).Formula = assignIsoCWFormula(EP.E_OUT_TR.E_OUT_TR_COL_S_CW)
                    
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_COL_S_YEAR).Value = assignYearFOrmula(EP.E_OUT_TR.E_OUT_TR_COL_S_DATE)
                    ' ----------------------------------------------------------------------------------------------------------
                    
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_ROW).Value = ik.data.row
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_ARTICLE).Value = ik.data.pn
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_INDICE).Value = ik.data.indice
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_SGR_FAB).Value = ik.data.sgrFab
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_DESC).Value = ik.data.des
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_DOMAIN).Value = ik.data.domain
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_DECISION).Value = ik.data.soreppDecision
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_SOUS_PROJET_2).Value = ik.data.sousProjetColumn
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_SYNTHESIS).Value = ik.data.synteza
                    
                    
                    ' NEW BA
                    ' ----------------------------------------------------------------------------------------------------------
                    If IsNumeric(ik.data.ba) Then
                        .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_SYNTHESIS + 3).Value = CLng(ik.data.ba)
                    End If
                    ' ----------------------------------------------------------------------------------------------------------
                    
                    ' NEW COFOR
                    ' ----------------------------------------------------------------------------------------------------------
                    .Cells(wiersz, EP.E_OUT_TR.E_OUT_TR_COFOR).Value = CStr(ik.data.cofor)
                    ' ----------------------------------------------------------------------------------------------------------
                    
                    
                    wiersz = wiersz + 1
                End If
            End If
        End With
        
        
    Next ik
    
    
    dataCounter = dataCounter + wiersz
    
    Debug.Print "dataCounter: " & dataCounter
End Sub



Private Function assignMonthFromFriday(d1 As Date) As Integer
    
    assignMonthFromFriday = -1
    
    Dim a As Integer
    a = Weekday(d1, vbMonday)
    
    Dim d2 As Date
    d2 = d1 - (a - 1) + 4
    
    ' Debug.Print d2
    
    assignMonthFromFriday = Month(d2)
    
End Function


Private Function assignYearFOrmula(e1 As E_OUT_TR) As String
    
    assignYearFOrmula = ""
    
    Dim str As String
    ' this is template for year formula YEAR(X)
    str = ThisWorkbook.Sheets("register").Range("A201").Formula
    
    If e1 = E_OUT_TR_COL_S_DATE Then
        
        str = Replace(str, "X", "RC[-3]")
    End If
    
    assignYearFOrmula = str
    
End Function

Private Function assignIsoCWFormula(e1 As E_OUT_TR) As String
    
    assignIsoCWFormula = ""
    
    Dim str As String
    str = ThisWorkbook.Sheets("register").Range("A200").Formula
    
    ' Debug.Print str
    ' EP.E_OUT_TR.E_OUT_TR_CW
    If e1 = E_OUT_TR_CW Then
        str = Replace(str, "X", "RC[-4]")
    ElseIf e1 = E_OUT_TR_COL_S_CW Then
        str = Replace(str, "X", "RC[-1]")
    End If
    
    assignIsoCWFormula = str
End Function


Private Function findFirstEmpty(sh1 As Worksheet) As Long
    
    findFirstEmpty = 2
    
    Dim r As Range
    Set r = sh1.Range("A1048576").End(xlUp).offset(1, 0)
    
    findFirstEmpty = CLng(r.row)
    
End Function


Public Sub addNewKnottedItemFromArchive(a As TransferFileArchItemHandler)

    Dim ki As Variant, tbi As TransferFileTableItemHandler
        
    Set ik = Nothing
    
    ' this loop covers resp for matching def of the wave project for each line in the raw data from
    ' the transfer list
    ' -------------------------------------------------------------
    For Each ki In dic.Keys
        Set tbi = dic(ki)
        If CStr(tbi.sousProjet) = CStr(a.sousProjet) Then
    
            Set ik = New TransferFileKnottedItemHandler
            Set ik.data2 = a ' data2 stands for var that stores content from Archive worksheet
            Set ik.tb = tbi
    
            Exit For
        End If
    Next
    ' -------------------------------------------------------------
    
    
    ' it means that there no match on data
    ' in previous version of the tool
    ' we have here quite a issue becuase all items was (not necessarly good)
    If ik Is Nothing Then
        
        Set ik = New TransferFileKnottedItemHandler
        Set ik.data = a
        ' imho this is better solution
        ' nothing only on stuf that you are really missing
        Set ik.tb = Nothing
    End If
    
    ' now we are adding to collection much more data
    collectionOfKnottedItems.Add ik
    
End Sub


Public Sub addNewKnottedItem(i As TransferFileListItemHandler)

    Dim ki As Variant, tbi As TransferFileTableItemHandler
        
    Set ik = Nothing
    
    ' this loop covers resp for matching def of the wave project for each line in the raw data from
    ' the transfer list
    ' -------------------------------------------------------------
    For Each ki In dic.Keys
        Set tbi = dic(ki)
        If CStr(tbi.sousProjet) = CStr(i.sousProjetColumn) Then
    
            Set ik = New TransferFileKnottedItemHandler
            Set ik.data = i
            Set ik.tb = tbi
    
            Exit For
        End If
    Next
    ' -------------------------------------------------------------
    
    
    ' it means that there no match on data
    ' in previous version of the tool
    ' we have here quite a issue becuase all items was (not necessarly good)
    If ik Is Nothing Then
        
        Set ik = New TransferFileKnottedItemHandler
        Set ik.data = i
        ' imho this is better solution
        ' nothing only on stuf that you are really missing
        Set ik.tb = Nothing
    End If
    
    ' now we are adding to collection much more data
    collectionOfKnottedItems.Add ik
    
End Sub


Public Sub getTheDicFromTable(d1 As Dictionary)
    
    Set dic = d1
End Sub

Private Sub Class_Initialize()
    Set dic = Nothing
    Set collectionOfKnottedItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set dic = Nothing
    Set collectionOfKnottedItems = Nothing
End Sub
