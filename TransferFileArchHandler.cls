VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferFileArchHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Private sh As Worksheet
Public listColumns As TransferFileArchColumnsHandler
' Private kolekcjaDanych As Collection
Private ajtem As TransferFileArchItemHandler



Public Sub takeDataFromListingAndBindEverything(ByRef kntd As TransferFileKnottedHandler)


    ' at the end taking knotted data as kntd is overworking a bit,
    ' but at the end you see DIRECTLY
    ' also in kntd you already haf dic from the tb data
    ' so you have HERE in this sub all scope organised with SOME quality!


    ' now this sub is wierd cause I want to have bind the from this object
    ' and general characteristics object from table items
    ' the wierdo stuff is becuase this binding going here: inside the component
    ' of the knotted object which should by intuition "outside"
    ' anyway fair enough becuase all 3 object are paralel in layers
    ' just some play around from my side:
    '
    'Public tb As TransferFileTableHandler
    'Public list As TransferFileListHandler
    '
    'Public knottedData As TransferFileKnottedHandler
    '
    ' how it is going on in doc info handler sub object few steps behind this one:
    ' starting from filling tb - for major data in Table worksheet inside transfer file
    ' later taking only dic from the table data and passing it as param for knotted data
    ' from this point we can "forget" about  tb object - because all important content
    ' is already inside knotted object
    ' now we starting taking data from listing - and this is a bit much more data
    ' after this the idea to not doubling the loop in this amount
    ' ill make 2 action in one during going through loop
    ' filling the list from listing and also grouping the data by content from dic
    ' when I'm writing it sounds reasonable- will se how it goes ...
    '
    ' br forrest from the past
    
    
    If listColumns.checkIfAllColumnsAreIn() Then
            
        Dim startingRow As Long, tmp As Range
        startingRow = 3
        
        Set tmp = sh.Range("A3")
        
        Do
            
            Set ajtem = New TransferFileArchItemHandler
            With ajtem
                .dateColumn1 = CStr(sh.Cells(tmp.row, listColumns.dateColumn1.Column).Value)
                .pn = CStr(sh.Cells(tmp.row, listColumns.pn.Column).Value)
                .description = CStr(sh.Cells(tmp.row, listColumns.description.Column).Value)
                .sousProjet = CStr(sh.Cells(tmp.row, listColumns.sousProjet.Column).Value)
                .ru = CStr(sh.Cells(tmp.row, listColumns.ru.Column).Value)
                .date1 = CStr(sh.Cells(tmp.row, listColumns.date1.Column).Value)
                .cmnt1 = CStr(sh.Cells(tmp.row, listColumns.cmnt1.Column).Value)
                .ba = CStr(sh.Cells(tmp.row, listColumns.ba.Column).Value)
                
                .row = tmp.row
            End With
            
            ' going a lot here - check it (not only adding data, but also binding with tb data)
            ' ------------------------------------------------------------------------------------
            kntd.addNewKnottedItemFromArchive ajtem
            ' ------------------------------------------------------------------------------------
            
            ' Debug.Print tmp.Address & " " & startingRow
            
            startingRow = startingRow + 1
            Set tmp = tmp.offset(1, 0)
        Loop Until Trim(tmp.Value) = ""
        
        
        Debug.Print " startingRow:  " & startingRow
    End If
End Sub


Public Sub findAllStaticColumns(sh1 As Worksheet)


    Set sh = sh1

    If Not sh Is Nothing Then
        
        Dim reg As Worksheet
        Set reg = ThisWorkbook.Sheets("register")
        
        With listColumns
            ' Public Function lookForColumn(srcSh As Worksheet, _
            '    labelRow As Integer, _
            '    namePattern As String, _
            '    wildcardRef As Range, _
            '    Optional col1 As Integer, Optional col2 As Integer, _
            '    Optional bgColor As Long, Optional fontColor As Long) As Range
            Set .dateColumn1 = lookForColumn(sh, 2, "*", sh1.Range("A2"))
            Set .pn = lookForColumn(sh, 2, "*", sh1.Range("B2"))
            Set .description = lookForColumn(sh, 2, "*", sh1.Range("C2"))
            Set .sousProjet = lookForColumn(sh, 2, "*", sh1.Range("D2"))
            Set .ru = lookForColumn(sh, 2, "*", sh1.Range("E2"))
            Set .date1 = lookForColumn(sh, 2, "*", sh1.Range("F2"))
            Set .cmnt1 = lookForColumn(sh, 2, "*", sh1.Range("G2"))
            ' column H Metier - ignore
            Set .ba = lookForColumn(sh, 2, "*", sh1.Range("I2"))
            
            

        End With
    End If
End Sub

Private Sub Class_Initialize()
    Set sh = Nothing
    Set listColumns = New TransferFileArchColumnsHandler
    Set ajtem = Nothing
End Sub

Private Sub Class_Terminate()
    Set sh = Nothing
    Set listColumns = Nothing
    Set ajtem = Nothing
End Sub

