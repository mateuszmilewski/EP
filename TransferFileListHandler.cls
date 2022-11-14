VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferFileListHandler"
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


Private sh As Worksheet
Public listColumns As TransferFileListColumnsHandler
' Private kolekcjaDanych As Collection
Private ajtem As TransferFileListItemHandler



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
            
            Set ajtem = New TransferFileListItemHandler
            With ajtem
                .des = CStr(sh.Cells(tmp.row, listColumns.des.Column).Value)
                .domain = CStr(sh.Cells(tmp.row, listColumns.domain.Column).Value)
                .fltr = CStr(sh.Cells(tmp.row, listColumns.fltr.Column).Value)
                .indice = CStr(sh.Cells(tmp.row, listColumns.indice.Column).Value)
                .pn = CStr(sh.Cells(tmp.row, listColumns.pn.Column).Value)
                .row = tmp.row
                .sgrFab = CStr(sh.Cells(tmp.row, listColumns.sgrFab.Column).Value)
                .soreppDecision = CStr(sh.Cells(tmp.row, listColumns.soreppDecision.Column).Value)
                .sousProjetColumn = CStr(sh.Cells(tmp.row, listColumns.sousProjetColumn.Column).Value)
                .synteza = CStr(sh.Cells(tmp.row, listColumns.synteza.Column).Value)
                
                .ba = CStr(sh.Cells(tmp.row, listColumns.ba.Column).Value)
                .cofor = CStr(sh.Cells(tmp.row, listColumns.cofor.Column).Value)
                
                On Error Resume Next
                .columnS = CDate(sh.Cells(tmp.row, listColumns.columnS.Column).Value)
            End With
            
            ' going a lot here - check it (not only adding data, but also binding with tb data)
            ' ------------------------------------------------------------------------------------
            kntd.addNewKnottedItem ajtem
            ' ------------------------------------------------------------------------------------
            
            ' Debug.Print tmp.Address & " " & startingRow
            
            startingRow = startingRow + 1
            Set tmp = tmp.offset(1, 0)
        Loop Until Trim(tmp.Value) = ""
        
        
        Debug.Print " startingRow:  " & startingRow
    End If
End Sub


Public Sub findAllColumns(sh1 As Worksheet)


    Set sh = sh1

    If Not sh Is Nothing Then
        
        Dim reg As Worksheet
        Set reg = ThisWorkbook.Sheets("register")
        
        With listColumns
        
            Set .pn = lookForColumn(sh, 2, "*NUM*PRODU*", reg.Range("AA3"), 1, 1, RGB(204, 255, 204), RGB(0, 0, 255))
            Debug.Print .pn.Value & " " & .pn.Address
            
            Set .sgrFab = lookForColumn(sh, 2, "*SGR*FAB*", reg.Range("AA4"), 2, 2, RGB(204, 255, 204), RGB(0, 0, 255))
            Debug.Print .sgrFab.Value & " " & .sgrFab.Address
            
            Set .indice = lookForColumn(sh, 2, "*IND*", reg.Range("AA5"), 3, 3, RGB(204, 255, 204), RGB(0, 0, 255))
            Debug.Print .indice.Value & " " & .indice.Address
            
            Set .des = lookForColumn(sh, 2, "*D?SIGN*L*", reg.Range("AA6"), 8, 10, RGB(204, 255, 204), RGB(0, 0, 255))
            Debug.Print .des.Value & " " & .des.Address
            
            Set .domain = lookForColumn(sh, 2, "*DOM*SIGAPP*", reg.Range("AA7"), 8, 10, RGB(204, 255, 204), RGB(0, 0, 255))
            Debug.Print .domain.Value & " " & .domain.Address
            
            Set .sousProjetColumn = lookForColumn(sh, 2, "*PRO*SIGAPP", reg.Range("AA2"), 11, 11, RGB(204, 255, 204), RGB(0, 0, 255))
            Debug.Print .sousProjetColumn.Value & " " & .sousProjetColumn.Address
            
            
            
            Set .soreppDecision = lookForColumn(sh, 2, "*CISI*", reg.Range("AA8"), 29, 31, RGB(204, 153, 255), RGB(0, 0, 255))
            Debug.Print .soreppDecision.Value & " " & .soreppDecision.Address
            
            Set .synteza = lookForColumn(sh, 2, "*SYN*", reg.Range("AA9"), 30, 35, RGB(128, 128, 128), RGB(255, 0, 0))
            Debug.Print .synteza.Value & " " & .synteza.Address
            
            Set .fltr = lookForColumn(sh, 2, "*FILT*", reg.Range("AA10"), 100, 105, RGB(192, 192, 192), RGB(0, 0, 0))
            Debug.Print "fltr: " & .fltr.Value & " " & .fltr.Address
            
            
            Set .columnS = lookForColumn(sh, 2, "*TI*DOM*", reg.Range("AA11"), 15, 20, RGB(153, 204, 255), RGB(0, 0, 255))
            Debug.Print "columns S date: " & .columnS.Value & " " & .columnS.Address
            
            
             ' BA!
             Set .ba = lookForColumn(sh, 2, "*BA*", reg.Range("AA12"), 40, 50)
             Debug.Print "BA col: " & .ba.Value & " " & .ba.Address
             
             
             ' COFOR
             Set .cofor = lookForColumn(sh, 2, "*COD*VE*-*", reg.Range("AA13"), 20, 30, RGB(255, 255, 153), RGB(0, 0, 255))
             Debug.Print "COFOR col: " & .cofor.Value & " " & .cofor.Address
            
        End With
    End If
End Sub

Private Sub Class_Initialize()
    Set sh = Nothing
    Set listColumns = New TransferFileListColumnsHandler
    Set ajtem = Nothing
End Sub

Private Sub Class_Terminate()
    Set sh = Nothing
    Set listColumns = Nothing
    Set ajtem = Nothing
End Sub
