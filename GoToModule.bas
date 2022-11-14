Attribute VB_Name = "GoToModule"
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


Public Sub goToPertu()
    ActiveWindow.ScrollRow = 9
    

    With ThisWorkbook.Sheets("DASHBOARD").Shapes.Range(Array("PLT 1"))
        .Left = 1
        .Top = 150 ' heurystycznie
    End With
End Sub

Public Sub gotoPerfoRetard()
    ActiveWindow.ScrollRow = 35
    

    With ThisWorkbook.Sheets("DASHBOARD").Shapes.Range(Array("PLT 1"))
        .Left = 1
        .Top = 540 ' heurystycznie
    End With
End Sub

Public Sub gotoPerfoManq()
    ActiveWindow.ScrollRow = 75
    

    With ThisWorkbook.Sheets("DASHBOARD").Shapes.Range(Array("PLT 1"))
        .Left = 1
        .Top = 1120 ' heurystycznie
    End With
End Sub

Public Sub gotoTransfer()
    ActiveWindow.ScrollRow = 105
    

    With ThisWorkbook.Sheets("DASHBOARD").Shapes.Range(Array("PLT 1"))
        .Left = 1
        .Top = 1570 ' heurystycznie
    End With
End Sub

