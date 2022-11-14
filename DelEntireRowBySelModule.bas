Attribute VB_Name = "DelEntireRowBySelModule"
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



Public Sub beCarefulDeleteEntireRowBySelection()
    
    Dim r As Range
    Set r = Selection
    r.SpecialCells(xlCellTypeVisible).EntireRow.Delete xlShiftUp
End Sub


Public Sub removeEmptyLines()
    
    Dim r As Range
    Set r = Range("A1")
    
    Application.Calculation = xlCalculationManual
    
    Do
        If Trim(r.Value) = "" Or Left(r.Value, 4) = "    " Then
            r.Select
            
            If Application.WorksheetFunction.CountA(Range(Selection, Cells(1048576, 1))) > 0 Then
            
                r.EntireRow.Delete xlShiftUp
                Set r = Selection
            Else
                Exit Do
            End If
        Else
            Set r = r.offset(1, 0)
        End If
        
    Loop Until r.row = 1048576
    
    
    Application.Calculation = xlCalculationAutomatic
End Sub
