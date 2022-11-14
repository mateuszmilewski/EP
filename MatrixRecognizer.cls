VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatrixRecognizer"
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



' this function go through all sheet and return collection of
' potential worksheets
Public Function processAndRetrieve(sh As Worksheet) As Collection



    ' this sub should in smart way recognise if the data having rectangular info stored
    ' rectangular means that it look like some kinda table
    
    Set processAndRetrieve = Nothing
    
    Dim tmpColl As New Collection
    Dim m As IMatrix
    
    
    Dim x As Variant, y As Variant
    
    
    'ThisWorkbook.Activate
    'sh.Activate
    
    '  Cells(4, 6).Select
    
    Dim r As Range, ir As Range, line As Long, c As Range, prevRng As Range
    Set r = sh.UsedRange
    
    Debug.Print "used range: " & r.Address & " " & r.columnS.Count & " " & r.Rows.Count
    
    If r.Count > 0 Then
    
    
        ' Set m = New Matrix
    
        For Each ir In r.Rows
        
            line = ir.row
            ' Debug.Print ir.Address ' all row
            
            Debug.Print Application.WorksheetFunction.CountBlank(ir)
            Debug.Print Application.WorksheetFunction.CountA(ir)
            
            Dim ratio As Double
            ratio = CDbl(Application.WorksheetFunction.CountA(ir) / (1# * Application.WorksheetFunction.CountBlank(ir)))
            
            Debug.Print "ratio avail: " & ratio
            
            If CLng(Application.WorksheetFunction.CountBlank(ir)) = CLng(r.columnS.Count) Then
                ' all line empty
            ElseIf Application.WorksheetFunction.CountA(ir) = 1 Then
                
                ' potential title
                Debug.Print " potential title "
                
            Else
            
                Set m = New Matrix
            
                For Each c In ir.columnS
                
                    Debug.Print "c: " & line & " " & c.Address
                    If m.isEmpty() Then
                        If c.Value <> "" Then
                            m.initialFill c
                        End If
                    Else

                    End If
                    
                    Set prevRng = c
                Next c
            End If
        Next ir
    End If
End Function

