Attribute VB_Name = "AddRawDataModule"
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


Public Sub addRawData(ictrl As IRibbonControl)
    addToRawDataNewData
End Sub


Public Sub addToRawDataNewData()
Attribute addToRawDataNewData.VB_ProcData.VB_Invoke_Func = "W\n14"
    
    DefineRange.RefEditContent.Value = ""
    DefineRange.RefEditLabels.Value = ""
    
    DefineRange.ComboBox1.Clear
    ' BIG NOK - to nie jest on nawet
    ' DefineRange.ComboBox1.ControlSource = Array("P1", "P2", "R1", "R2", "M1", "M2", "T1", "T2")
    With DefineRange.ComboBox1
        Dim x As Variant, xx As Variant
        x = Array("P1", "P2", "R1", "R2", "R3", "M1", "M2", "M3", "T1", "T2", "T3", "S1")
        For Each xx In x
            .addItem xx
        Next xx
        
        
    End With
    
    DefineRange.show
End Sub
