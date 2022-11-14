Attribute VB_Name = "TestPowerPointModule"
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


' obsolete test
Private Sub testPowerPoint001()
    
    Dim pp1 As New pp
    pp1.connectWithSourceAndCreatNewInstance
    
    Dim dsh As Worksheet
    Set dsh = ThisWorkbook.Sheets("DASHBOARD")
    
    Dim sz As Shape
    Set sz = dsh.Shapes("ChartPertu1")
    
    
    Debug.Print sz.name
    
    
    pp1.justCopyToNewSlideAsItIs sz, "Perturbation Chart #1", 100, 50
    

    
End Sub


' obsolete test
Private Sub anotherPowerPointTestToExisting()

    Dim pp1 As New pp
    pp1.copyToExisting ThisWorkbook.Sheets("DASHBOARD").Shapes("ChartPertu1"), _
        "Perturbation Chart #1", 100, 50
    


End Sub


Public Sub testMakeNewInstanceOfTemplatedPowerPointPresentation()
    Dim pp1 As New pp
    pp1.connectWithSourceAndCreatNewInstance
    
    
    Dim pvts As Worksheet
    Set pvts = ThisWorkbook.Sheets("PIVOTS")
    
    
    pp1.copyToNewSlide pvts.Range("B1"), "slide title", 60, 10
    pp1.copyToCurrentSlide pvts.Range("B24"), 340, 470
    
    
    pvts.Activate
End Sub
