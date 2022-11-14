Attribute VB_Name = "ResetModule"
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


Public Sub clearOldSideSheets()
    reset
End Sub


Public Sub reset()
    ActiveWindow.Visible = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Dim dih As New DocInfoHandler
    dih.closeAllPjWorkbooks
    dih.clearAllOldSheetsFromPrevDownload
    
    
    ' clear main status
    With ThisWorkbook.Sheets("MAIN").Range("C3:C6")
        .Value = ""
        .Font.Color = RGB(0, 0, 0)
    End With
End Sub


Public Sub justPureResetOnCofiguration()
    ActiveWindow.Visible = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.Calculate
End Sub
