Attribute VB_Name = "MainModule"
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





' -------------------------------


Public Sub takeTransferFilesFromDocInfo()
Attribute takeTransferFilesFromDocInfo.VB_ProcData.VB_Invoke_Func = "T\n14"


    Application.ScreenUpdating = False
    
    Dim dih As New DocInfoHandler
    dih.closeAllPjWorkbooks
    ' dih.clearAllOldSheetsFromPrevDownload
    dih.collectTransferLists
    
    
    ' go back to main
    ThisWorkbook.Sheets("MAIN").Activate
    
    Application.ScreenUpdating = True
End Sub



' -------------------------------





Public Sub takeDataFromDocInfo()


    Application.ScreenUpdating = False
    
    Dim dih As New DocInfoHandler
    dih.closeAllPjWorkbooks
    dih.clearAllOldSheetsFromPrevDownload
    dih.collectWorkbooksFromList
    
    showModelessPanel
    
    
    
    ' go back to main
    ThisWorkbook.Sheets("MAIN").Activate
    
    Application.ScreenUpdating = True
End Sub


Public Sub showQuickPanel(ictrl As IRibbonControl)
    showModelessPanel
End Sub

Public Sub showModelessPanel()
Attribute showModelessPanel.VB_ProcData.VB_Invoke_Func = "Q\n14"
    ControlPanel.show vbModeless
End Sub




Public Sub sythesizeData()
    
    ' thsis sub taking all bits and pieces from this workbook and trying to put into one pager
    
    ' new one pager sheet
    Dim newOnePager As Worksheet
    Dim c As New Common
    
    Set newOnePager = ThisWorkbook.Sheets.Add
    newOnePager.name = c.tryToRenameIt(newOnePager, "ONE_PAGER_", "")
    
    c.colorSheetTab newOnePager, xlThemeColorAccent6
    
    
    Set newOnePager = Nothing
    Set c = Nothing
End Sub



