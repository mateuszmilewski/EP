VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocInfoHandler"
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


Private perturbationWorkbook As Workbook
Private perfoWorkbook As Workbook
Private majorTransferWorkbook As Workbook
Private stockWorkbook As Workbook


' transfer file part
Private trasferFileWorkbook As Workbook
Private tfh As TransferFileHandler


Private mainRef As Range

Public Sub collectWorkbooksFromList()


    Application.ScreenUpdating = False

    
    
    ' try to open each file
    tryToOpen (E_TOFF_PERTURBATION - 1), perturbationWorkbook
    doSomeOperationsAndCloseIt EP.E_TOFF_PERTURBATION, perturbationWorkbook
    
    tryToOpen (E_TOFF_PERFO_APPRO - 1), perfoWorkbook
    doSomeOperationsAndCloseIt EP.E_TOFF_PERFO_APPRO, perfoWorkbook
    
    tryToOpen (E_TOFF_TRANSFER_MAJOR_FILE - 1), majorTransferWorkbook
    doSomeOperationsAndCloseIt EP.E_TOFF_TRANSFER_MAJOR_FILE, majorTransferWorkbook
    
    tryToOpen (EP.E_TOFF_KPI_STOCK - 1), stockWorkbook
    doSomeOperationsAndCloseIt EP.E_TOFF_KPI_STOCK, stockWorkbook
    
    
    ' this concept will fail - because every file
    ' on doc-info has the same name
    ' so I m forced to work step by step
    'If checkMyWorkbooks() Then ' logs and boolean
    'End If
    
    
    
    
    
    
    
    Application.ScreenUpdating = True
    
End Sub

Public Sub collectTransferLists()
    Application.ScreenUpdating = False
    
    
    ' new feature
    ' all the transfer files!
    
    Dim rangeOfTransferFileUrls As Range
    Set rangeOfTransferFileUrls = defineRangeOfUrls()
    
    
    Dim ir As Range
    For Each ir In rangeOfTransferFileUrls
        
        ' -----------------------------------------------------
        
        'Debug.Print "URL: " & ir.Value
        'Debug.Print "STATUS: " & ir.offset(0, 1).Value
        'Debug.Print "RUN: " & ir.offset(0, 2).Value
        
        
        If UCase(CStr(ir.offset(0, EP.E_TFL.E_FTL_RUN - 2).Value)) = "X" Then
            tryToOpen ir.row - 3, trasferFileWorkbook
            doSomeOperationsAndCloseIt ir.row - 3, trasferFileWorkbook, ir
        End If
        
        ' -----------------------------------------------------
    Next ir
    
    
    
    Application.ScreenUpdating = True
End Sub

Private Function defineRangeOfUrls() As Range
    Set defineRangeOfUrls = Nothing
    
    Dim turl As Range
    Set turl = ThisWorkbook.Sheets("MAIN").Range("URL_T")
    
    Set turl = turl.offset(1, 0)
    Set turl = turl.Parent.Range(turl, turl.End(xlDown))
    
    Set defineRangeOfUrls = turl
End Function

Public Sub closeAllPjWorkbooks()
    
    Application.DisplayAlerts = False
    
    Dim w As Workbook
    For Each w In Workbooks
        If UCase(w.name) Like "*PJ*" Then
            
            Dim ans As Variant
            
            ans = MsgBox("Before starting you need to close all doc-info related files, OK (close without save)?", vbYesNo)
            
            If ans = vbYes Then
                w.Close False
                
                ActiveWindow.Visible = True
            End If
        End If
    Next w
    
    ActiveWindow.Visible = True
    Application.DisplayAlerts = True
End Sub


Public Sub clearAllOldSheetsFromPrevDownload()
    
    Dim sh As Worksheet, x As Variant
    Application.DisplayAlerts = False

    x = 1
    Do
        If checkIfYouCantDelete(ThisWorkbook.Sheets(x)) Then
            x = x + 1
        Else
            Sheets(x).Delete
        End If
    Loop Until x > Sheets.Count

    Application.DisplayAlerts = True
End Sub



Private Function checkIfYouCantDelete(sh1 As Worksheet) As Boolean
    checkIfYouCantDelete = False
    
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets("register").Cells(2, 1)
    
    Do
    
        If sh1.name = r.Value Then
            checkIfYouCantDelete = True
            Exit Do
        End If
        
        Set r = r.offset(1, 0)
    Loop Until Trim(r) = ""
End Function


' obsolete
Private Function checkMyWorkbooks() As Boolean

    checkMyWorkbooks = False
    
    check_ EP.E_TOFF_PERTURBATION, perturbationWorkbook, checkMyWorkbooks
    check_ EP.E_TOFF_PERFO_APPRO, perfoWorkbook, checkMyWorkbooks
    check_ EP.E_TOFF_TRANSFER_MAJOR_FILE, majorTransferWorkbook, checkMyWorkbooks
    check_ EP.E_TOFF_KPI_STOCK, stockWorkbook, checkMyWorkbooks
End Function

Private Sub check_(e1 As E_TOFF, wrk1 As Workbook, ByRef bool As Boolean)
    
    If wrk1 Is Nothing Then
        ' for now nothing agressive - workbook is just empty continue reporting
        ' without this oen
        Debug.Print "workbook is nothing! Reporting will be continued but without this one!"
        bool = True
    Else
        
        Debug.Print "there is real workbook here! Name of this workbook: " & CStr(wrk1.name)
    End If
End Sub

Private Sub tryToOpen(ofst As Integer, ByRef wrk1 As Workbook)


    Application.ScreenUpdating = False
    
    ' perturbartion = 0
    ' perfo = 1
    ' major transfer = 2
    ' kpi stock = 4 -> but this one is still not available so be careful
    Dim tmpR As Range
    Set tmpR = mainRef.offset(Int(ofst), 0)
    
    ' Debug.Print tmpR.Value; tmpR.Address
    
    With tmpR.offset(Int(ofst), 1)
        .Value = ""
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    If Trim(tmpR.Value) <> "" Then
    
        If UCase(tmpR.Value) Like "HTTP://*" Then
            ' ok we have link - so it is good to fo furhter
            Application.DisplayAlerts = False
            Set wrk1 = Application.Workbooks.Open(tmpR.Value, True, True)
            
            ' ActiveWindow.Visible = False
            
            
            Do
                DoEvents
            Loop While wrk1 Is Nothing
            
            If Not wrk1 Is Nothing Then
            
                ThisWorkbook.Activate
                Application.ScreenUpdating = True
                Debug.Print wrk1.name
                
                
                With tmpR.offset(0, 1)
                    .Value = "OK"
                    .Font.Color = RGB(0, 240, 0)
                    .Font.Bold = True
                End With
                
                
                Application.ScreenUpdating = False
            Else
                With tmpR.offset(0, 1)
                    .Value = "NOK"
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = True
                End With
            End If
            
            Application.DisplayAlerts = True
        Else
        
        
            ' there is some content but this is not a proper URL
            MsgBox "missing URL!"
        End If
        
    Else
        ' no link = no action
        Application.ScreenUpdating = True
        
        With tmpR.offset(0, 1)
            .Value = "NOK"
            .Font.Color = RGB(225, 0, 0)
            .Font.Bold = True
        End With
        
        Application.ScreenUpdating = False
        
    End If
    
    
    Application.ScreenUpdating = True
End Sub


Private Sub doSomeOperationsAndCloseIt(ofst As Integer, ByRef wrk1 As Workbook, Optional rRef As Range)


    If Not wrk1 Is Nothing Then
        
    
        Application.DisplayAlerts = False
        Application.ScreenUpdating = True
        ThisWorkbook.Activate
        Application.ScreenUpdating = False
        
        
        Dim tmpR As Range
        Set tmpR = mainRef.offset(Int(ofst), 1)
        tmpR.Value = ""
        tmpR.Font.Color = RGB(0, 0, 0)
        
        
        
        
        ' simple example of coping data from source sheet
        'Sheets("Perturbations").Select
        'Sheets("Perturbations").Copy After:=Workbooks("threePieces.xlsm").Sheets(3)
        Dim arr() As Variant
        
        Dim sh As Worksheet, shList As Worksheet, shArchive As Worksheet
        
        arr = EP.getSourceWorksheetNames(ofst)
        
        If ofst >= EP.E_TOFF_TRANSFER_FILE_URL Then
            ' Array(G_KPI_TRANSFER_FILE_SH_NM_TABLE, G_KPI_TRANSFER_FILE_SH_NM_LISTES)
            
            ' main table info
            Set sh = Nothing
            On Error Resume Next
            Set sh = wrk1.Sheets(CStr(arr(LBound(arr))))
            
            ' G_KPI_TRANSFER_FILE_SH_NM_TABLE
            Set tfh = New TransferFileHandler
            tfh.tb.init sh
            
            ' tfh.tb.testDic ' OK
            
            tfh.knottedData.getTheDicFromTable tfh.tb.dic
            
            
            ' UPDATE IN IMPLEMENTATION
            Set shArchive = Nothing
            On Error Resume Next
            'Set shArchive = wrk1.Sheets(CStr(arr(UBound(arr))))
            'tfh.arch.findAllStaticColumns shArchive
            
            
            
            Set shList = Nothing
            On Error Resume Next
            Set shList = wrk1.Sheets(CStr(arr(UBound(arr) - 1)))
            tfh.list.findAllColumns shList
            
            
            
            
            
            ' tutaj jest ciekawie
            ' sibling "polyka" siblinga jako parametr - chowanie operacji zewnetrznej do brata/siostry najsss...
            ' wewnatrz komponentu komponentu
            ' dobra, wyglada to tak:
            ' main
            '   -> DocInfoHandler
            '       -> sub init TransferFileHandler
            '           -> 3 fields
            '               > list obj
            '               > tab obj
            '               > knotting object for above
            
            ' HERE PUT DATA FROM LISTES
            tfh.list.takeDataFromListingAndBindEverything tfh.knottedData
            
            ' same logic for archive lines - the obejct it self is quite different but fortunanetly knotted data stores info in inner variables
            ' so in worst case scenario you can simply understand that if specific var is nothing means that we have other type of data here!
            '
            ' HERE PUT DATA FROM ARCHIVE
            'tfh.arch.takeDataFromListingAndBindEverything tfh.knottedData
            
            'tfh.knottedData.showCollection
            tfh.knottedData.putCollectionOfDataIntoWorksheet rRef
            
        Else
        
            ' this is just simple copy sh from one wrk to another wrk
            copyWorksheetsFromSource sh, arr, ofst, wrk1, tmpR
        End If
        
        wrk1.Close False
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        
        
    End If
End Sub

Private Sub copyWorksheetsFromSource(sh As Worksheet, arr() As Variant, ofst As Integer, ByRef wrk1 As Workbook, tmpR As Range)
    
    Dim x As Variant
    For x = LBound(arr) To UBound(arr)
    
    
        Debug.Print "doSomeOperationsAndCloseIt: " & arr(x)
        
        Set sh = Nothing
        On Error Resume Next
        Set sh = wrk1.Sheets(CStr(arr(x)))
        
        If sh Is Nothing Then
            ' it means that name of the worksheet is a little bit more soph.
            Set sh = tryToFindAgainWorksheet(wrk1, CStr(arr(x)))
        End If
        
        
        Application.ScreenUpdating = True
        ThisWorkbook.Activate
        
        
        
        If Not sh Is Nothing Then
        
            tmpR.Value = "OK"
            tmpR.Font.Color = RGB(0, 255, 0)
            
            
            
            
            If ofst > EP.E_TOFF_TRANSFER_FILE_URL Then
                ' Array(G_KPI_TRANSFER_FILE_SH_NM_TABLE, G_KPI_TRANSFER_FILE_SH_NM_LISTES)
            Else
            
                ' copy only for direct KPI files
                sh.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                
            End If
            
        Else
            MsgBox "No standarised worksheets in this doc-info file!", vbExclamation
            tmpR.Value = "NO DATA"
            tmpR.Font.Color = RGB(255, 255, 0)
        End If
        
        
        Application.ScreenUpdating = False
        
    Next x
End Sub

Private Function tryToFindAgainWorksheet(wrk1 As Workbook, nm As String) As Worksheet
    
    Set tryToFindAgainWorksheet = Nothing
    
    
    Dim tmpSh As Worksheet
    For Each tmpSh In wrk1.Sheets
        If UCase(tmpSh.name) Like "*" & UCase(nm) & "*" Then
            Set tryToFindAgainWorksheet = tmpSh
            Exit Function
        End If
    Next tmpSh
End Function

Private Sub Class_Initialize()
    Set perturbationWorkbook = Nothing
    Set perfoWorkbook = Nothing
    Set majorTransferWorkbook = Nothing
    Set stockWorkbook = Nothing
    
    
    Set mainRef = ThisWorkbook.Sheets("MAIN").Range("B3")
End Sub

Private Sub Class_Terminate()
    Set perturbationWorkbook = Nothing
    Set perfoWorkbook = Nothing
    Set majorTransferWorkbook = Nothing
    Set stockWorkbook = Nothing
End Sub
