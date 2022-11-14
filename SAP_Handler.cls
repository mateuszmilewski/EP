VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP_Handler"
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


' OK - we need to change it into late binging becuase EVO users
' might do not have any SAP app

Private sapGuiAuto As Object
'Private sapGui As SAPFEWSELib.GuiApplication
'Private cnn As SAPFEWSELib.GuiConnection
'Private sess As SAPFEWSELib.GuiSession
'Private window As SAPFEWSELib.GuiModalWindow
'Private c As SAPFEWSELib.GuiCollection
'Private item2 As SAPFEWSELib.GuiComponent
'Private Btn As SAPFEWSELib.GuiButton
'Private table2 As SAPFEWSELib.GuiTableControl ' ISapTableControlTarget
'Private grid As SAPFEWSELib.GuiGridView
'Private gv As SAPFEWSELib.GuiGridView
'Private myTableRow As SAPFEWSELib.GuiTableRow
'Private coll As SAPFEWSELib.GuiCollection

'Private chbx As SAPFEWSELib.GuiCheckBox
'Private txt As SAPFEWSELib.GuiTextedit

'Private tf As SAPFEWSELib.GuiTextField
'Private lt As SAPFEWSELib.GuiLabel

    
Private sapGui As Object
Private cnn As Object
Private sess As Object
Private window As Object
Private c As Object
Private item As Object
Private Btn As Object
Private table As Object
Private grid As Object
Private gv As Object
Private myTableRow As Object
Private coll As Object
    
Private chbx As Object
Private txt As Object


' !!
Private cols As Variant

' params for table inside sap
Private y As Variant
Private x As Variant
Private x2 As Variant
Private x3 As Variant




Private Sub Class_Initialize()

    If GlobalSapModule.sapGuiAuto Is Nothing Then
        Set sapGuiAuto = GetObject("SAPGUI")
        Set GlobalSapModule.sapGuiAuto = sapGuiAuto
        Set sapGui = sapGuiAuto.GetScriptingEngine
        Set GlobalSapModule.sapGui = sapGui
    Else
        Set sapGuiAuto = GlobalSapModule.sapGuiAuto
        Set sapGui = GlobalSapModule.sapGui
    End If
    
    Set cnn = sapGui.Connections(0)
    
    Debug.Print cnn.ConnectionString
    Debug.Print cnn.Sessions.Count

    
    Set sess = cnn.Children(0)
    Set item = sess.Children(0)
    Debug.Print item.name
    
    Debug.Print sess.Children.Count
    
    
    ' Set item = sess.Children(0)
    Set item = sess.FindById("wnd[0]/usr")
    Debug.Print item.Children.Count

    
    
    ' sess.FindById("wnd[0]").Maximize
End Sub

Private Sub Class_Terminate()
    Set sapGuiAuto = Nothing
    Set sapGui = Nothing
    Set cnn = Nothing
    Set sess = Nothing
    Set item = Nothing
End Sub




Private Function tryToAssignFirstBatchOfData(sh As Worksheet) As Range
    Set tryToAssignFirstBatchOfData = sh.Range("A2:A500")
End Function



Private Function tryToAssignNewRange(sh As Worksheet, data As Range) As Range
    Set tryToAssignNewRange = Nothing
    
    
    Dim tmp As Range
    
    Set tmp = data.offset(499, 0)
    
    ' Debug.Print tmp.Address
    
    Set tryToAssignNewRange = tmp
End Function

Private Function youDryOutRangeOfDataForSq01(data As Variant) As Boolean

    youDryOutRangeOfDataForSq01 = False
    
   
End Function

Private Function findMyButton() As Variant ' SAPFEWSELib.GuiButton
    
    Set findMyButton = Nothing
    
    ' top
    '51
    '73
    '95
    '29
    
    Dim tmpButton As Variant ' SAPFEWSELib.GuiButton
    
    Dim x As Variant
    For x = 1 To 4
        ' Set btn = .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH")
        Set tmpButton = sess.FindById("wnd[0]/usr/btn%_SP$0000" & CStr(x) & "_%_APP_%-VALU_PUSH")
        
        If (tmpButton.Top > 70 And tmpButton.Top < 80) Or (tmpButton.Top = 73) Then
            ' this is it
            Set findMyButton = tmpButton
            Exit For
        End If
    Next x
End Function


Private Function findMyButton2(e2 As E_CFG_SQ01_QUERY) As Variant ' SAPFEWSELib.GuiButton
    
    Set findMyButton2 = Nothing
    
    ' top
    '51
    '73
    '95
    '29
    
    Dim tmpButton As Variant ' SAPFEWSELib.GuiButton
    
    Dim x As Variant
    For x = 1 To 4
        ' Set btn = .FindById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH")
        Set tmpButton = sess.FindById("wnd[0]/usr/btn%_SP$0000" & CStr(x) & "_%_APP_%-VALU_PUSH")
        
        
        Debug.Print "tmpButton.Top: " & tmpButton.Top
        
        
        ' button pod numery
        If e2 = E_NUMERO Then
            If (tmpButton.Top > 70 And tmpButton.Top < 80) Or (tmpButton.Top = e2) Then
                Set findMyButton2 = tmpButton
                Exit For
            End If
        End If
        
        
        ' achat
        If e2 = E_TYPE_DE_DOC_ACHAT Then
            If (tmpButton.Top > 20 And tmpButton.Top < 30) Or (tmpButton.Top = e2) Then
                Set findMyButton2 = tmpButton
                Exit For
            End If
        End If
        
        
        ' code
        If e2 = E_CODE_DE_SUPP Then
            If (tmpButton.Top > 50 And tmpButton.Top < 60) Or (tmpButton.Top = e2) Then
                Set findMyButton2 = tmpButton
                Exit For
            End If
        End If
        
        
        If e2 = E_DIV Then
            If (tmpButton.Top > 90 And tmpButton.Top < 100) Or (tmpButton.Top = e2) Then
                Set findMyButton2 = tmpButton
                Exit For
            End If
        End If
    Next x
End Function


Private Sub runLoopsWithStaticText375(stringReference As String)


    Dim wynik As String

    Set grid = sess.FindById(CStr(stringReference))
    Set cols = grid.ColumnOrder
    
    wynik = "TP04_375"
    
    For x = 0 To grid.RowCount - 1
        ' Debug.Print CStr(grid.getCellValue(x, cols(0))) ' glowna lista z SQ01
        If CStr(wynik) = CStr(grid.getCellValue(x, cols(0))) Then
        
            With grid
                .CurrentCellRow = x
                .SelectedRows = "" & CStr(x)
                .DoubleClickCurrentCell
            End With
            
            Exit For
            
        End If
    Next x

End Sub


Public Sub runMainLogicForSQ01__forEP(inputSheet As Worksheet, startingPoint As String, refOutput2_SH As Worksheet, tp04_static_string As String)



    EP.SuppressingMessageModule.KillMessageFilter
    
    Dim sel As Range
    ' im not proud of this "hack"
    ' ----------------------------------------------------
    Dim x17 As Variant
    For x17 = 0 To 10
        On Error Resume Next
        sess.FindById("wnd[0]/tbar[0]/btn[12]").press
    Next x17
    ' ----------------------------------------------------


    sess.FindById("wnd[0]/tbar[0]/okcd").Text = "SQ01"
    sess.FindById("wnd[0]").sendVKey 0

    runLoopsWithStaticText375 "wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell"
    
    'sess.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").DoubleClickCurrentCell
    sess.FindById("wnd[0]/tbar[1]/btn[8]").press
    

    
    Dim btn3 As Variant
    Set btn3 = Nothing
    On Error Resume Next
    Set btn3 = findMyButton2(E_TYPE_DE_DOC_ACHAT)
    If Not btn3 Is Nothing Then
        btn3.press
        sess.FindById("wnd[1]/tbar[0]/btn[16]").press
        sess.FindById("wnd[1]").sendVKey 4
        sess.FindById("wnd[2]/usr").verticalScrollbar.Position = 48
        sess.FindById("wnd[2]/usr/lbl[5,23]").SetFocus
        sess.FindById("wnd[2]/usr/lbl[5,23]").caretPosition = 2
        sess.FindById("wnd[2]").sendVKey 2
        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        sess.FindById("wnd[1]/tbar[0]/btn[8]").press
    End If
    
    
    Set btn3 = Nothing
    On Error Resume Next
    Set btn3 = findMyButton2(E_CODE_DE_SUPP)
    If Not btn3 Is Nothing Then
    
        btn3.press
        sess.FindById("wnd[1]/tbar[0]/btn[16]").press
        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        sess.FindById("wnd[1]/tbar[0]/btn[8]").press
    End If
    
    Set btn3 = Nothing
    On Error Resume Next
    Set btn3 = findMyButton2(E_DIV)
    If Not btn3 Is Nothing Then
    
        btn3.press
        sess.FindById("wnd[1]/tbar[0]/btn[16]").press
        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        sess.FindById("wnd[1]/tbar[0]/btn[8]").press
    End If
    
    
    Dim iter As Long, s1 As Long, endzik As Long, sajz As Long
    iter = 1
    sajz = 500
    
    Dim refOutput2 As Range
    Set refOutput2_SH = ThisWorkbook.Sheets.Add
    Set refOutput2 = refOutput2_SH.Range(startingPoint)
    
    Dim a As String
    a = Left(startingPoint, 1)
        
    Do

        ' foo should be but sth is not working propetly
        ' ----------------------------------------------------------------------
        If iter = 1 Then
            s1 = 2 + sajz * (iter - 1)
        Else
            s1 = 1 + sajz * (iter - 1)
        End If
        endzik = sajz + sajz * (iter - 1)
        Set sel = inputSheet.Range(a & CStr(s1) & ":" & a & CStr(endzik))
        ' ----------------------------------------------------------------------
    
    
        Set btn3 = findMyButton2(E_NUMERO)
        btn3.press
    
    
        sess.FindById("wnd[1]/tbar[0]/btn[16]").press
        
        sel.Copy
        sess.FindById("wnd[1]/tbar[0]/btn[24]").press
        
        
        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
        sess.FindById("wnd[1]/tbar[0]/btn[8]").press
        sess.FindById("wnd[0]/tbar[1]/btn[8]").press
        
        
        Set grid = Nothing
        On Error Resume Next
        Set grid = sess.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        
        If Not grid Is Nothing Then
            Set cols = grid.ColumnOrder
        

            
            
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
            
            
            For x = 0 To grid.RowCount - 1
                For y = 0 To grid.ColumnCount - 1
                
                    ' Debug.Print CStr(grid.getCellValue(x, cols(y)))
                    refOutput2.offset(0, y).Value = CStr(grid.getCellValue(x, cols(y)))
                    
                Next y
                
                If x Mod 50 = 0 Then
                
                    grid.FirstVisibleRow = x
                    grid.CurrentCellRow = x
                    
            
                    
                End If
                
                
                Set refOutput2 = refOutput2.offset(1, 0)
            Next x
            
            Application.EnableEvents = True
            Application.Calculation = xlCalculationAutomatic
            
            
            'step back in SIGAPP
            sess.FindById("wnd[0]/tbar[0]/btn[3]").press
            
            
            
        End If
        
        iter = iter + 1
        
        
    Loop While dataAvailable(sel, inputSheet)
    


    EP.SuppressingMessageModule.RestoreMessageFilter
    

        

End Sub


Private Function dataAvailable(r As Range, ish As Worksheet) As Boolean
    
    Dim lt As Range
    
    Set lt = r.Cells(r.Count)
    
    
    If Trim(lt.Value) <> "" Then
        dataAvailable = True
    Else
        dataAvailable = False
    End If
End Function







Private Sub makeDoubleListWithDashAndStar(refR As Range)
    
    Dim doubleR As Range, lastOrgRng As Range
    Set doubleR = refR.End(xlDown).End(xlDown).End(xlUp).offset(1, 0)
    
    If refR.offset(1, 0) <> "" Then
        
        Set lastOrgRng = refR.End(xlDown)
    Else
        Set lastOrgRng = refR
    End If
    
    Dim ir As Range
    For Each ir In refR.Parent.Range(refR, lastOrgRng)
        doubleR.Value = ir.Value & "-*"
        doubleR.offset(0, 1).Value = "*-*"
        Set doubleR = doubleR.offset(1, 0)
    Next ir
End Sub





Public Sub runMainLogicFor__ME33K_EP(sh1 As Worksheet)
    

    
    
    If sh1.Range("A1").Value = "DIV" And sh1.Range("B1").Value = "A" Then
    

            
             '------------------------------------------------------
             
             Dim ir As Range, r As Range, lt As Variant, tf As Variant
             Dim idForMag As String, idForArticle As String, idForDiv As String
             Dim articleWoIndex As String, domain As String, mag As String
             
             Dim x As Variant, y As Variant
             
             Dim innerColl As Collection
             Dim tmpStr As String, prevStr As String
             
             Dim directDataRow As Variant, el As Variant
             
             Dim txtForDiv As String, txtForDomain As String
             
             Set r = sh1.Range(sh1.Cells(2, 1), sh1.Cells(2, 1).End(xlDown))
             
             For Each ir In r
             
             
             
                If ir.EntireRow.Hidden = True Then
                    ' no operation for filtered out
                Else
                
                                    
                    idForMag = _
                        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]"
                    idForArticle = _
                        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[5,24]"
                        
                    idForDiv = _
                        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[3,24]"
                        
                    
                    
                    articleWoIndex = CStr(Split(CStr(ir.offset(0, 1).Value), "-")(0))
                    domain = Trim(ir.offset(0, 2).Value)
                    mag = domain & "0"
                    txtForDiv = CStr(ir.Value)
                    
                    sess.FindById("wnd[0]/tbar[0]/okcd").Text = "ME33K"
                    sess.FindById("wnd[0]").sendVKey 0
                    sess.FindById("wnd[0]").sendVKey 4
                    
                    ' tab 20 == possibility to put articles!
                    sess.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020").Select
                    ' sess.findById("wnd[0]/tbar[1]/btn[8]").press
                    
                    If Trim(mag) = "0" Then
                        sess.FindById(idForMag).Text = "*"
                    Else
                        sess.FindById(idForMag).Text = mag
                    End If
                    ' sess.findById(idForArticle).Text = "98398447XT"
                    sess.FindById(idForArticle).Text = articleWoIndex & "*"
                    
                    
                    ' new for 107
                    sess.FindById(idForDiv).Text = CStr(txtForDiv)
                    
                    'sess.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").SetFocus
                    'sess.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB020/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").caretPosition = 4
                    sess.FindById("wnd[1]/tbar[0]/btn[0]").press
                    
                    
                    
                    
                    x = 1
                    y = 1
                    
                    Set lt = Nothing
                    On Error Resume Next
                    Set lt = sess.FindById("wnd[1]/usr/lbl[" & CStr(y) & "," & CStr(x) & "]")
                    
                    If lt Is Nothing Then ' still
                        sess.FindById(idForMag).Text = "*"
                        sess.FindById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                    
                    
                    ' example of output
                    
                    ' row 8 as first ' v on 21 col
                    ' 1, 1 "division" ' v on 21 col
                    ' 2, 1"mag" ' v from mag starts like everything else on 21 col
                    ' Catég. doc. d'achat x 3 y 1 'K x 3 y 21
                    ' Type document achat x 4 y 1 ' YPRS x 4 y 21
                    
                    ' x = 5 empty
                    
                    ' beg of table
                    ' Doc achat x 6 y 1
                    ' Poste x 6 y 12
                    ' Article x 6 y 18
                    ' Fourn. x 6 y 37
                    
                    ' x = 7 empty
                    
                    ' x = 8 content starting
                    '
                    ' 3939452822 x 8 y 1 ( doc achat )
                    ' 00010 x 8 y 12 ( poste )
                    ' 9830943980 x 8 y 18 ( article )
                    ' 50871U  01 x 8 y 37 ( cofor)
                    
                    
                    
                    Set innerColl = Nothing
                    Set innerColl = New Collection
                    
                    tmpStr = ""
                    prevStr = ""
                    
                    
                    For x = 1 To 20
                        
                        tmpStr = ""
                        prevStr = ""
                    
    
                        
                        For y = 1 To 100
                        
                            If y = 1 Then prevStr = ""
                            
                        
                            Set lt = Nothing
                            On Error Resume Next
                            Set lt = sess.FindById("wnd[1]/usr/lbl[" & CStr(y) & "," & CStr(x) & "]")
                            If Not lt Is Nothing Then
                                ' Debug.Print lt.Text & " x " & x & " y " & y
                                
                                tmpStr = tmpStr & lt.Text & ";"
                                prevStr = lt.Text
                                
                            End If
                        Next y
                        
                        innerColl.Add tmpStr
                    Next x
                    
                    

                    Dim pozycja As Integer ' , lesposib
                    pozycja = 1
                    
                    For Each el In innerColl
                        
                        If CStr(el) Like "*" & articleWoIndex & "*" Then
                            If CStr(el) Like "*YPRS*" Then
                                Exit For
                            Else
                                Debug.Print el
                                Exit For
                            End If
                        End If
                        
                        pozycja = pozycja + 1
                        
                    Next el
                    
                    If pozycja > innerColl.Count Then pozycja = 3
                    
                    
                    Debug.Print "pozycja: " & CStr(pozycja)
                    directDataRow = pozycja
                    sess.FindById("wnd[1]/usr/lbl[1," & CStr(directDataRow) & "]").SetFocus
                    sess.FindById("wnd[1]/usr/lbl[1," & CStr(directDataRow) & "]").caretPosition = 1
                    
                    
                    
                    sess.FindById("wnd[1]/tbar[0]/btn[0]").press
                    'sess.findById("wnd[0]").sendVKey 0
                    
                    sess.FindById("wnd[0]/tbar[1]/btn[5]").press
                    
                    
                    ' ISapCTextField
                    ' Debug.Print TypeName(sess.findByid("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-BPRME[9,1]"))
                    
                    ' session.findById("wnd[0]/usr/ctxtRM06E-EVRTN")
                    Debug.Print TypeName(sess.FindById("wnd[0]/usr/ctxtRM06E-EVRTN"))
                    Debug.Print TypeName(sess.FindById("wnd[0]/usr/ctxtEKKO-LIFNR"))
                        
                    Set tf = Nothing
                        
                    On Error Resume Next
                    Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-BPRME[9,1]")
                    
                    Debug.Print tf.Text
                    Debug.Print tf.id
                    
                    
                    
                    ' very first field:  contract
                    Set tf = Nothing
                    Set tf = sess.FindById("wnd[0]/usr/ctxtRM06E-EVRTN")
                    Debug.Print tf.Text
                    Debug.Print tf.id
                    
                    
                    
                    ' fournisseur
                    Set tf = Nothing
                    Set tf = sess.FindById("wnd[0]/usr/ctxtEKKO-LIFNR")
                    Debug.Print tf.Text
                    Debug.Print tf.id
                    
                    
                    
                    ' date contr
                    Set tf = Nothing
                    Set tf = sess.FindById("wnd[0]/usr/ctxtRM06E-VEDAT")
                    Debug.Print tf.Text
                    Debug.Print tf.id
                    
                    
                    
                    ' type de contrat
                    Set tf = Nothing
                    Set tf = sess.FindById("wnd[0]/usr/ctxtRM06E-EVART")
                    Debug.Print tf.Text
                    Debug.Print tf.id
                    
                    
                    ' devise / currency
                    Set tf = Nothing
                    Set tf = sess.FindById("wnd[0]/usr/ctxtEKKO-WAERS")
                    Debug.Print tf.Text
                    Debug.Print tf.id
                    
                    
                    
                    ' 3 or 4 column jako artykul zero bez indexu
                    Set tf = Nothing
                    Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]")
                    Debug.Print tf.Text
                    Debug.Print tf.id
                    
                    
                    ' 3 or 4 column jako artykul - jeden nizej z indexem
                    Set tf = Nothing
                    Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,1]")
                    Debug.Print tf.Text
                    Debug.Print tf.id
                    
                    ' jesli tylko dwa, to to powinno byc puste!
                    Set tf = Nothing
                    Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,2]")
                    Debug.Print tf.Text & " tutaj powinno byc pusto!" ' al jest cos takiego: "__________________"
                    Debug.Print tf.id
                    
                    Dim underscores As String
                    underscores = "__________________"
                    
                    
                    Dim ajtem As ME33K_Item
                    Dim collectionOfItems As Collection
                    Set collectionOfItems = Nothing
                    Set collectionOfItems = New Collection
                    
                    For x = 0 To 12
                    
                    
                        Set ajtem = New ME33K_Item
                    
                        Set tf = Nothing
                        y = 3
                        On Error Resume Next
                        Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[" & CStr(y) & "," & CStr(x) & "]")
                            
                        If Not tf Is Nothing Then
                            If tf.Text = "__________________" Then
                                Exit For
                            Else
                            
                                ajtem.article = tf.Text
                                ' session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-NETPR[7,0]")
                                
                                
                                ' EUR
                                ' ====================================================
                                Set tf = Nothing
                                Set tf = sess.FindById("wnd[0]/usr/ctxtEKKO-WAERS")
                                On Error Resume Next
                                ajtem.curr = tf.Text
                                ' ====================================================
                                
                                
                                
                                ' PRICE
                                ' ====================================================
                                ' ====================================================
                                Set tf = Nothing
                                y = 7
                                On Error Resume Next
                                Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-NETPR[" & CStr(y) & "," & CStr(x) & "]")
                                
                                If Not tf Is Nothing Then
                                    ajtem.strPrice = tf.Text
                                    ajtem.parsePrice
                                    
                                    
                                    Set tf = Nothing
                                    y = 9
                                    ' On Error Resume Next
                                    Set tf = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-BPRME[" & CStr(y) & "," & CStr(x) & "]")
                                    If Not tf Is Nothing Then ajtem.up = tf.Text
                                    
                                    ajtem.makePriceByUnit
                                    
                                    
                                    collectionOfItems.Add ajtem
                                End If
                                ' ====================================================
                                ' ====================================================
                                
                            End If
                        End If
                    Next x
                    
                    
                    
                    
                    ' ISapTableControlTarget
                    ' Debug.Print TypeName(sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220"))
                    
                    'Set table2 = sess.FindById("wnd[0]/usr/tblSAPMM06ETC_0220")
                    '
                    '
                    'x = 3
                    'y = 1
                    'Set item2 = table2.GetCell(x, y)
                    '
                    'Debug.Print item2.id & " " & TypeName(item2)
                    '
                    'Set tf = item2
                    'Debug.Print item2.Text
                    
                    
                    'For Each ajtem In collectionOfItems
                    '    Debug.Print "ajtem: " & ajtem.article & " " & ajtem.price
                    'Next ajtem
                    
                    
                    
                    ' ir jako column A
                    ' final green light re-check price

                    
                    Dim preSerialPriceRange As Range, initialSerialPriceRange As Range
                    
                    Set preSerialPriceRange = ir.offset(0, 3)
                    
                    'preSerialPriceRange.Font.Bold = True
                    'preSerialPriceRange.Font.Color = RGB(0, 0, 0)
                    'preSerialPriceRange.Interior.Color = RGB(255, 255, 255)
                    
                    Set initialSerialPriceRange = ir.offset(0, 4)
                    
                    'initialSerialPriceRange.Font.Bold = True
                    'initialSerialPriceRange.Font.Color = RGB(0, 0, 0)
                    'initialSerialPriceRange.Interior.Color = RGB(255, 255, 255)
                    
                    'Debug.Print "pre: " & preSerialPriceRange.Address & " " & preSerialPriceRange.Value
                    'Debug.Print "ini: " & initialSerialPriceRange.Address & " " & initialSerialPriceRange.Value
                    
                    
                    If collectionOfItems.Count > 0 Then
                    
                    
                        
                        Dim txtForPreSerialFromME33K As String, xtraComment As String
                        
                        'txtForPreSerialFromME33K = ""
                        'xtraComment = ""
                        '
                        'txtForPreSerialFromME33K = txtForPreSerialFromME33K & CStr(Date) & Chr(10) & _
                        '    "data from ME33K: " & Chr(10)
                    
                        For Each ajtem In collectionOfItems
                            'txtForPreSerialFromME33K = txtForPreSerialFromME33K & Chr(10) & _
                            '    ": " & ajtem.article & " " & ajtem.price
                                
                                
                            If CStr(ajtem.article) = CStr(ir.offset(0, 1).Value) Then
                                ' w index
                                preSerialPriceRange.Value = ajtem.price
                            End If
                            
                            If CStr(ajtem.article) = CStr(Split(ir.offset(0, 1).Value, "-")(0)) Then
                                ' wo index
                                initialSerialPriceRange.Value = ajtem.price
                            End If

                        Next ajtem
                    

                    End If
                
                
                
                
                    sess.FindById("wnd[0]/tbar[0]/btn[15]").press
                
                End If
                
            Next ir
        End If
             

End Sub

