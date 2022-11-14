Attribute VB_Name = "SoftLookupsModule"
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



Public Function lookForColumn(srcSh As Worksheet, _
    labelRow As Integer, _
    namePattern As String, _
    wildcardRef As Range, _
    Optional col1 As Integer, Optional col2 As Integer, _
    Optional bgColor As Long, Optional fontColor As Long) As Range
    
    Set lookForColumn = Nothing
    
    Dim tmp As Range
    Set tmp = srcSh.Range("A" & CStr(labelRow))
    
    
    
    Dim matchProbability As Byte
    Dim collectionOfGoodResults As Collection
    
    matchProbability = 0
    
    ' overflow byte - 0 -255
    ' matchProbability = 300
    ' matchProbability = 255
    
    Set collectionOfGoodResults = New Collection
    
    Do
    
    
        If Trim(tmp.Value) <> "" Then
    
            matchProbability = 0
        
            If tmp.Interior.Color = bgColor Then
                matchProbability = matchProbability + 10
            End If
            
            If CLng(tmp.Font.Color) = fontColor Then
                matchProbability = matchProbability + 10
            End If
            
            If UCase(tmp.Value) Like UCase(namePattern) Then
                matchProbability = matchProbability + 30
            End If
            
            If tmp.Column >= col1 And tmp.Column <= col2 Then
                matchProbability = matchProbability + 10
            End If
            
            If wildcardRef.Value Like "*" & CStr(tmp.Value) & "*" Then
                matchProbability = matchProbability + 100
            End If
            
            
            If matchProbability > 40 Then
                collectionOfGoodResults.Add tmp
            End If
            
            If matchProbability > 140 Then
                ' top result
                Set lookForColumn = tmp
                Exit Function
            End If
        End If
        
        
        Set tmp = tmp.offset(0, 1)
    Loop While tmp.Column < 1000
    
    
    If lookForColumn Is Nothing Then
        
        If collectionOfGoodResults.Count > 0 Then
            
            Decision.ListBox1.Clear
            Decision.Caption = CStr(namePattern) & " " & Trim(wildcardRef.Value) & " " & Trim(tmp.Value)
            For Each tmp In collectionOfGoodResults
                If Trim(tmp.Value) <> "" Then
                    Decision.ListBox1.addItem tmp.Address & " __;__ " & tmp.Value
                End If
            Next tmp
            
            If collectionOfGoodResults.Count = 1 Then
                
                For Each tmp In collectionOfGoodResults
                    
                    Set lookForColumn = tmp
                    Exit Function
                Next tmp
                
            ElseIf collectionOfGoodResults.Count = 2 Then
            
            
                Set lookForColumn = Nothing
            
                Dim ifOnlyOne As Boolean
                ifOnlyOne = False
                
                
                Debug.Print collectionOfGoodResults.item(2)
                Debug.Print collectionOfGoodResults.item(1)
                
                ' very very specific
                If collectionOfGoodResults.item(1).Value = Trim("Date 1ère uti. domaine") Then
                    ifOnlyOne = True
                    Set lookForColumn = collectionOfGoodResults.item(1)
                End If
                
                
                ' if still nothing
                If lookForColumn Is Nothing Then
                    Decision.show
                    Set lookForColumn = srcSh.Range(Split(Decision.ListBox1.Value, " __;__ ")(0))
                End If
                
            Else
            
            
                Decision.show
                Set lookForColumn = srcSh.Range(Split(Decision.ListBox1.Value, " __;__ ")(0))
            End If
        End If
    End If
    
End Function


Public Sub lookUpTest()
    
    Dim sousProjetColumn As Range, pn As Range, sgrFab As Range, indice As Range, des As Range, domain As Range
    Dim soreppDecision As Range, synteza As Range
    Dim fltr As Range
    
    Dim sh As Worksheet, reg As Worksheet
    Set sh = ActiveSheet
    Set reg = ThisWorkbook.Sheets("register")
    
    Set pn = lookForColumn(sh, 2, "*NUM*PRODU*", reg.Range("AA3"), 1, 1, RGB(204, 255, 204), RGB(0, 0, 255))
    Debug.Print pn.Value & " " & pn.Address
    
    Set sgrFab = lookForColumn(sh, 2, "*SGR*FAB*", reg.Range("AA4"), 2, 2, RGB(204, 255, 204), RGB(0, 0, 255))
    Debug.Print sgrFab.Value & " " & sgrFab.Address
    
    Set indice = lookForColumn(sh, 2, "*IND*", reg.Range("AA5"), 3, 3, RGB(204, 255, 204), RGB(0, 0, 255))
    Debug.Print indice.Value & " " & indice.Address
    
    Set des = lookForColumn(sh, 2, "*DES*L*", reg.Range("AA6"), 8, 10, RGB(204, 255, 204), RGB(0, 0, 255))
    Debug.Print des.Value & " " & des.Address
    
    Set domain = lookForColumn(sh, 2, "*DOM*SIGAPP*", reg.Range("AA7"), 8, 10, RGB(204, 255, 204), RGB(0, 0, 255))
    Debug.Print domain.Value & " " & domain.Address
    
    Set sousProjetColumn = lookForColumn(sh, 2, "*PRO*SIGAPP", reg.Range("AA2"), 11, 11, RGB(204, 255, 204), RGB(0, 0, 255))
    Debug.Print sousProjetColumn.Value & " " & sousProjetColumn.Address
    
    
    
    Set soreppDecision = lookForColumn(sh, 2, "*CISI*", reg.Range("AA8"), 29, 31, RGB(204, 153, 255), RGB(0, 0, 255))
    Debug.Print soreppDecision.Value & " " & soreppDecision.Address
    
    Set synteza = lookForColumn(sh, 2, "*SYN*", reg.Range("AA9"), 30, 35, RGB(128, 128, 128), RGB(255, 0, 0))
    Debug.Print synteza.Value & " " & synteza.Address
    
    Set fltr = lookForColumn(sh, 2, "*FILT*", reg.Range("AA10"), 100, 105, RGB(192, 192, 192), RGB(0, 0, 0))
    Debug.Print "fltr: " & fltr.Value & " " & fltr.Address
End Sub
