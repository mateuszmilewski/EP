VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferFileTableHandler"
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


Public dic As Dictionary


Private srcSh As Worksheet
Private srcRange As Range


Private tfti As TransferFileTableItemHandler

Private Sub Class_Initialize()
    Set dic = New Dictionary
    
    Set srcSh = Nothing
    
    
End Sub


Public Sub init(sh1 As Worksheet)


    Dim ki As String

    On Error Resume Next
    Set srcSh = sh1
    
    If Not srcSh Is Nothing Then
    
        ' ---------------------------------
        Set srcRange = srcSh.Range("A3")
        
        
        Do
        
            Set tfti = New TransferFileTableItemHandler
        
            With tfti
                .str = srcRange.Value
                .sousProjet = srcRange.offset(0, 1).Value
                .indicateurs = srcRange.offset(0, 2).Value
                .EL1 = parsetodate(srcRange.offset(0, 3).Value)
                .EL2 = parsetodate(srcRange.offset(0, 4).Value)
                .MDL = parsetodate(srcRange.offset(0, 5).Value)
                .AMC = parsetodate(srcRange.offset(0, 6).Value)
                
            End With
            
            ki = tfti.str & "---" & tfti.sousProjet
            dic.Add ki, tfti
            
            
        
            Set srcRange = srcRange.offset(1, 0)
        Loop Until Trim(srcRange.Value) = ""
        ' ---------------------------------
    End If
End Sub

Private Sub Class_Terminate()
    Set dic = Nothing
End Sub



Public Sub testDic()
    
    Dim ki As Variant
    Debug.Print "dic.Count: " & dic.Count
    For Each ki In dic.Keys
    
        Set tfti = dic(ki)
        Debug.Print tfti.str
        Debug.Print tfti.sousProjet
        Debug.Print tfti.EL1
        Debug.Print tfti.MDL
        
    Next
End Sub



Private Function parsetodate(strDate As String) As Date
    
    If IsDate(strDate) Then
        parsetodate = CDate(strDate)
    Else
        parsetodate = CDate(0)
    End If
End Function
