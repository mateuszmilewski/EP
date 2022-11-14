VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AmpedData"
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

Implements IAmpedData



Private ai As AmpedItem
Private mi As IMatrix


Private chaoticCollectionOfData As Collection


Private Sub Class_Initialize()
    Set ai = Nothing
    Set mi = Nothing
    
    Set chaoticCollectionOfData = Nothing
End Sub

Private Sub Class_Terminate()
    Set ai = Nothing
    Set mi = Nothing
    
    Set chaoticCollectionOfData = Nothing
End Sub

Private Property Set IAmpedData_ai(ByVal RHS As AmpedItem)
    Set ai = RHS
End Property

Private Property Get IAmpedData_ai() As AmpedItem
    Set IAmpedData_ai = ai
End Property

Private Property Set IAmpedData_dataset(ByVal RHS As Collection)
    Set chaoticCollectionOfData = RHS
End Property

Private Property Get IAmpedData_dataset() As Collection
    Set IAmpedData_dataset = chaoticCollectionOfData
End Property

Private Sub IAmpedData_make(m1 As IMatrix, whatYouWantToAmp As E_AMP)



    Dim tlc As Integer, trc As Integer, blc As Integer, brc As Integer
    
    tlc = Int(m1.topLeft.Column)
    trc = Int(m1.topRight.Column)
    blc = Int(m1.bottomLeft.Column)
    brc = Int(m1.bottomRight.Column)
    
    Dim tlr As Integer, trr As Integer, blr As Integer, brr As Integer
    
    tlr = Int(m1.topLeft.row)
    trr = Int(m1.topRight.row)
    blr = Int(m1.bottomLeft.row)
    brr = Int(m1.bottomRight.row)
    
    
    If whatYouWantToAmp = E_AMP_PERTURBATION Or _
        whatYouWantToAmp = E_AMP_RETARD Or _
        whatYouWantToAmp = E_AMP_MANQ Or _
        whatYouWantToAmp = E_AMP_TRANSFER Or _
        whatYouWantToAmp = E_AMP_STOCK Then
        
        
            ' Debug.Assert whatYouWantToAmp <> E_AMP_STOCK
    
        
            If chaoticCollectionOfData Is Nothing Then
                Set chaoticCollectionOfData = New Collection
            End If

    
            Dim r As Range
        
            For Each r In m1.area
                ' Debug.Print "r from area: " & r
                
                If r.row = tlr And r.Column = tlc Then
                    ' empty
                Else
                
                    If r.row > tlr Then
                    
                        If r.Column > tlc Then
                        
                            
                            Set ai = Nothing
                            Set ai = New AmpedItem
                            With ai
                                .plantName = r.Parent.Cells(r.row, tlc).Value
                                .dataType = m1.name
                                .rawTiming = r.Parent.Cells(tlr, r.Column).Value
                                
                                .v = CStr(r.Value)
                                
                                .std whatYouWantToAmp
                                
                            End With
                            
                            
                            chaoticCollectionOfData.Add ai
                            
                            
                        End If
                    
                    End If
                    
                End If
            Next r
    
    ElseIf whatYouWantToAmp = E_AMP_TRANSFER_FILE Then
        
        ' to jest calkiem nowe i nieco inne
        '===================================================
        '===================================================
    End If
    
End Sub

Private Property Set IAmpedData_mi(ByVal RHS As IMatrix)
    Set mi = RHS
End Property

Private Property Get IAmpedData_mi() As IMatrix
    Set IAmpedData_mi = mi
End Property

Private Sub IAmpedData_setMatrix(m1 As IMatrix)
    Set mi = m1
End Sub

Private Sub IAmpedData_testReturnData()


    If Not chaoticCollectionOfData Is Nothing Then
        If chaoticCollectionOfData.Count > 0 Then

            For Each ai In chaoticCollectionOfData
                Debug.Print "AI: "; ai.stdDataType & " " & ai.stdPlt & " " & ai.stdTiming & " value: " & ai.v
            Next ai
        Else
            MsgBox "chaotic collection is empty!"
        End If
    Else
        MsgBox "no collection!", vbCritical
    End If
End Sub
