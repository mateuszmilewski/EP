VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferFileListColumnsHandler"
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


Public sousProjetColumn As Range
Public pn As Range
Public sgrFab As Range
Public indice As Range
Public des As Range
Public domain As Range
Public soreppDecision As Range
Public synteza As Range
Public fltr As Range

Public ba As Range
Public cofor As Range

Public columnS As Range


Public Function checkIfAllColumnsAreIn() As Boolean
    
    Dim tmpB As Boolean
    tmpB = False
    
    If Not Me.sousProjetColumn Is Nothing Then
        If Not Me.pn Is Nothing Then
            If Not Me.sgrFab Is Nothing Then
                If Not Me.indice Is Nothing Then
                    If Not Me.des Is Nothing Then
                        If Not Me.domain Is Nothing Then
                            If Not Me.soreppDecision Is Nothing Then
                                If Not Me.synteza Is Nothing Then
                                    If Not Me.fltr Is Nothing Then
                                        If Not Me.columnS Is Nothing Then
                                            If Not Me.ba Is Nothing Then
                                                If Not Me.cofor Is Nothing Then
                                            
                                                    tmpB = True
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    
    
    '
    
    checkIfAllColumnsAreIn = tmpB
    
End Function

Private Sub Class_Initialize()
    Set sousProjetColumn = Nothing
    Set pn = Nothing
    Set sgrFab = Nothing
    Set indice = Nothing
    Set des = Nothing
    Set domain = Nothing
    Set soreppDecision = Nothing
    Set synteza = Nothing
    Set fltr = Nothing
    
    Set ba = Nothing
    Set cofor = Nothing
    
    Set columnS = Nothing
End Sub

Private Sub Class_Terminate()
    Set sousProjetColumn = Nothing
    Set pn = Nothing
    Set sgrFab = Nothing
    Set indice = Nothing
    Set des = Nothing
    Set domain = Nothing
    Set soreppDecision = Nothing
    Set synteza = Nothing
    Set fltr = Nothing
    
    Set ba = Nothing
    Set cofor = Nothing
    
    Set columnS = Nothing
End Sub
