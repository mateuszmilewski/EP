VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferFileArchColumnsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Public dateColumn1 As Range
Public pn As Range
Public description As Range
Public sousProjet As Range
Public ru As Range
Public date1 As Range
Public cmnt1 As Range


Public ba As Range


Public Function checkIfAllColumnsAreIn() As Boolean
    
    Dim tmpB As Boolean
    tmpB = False
    
    If Not Me.ba Is Nothing Then
        If Not Me.dateColumn1 Is Nothing Then
            If Not Me.pn Is Nothing Then
                If Not Me.description Is Nothing Then
                    If Not Me.sousProjet Is Nothing Then
                        If Not Me.ru Is Nothing Then
                            If Not Me.date1 Is Nothing Then
                                If Not Me.cmnt1 Is Nothing Then
                                    
                                    tmpB = True
                                    
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
    Set dateColumn1 = Nothing
    Set pn = Nothing
    Set description = Nothing
    Set sousProjet = Nothing
    Set ru = Nothing
    Set date1 = Nothing
    Set cmnt1 = Nothing
    
    
    Set ba = Nothing

End Sub

Private Sub Class_Terminate()
    Set dateColumn1 = Nothing
    Set pn = Nothing
    Set description = Nothing
    Set sousProjet = Nothing
    Set ru = Nothing
    Set date1 = Nothing
    Set cmnt1 = Nothing
    
    
    Set ba = Nothing
End Sub

