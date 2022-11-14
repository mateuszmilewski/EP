VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerturbationListItem"
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


' PerturbationListItem

Public id As Integer
Public site As String
Public strDate As String
Public strMonth As String
Public d As Date
Public veh As String
Public dan As String
Public perturbation As String

' accounting
' -----------------------------------
Public accounting As String
Public bacc As Boolean
' -----------------------------------

Public causePerturbation As String
Public responsibility As String
Public Disturbance_origin_BOM As String
Public str_Nr_of_affected_cars As String
Public int_Nr_of_affected_cars As Long

' -----------------------------------
Public b5w As Boolean
Public fivewhy As String
Public fivewhyDescription As String
' -----------------------------------

Public comment As String

Public log As String
Public quality As Integer

Public eSrcType As E_PERTU_SRC_TYPE


Public Sub calculateQuality()
    
    If site <> "" Then incq
    If strDate <> "" Then incq
    If Me.veh <> "" Then incq
    If Me.dan <> "" Then incq
    If Me.perturbation <> "" Then incq
    If Me.accounting <> "" Then incq
    If Me.causePerturbation <> "" Then incq
    If Me.responsibility <> "" Then incq
    If Me.Disturbance_origin_BOM <> "" Then incq
    If Me.str_Nr_of_affected_cars <> "" Then incq
    If Me.b5w Then incq
    
    If CLng(Me.d) = 0 Then
        decreaseBecuaseWrongDate
    End If
    
End Sub

Private Sub decreaseBecuaseWrongDate()
    quality = quality - 5
End Sub

Private Sub incq()
    quality = quality + 1
End Sub

Private Sub Class_Initialize()
    quality = 0
End Sub
