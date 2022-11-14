VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
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


Implements IMatrix


Private id As String
Private nm As String
Private src As Worksheet

Private tl As Range
Private tr As Range

Private bl As Range
Private br As Range


Private a As Range



Private Sub Class_Initialize()
    Set tl = Nothing
    Set tr = Nothing
    
    Set bl = Nothing
    Set br = Nothing
    
    Set a = Nothing
    
End Sub

Private Sub Class_Terminate()
    Set tl = Nothing
    Set tr = Nothing
    
    Set bl = Nothing
    Set br = Nothing
    
    Set a = Nothing
End Sub

Private Property Set IMatrix_area(ByVal RHS As Range)
    Set a = RHS
End Property

Private Property Get IMatrix_area() As Range
    Set IMatrix_area = a
End Property

Private Property Set IMatrix_bottomLeft(ByVal RHS As Range)
    Set bl = RHS
End Property

Private Property Get IMatrix_bottomLeft() As Range
    Set IMatrix_bottomLeft = bl
End Property

Private Property Set IMatrix_bottomRight(ByVal RHS As Range)
    Set br = RHS
End Property

Private Property Get IMatrix_bottomRight() As Range
    Set IMatrix_bottomRight = br
End Property

Private Property Let IMatrix_id(ByVal RHS As String)
    id = RHS
End Property

Private Property Get IMatrix_id() As String
    IMatrix_id = id
End Property

Private Sub IMatrix_initialFill(r As Range)
    
    Set tl = r
    Set tr = r
    Set bl = r
    Set br = r
    
    Set a = r
    
    Set src = r.Parent
End Sub

Private Function IMatrix_isEmpty() As Boolean
    
    If tl Is Nothing Then
        If tr Is Nothing Then
            If bl Is Nothing Then
                If br Is Nothing Then
                    If a Is Nothing Then
                        IMatrix_isEmpty = True
                    End If
                End If
            End If
        End If
    End If
                
End Function

Private Property Let IMatrix_name(ByVal RHS As String)
    nm = RHS
End Property

Private Property Get IMatrix_name() As String
    IMatrix_name = nm
End Property

Private Sub IMatrix_simpleMake()
    Set tr = tl.End(xlToRight)
    Set bl = tl.End(xlDown)
    Set br = src.Cells(bl.row, tr.Column)
    
    
    Set a = src.Range(tl, br)
    
    Debug.Print "a adr simple Make: " & a.Address
End Sub

Private Property Set IMatrix_src(ByVal RHS As Worksheet)
    Set src = RHS
End Property

Private Property Get IMatrix_src() As Worksheet
    Set IMatrix_src = src
End Property

Private Property Set IMatrix_topLeft(ByVal RHS As Range)
    Set tl = RHS
End Property

Private Property Get IMatrix_topLeft() As Range
    Set IMatrix_topLeft = tl
End Property

Private Property Set IMatrix_topRight(ByVal RHS As Range)
    Set tr = RHS
End Property

Private Property Get IMatrix_topRight() As Range
    Set IMatrix_topRight = tr
End Property
