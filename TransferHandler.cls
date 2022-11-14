VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferHandler"
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


' PerturbationHandler


Implements ISourceProxy


Private ad As IAmpedData
Private m As IMatrix

Private Sub Class_Initialize()
    Set ad = Nothing
    Set m = Nothing
End Sub

Private Sub Class_Terminate()
    Set ad = Nothing
    Set m = Nothing
End Sub



Private Property Set ISourceProxy_AmpedData(ByVal RHS As IAmpedData)
    Set ad = RHS
End Property

Private Property Get ISourceProxy_AmpedData() As IAmpedData
    Set ISourceProxy_AmpedData = ad
End Property

Private Sub ISourceProxy_amplify(dataset As Collection)
    If Not m Is Nothing Then
        
        
        Set ad = New AmpedData
        Set ad.dataset = dataset
        ' ad.setMatrix m
        ad.make m, E_AMP_TRANSFER
    End If
End Sub

Private Property Set ISourceProxy_m(ByVal RHS As IMatrix)
    Set m = RHS
End Property

Private Property Get ISourceProxy_m() As IMatrix
    Set ISourceProxy_m = m
End Property


