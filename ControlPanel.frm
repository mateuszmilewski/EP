VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControlPanel 
   Caption         =   "Go To"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2265
   OleObjectBlob   =   "ControlPanel.frx":0000
End
Attribute VB_Name = "ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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


Private shToActivate As Worksheet



Private Sub ImageMAIN_Click()
    Set shToActivate = ThisWorkbook.Sheets("MAIN")
    shToActivate.Activate
    shToActivate.Cells(1, 1).Select
End Sub


Private Sub ImageDASHBOARD_Click()
    Set shToActivate = ThisWorkbook.Sheets("PIVOTS")
    shToActivate.Activate
    shToActivate.Cells(1, 1).Select
End Sub


Private Sub ImagePerfoAppro_Click()
    Set shToActivate = Nothing
    On Error Resume Next
    Set shToActivate = ThisWorkbook.Sheets(EP.G_PERFO_APPRO_SH_NM_MANQUANTS)
    
    If shToActivate Is Nothing Then
        For Each shToActivate In ThisWorkbook.Sheets
            If UCase(shToActivate.name) Like "*" & UCase(EP.G_PERFO_APPRO_SH_NM_MANQUANTS) & "*" Then
                shToActivate.Activate
                shToActivate.Cells(1, 1).Select
                Exit For
            End If
        Next shToActivate
    Else
        shToActivate.Activate
        shToActivate.Cells(1, 1).Select
    End If
End Sub

Private Sub ImagePertu1_Click()
    Set shToActivate = Nothing
    On Error Resume Next
    Set shToActivate = ThisWorkbook.Sheets(EP.G_PERTURBATION_SH_NM_DONNES_PERT)
    
    If shToActivate Is Nothing Then
        For Each shToActivate In ThisWorkbook.Sheets
            If UCase(shToActivate.name) Like "*" & UCase(EP.G_PERTURBATION_SH_NM_DONNES_PERT) & "*" Then
                shToActivate.Activate
                shToActivate.Cells(1, 1).Select
                Exit For
            End If
        Next shToActivate
    Else
        shToActivate.Activate
        shToActivate.Cells(1, 1).Select
    End If
End Sub

Private Sub ImagePertu2_Click()
    Set shToActivate = Nothing
    On Error Resume Next
    Set shToActivate = ThisWorkbook.Sheets(EP.G_PERTURBATION_SH_NM_PERTURBATION)
    
    If shToActivate Is Nothing Then
        For Each shToActivate In ThisWorkbook.Sheets
            If UCase(shToActivate.name) Like "*" & UCase(EP.G_PERTURBATION_SH_NM_PERTURBATION) & "*" Then
                shToActivate.Activate
                shToActivate.Cells(1, 1).Select
                Exit For
            End If
        Next shToActivate
    Else
        shToActivate.Activate
        shToActivate.Cells(1, 1).Select
    End If
End Sub


Private Sub ImageStock_Click()
    Set shToActivate = Nothing
    On Error Resume Next
    Set shToActivate = ThisWorkbook.Sheets(EP.G_KPI_STOCK_SH_NM_BILAN)
    
    If shToActivate Is Nothing Then
        For Each shToActivate In ThisWorkbook.Sheets
            If UCase(shToActivate.name) Like "*" & UCase(EP.G_KPI_STOCK_SH_NM_BILAN) & "*" Then
                shToActivate.Activate
                shToActivate.Cells(1, 1).Select
                Exit For
            End If
        Next shToActivate
    Else
        shToActivate.Activate
        shToActivate.Cells(1, 1).Select
    End If
End Sub

Private Sub ImageTransfer_Click()
    Set shToActivate = Nothing
    On Error Resume Next
    Set shToActivate = ThisWorkbook.Sheets(EP.G_KPI_MAJOR_TRANSFER_SH_NM_DONNES_2)
    
    If shToActivate Is Nothing Then
        For Each shToActivate In ThisWorkbook.Sheets
            If UCase(shToActivate.name) Like "*" & UCase(EP.G_KPI_MAJOR_TRANSFER_SH_NM_DONNES_2) & "*" Then
                shToActivate.Activate
                shToActivate.Cells(1, 1).Select
                Exit For
            End If
        Next shToActivate
    Else
        shToActivate.Activate
        shToActivate.Cells(1, 1).Select
    End If
End Sub
