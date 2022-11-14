Attribute VB_Name = "GlobalModule"
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



' global fetch on current presention
' ------------------------------------------------------------------
Global G_CURRENT_PRESENTATION As PowerPoint.Presentation



' all about perturbation stuff
' ------------------------------------------------------------------
Global Const G_PERTURBATION_SH_NM_DONNES_PERT = "Donnés Pert"
Global Const G_PERTURBATION_SH_NM_PERTURBATION = "Perturbations"
' ------------------------------------------------------------------
' normalized data
Global Const G_SH_NM_NORM_PERTU_LIST = "NormPertuList"
' ------------------------------------------------------------------
' ------------------------------------------------------------------



' register global
' -------------------------------------------------------
Global Const G_PLT_REF = "N2"
Global Const G_TIMING_REF = "E2"
' -------------------------------------------------------


' -------------------------------------------------------
Global Const G_PERFO_APPRO_SH_NM_RETARD = "Retard"
Global Const G_PERFO_APPRO_SH_NM_MANQUANTS = "Manquants"
' -------------------------------------------------------



' -------------------------------------------------------
Global Const G_KPI_MAJOR_TRANSFER_SH_NM_DONNES = "Donnes_Data 2020"
Global Const G_KPI_MAJOR_TRANSFER_SH_NM_IMPACT = "Impact Cout_Cost Impact 2020"
Global Const G_KPI_MAJOR_TRANSFER_SH_NM_DONNES_2 = "Donnes_Data 2021"
Global Const G_KPI_MAJOR_TRANSFER_SH_NM_IMPACT_2 = "Impact Cout_Cost Impact 2021"
' -------------------------------------------------------



' -------------------------------------------------------

Global Const G_KPI_STOCK_SH_NM_TABLE = "Table"
Global Const G_KPI_STOCK_SH_NM_LISTE = "Liste"
Global Const G_KPI_STOCK_SH_NM_BILAN = "Bilan"

' -------------------------------------------------------



' -------------------------------------------------------
Global Const G_KPI_TRANSFER_FILE_SH_NM_TABLE = "Table"
Global Const G_KPI_TRANSFER_FILE_SH_NM_LISTES = "Listes"
Global Const G_KPI_TRANSFER_FILE_SH_NM_ARCHIVE = "Archive"
' -------------------------------------------------------


Function getSourceWorksheetNames(ofst As Integer) As Variant

    Set getSourceWorksheetNames = Nothing

    If ofst = E_TOFF_PERTURBATION Then
        getSourceWorksheetNames = Array(G_PERTURBATION_SH_NM_DONNES_PERT, _
            G_PERTURBATION_SH_NM_PERTURBATION)
    End If
    
    
    If ofst = EP.E_TOFF.E_TOFF_PERFO_APPRO Then
        getSourceWorksheetNames = Array(G_PERFO_APPRO_SH_NM_RETARD, _
            G_PERFO_APPRO_SH_NM_MANQUANTS)
    End If
        
    If ofst = EP.E_TOFF.E_TOFF_TRANSFER_MAJOR_FILE Then
        getSourceWorksheetNames = Array(G_KPI_MAJOR_TRANSFER_SH_NM_DONNES, _
            G_KPI_MAJOR_TRANSFER_SH_NM_IMPACT, _
            G_KPI_MAJOR_TRANSFER_SH_NM_DONNES_2, _
            G_KPI_MAJOR_TRANSFER_SH_NM_IMPACT_2)
    End If
    
    If ofst = EP.E_TOFF_KPI_STOCK Then
        getSourceWorksheetNames = Array(G_KPI_STOCK_SH_NM_TABLE, _
            G_KPI_STOCK_SH_NM_LISTE, G_KPI_STOCK_SH_NM_BILAN)
    End If
    
    If ofst >= EP.E_TOFF_TRANSFER_FILE_URL Then
        getSourceWorksheetNames = Array(G_KPI_TRANSFER_FILE_SH_NM_TABLE, _
            G_KPI_TRANSFER_FILE_SH_NM_LISTES, G_KPI_TRANSFER_FILE_SH_NM_ARCHIVE)
    End If
End Function


Public Function monthToString(r As Range, yr As Range) As Variant
    monthToString = ""
    
    If IsNumeric(r.Value) Then
    
        Dim ref As Range
        Set ref = ThisWorkbook.Sheets("register").Range("G1")
        
        Dim i As Integer
        i = 100 + Int(r.Value)
    
    
        If Int(r.Value) = 0 Then
            monthToString = Right(CStr(yr.Value), 2) & "-" & CStr(Right(CStr(i), 2)) & "()"
        Else
            monthToString = Right(CStr(yr.Value), 2) & "-" & CStr(Right(CStr(i), 2)) & " (" & CStr(ref.offset(Int(r.Value), 0).Value) & ")"
        End If
    Else
        monthToString = "0"
    End If
    
End Function


Public Function monthToString2(mm As Integer, yyyy As Integer) As Variant
    monthToString2 = ""
    
    If IsNumeric(mm) Then
    
        Dim ref As Range
        Set ref = ThisWorkbook.Sheets("register").Range("G1")
        
        Dim i As Integer
        i = 100 + Int(mm)
    
    
        If Int(mm) = 0 Then
            monthToString2 = Right(CStr(yyyy), 2) & "-" & CStr(Right(CStr(i), 2)) & "()"
        Else
            monthToString2 = Right(CStr(yyyy), 2) & "-" & CStr(Right(CStr(i), 2)) & " (" & CStr(ref.offset(Int(mm), 0).Value) & ")"
        End If
    Else
        monthToString2 = "0"
    End If
    
End Function



Public Function calcUnSpecial(param As Variant) As Double
    calcUnSpecial = 1#
    
    Dim regRef As Range
    Set regRef = ThisWorkbook.Sheets("register").Range("UN_REF")
    
    Do
        If CStr(regRef.Value) = CStr(param) Then
            calcUnSpecial = CDbl(regRef.offset(0, 1).Value)
            Exit Do
        End If
        Set regRef = regRef.offset(1, 0)
    Loop Until Trim(regRef.Value) = ""
End Function
