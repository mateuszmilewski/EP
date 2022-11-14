VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControlPanel2 
   Caption         =   "ControlPanel"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3480
   OleObjectBlob   =   "ControlPanel2.frx":0000
End
Attribute VB_Name = "ControlPanel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private shToActivate As Worksheet

Private Sub BtnDonnesData_Click()
    Set shToActivate = Nothing
    On Error Resume Next
    Set shToActivate = ThisWorkbook.Sheets(EP.G_KPI_MAJOR_TRANSFER_SH_NM_DONNES)
    
    If shToActivate Is Nothing Then
        For Each shToActivate In ThisWorkbook.Sheets
            If UCase(shToActivate.name) Like "*" & UCase(EP.G_KPI_MAJOR_TRANSFER_SH_NM_DONNES) & "*" Then
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

Private Sub BtnDonnesPert_Click()

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

Private Sub BtnImpact_Click()
    Set shToActivate = Nothing
    On Error Resume Next
    Set shToActivate = ThisWorkbook.Sheets(EP.G_KPI_MAJOR_TRANSFER_SH_NM_IMPACT)
    
    If shToActivate Is Nothing Then
        For Each shToActivate In ThisWorkbook.Sheets
            If UCase(shToActivate.name) Like "*" & UCase(EP.G_KPI_MAJOR_TRANSFER_SH_NM_IMPACT) & "*" Then
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

Private Sub BtnMain_Click()
    Set shToActivate = ThisWorkbook.Sheets("MAIN")
    shToActivate.Activate
    shToActivate.Cells(1, 1).Select
End Sub



Private Sub BtnManq_Click()
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

Private Sub BtnPerturbation_Click()
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

Private Sub BtnRetard_Click()


    Set shToActivate = Nothing
    On Error Resume Next
    Set shToActivate = ThisWorkbook.Sheets(EP.G_PERFO_APPRO_SH_NM_RETARD)
    
    If shToActivate Is Nothing Then
        For Each shToActivate In ThisWorkbook.Sheets
            If UCase(shToActivate.name) Like "*" & UCase(EP.G_PERFO_APPRO_SH_NM_RETARD) & "*" Then
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

Private Sub UserForm_Initialize()
    Me.Top = 10 ' Application.Top + (Application.UsableHeight) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.UsableWidth) - (Me.Width) - 10
End Sub
