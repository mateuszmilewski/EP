VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DefineRange 
   Caption         =   "UserForm1"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5235
   OleObjectBlob   =   "DefineRange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DefineRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public lr As Range
Public cr As Range


Private Sub BtnAdd_Click()

    
    Set lr = Range(Me.RefEditLabels.Value)
    Set cr = Range(Me.RefEditContent.Value)
    
    Debug.Print lr.columnS.Count & " " & lr.Address
    Debug.Print cr.columnS.Count & " " & cr.Address
    
    Debug.Print lr.Parent.name & " " & cr.Parent.name
    
    
    If lr.Parent.name = cr.Parent.name Then
        If lr.columnS.Count = cr.columnS.Count Then
        
        
            Dim rf As Range
            Set rf = ThisWorkbook.Sheets("RAW_DATA").Range("A1048576").End(xlUp).offset(4, 0)
            Debug.Print rf.Address
            
            
            rf.Value = "ptrn" & "" & CStr(Me.ComboBox1.Value)
            
            Dim ir As Range, iter As Integer, c As Range, line As Integer
            iter = 0
            For Each ir In lr
                If iter = 0 And ir.Value = "" Then
                    rf.offset(1, iter).Value = "PLT"
                Else
                    rf.offset(1, iter).Value = ir.Value
                End If
                
                iter = iter + 1
            Next ir
            
            
            
            
            Set rf = rf.offset(2, 0)
            
            
            iter = 0
            line = 0
            ' ' '
            For Each ir In cr.Rows
                
                
                For Each c In ir.columnS
                    rf.offset(line, iter).Value = c.Value
                    iter = iter + 1
                Next c
                
                iter = 0
                line = line + 1
            Next ir
            
        Else
            MsgBox "addresses from labels and content do not match!", vbCritical
        End If
    Else
        MsgBox "Labels and content need to be from same sheet!", vbCritical
    End If
End Sub
