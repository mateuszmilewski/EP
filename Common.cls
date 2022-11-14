VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function tryToRenameIt(sh1 As Worksheet, pf As String, Optional xtraChar As String) As String


    ' sh1 is sheet for change name
    ' pf is prefix for new name
    
    tryToRenameIt = sh1.name
    
    Dim tmp As String, yyyymmdd As Long
    
    yyyymmdd = CLng(Year(Date)) * 10000 + CLng(Month(Date)) * 100 + CLng(Day(Date))
    
    tmp = "" & CStr(pf) & CStr(yyyymmdd) & "_" & xtraChar
    
    
    Dim alreadyExisting As Boolean
    alreadyExisting = False
    
    Dim ish As Worksheet
    For Each ish In ThisWorkbook.Sheets
        
        If ish.name = tmp Then
            alreadyExisting = True
            Exit For
        End If
    Next ish
    
    If alreadyExisting Then
        tmp = tryToRenameIt(sh1, pf, xtraChar & "x")
    End If
    
    On Error Resume Next
    tryToRenameIt = tmp
    
End Function


Public Sub colorSheetTab(sh1 As Worksheet, et As XlThemeColor)
    
    
    With sh1.Tab
        .ThemeColor = et
        .TintAndShade = 0
    End With
End Sub





Public Function collectMatrix(sh1 As Worksheet, ptrn As String) As IMatrix
    
    
    Set collectMatrix = Nothing
    
    Dim r As Range
    Set r = sh1.Range("A10000")
    
    
    
    Do
        If r.Value = "ptrn" & CStr(ptrn) Then
        
            Dim tmpM As IMatrix
            Set tmpM = New Matrix
            
            tmpM.initialFill r.offset(1, 0)
            tmpM.simpleMake
            tmpM.name = ptrn
            Set collectMatrix = tmpM
            Exit Function
            
        End If
        
        Set r = r.offset(-1, 0)
    Loop While r.row > 1
End Function
