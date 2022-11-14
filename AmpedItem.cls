VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AmpedItem"
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


Public plantName As String
Public rawTiming As String
Public dataType As String

Public v As String


Public stdPlt As String
Public stdTiming As String
Public stdYear As String
Public stdMonth As String
Public referenceDate As Date
'
'

Public inputDate As Date
Public resultDate As Date
Public visibleDate As Date

'
'
Public stdDataType As String
Public stdImprovedDataType As String
'
'
' 21-02 (FEB)
Public strMonth As String


' in register datatype defines boundaries
Public yellowTrigger As Double
Public greenTrigger As Double




Public Sub std(whatYouWantToAmp As E_AMP)
    
    innerStd E_STD_PLT
    innerStd E_STD_TIMING, whatYouWantToAmp
    innerStd E_STD_DATATYPE
End Sub


Private Function convertIntoImprovedDataType(dt As String) As String
    convertIntoImprovedDataType = dt
    
    
    ' Debug.Assert dt <> "S1"
    
    Dim rr As Range
    Set rr = ThisWorkbook.Sheets("register").Range("A100")
    
    Do
    
        If CStr(rr.Value) = CStr(dt) Then
            convertIntoImprovedDataType = CStr(rr.offset(0, 1).Value)
            Exit Function
        End If
        Set rr = rr.offset(1, 0)
    Loop Until Trim(rr) = ""
End Function


Private Function becauseInRegisterIsInTheSamePlaceTakeGreenTrigger(dt As String) As Double
    
    
    Dim rr As Range
    Set rr = ThisWorkbook.Sheets("register").Range("A100")
    
    Do
    
        If CStr(rr.Value) = CStr(dt) Then
            On Error Resume Next
            becauseInRegisterIsInTheSamePlaceTakeGreenTrigger = CDbl(rr.offset(0, 3).Value)
            Exit Function
        End If
        Set rr = rr.offset(1, 0)
    Loop Until Trim(rr) = ""
End Function


Private Function becauseInRegisterIsInTheSamePlaceTakeYellowTrigger(dt As String) As Double
    
    
    Dim rr As Range
    Set rr = ThisWorkbook.Sheets("register").Range("A100")
    
    Do
    
        If CStr(rr.Value) = CStr(dt) Then
            On Error Resume Next
            becauseInRegisterIsInTheSamePlaceTakeYellowTrigger = CDbl(rr.offset(0, 4).Value)
            Exit Function
        End If
        Set rr = rr.offset(1, 0)
    Loop Until Trim(rr) = ""
End Function


Private Sub innerStd(e As E_STD, Optional whatYouWantToAmp As E_AMP)

    ' note 2021-01-28
    '
    ' the 3 dates approach - that is why have: Optional whatYouWantToAmp As E_AMP
    '
    
    ' note 2021-02-04
    ' ' new level of abstraction to have inside innerStd just basic name of the instruction

    
    If e = E_STD_DATATYPE Then
        
        
        innerInnerStdDataType
    End If
    
    
    If e = E_STD_PLT Then
        
        
        innerInnerStdPlt
    End If
    
    If e = E_STD_TIMING Then
    
    
        innerInnerStdTiming whatYouWantToAmp
    End If
End Sub



Private Sub innerInnerStdDataType()


    stdDataType = dataType
    stdImprovedDataType = convertIntoImprovedDataType(stdDataType)
    
    
    yellowTrigger = 0
    greenTrigger = 0
    
    yellowTrigger = becauseInRegisterIsInTheSamePlaceTakeYellowTrigger(stdDataType)
    greenTrigger = becauseInRegisterIsInTheSamePlaceTakeGreenTrigger(stdDataType)
    
    On Error Resume Next
    strMonth = CStr(monthToString2(Int(stdMonth), Int(stdYear)))
End Sub


Private Sub innerInnerStdPlt()

    stdPlt = plantName
    
    Dim pltRef As Range
    Set pltRef = ThisWorkbook.Sheets("register").Range(G_PLT_REF)
    
    Do
        If stdPlt Like "*" & pltRef.Value & "*" Then
            stdPlt = pltRef.Value
            Exit Sub
        End If
        
        Set pltRef = pltRef.offset(1, 0)
    Loop Until Trim(pltRef.Value) = ""
    
    
    ' if still check wildcards
    If stdPlt = plantName Then
        
        Set pltRef = ThisWorkbook.Sheets("register").Range(G_PLT_REF)
        Dim wild As String
        
        Do
            ' offset col 5 for wildcard column
            wild = pltRef.offset(0, 3).Value
            If UCase(wild) Like "*" & UCase(stdPlt) & "*" Then
                stdPlt = pltRef.Value
                Exit Sub
            End If
            
            Set pltRef = pltRef.offset(1, 0)
        Loop Until Trim(pltRef.Value) = ""
        
    End If
End Sub


Private Sub innerInnerStdTiming(wywta As E_AMP)
    '
    ' only for timing we are using this opszonal: whatYouWantToAmp
    '
    ' wywta == whatYouWantToAmp
    
    ' because for my mental health I want name "wytwa"
    Dim wytwa As E_AMP
    wytwa = wywta
    
    stdTiming = rawTiming
    
    Dim timingRef As Range
    Set timingRef = ThisWorkbook.Sheets("register").Range(G_TIMING_REF)
    
    Do
        If stdTiming = timingRef.Value Then
            stdTiming = "Y" & Year(Date) & "M" & timingRef.offset(0, -1).Value
            'Exit Sub
        End If
        
        Set timingRef = timingRef.offset(1, 0)
    Loop Until Trim(timingRef.Value) = ""
    
    ' if still check wildcards
    If stdTiming = rawTiming Then
        
        Set timingRef = ThisWorkbook.Sheets("register").Range(G_TIMING_REF)
        Dim wildT As String
        
        Do
            ' offset col 5 for wildcard column
            wildT = timingRef.offset(0, 3).Value
            
            'Debug.Assert "AOÛT" <> stdTiming ' OK
            'Debug.Assert "AOUT" <> stdTiming ' OK
            
            If UCase(wildT) Like "*" & UCase(stdTiming) & "*" Then
                stdTiming = "Y" & Year(Date) & "M" & timingRef.offset(0, -1).Value
                'Exit Sub
            End If
            
            Set timingRef = timingRef.offset(1, 0)
        Loop Until Trim(timingRef.Value) = ""
        
    End If
    
    
    ' Debug.Assert Not stdTiming Like "*Y*"
    
    ' now rest of the timing
    If IsNumeric(stdTiming) = True And Len(stdTiming) = 4 Then
        stdYear = stdTiming
        stdMonth = 1
    ElseIf stdTiming Like "Y????M*" Then
        On Error Resume Next
        stdYear = Mid(stdTiming, 2, 4)
        On Error Resume Next
        stdMonth = Split(stdTiming, "M")(1)
    End If
    
    
    
    ' here we have assigned:
    ' Public stdTiming As String
    ' Public stdYear As String
    ' Public stdMonth As String
    ' stdTiming = YyyyyMmm
    ' stdYear = yyyy
    ' stdMonth = m
    
    ' now to fill:
    'Public referenceDate As Date, Public inputDate As Date
    ' Public resultDate As Date, Public visibleDate As Date
    ' the idea is that reference date will be always visible date
    ' due to legacy download - that in previous kpi data was taken
    ' directly from the interface, which was both back-end and front-end
    
    
    ' really important stuff - reference date will be always connected with visiblity date!
    
    '
    
    ' this is new logic that provides extra dates
    ' to understand better the logic that stays behind transfer logic
    ' which is a little bit tricky
    ' ==============================================================================
    ' whatYouWantToAmp == wywta == wytwa
    layer_3rd_onDiffDatesInsideData wytwa
    
    ' ==============================================================================
End Sub


Private Sub layer_3rd_onDiffDatesInsideData(w As E_AMP)



    ' ' almost "smooth", but now filling muuuuch better with this so called alias!
    ' ' whatYouWantToAmp == wywta == wytwa == w
    
    
    ' now here where is the french magic happening
    ' they tring hard to make everything looks more complicasted
    ' that it really is - funny stuff!
    '
    ' this tabel showing the behaviour of the dates
    ' we have input, result & visiblity date - also
    ' i have my own date called reference - and at
    ' the beginig was really simple becuase i was working
    ' is what you see is what you get WYSIWYG !!!
    ' of course french guys do not understand this simple
    ' rule :)

    '+--------------+-------+---------+-----------+
    '|     kpi      | input |  result | visiblity |
    '+--------------+-------+---------+-----------+
    '| transfer     | I     | III     | III       |
    '| perturbation | III   | II      | III       |
    '| perfo appro  | II    | II      | III       |
    '+--------------+-------+---------+-----------+
    '
    '
    '

    If IsNumeric(stdYear) And IsNumeric(stdMonth) Then
        Me.referenceDate = CDate(DateSerial(Int(stdYear), Int(stdMonth), 1))
        
        ' pertu
        ' whatYouWantToAmp == wywta == wytwa == w   ( just a reminder )
        If (w = E_AMP_PERTURBATION) Then
            Me.visibleDate = CDate(DateSerial(Int(stdYear), Int(stdMonth), 1))
            Me.resultDate = CDate(DateSerial(Int(stdYear), Int(stdMonth) - 1, 1))
            Me.inputDate = CDate(DateSerial(Int(stdYear), Int(stdMonth), 1))
        
        ' transfer
        ElseIf w = E_AMP_TRANSFER Then
            Me.visibleDate = CDate(DateSerial(Int(stdYear), Int(stdMonth), 1))
            Me.resultDate = CDate(DateSerial(Int(stdYear), Int(stdMonth), 1))
            Me.inputDate = CDate(DateSerial(Int(stdYear), Int(stdMonth) - 2, 1))
        
        ' perfo
        ElseIf (w = E_AMP_MANQ) Or (w = E_AMP_RETARD) Then
            Me.visibleDate = CDate(DateSerial(Int(stdYear), Int(stdMonth), 1))
            Me.resultDate = CDate(DateSerial(Int(stdYear), Int(stdMonth) - 1, 1))
            Me.inputDate = CDate(DateSerial(Int(stdYear), Int(stdMonth) - 1, 1))
            
        Else
            Me.referenceDate = CDate(DateSerial(Int(stdYear), Int(stdMonth), 1))
            Me.referenceDate = CDate(DateSerial(Int(stdYear), Int(stdMonth), 1))
            Me.referenceDate = CDate(DateSerial(Int(stdYear), Int(stdMonth), 1))
        End If
    Else
        Me.referenceDate = CDate(DateSerial(Year(Date), 1, 1))
        Me.inputDate = CDate(DateSerial(Year(Date), 1, 1))
        Me.visibleDate = CDate(DateSerial(Year(Date), 1, 1))
        Me.resultDate = CDate(DateSerial(Year(Date), 1, 1))
    End If
End Sub
