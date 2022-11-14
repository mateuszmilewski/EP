VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2021 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Private ppApp As PowerPoint.Application
Private pp As Presentation
Private ppSlide As PowerPoint.Slide
Private myCustomLayout As PowerPoint.CustomLayout

Private srcSh As Worksheet

Private st As Object
Private sz As Shape
Private shr As ShapeRange
Private srcP As Presentation
Private p1 As PowerPoint.Application
Private pp1 As Presentation


' slide sizes
' ------------------------------------------
Private s_id As Variant
Private s_h As Variant
Private s_w As Variant
' ------------------------------------------

' title1 position
' ================================================
Private sz_pos_top As Variant
Private sz_pos_left As Variant
' ================================================


Private s As Slide
Private currentSlide As PowerPoint.Slide
Private gobj As Object


Private templateFilename As String


Public Function getCurrentSlide() As PowerPoint.Slide
    Set getCurrentSlide = currentSlide
End Function


Public Sub justCopyToNewSlideAsItIs(sz As Shape, title1 As String, mTop As Integer, mLeft As Integer)

    sz.Copy
    
    Set currentSlide = Me.getPP1.Slides(1).Duplicate(Me.getPP1.Slides.Count)
    currentSlide.Select
    currentSlide.Shapes("Title1").TextFrame2.TextRange.Text = title1
    
    Set gobj = currentSlide.Shapes.PasteSpecial(ppPasteJPG)
    'Debug.Print "TypeName(gobj): " & TypeName(gobj)
    'Debug.Print "sz.name: " & sz.name
    
    gobj.Top = mTop
    gobj.Left = mLeft
End Sub


Public Sub copyToNewSlide(refRng As Range, slideTitle As String, mTop As Integer, mLeft As Integer)

    If pp1 Is Nothing Then
        Set EP.G_CURRENT_PRESENTATION = getPresentationBasedOnRegisterMetaData()
        ppApp.Activate
        Set pp1 = EP.G_CURRENT_PRESENTATION
    End If
    
    If Not pp1 Is Nothing Then
        
        ' main logic body
        ' ----------------------------------------------------------------------------
        ''
        '
        
        
        Debug.Print "pp1 slides: " & pp1.Slides.Count & " " & TypeName(pp1)
        ' Set currentSlide = pp1.Slides(1).Duplicate
        Set currentSlide = pp1.Slides(1).Duplicate(1)
        currentSlide.Select
        currentSlide.Shapes("Title1").TextFrame2.TextRange.Text = CStr(slideTitle)
        
        currentSlide.MoveTo pp1.Slides.Count
        
        Dim area As Range
        
        Set area = defineArea(refRng)
        If Not area Is Nothing Then
            area.Copy
        End If
        
        Set gobj = currentSlide.Shapes.PasteSpecial(ppPasteDefault)
        Debug.Print "TypeName(gobj): " & TypeName(gobj)
        Debug.Print "sz.name: " & sz.name
        gobj.Top = mTop
        gobj.Left = mLeft
        
        
        ' ----------------------------------------------------------------------------
    End If
End Sub

Public Sub copyToCurrentSlide(refRng As Range, mTop As Integer, mLeft As Integer)

        Dim area As Range
        
        Set area = defineArea(refRng)
        If Not area Is Nothing Then
            area.Copy
        End If
        
        
        Set gobj = currentSlide.Shapes.PasteSpecial(ppPasteDefault)
        Debug.Print "TypeName(gobj): " & TypeName(gobj)
        Debug.Print "sz.name: " & sz.name
        

        gobj.Top = mTop
        gobj.Left = mLeft
End Sub


Public Sub copyShapeToCurrentSlide(sz As Shape, mTop As Integer, mLeft As Integer)

        sz.Copy
        
        
        Set gobj = currentSlide.Shapes.PasteSpecial(ppPasteDefault)
        Debug.Print "TypeName(gobj): " & TypeName(gobj)
        Debug.Print "sz.name: " & sz.name
        

        gobj.Top = mTop
        gobj.Left = mLeft
End Sub


Private Function defineArea(rr As Range) As Range
    Set defineArea = Nothing
    
    Dim tl As Range, br As Range
    Set tl = rr
    
    
    Dim tmp As Range, mostRightColumn As Integer
    mostRightColumn = rr.Column
    
    Set tmp = rr
    
    Do
        Do
        
            Set tmp = tmp.offset(0, 1)
        Loop Until Trim(tmp.Value) = ""
        
        If Int(tmp.Column) > mostRightColumn Then
            mostRightColumn = Int(tmp.Column)
        End If
        
        Set tmp = tmp.offset(1, 0)
        Set tmp = tmp.Parent.Cells(tmp.row, rr.Column)
        
    Loop Until (Trim(tmp.Value) = "") Or (Trim(tmp.Value) = "!x")
    
    
    Set br = tmp.Parent.Cells(tmp.row - 1, mostRightColumn - 1)
    
    Set defineArea = rr.Parent.Range(tl, br)
    
    Debug.Print "defineArea: " & defineArea.Address
    
End Function

Public Sub copyToExisting(sz As Shape, title1 As String, mTop As Integer, mLeft As Integer)
    
    If pp1 Is Nothing Then
        Set EP.G_CURRENT_PRESENTATION = getPresentationBasedOnRegisterMetaData()
        ppApp.Activate
        Set pp1 = EP.G_CURRENT_PRESENTATION
    End If
    
    If Not pp1 Is Nothing Then
        ' now it should be fine
        sz.Copy
        
        Debug.Print "pp1 slides: " & pp1.Slides.Count & " " & TypeName(pp1)
        ' Set currentSlide = pp1.Slides(1).Duplicate
        Set currentSlide = pp1.Slides(1).Duplicate(1)
        currentSlide.Select
        currentSlide.Shapes("Title1").TextFrame2.TextRange.Text = title1
        
        Set gobj = currentSlide.Shapes.PasteSpecial(ppPasteJPG)
        'Debug.Print "TypeName(gobj): " & TypeName(gobj)
        'Debug.Print "sz.name: " & sz.name
        
        gobj.Top = mTop
        gobj.Left = mLeft
    Else
        Debug.Print "There is still sth wrong with the presentation!"
    End If
End Sub

Private Function getPresentationBasedOnRegisterMetaData() As PowerPoint.Presentation
    Set getPresentationBasedOnRegisterMetaData = Nothing
    
    
    Dim rsh As Worksheet
    Set rsh = ThisWorkbook.Sheets("register")
    
    Set ppApp = Nothing
    
    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    
    If Not ppApp Is Nothing Then
        For Each pp In ppApp.Presentations
            
            If CStr(pp.name) = CStr(rsh.Range("presentationName").Value) Then
                If CStr(pp.FullName) = CStr(rsh.Range("presentationName").offset(1, 0).Value) Then
                    
                    Set getPresentationBasedOnRegisterMetaData = pp
                    Exit Function
                    
                End If
            End If
        Next pp
    End If
    
    
End Function

Public Function getPP1() As Presentation

    If Not pp1 Is Nothing Then
        Set getPP1 = pp1
    Else
        MsgBox "No presentation avail.!", vbCritical
    End If
End Function



Private Sub Class_Initialize()
    Set srcSh = ThisWorkbook.Sheets("stellaPP_template1")
    srcSh.Visible = xlSheetVisible
    srcSh.Activate
    templateFilename = "StellaTemplate1"
    
    Set pp1 = Nothing
    Set EP.G_CURRENT_PRESENTATION = Nothing
    
End Sub

Private Sub init()
    Set srcSh = ThisWorkbook.Sheets("stellaPP_template1")
    srcSh.Visible = xlSheetVisible
    srcSh.Activate
    templateFilename = "StellaTemplate1"
    
    Set pp1 = Nothing
    Set EP.G_CURRENT_PRESENTATION = Nothing
End Sub


Public Sub connectWithSourceAndCreatNewInstance()
    
    
    ' Set st = ThisWorkbook.Sheets("Sheet1").Shapes("StellaTempalte1")
    For Each sz In srcSh.Shapes
        Debug.Print sz.name
        
        If sz.name = templateFilename Then
            Set st = sz
            Exit For
        End If
    Next sz
    
    st.Select
    Selection.Verb Verb:=xlOpen
    
    
    Set srcP = Selection.Object
    
    Set p1 = New PowerPoint.Application
    
    Set pp1 = p1.Presentations.Add
    
    With ThisWorkbook.Sheets("register")
        .Range("presentationName").Value = CStr(pp1.name)
        .Range("presentationName").offset(1, 0).Value = CStr(pp1.FullName)
        ' .Range("presentationName").offset(2, 0).Value = CStr(pp1.Tags.Value(1))
    End With
    Set EP.G_CURRENT_PRESENTATION = pp1
    

    Set s = srcP.Slides(1)
    s.Copy
    pp1.Slides.Paste
    
    srcP.Close
    
    p1.Activate
    
    
    ' init some size and layout figures!
    ' ---------------------------------------------------------------------
    
    'new presentation is created - check now parameters!
    initSomeSizeAndLayoutFor p1, pp1
    
    ' ---------------------------------------------------------------------
End Sub

Private Sub initSomeSizeAndLayoutFor(ByRef mp As PowerPoint.Application, ByRef mpp As Presentation)
    
    s_h = mpp.PageSetup.SlideHeight
    s_w = mpp.PageSetup.SlideWidth
    
    Set s = mpp.Slides(1)
    Set szr = s.Shapes("Title1")
    
    sz_pos_top = sz.Top
    sz_pos_left = sz.Left
    
    
End Sub


Private Sub Class_Terminate()
    Set pp1 = Nothing
    Set EP.G_CURRENT_PRESENTATION = Nothing
End Sub
