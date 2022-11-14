Attribute VB_Name = "Module1"
Option Explicit

Sub Macro2PertuPivot()
Attribute Macro2PertuPivot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2PertuPivot Macro
'

'
    With ActiveWorkbook.SlicerCaches("Slicer_STR_MONTH2")
        .SlicerItems("21-01 (JAN)").Selected = True
        .SlicerItems("10-00()").Selected = False
        .SlicerItems("11-00()").Selected = False
        .SlicerItems("12-00()").Selected = False
        .SlicerItems("13-00()").Selected = False
        .SlicerItems("14-00()").Selected = False
        .SlicerItems("15-00()").Selected = False
        .SlicerItems("16-00()").Selected = False
        .SlicerItems("17-00()").Selected = False
        .SlicerItems("18-00()").Selected = False
        .SlicerItems("19-00()").Selected = False
        .SlicerItems("20-00()").Selected = False
        .SlicerItems("20-01 (JAN)").Selected = False
        .SlicerItems("20-02 (FEB)").Selected = False
        .SlicerItems("20-03 (MAR)").Selected = False
        .SlicerItems("20-04 (APR)").Selected = False
        .SlicerItems("20-05 (MAY)").Selected = False
        .SlicerItems("20-06 (JUN)").Selected = False
        .SlicerItems("20-07 (JUL)").Selected = False
        .SlicerItems("20-09 (SEP)").Selected = False
        .SlicerItems("20-10 (OCT)").Selected = False
        .SlicerItems("20-11 (NOV)").Selected = False
        .SlicerItems("20-12 (DEC)").Selected = False
        .SlicerItems("21-02 (FEB)").Selected = False
        .SlicerItems("21-03 (MAR)").Selected = False
        .SlicerItems("MAR").Selected = False
        .SlicerItems("20-08 (AUG)").Selected = False
        .SlicerItems("21-04 (APR)").Selected = False
        .SlicerItems("21-05 (MAY)").Selected = False
        .SlicerItems("21-06 (JUN)").Selected = False
        .SlicerItems("21-07 (JUL)").Selected = False
        .SlicerItems("21-08 (AUG)").Selected = False
        .SlicerItems("21-09 (SEP)").Selected = False
        .SlicerItems("21-10 (OCT)").Selected = False
        .SlicerItems("21-11 (NOV)").Selected = False
        .SlicerItems("21-12 (DEC)").Selected = False
        .SlicerItems("APR").Selected = False
        .SlicerItems("AUG").Selected = False
        .SlicerItems("DEC").Selected = False
        .SlicerItems("FEB").Selected = False
        .SlicerItems("JAN").Selected = False
        .SlicerItems("JUL").Selected = False
        .SlicerItems("JUN").Selected = False
        .SlicerItems("MAY").Selected = False
        .SlicerItems("NOV").Selected = False
        .SlicerItems("OCT").Selected = False
        .SlicerItems("SEP").Selected = False
    End With
    With ActiveSheet.PivotTables("PivotTablePertu1").PivotFields("STR_MONTH")
        .PivotItems("21-02 (FEB)").Visible = True
    End With
End Sub
