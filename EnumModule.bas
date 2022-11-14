Attribute VB_Name = "EnumModule"
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


' type of feeding file
Public Enum E_TOFF
    E_TOFF_UNDEF = 0
    E_TOFF_PERTURBATION
    E_TOFF_PERFO_APPRO
    E_TOFF_TRANSFER_MAJOR_FILE
    E_TOFF_KPI_STOCK
    E_TOFF_EMPTY_1
    E_TOFF_EMPTY_2
    E_TOFF_EMPTY_3
    E_TOFF_EMPTY_4
    E_TOFF_TRANSFER_FILE_URL
End Enum


' transfer files library in MAIN WORKSHEET
Public Enum E_TFL
    E_FTL_PLT = 1
    E_FTL_URL
    E_FTL_STATUS
    E_FTL_RUN
    E_FTL_COMMENT
End Enum


Public Enum E_STD
    E_STD_PLT
    E_STD_TIMING
    E_STD_DATATYPE
End Enum

Public Enum E_AMP
    E_AMP_PERTURBATION
    E_AMP_RETARD
    E_AMP_MANQ
    E_AMP_TRANSFER
    E_AMP_TRANSFER_FILE
    E_AMP_STOCK
End Enum


Public Enum E_DATA
    E_DATA_PLT = 1
    E_DATA_T
    E_DATA_Y
    E_DATA_M
    E_DATA_STR_MONTH
    E_DATA_REF_DATE
    E_DATA_INPUT_DATE
    E_DATA_RESULT_DATE
    E_DATA_VISIBILITY_DATE
    E_DATA_COMMON_DATATYPE_CODE
    E_DATA_SPECIFIC_DATATYPE_CODE
    E_DATA_STR_DATATYPE_CODE
    E_DATA_V
    E_DATA_YELLOW_TRIGGER
    E_DATA_GREEN_TRIGGER
End Enum



' .Cells(wiersz, 1).Value = CDate(ik.photoDate)
'.Cells(wiersz, 1).Value = ik.tb.str
'.Cells(wiersz, 2).Value = ik.tb.sousProjet
'.Cells(wiersz, 3).Value = CDate(ik.tb.MDL)
'.Cells(wiersz, 4).Value = CDate(ik.tb.AMC)
'.Cells(wiersz, 5).Value = CLng(ik.tb.MDL - ik.photoDate)
'.Cells(wiersz, 5).Value = CLng(ik.tb.AMC - ik.photoDate)
'
''
'
'.Cells(wiersz, 5).Value = ik.data.row
'.Cells(wiersz, 6).Value = ik.data.pn
'.Cells(wiersz, 7).Value = ik.data.indice
'.Cells(wiersz, 8).Value = ik.data.sgrFab
'.Cells(wiersz, 9).Value = ik.data.des
'.Cells(wiersz, 10).Value = ik.data.domain
'.Cells(wiersz, 11).Value = ik.data.soreppDecision
'.Cells(wiersz, 12).Value = ik.data.sousProjetColumn
'.Cells(wiersz, 13).Value = ik.data.synteza
Public Enum E_OUT_TR
    E_OUT_TR_PLT = 1
    E_OUT_TR_PHOTO_DATE
    E_OUT_TR_PROJECT
    E_OUT_TR_SOUS_PROJET
    E_OUT_TR_MDL
    E_OUT_TR_AMC
    E_OUT_TR_DIFF1
    E_OUT_TR_DIFF2
    E_OUT_TR_SCOPE
    E_OUT_TR_CW
    E_OUT_TR_MONTH
    E_OUT_TR_ROW
    E_OUT_TR_COL_S_DATE
    E_OUT_TR_COL_S_CW
    E_OUT_TR_COL_S_MONTH
    E_OUT_TR_COL_S_YEAR ' formula here
    E_OUT_TR_ARTICLE
    E_OUT_TR_INDICE
    E_OUT_TR_SGR_FAB
    E_OUT_TR_DESC
    E_OUT_TR_DOMAIN
    E_OUT_TR_DECISION
    E_OUT_TR_SOUS_PROJET_2
    E_OUT_TR_SYNTHESIS
    E_OUT_TR_PRESERIAL
    E_OUT_TR_SERIAL
    E_OUT_TR_BA
    E_OUT_TR_COFOR
End Enum







' SCOPE PERTUBRATION NORMALISATION
' -------------------------------------------------------------------
' -------------------------------------------------------------------
''
'

Public Enum E_PERTU_LIST
    E_PERTU_LIST_SITE = 2
    E_PERTU_LIST_DATE
    E_PERTU_LIST_VEH
    E_PERTU_LIST_DAN
    E_PERTU_LIST_PERTURBATION
    E_PERTU_LIST_ACCOUNTING
    E_PERTU_LIST_CAUSE
    E_PERTU_LIST_RESP
    E_PERTU_LIST_ORIGIN_BOM
    E_PERTU_LIST_NUM_OF_AFFECTED_CARS
    E_PERTU_LIST_5WHY
End Enum


Public Enum E_PERTU_SRC_TYPE
    E_PERTU_SRC_TYPE_0
End Enum



Public Enum E_PERTU_OUT_LIST
    E_PERTU_OUT_LIST_SITE = 1
    E_PERTU_OUT_LIST_STR_DATE
    E_PERTU_OUT_LIST_DATE
    E_PERTU_OUT_LIST_VEH
    E_PERTU_OUT_LIST_DAN
    E_PERTU_OUT_LIST_PERTURBATION
    E_PERTU_OUT_LIST_ACCOUNTING
    E_PERTU_OUT_LIST_BACC
    E_PERTU_OUT_LIST_CAUSE
    E_PERTU_OUT_LIST_RESP
    E_PERTU_OUT_LIST_ORIGIN_BOM
    E_PERTU_OUT_LIST_NUM_OF_AFFECTED_CARS
    E_PERTU_OUT_LIST_5WHY
    E_PERTU_OUT_LIST_B_5WHY
    E_PERTU_OUT_LIST_Q
    E_PERTU_OUT_LIST_STR_MONTH
End Enum

'
''
' -------------------------------------------------------------------
' -------------------------------------------------------------------
