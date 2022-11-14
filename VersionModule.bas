Attribute VB_Name = "VersionModule"
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



'v001
' -
' EP - stands for 3P - THREE PIECES
' rep based on other 3 reports but synthesis that we all like!
'
' main class EPC - 3 pieces class - the main class!
' -


' v002
' very basic version that downloading all data from doc-info


' v003
' starting point for unpivot stuff


' 004
' initial concat for p1 pattern and p2 pattern

' v005
' pattern for flat table: PLT, TIMING, DATATYPE, V

' v006
' extra fields for % data triggers for yellow and green


' v010
' version before was big bug!


'v011
' on dashboard extra button to auto scroll + protect dashboard sheet


' v012
' std in amped data and item diff on type of KPI to adjust
'
' Public Sub std(whatYouWantToAmp As E_AMP) AmpedItem - know have knowledge on which KPI we are
' so std sub have extra param ENUM!


' v013
' re-test 2021-01-31

'v014
' save backup copy on change - new dates (all 4): ref, in, result, visiblity
' amped item std -> innerStd particullary for the date
' as a good starting point to make more DRY approach!


'v015
' transfer file scaffold!



'v016
' in version 015 first succ on taking data from one transfer file - copy paste worksheet - but...
' im not sure if we want to copy all the data like this 'cause file will be heavy as hell..
' after the  transfer file is open there should be already some operations connected with optim of data
' so... for transfer file there will be copy worksheets into thisworkbook but after operations
' we need some deletion maneuver



'017
' backup....
' transfer file new handlers
' big change on apporach - when data from transfer file we are not coping all worksheets
' too much data - already providing some operations on DocInfoHandler already
' transferFileList classes


'018
' issue with loop in da loop - need to change'
' change in takeDataFromListingAndBindEverything - strange stuff inside transfer file list


'019
' transfer items general step closure
' issue with items without connection with sous projet closed
' you will see  anyway on raw list but q if should be in one pager


'020
' build back-up
' spread the foundation of the project - data need to have better quality before
' ill go with marriage stuff with input and output with blaytiful interface
' check on non std file - working strange - but do not make any fail ... so good.


' 023 perturbation list normalisation
' 024 power point initialimplementation of helper

' 025 get object feature

' 026 backup copy


' 030 big change - new presentation mode!
' 030 part II pivots sheet4 - new layout of output dashboard!
' 030 perturbation normalisation list
' 030 archive from 20210426 - added sheet from transfer file - doc info handler doSomeOperation major changes.

' 031 BA in listes
' 031 G200 dodane kody plantow



'033 - added ME33K logic for BA qty


'034 - 20210901 - version with some small fixes on lookup for columns
' from 30 to If matchProbability > 40 Then - more strict on soft lookup
