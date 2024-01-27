' This block of code is an example of how optimisation is checked

Public Function Optimisation(ByVal Cell As Range) As Boolean

'__________________________________________________________________________________________________________________________

'[Declaration Block]

Dim Row As Long
Dim Column As Long

Static Tries As Long


Dim TargetRow1 As Integer
Dim TargetRow2 As Integer
Dim TargetRow3 As Integer
Dim TargetRow4 As Integer

Dim TargetColm As Integer

Dim Sheet1IndxA As Boolean
Dim Sheet1IndxB As Boolean
Dim Sheet1IndxC As Boolean
Dim Sheet1IndxD As Boolean

Dim Sheet2IndxA As Boolean
Dim Sheet2IndxB As Boolean
Dim Sheet2IndxC As Boolean
Dim Sheet2IndxD As Boolean

Dim Sheet3IndxA As Boolean
Dim Sheet3IndxB As Boolean
Dim Sheet3IndxC As Boolean
Dim Sheet3IndxD As Boolean

Dim Sheet4IndxA As Boolean
Dim Sheet4IndxB As Boolean
Dim Sheet4IndxC As Boolean
Dim Sheet4IndxD As Boolean

Dim Sheet5IndxA As Boolean
Dim Sheet5IndxB As Boolean
Dim Sheet5IndxC As Boolean
Dim Sheet5IndxD As Boolean

'Dim SheetPCUP As Boolean     '<----Boolean Variable checking if invigilator has lesson 1 slot earlier (on hold)
'Dim SheetPCDN As Boolean     '<----Boolean Variable checking if invigilator has lesson 1 slot later   (on hold)


Dim ConditionOne As Boolean  '<----Condition that invigilator has been used 1 slot earlier across all sheets
Dim ConditionTwo As Boolean  '<----Condition that invigilator has been used 2 slot earlier across all sheets
Dim ConditionThree As Boolean '<----Condition that invigilation has been used 1 slot LATER across all sheets
Dim ConditionFour As Boolean  '<----Condition that invigilator has been used 2 slot LATER across all sheets
Dim VerticalOne As Boolean   '<----Condition that invigilator has been used 1 slot directly above same sheet
Dim VerticalTwo As Boolean   '<----Condition that invigilator has been used 2 Slot directly above same sheet
Dim VerticalThree As Boolean '<----Condition that invigilator has been used 1 Slot Below above same sheet
Dim VerticalFour As Boolean  '<----Condition that invigilator has been used 2 Slot Below above same sheet

'__________________________________________________________________________________________________________________________
'[Assignment Block]

'Row = ActiveCell.Row
Row = Range(Cell.Address).Row
Column = Range(Cell.Address).Column


TargetRow1 = Row - 1 '<-- Row Index of the row of slots directly above current slot
TargetRow2 = Row - 2 '<-- Row Index of the row of slots 2 rows directly above the current slot
TargetRow3 = Row + 1 '<-- Row Index of the row of slots directly below current slot
TargetRow4 = Row + 2 '<-- Row Index of the row of slots 2 rows directly below the current slot

'VerticalOne True implies invigilator used in one slot directly above current cell in the same sheet
'VerticalTwo True implies invigilator used in two slots directly above current cell in the same sheet
'VerticalThree True implies invigilator used in one slots directly below current cell in the same sheet
'VerticalFour True implies invigilator used in two slots directly below current cell in the same sheet

'If Cells(TargetRow1, Column).Value <> Cell.Value Then VerticalOne = False
If Cells(TargetRow1, Column).Value = Cell.Value Then VerticalOne = True
'If Cells(TargetRow2, Column).Value <> Cell.Value Then VerticalTwo = False
If Cells(TargetRow2, Column).Value = Cell.Value Then VerticalTwo = True
'If Cells(TargetRow3, Column).Value <> Cell.Value Then VerticalThree = False
If Cells(TargetRow3, Column).Value = Cell.Value Then VerticalThree = True
'If Cells(TargetRow4, Column).Value <> Cell.Value Then VerticalFour = False
If Cells(TargetRow4, Column).Value = Cell.Value Then VerticalFour = True

'MsgBox Row
'MsgBox Column
'MsgBox RowScope1

'__________________________________________________________________________________________________________________________

'[Checking Sheet1]


Sheet1IndxA = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec1.Cells(TargetRow1, TargetColm).Value = Cell.Value Then
        Sheet1IndxA = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet1IndxB = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec1.Cells(TargetRow2, TargetColm).Value = Cell.Value Then
        Sheet1IndxB = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet1IndxC = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec1.Cells(TargetRow3, TargetColm).Value = Cell.Value Then
        Sheet1IndxC = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet1IndxD = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec1.Cells(TargetRow4, TargetColm).Value = Cell.Value Then
        Sheet1IndxD = False
        Exit For
    End If
    Next TargetColm


'__________________________________________________________________________________________________________________________

'[Checking Sheet2]

Sheet2IndxA = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec2.Cells(RowScope1, Sheet1Col).Value
    If SheetSec2.Cells(TargetRow1, TargetColm).Value = Cell.Value Then
        Sheet2IndxA = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet2IndxB = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec2.Cells(TargetRow2, TargetColm).Value = Cell.Value Then
        Sheet2IndxB = False
        Exit For
    End If
    Next TargetColm


Sheet2IndxC = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec2.Cells(TargetRow3, TargetColm).Value = Cell.Value Then
        Sheet2IndxC = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet2IndxD = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec2.Cells(TargetRow4, TargetColm).Value = Cell.Value Then
        Sheet2IndxD = False
        Exit For
    End If
    Next TargetColm

'__________________________________________________________________________________________________________________________


'[Checking Sheet3]

Sheet3IndxA = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec2.Cells(RowScope1, Sheet1Col).Value
    If SheetSec3.Cells(TargetRow1, TargetColm).Value = Cell.Value Then
        Sheet3IndxA = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet3IndxB = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec3.Cells(TargetRow2, TargetColm).Value = Cell.Value Then
        Sheet3IndxB = False
        Exit For
    End If
    Next TargetColm


Sheet3IndxC = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec3.Cells(TargetRow3, TargetColm).Value = Cell.Value Then
        Sheet3IndxC = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet3IndxD = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec3.Cells(TargetRow4, TargetColm).Value = Cell.Value Then
        Sheet3IndxD = False
        Exit For
    End If
    Next TargetColm

'__________________________________________________________________________________________________________________________


'[Checking Sheet4]

Sheet4IndxA = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec4.Cells(RowScope1, Sheet1Col).Value
    If SheetSec4.Cells(TargetRow1, TargetColm).Value = Cell.Value Then
        Sheet4IndxA = False
        Exit For
    End If
    Next TargetColm
    
      
Sheet4IndxB = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec4.Cells(TargetRow2, TargetColm).Value = Cell.Value Then
        Sheet4IndxB = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet4IndxC = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec4.Cells(TargetRow3, TargetColm).Value = Cell.Value Then
        Sheet4IndxC = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet4IndxD = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec4.Cells(TargetRow4, TargetColm).Value = Cell.Value Then
        Sheet4IndxD = False
        Exit For
    End If
    Next TargetColm


'__________________________________________________________________________________________________________________________

'[Checking Sheet5]

Sheet5IndxA = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec5.Cells(RowScope1, Sheet1Col).Value
    If SheetSec5.Cells(TargetRow1, TargetColm).Value = Cell.Value Then
        Sheet5IndxA = False
        Exit For
    End If
    Next TargetColm
    
    
Sheet5IndxB = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec5.Cells(TargetRow2, TargetColm).Value = Cell.Value Then
        Sheet5IndxB = False
        Exit For
    End If
    Next TargetColm
    

Sheet5IndxC = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec5.Cells(TargetRow3, TargetColm).Value = Cell.Value Then
        Sheet5IndxC = False
        Exit For
    End If
    Next TargetColm
    
    
    
Sheet5IndxD = True '<---- Initialise to TRUE condition First
For TargetColm = 3 To 14
    'MsgBox SheetSec1.Cells(RowScope1, Sheet1Col).Value
    If SheetSec5.Cells(TargetRow4, TargetColm).Value = Cell.Value Then
        Sheet5IndxD = False
        Exit For
    End If
    Next TargetColm
'_________________________________________________________________________________________________________________________
'[Checking for adjacent Lesson Blocks from time-table]

 'To be inserted if there is a need to   <----Decision aborted 7th June'2012 as entering the check would drastically
                                              'decrease deployement sample space.

'__________________________________________________________________________________________________________________________

'[Summary of Conditions across the 5 sheets ]   Revisited 4/6/2019

'ConditionOne True implies invigilator not allocated anywhere in ONE slot earlier for the 5 levels
'ConditionTwo True Implies invigilator not allocated anywhere in TWO slots earlier for the 5 levels
'ConditionThree True Implies invigilator not allocated anywhere in ONE slots later for the 5 levels
'ConditionFour True Implies invigilator not allocated anywhere in TWO slots later for the 5 levels

ConditionOne = Sheet1IndxA And Sheet2IndxA And Sheet3IndxA And Sheet4IndxA And Sheet5IndxA
ConditionTwo = Sheet1IndxB And Sheet2IndxB And Sheet3IndxB And Sheet4IndxB And Sheet5IndxB
ConditionThree = Sheet1IndxC And Sheet2IndxC And Sheet3IndxC And Sheet4IndxC And Sheet5IndxC
ConditionFour = Sheet1IndxD And Sheet2IndxD And Sheet3IndxD And Sheet4IndxD And Sheet5IndxD


Optimisation = True '<--- Initialise to False (Temporary change to True for testing 4 Sept)


If ConditionOne And ConditionThree Then                '[If the Staff randomly chosen is deployed 1 slot earlier and one slot later is in the OTHER level]
    Optimisation = True

    Else
        If VerticalOne And VerticalThree Then  '[ If the Staff randomly chosen has been deployed directly 1 slot earlier and 1 slot later in SAME levels ]
            Optimisation = True

            'Else
                'If ConditionOne And ConditionThree And NotConditionTwo Then   '[If the Staff has Teaching block 1 slot earlier and 1 slot later]
                    'Optimisation = True

                'End If
        End If
End If


End Function
