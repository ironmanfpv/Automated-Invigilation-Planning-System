'This vba code is an example of the algorithm used to swap invigilators across the same lvl same time slot 

Sub InvigilatorSwap1A()
'   Fill the invigilation roster by looping through cells and matching number with names

    Dim CellsDown As Long, CellsAcross As Integer, HorizontalScan As Integer, HorizontalScanEnd As Integer
    Dim CurrRow As Long, CurrCol As Integer, RowScan As Integer
    Dim StartTime As Double
    Dim Response As String
    Dim Confirmation As Boolean
    Dim GothruSeed As Integer
       
'   Get the dimensions
    'CellsDown = Val(InputBox("How many cells down?(Integers: 1 to 25)"))
    'CellsAcross = Val(InputBox("How many cells across?(Integers: 1 to 12)"))
    'Response = Val(InputBox("Do you wish to show the cells updating as the program generates possible invigilators ?(Reply: Y for Yes and N for No)"))
     Response = Application.InputBox(prompt:="Do you wish to show the cells updating as the program verifies the staff allocation?(Reply: Y for Yes and N for No)", Type:=2)
    
    'MsgBox Response
     
     CellsDown = 25
     CellsAcross = 12
     HorizontalScanEnd = 12
     
'   Record starting time
    StartTime = Timer

'   Loop through cells and insert values and Match invigilator
'   Application.ScreenUpdating if set to True shows the whole allocation process in action
'   Application.ScreenUpdating = True
    
   'If CellsDown <= 25 And CellsAcross <= 12 Then
             
    If Response <> "Y" And Response <> "y" And Response <> "N" And Response <> "n" Then
       MsgBox "Invalid Response !!"
       GothruSeed = 0
    Else
        If Response = "Y" Or Response = "y" Then
        Confirmation = True
        GothruSeed = 1
    Else
        If Response = "N" Or Response = "n" Then
        Confirmation = False
        GothruSeed = 1
    End If
    End If
    End If
    
    Application.ScreenUpdating = Confirmation
         
    If GothruSeed = 1 Then
         
'   When iterating through the cells, scans the next immediate row for a same invigilator and do a swap
    
    For CurrRow = 1 To CellsDown
        For CurrCol = 1 To CellsAcross
        
        If CurrRow < CellsDown Then
                
            For HorizontalScan = 1 To HorizontalScanEnd
            
                        If SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Interior.ColorIndex <> xlNone Then
                            Exit For
                            Exit For
                        Else
                            If SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1) = SheetSec1.Range("C22").Offset(CurrRow, HorizontalScan - 1) Then
                            
                            SheetSec1.Range("Z22").Offset(CurrRow, 0) = SheetSec1.Range("C22").Offset(CurrRow, HorizontalScan - 1)           'A to Temp
                            SheetSec1.Range("C22").Offset(CurrRow, HorizontalScan - 1) = SheetSec1.Range("C22").Offset(CurrRow, CurrCol - 1) ' Scan cell to A
                            SheetSec1.Range("C22").Offset(CurrRow, CurrCol - 1) = SheetSec1.Range("Z22").Offset(CurrRow, 0)                  ' Temp to Scan cell
                            SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Interior.ColorIndex = 8                                  ' Colour the current iterated cell blue
                            End If
                        End If
            Next HorizontalScan
    End If
           
       
       Next CurrCol
    Next CurrRow
     
     
' Paint all consecutive invigilation slots Blue
     
       For CurrRow = 1 To CellsDown
            For CurrCol = 1 To CellsAcross
        
                If CurrRow < CellsDown Then
                    If SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1) = SheetSec1.Range("C22").Offset(CurrRow, CurrCol - 1) And _
                        SheetSec1.Range("C22").Offset(CurrRow, CurrCol - 1).Interior.ColorIndex = xlNone Then
                        
                        SheetSec1.Range("C22").Offset(CurrRow, CurrCol - 1).Interior.ColorIndex = 8
                    End If
                End If
       
            Next CurrCol
        Next CurrRow
     
' Clears all the INVG SWAP cells
     
        For CurrRow = 1 To CellsDown
        
            SheetSec1.Range("Z22").Offset(CurrRow, 0) = ""
     
        Next CurrRow
     
     
'Ends the whole iteration

    End If
    
         
'   Display elapsed time
'   Application.ScreenUpdating = True
    MsgBox Format(Timer - StartTime, "00.00") & " seconds"
    

End Sub







