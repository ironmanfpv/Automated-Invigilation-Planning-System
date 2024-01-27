' This is an example of the module that scans for optimisation of deployment

Sub ScanOptimisation1A()
'   Fill the invigilation roster by looping through cells and matching number with names

    Dim CellsDown As Long, CellsAcross As Integer
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
         
    For CurrRow = 1 To CellsDown
        For CurrCol = 1 To CellsAcross
            
               
            Dim ScanTest As Boolean   '<--- Boolean Variable to check 4 other levels if Staff has been deployed relief(6 June 2019)
            Dim ScanTestTT As Boolean '<--- Boolean Variable to check TT if Staff has lesson before or after relief (6 June 2019)
               
            ScanTest = Optimisation(SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1))
            ScanTestTT = OptimisationTT(SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1))
        
                If SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Interior.ColorIndex = xlNone And _
                ScanTestTT And ScanTest Then
     
                    SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Interior.ColorIndex = 40
        
                End If
        
       Next CurrCol
    Next CurrRow
     
    End If
     
    'Else: If CellsDown > 25 Or CellsAcross > 12 Then MsgBox "Error!!"
         
'   Display elapsed time
    'Application.ScreenUpdating = True
    MsgBox Format(Timer - StartTime, "00.00") & " seconds"
    

End Sub




