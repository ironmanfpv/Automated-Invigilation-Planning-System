'This block of code is an example of the generating function for one of the days

Sub AutoMatch1A()
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
     Response = Application.InputBox(prompt:="Do you wish to show the cells updating as the program generates possible invigilators ?(Reply: Y for Yes and N for No)", Type:=2)
    
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
            
               Dim TempNum As Variant
               Dim CoordinatorClearIndex As Boolean
               Dim FairDistribution As Boolean
               Dim FindOptimisation As Boolean
               Dim CheckMonDailyLimit As Boolean   '<--- Boolean Variable to check Mon daily limit (13 May 2019)
               Dim CheckTwoWeekLimit As Boolean '<--- Boolean Variable to check 2 Weekly Limit (13 May 2019)
                
               
     If SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Interior.ColorIndex = xlNone Then
     
     Do
     
     SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = VBA.Math.Round((Rnd() * 120), 0)
 
     TempNum = SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value
    'MsgBox TempNum
    
     Select Case TempNum
     Case 1: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B17").Value
     Case 2: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B18").Value
     Case 3: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B19").Value
     Case 4: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B20").Value
     Case 5: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B21").Value
     Case 6: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B22").Value
     Case 7: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B23").Value
     Case 8: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B24").Value
     Case 9: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B25").Value
     Case 10: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B26").Value
     Case 11: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B27").Value
     Case 12: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B28").Value
     Case 13: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B29").Value
     Case 14: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B30").Value
     Case 15: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B31").Value
     Case 16: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B32").Value
     Case 17: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B33").Value
     Case 18: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B34").Value
     Case 19: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B35").Value
     Case 20: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B36").Value
     Case 21: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B37").Value
     Case 22: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B38").Value
     Case 23: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B39").Value
     Case 24: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B40").Value
     Case 25: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B41").Value
     Case 26: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B42").Value
     Case 27: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B43").Value
     Case 28: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B44").Value
     Case 29: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B45").Value
     Case 30: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B46").Value
     Case 31: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B47").Value
     Case 32: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B48").Value
     Case 33: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B49").Value
     Case 34: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B50").Value
     Case 35: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B51").Value
     Case 36: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B52").Value
     Case 37: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B53").Value
     Case 38: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B54").Value
     Case 39: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B55").Value
     Case 40: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B56").Value
     Case 41: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B57").Value
     Case 42: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B58").Value
     Case 43: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B59").Value
     Case 44: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B60").Value
     Case 45: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B61").Value
     Case 46: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B62").Value
     Case 47: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B63").Value
     Case 48: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B64").Value
     Case 49: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B65").Value
     Case 50: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B66").Value
     Case 51: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B67").Value
     Case 52: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B68").Value
     Case 53: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B69").Value
     Case 54: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B70").Value
     Case 55: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B71").Value
     Case 56: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B72").Value
     Case 57: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B73").Value
     Case 58: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B74").Value
     Case 59: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B75").Value
     Case 60: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B76").Value
     Case 61: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B77").Value
     Case 62: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B78").Value
     Case 63: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B79").Value
     Case 64: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B80").Value
     Case 65: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B81").Value
     Case 66: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B82").Value
     Case 67: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B83").Value
     Case 68: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B84").Value
     Case 69: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B85").Value
     Case 70: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B86").Value
     Case 71: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B87").Value
     Case 72: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B88").Value
     Case 73: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B89").Value
     Case 74: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B90").Value
     Case 75: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B91").Value
     Case 76: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B92").Value
     Case 77: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B93").Value
     Case 78: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B94").Value
     Case 79: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B95").Value
     Case 80: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B96").Value
     Case 81: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B97").Value
     Case 82: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B98").Value
     Case 83: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B99").Value
     Case 84: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B100").Value
     Case 85: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B101").Value
     Case 86: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B102").Value
     Case 87: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B103").Value
     Case 88: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B104").Value
     Case 89: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B105").Value
     Case 90: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B106").Value
     Case 91: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B107").Value
     Case 92: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B108").Value
     Case 93: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B109").Value
     Case 94: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B110").Value
     Case 95: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B111").Value
     Case 96: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B112").Value
     Case 97: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B113").Value
     Case 98: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B114").Value
     Case 99: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B115").Value
     Case 100: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B116").Value
     Case 101: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B117").Value
     Case 102: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B118").Value
     Case 103: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B119").Value
     Case 104: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B120").Value
     Case 105: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B121").Value
     Case 106: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B122").Value
     Case 107: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B123").Value
     Case 108: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B124").Value
     Case 109: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B125").Value
     Case 110: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B126").Value
     Case 111: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B127").Value
     Case 112: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B128").Value
     Case 113: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B129").Value
     Case 114: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B130").Value
     Case 115: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B131").Value
     Case 116: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B132").Value
     Case 117: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B133").Value
     Case 118: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B134").Value
     Case 119: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B135").Value
     Case 120: SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value = SheetIndx.Range("B136").Value
           
        
        
     End Select
      
     CoordinatorClearIndex = CoordinatorScan(SheetP_C.Range("C22").Offset(CurrRow - 1, CurrCol - 1))
     FairDistribution = Equalizer(SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1))
     'MsgBox CoordinatorClearIndex <--- This checks for correct boolean value as the co-ordinator scan is invoked
     'MsgBox Sheet28.Range("C16").Offset(1, 14).Value   <-- This checks if I scan the right coordinator list
     FindOptimisation = Optimisation(SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1))
     CheckMonDailyLimit = MonDailyLimit(SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1)) '<-- Boolean Variable taking fn Val
     CheckTwoWeekLimit = TwoWeekLimit(SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1)) '<-- Boolean Variable taking fn Val
     
     Loop Until SheetSec1.Range("C22").Offset(CurrRow - 1, 13).Value = "NO" And _
     SheetSec1.Range("C22").Offset(CurrRow - 1, 14).Value = "NO" And _
     TempNum <> 0 And CoordinatorClearIndex And FindOptimisation And _
     SheetSec1.Range("P5").Value = "NO" And _
     SheetSec1.Range("P8").Value = "NO" And _
     SheetSec1.Range("P11").Value = "NO" And _
     SheetSec1.Range("P14").Value = "NO" And _
     SheetSec1.Range("C22").Offset(CurrRow - 1, CurrCol - 1).Value <> "Not Deployed" And _
     CheckMonDailyLimit And CheckTwoWeekLimit '(13 May 2019)'And FairDistribution
     
     'MsgBox SheetM_S_D.Range("AE4").Offset(Counter, 0).Value
     'MsgBox SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
     
     End If
     
       Next CurrCol
     Next CurrRow
     
     End If
     
    'Else: If CellsDown > 25 Or CellsAcross > 12 Then MsgBox "Error!!"
         
'   Display elapsed time
    'Application.ScreenUpdating = True
    MsgBox Format(Timer - StartTime, "00.00") & " seconds"
    

End Sub

Function CoordinatorScan(ByVal Cell As Range) As Boolean

    Dim count As Integer
   
        CoordinatorScan = True
        
    For count = 1 To 30
    
        If Cell.Value = SheetP_C.Range("C16").Offset(count - 1, 16).Value Then
        CoordinatorScan = False
        
        Exit For
        End If
         
    Next count

End Function