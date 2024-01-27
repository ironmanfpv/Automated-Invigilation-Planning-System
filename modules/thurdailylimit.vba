'This is function that checks for daily deployment limit for Thursday


Public Function ThurDailyLimit(ByVal Cell As Range) As Boolean

     
     Dim Counter As Integer
     ThurDailyLimit = True '<-- initialise the variable as TRUE
     'SheetSec1.Range("K160").Value = " "
     'SheetSec1.Range("K400").Value = " "
     'SheetSec2.Range("K160").Value = " "
     'SheetSec2.Range("K400").Value = " "
     'SheetSec3.Range("K160").Value = " "
     'SheetSec3.Range("K400").Value = " "
     'SheetSec4.Range("K160").Value = " "
     'SheetSec4.Range("K400").Value = " "
     'SheetSec5.Range("K160").Value = " "
     'SheetSec5.Range("K400").Value = " "
        
    For Counter = 1 To 120
    
        If Cell.Value = SheetM_S_D.Range("AE364").Offset(Counter, 0).Value And SheetM_S_D.Range("AK364").Offset(Counter, 0).Value = "YES" Then
        'Scrolls Through the 120 Staff Name in AE364 for a match with the : _
         current cell name and checks if the corresponding Daily limit indicator in AK364 is reached
                ThurDailyLimit = False '<-- The daily limit is reached and returns a FALSE value for Boolean
                SheetSec1.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec1.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                ' Displays the DailyLimit Indicator value in the cell for visual effect
            Exit For
            Else: ThurDailyLimit = True
            SheetSec1.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec1.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K160").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K400").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            ' Displays the DailyLimit Indicator value in the cell for visual effect
        End If
    Next Counter

End Function
