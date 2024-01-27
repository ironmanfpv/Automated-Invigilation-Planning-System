'This is function that checks for daily deployment limit for Tuesday

Public Function TueDailyLimit(ByVal Cell As Range) As Boolean

     
     Dim Counter As Integer
     TueDailyLimit = True '<-- initialise the variable as True
     'SheetSec1.Range("K64").Value = " "
     'SheetSec1.Range("K304").Value = " "
     'SheetSec2.Range("K64").Value = " "
     'SheetSec2.Range("K304").Value = " "
     'SheetSec3.Range("K64").Value = " "
     'SheetSec3.Range("K304").Value = " "
     'SheetSec4.Range("K64").Value = " "
     'SheetSec4.Range("K304").Value = " "
     'SheetSec5.Range("K64").Value = " "
     'SheetSec5.Range("K304").Value = " "
        
    For Counter = 1 To 120
    
        If Cell.Value = SheetM_S_D.Range("AE124").Offset(Counter, 0).Value And SheetM_S_D.Range("AK124").Offset(Counter, 0).Value = "YES" Then
        'Scrolls Through the 120 Staff Name in AE124 for a match with the : _
         current cell name and checks if the corresponding Daily limit indicator in AK124 is reached
                TueDailyLimit = False '<-- The daily limit is reached and returns a FALSE value for Boolean
                SheetSec1.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec1.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                ' Displays the DailyLimit Indicator value in the cell for visual effect
            Exit For
            Else: TueDailyLimit = True
            SheetSec1.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec1.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K64").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K304").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            ' Displays the DailyLimit Indicator value in the cell for visual effect
        End If
    Next Counter

End Function
