'This is function that checks for daily deployment limit for Friday

Public Function FriDailyLimit(ByVal Cell As Range) As Boolean

     
     Dim Counter As Integer
     FriDailyLimit = True '<-- initialise the variable as TRUE
     'SheetSec1.Range("K208").Value = " "
     'SheetSec1.Range("K448").Value = " "
     'SheetSec2.Range("K208").Value = " "
     'SheetSec2.Range("K448").Value = " "
     'SheetSec3.Range("K208").Value = " "
     'SheetSec3.Range("K448").Value = " "
     'SheetSec4.Range("K208").Value = " "
     'SheetSec4.Range("K448").Value = " "
     'SheetSec5.Range("K208").Value = " "
     'SheetSec5.Range("K448").Value = " "
        
    For Counter = 1 To 120
    
        If Cell.Value = SheetM_S_D.Range("AE484").Offset(Counter, 0).Value And SheetM_S_D.Range("AK484").Offset(Counter, 0).Value = "YES" Then
        'Scrolls Through the 120 Staff Name in AE484 for a match with the : _
         current cell name and checks if the corresponding Daily limit indicator in AK484 is reached
                FriDailyLimit = False '<-- The daily limit is reached and returns a FALSE value for Boolean
                SheetSec1.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec1.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                ' Displays the DailyLimit Indicator value in the cell for visual effect
            Exit For
            Else: FriDailyLimit = True
            SheetSec1.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec1.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K208").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K448").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            ' Displays the DailyLimit Indicator value in the cell for visual effect
        End If
    Next Counter

End Function
