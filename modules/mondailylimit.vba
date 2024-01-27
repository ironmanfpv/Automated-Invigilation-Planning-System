'This is function that checks for daily deployment limit for Monday

Public Function MonDailyLimit(ByVal Cell As Range) As Boolean

     
     Dim Counter As Integer
     MonDailyLimit = True '<-- initialise the variable as TRUE
     'SheetSec1.Range("K16").Value = " "
     'SheetSec1.Range("K257").Value = " "
     'SheetSec2.Range("K16").Value = " "
     'SheetSec2.Range("K257").Value = " "
     'SheetSec3.Range("K16").Value = " "
     'SheetSec3.Range("K257").Value = " "
     'SheetSec4.Range("K16").Value = " "
     'SheetSec4.Range("K257").Value = " "
     'SheetSec5.Range("K16").Value = " "
     'SheetSec5.Range("K257").Value = " "
        
    For Counter = 1 To 120
    
        If Cell.Value = SheetM_S_D.Range("AE4").Offset(Counter, 0).Value And SheetM_S_D.Range("AK4").Offset(Counter, 0).Value = "YES" Then
        'Scrolls Through the 120 Staff Name in AE4 for a match with the : _
         current cell name and checks if the corresponding Daily limit indicator in AK4 is reached
                MonDailyLimit = False '<-- The daily limit is reached and returns a FALSE value for Boolean
                SheetSec1.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec1.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                ' Displays the DailyLimit Indicator value in the cell for visual effect
                
            Exit For
            
            Else: MonDailyLimit = True
            SheetSec1.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec1.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K16").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K257").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            ' Displays the DailyLimit Indicator value in the cell for visual effect
            
        End If
    Next Counter

End Function


