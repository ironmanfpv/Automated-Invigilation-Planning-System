'This is function that checks for daily deployment limit for Wednesday


Public Function WedDailyLimit(ByVal Cell As Range) As Boolean

     
     Dim Counter As Integer
     WedDailyLimit = True '<-- initialise the variable as TRUE
     'SheetSec1.Range("K112").Value = " "
     'SheetSec1.Range("K352").Value = " "
     'SheetSec2.Range("K112").Value = " "
     'SheetSec2.Range("K352").Value = " "
     'SheetSec3.Range("K112").Value = " "
     'SheetSec3.Range("K352").Value = " "
     'SheetSec4.Range("K112").Value = " "
     'SheetSec4.Range("K352").Value = " "
     'SheetSec5.Range("K112").Value = " "
     'SheetSec5.Range("K352").Value = " "
        
    For Counter = 1 To 120
    
        If Cell.Value = SheetM_S_D.Range("AE244").Offset(Counter, 0).Value And SheetM_S_D.Range("AK244").Offset(Counter, 0).Value = "YES" Then
        'Scrolls Through the 120 Staff Name in AE244 for a match with the : _
         current cell name and checks if the corresponding Daily limit indicator in AK244 is reached
                WedDailyLimit = False '<-- The daily limit is reached and returns a FALSE value for Boolean
                SheetSec1.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec1.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec2.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec3.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec4.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                SheetSec5.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
                ' Displays the DailyLimit Indicator value in the cell for visual effect
            Exit For
            Else: WedDailyLimit = True
            SheetSec1.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec1.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec2.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec3.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec4.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K112").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            SheetSec5.Range("K352").Value = SheetM_S_D.Range("AK4").Offset(Counter, 0).Value
            ' Displays the DailyLimit Indicator value in the cell for visual effect
        End If
    Next Counter

End Function
