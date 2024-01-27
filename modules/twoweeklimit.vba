'This block of code checks for the hard cap limit of deployment slots over 2 weeks.

Public Function TwoWeekLimit(ByVal Cell As Range) As Boolean

     
     Dim Counter As Integer
   
     TwoWeekLimit = True '<-- initialise the variable as False
     'SheetSec1.Range("L16").Value = "NO "
     'SheetSec1.Range("L257").Value = "NO"
     'SheetSec2.Range("L16").Value = "NO"
     'SheetSec2.Range("L257").Value = "NO"
     'SheetSec3.Range("L16").Value = "NO"
     'SheetSec3.Range("L257").Value = "NO"
     'SheetSec4.Range("L16").Value = "NO"
     'SheetSec4.Range("L257").Value = "NO"
     'SheetSec5.Range("L16").Value = "NO"
     'SheetSec5.Range("L257").Value = "NO"
     
     'SheetSec1.Range("L64").Value = "NO"
     'SheetSec1.Range("L304").Value = "NO"
     'SheetSec2.Range("L64").Value = "NO"
     'SheetSec2.Range("L304").Value = "NO"
     'SheetSec3.Range("L64").Value = "NO"
     'SheetSec3.Range("L304").Value = "NO"
     'SheetSec4.Range("L64").Value = "NO"
     'SheetSec4.Range("L304").Value = "NO"
     'SheetSec5.Range("L64").Value = "NO"
     'SheetSec5.Range("L304").Value = "NO"
     
     'SheetSec1.Range("L112").Value = "NO"
     'SheetSec1.Range("L352").Value = "NO"
     'SheetSec2.Range("L112").Value = "NO"
     'SheetSec2.Range("L352").Value = "NO"
     'SheetSec3.Range("L112").Value = "NO"
     'SheetSec3.Range("L352").Value = "NO"
     'SheetSec4.Range("L112").Value = "NO"
     'SheetSec4.Range("L352").Value = "NO"
     'SheetSec5.Range("L112").Value = "NO"
     'SheetSec5.Range("L352").Value = "NO"
     
     'SheetSec1.Range("L160").Value = "NO"
     'SheetSec1.Range("L400").Value = "NO"
     'SheetSec2.Range("L160").Value = "NO"
     'SheetSec2.Range("L400").Value = "NO"
     'SheetSec3.Range("L160").Value = "NO"
     'SheetSec3.Range("L400").Value = "NO"
     'SheetSec4.Range("L160").Value = "NO"
     'SheetSec4.Range("L400").Value = "NO"
     'SheetSec5.Range("L160").Value = "NO"
     'SheetSec5.Range("L400").Value = "NO"
     
     'SheetSec1.Range("L208").Value = "NO"
     'SheetSec1.Range("L448").Value = "NO"
     'SheetSec2.Range("L208").Value = "NO"
     'SheetSec2.Range("L448").Value = "NO"
     'SheetSec3.Range("L208").Value = "NO"
     'SheetSec3.Range("L448").Value = "NO"
     'SheetSec4.Range("L208").Value = "NO"
     'SheetSec4.Range("L448").Value = "NO"
     'SheetSec5.Range("L208").Value = "NO"
     'SheetSec5.Range("L448").Value = "NO"
     
       
    For Counter = 1 To 120
    
        If Cell.Value = SheetM_S_D.Range("AE4").Offset(Counter, 0).Value And SheetM_S_D.Range("AJ4").Offset(Counter, 0).Value < 0 Then
                TwoWeekLimit = False '<-- The 2 week limit is not reached (True)and assigns False to TwoWeekLimit variable to fail the external AND loop
                SheetSec1.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                
                SheetSec1.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                
                SheetSec1.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                
                SheetSec1.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                
                SheetSec1.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                
                Exit For

           Else: TwoWeekLimit = True
           
 
                SheetSec1.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L16").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L257").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                
                SheetSec1.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L64").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L304").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                
                SheetSec1.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L112").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L352").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                
                SheetSec1.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L160").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L400").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                
                SheetSec1.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec1.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec2.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec3.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec4.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L208").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value
                SheetSec5.Range("L448").Value = SheetM_S_D.Range("AL4").Offset(Counter, 0).Value

        End If
    Next Counter

End Function