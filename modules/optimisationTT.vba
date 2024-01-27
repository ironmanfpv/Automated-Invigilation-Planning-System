'This block of code is where optimisation is also checked against the staff's schedule

Public Function OptimisationTT(ByVal Cell As Range) As Boolean

     Dim Counter As Integer
     Dim Row As Long
     Dim Column As Long
     Dim Periodnumber As Integer
     Dim Daynumber As Integer
     Dim RowPosition As Integer
    
     Row = Range(Cell.Address).Row                          'Obtains the Row of the current slot needing allocation
     Column = Range(Cell.Address).Column                    'Obtains the Column of the current slot needing allocation
     Periodnumber = Cell.Offset(0, -Column + 1)             'Obtains the period number of the current slot needing allocation
     'Daynumber = Cell.Offset(-Row + 12, -Column + 3).Value  'Obtains the day number of the current slot needing allocation
     
     'MsgBox Cell.Value
     'MsgBox Periodnumber
     'MsgBox Row
     
'Finds the Day number
     
     Select Case Row
     
     Case 22 To 46: Daynumber = Sheet2.Range("D3").Value
     Case 70 To 94: Daynumber = Sheet2.Range("D4").Value
     Case 118 To 142: Daynumber = Sheet2.Range("D5").Value
     Case 166 To 190: Daynumber = Sheet2.Range("D6").Value
     Case 214 To 238: Daynumber = Sheet2.Range("D7").Value
     Case 262 To 286: Daynumber = Sheet2.Range("D8").Value
     Case 310 To 334: Daynumber = Sheet2.Range("D9").Value
     Case 358 To 382: Daynumber = Sheet2.Range("D10").Value
     Case 406 To 430: Daynumber = Sheet2.Range("D11").Value
     Case 454 To 478: Daynumber = Sheet2.Range("D12").Value
     Case 502 To 526: Daynumber = Sheet2.Range("D13").Value
     
     End Select
     
     'MsgBox Daynumber
     
     Counter = 1                                    'initializes Counter to 1
     OptimisationTT = False                         'Initializes the Optimisation Indicator to False
    
'Matching the correct person

     For Counter = 1 To 120
                    RowPosition = 1
                    If (Cell.Value = SheetM_S_D.Range("AE4").Offset(Counter - 1, 0).Value) Then
                        RowPosition = Counter
                        Exit For
                    End If
     Next Counter
     
     'MsgBox Daynumber
     'MsgBox RowPosition
     
'Checking the TT of the corresponding person for lesson before and after

     Select Case Daynumber
     
            Case 1: 'Checks for Monday of Odd Week
            
            If (Periodnumber = 1) Then
                If (SheetM_S_D.Range("E4").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D.Range("E4").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D.Range("E4").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
            'MsgBox SheetM_S_D.Range("E4").Offset(RowPosition - 1, Periodnumber).Value
            
            Case 2: 'Checks for Tuesday of Odd Week
            
            If (Periodnumber = 1) Then
                If (SheetM_S_D.Range("E124").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D.Range("E124").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D.Range("E124").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
            'MsgBox SheetM_S_D.Range("E124").Offset(RowPosition - 1, Periodnumber).Value
            
            Case 3: 'Checks for Wednesday of Odd Week
            
            If (Periodnumber = 1) Then
                If (SheetM_S_D.Range("E244").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D.Range("E244").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D.Range("E244").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
            'MsgBox SheetM_S_D.Range("E244").Offset(RowPosition - 1, Periodnumber).Value
            
            Case 4: 'Checks for Thursday of Odd Week
            
            If (Periodnumber = 1) Then
                If (SheetM_S_D.Range("E364").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D.Range("E364").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D.Range("E364").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
            'MsgBox SheetM_S_D.Range("E364").Offset(RowPosition - 1, Periodnumber).Value
            
            Case 5: 'Checks for Friday of Odd Week
            
            If (Periodnumber = 1) Then
                If (SheetM_S_D.Range("E484").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D.Range("E484").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D.Range("E484").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
             'MsgBox SheetM_S_D.Range("E484").Offset(RowPosition - 1, Periodnumber).Value
            
            Case 21: 'Checks for Monday of Even Week
            
            If (Periodnumber = 1) Then
                If (SheetM_S_D1.Range("E4").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D1.Range("E4").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D1.Range("E4").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
            'MsgBox SheetM_S_D1.Range("E4").Offset(RowPosition - 1, Periodnumber).Value
            
            Case 22: 'Checks for Tuesday of Even Week
            
            If (Periodnumber = 1) Then
                If (SheetM_S_D1.Range("E124").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D1.Range("E124").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D1.Range("E124").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
            'MsgBox SheetM_S_D1.Range("E124").Offset(RowPosition - 1, Periodnumber).Value
            
            Case 23: 'Checks for Wednesday of Even Week
            
             If (Periodnumber = 1) Then
                If (SheetM_S_D1.Range("E244").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D1.Range("E244").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D1.Range("E244").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
            'MsgBox SheetM_S_D1.Range("E244").Offset(RowPosition - 1, Periodnumber).Value
            
            Case 24: 'Checks for Thursday of Even Week
            
             If (Periodnumber = 1) Then
                If (SheetM_S_D1.Range("E364").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D1.Range("E364").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D1.Range("E364").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
            'MsgBox SheetM_S_D1.Range("E364").Offset(RowPosition - 1, Periodnumber).Value
            
            Case 25: 'Checks for Friday of Even Week
            
             If (Periodnumber = 1) Then
                If (SheetM_S_D1.Range("E484").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                    OptimisationTT = True
                End If
            Else
                If (Periodnumber <> 1) Then
                    If (SheetM_S_D1.Range("E484").Offset(RowPosition - 1, Periodnumber - 2).Value = " " And _
                        SheetM_S_D1.Range("E484").Offset(RowPosition - 1, Periodnumber).Value = " ") Then
                            OptimisationTT = True
                    End If
                End If
            End If
            
            'MsgBox SheetM_S_D1.Range("E484").Offset(RowPosition - 1, Periodnumber).Value
            
            Case Else: MsgBox "Daynumber Not found" 'Error and Exception Handling
   
     End Select
     
     'OptimisationTT = True
     
End Function
