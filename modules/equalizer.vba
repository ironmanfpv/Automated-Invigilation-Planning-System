'This is an example of the balancing function during automatic deployment

Public Function Equalizer(ByVal Cell As Range) As Boolean

    Dim SlotCount As Integer
    Dim Meanval As Double
    Dim LowestFreq As Integer
    Dim HighestFreq As Integer
    Dim Difference As Integer
    Dim CountOfTrue As Integer
    Dim CountOfFalse As Integer
    Static DegreeOfFreedom As Integer
    Static LowestDOF As Integer
    
    
      Equalizer = False
      Meanval = SheetIndx.Range("E132").Value
      LowestFreq = SheetIndx.Range("E143").Value
      HighestFreq = SheetIndx.Range("E144").Value
      Difference = HighestFreq - LowestFreq
      
       
      If Cell.Value = SheetIndx.Range("B17").Value Then SlotCount = SheetIndx.Range("E17").Value
      If Cell.Value = SheetIndx.Range("B18").Value Then SlotCount = SheetIndx.Range("E18").Value
      If Cell.Value = SheetIndx.Range("B19").Value Then SlotCount = SheetIndx.Range("E19").Value
      If Cell.Value = SheetIndx.Range("B20").Value Then SlotCount = SheetIndx.Range("E20").Value
      If Cell.Value = SheetIndx.Range("B21").Value Then SlotCount = SheetIndx.Range("E21").Value
      If Cell.Value = SheetIndx.Range("B22").Value Then SlotCount = SheetIndx.Range("E22").Value
      If Cell.Value = SheetIndx.Range("B23").Value Then SlotCount = SheetIndx.Range("E23").Value
      If Cell.Value = SheetIndx.Range("B24").Value Then SlotCount = SheetIndx.Range("E24").Value
      If Cell.Value = SheetIndx.Range("B25").Value Then SlotCount = SheetIndx.Range("E25").Value
      If Cell.Value = SheetIndx.Range("B26").Value Then SlotCount = SheetIndx.Range("E26").Value
      If Cell.Value = SheetIndx.Range("B27").Value Then SlotCount = SheetIndx.Range("E27").Value
      If Cell.Value = SheetIndx.Range("B28").Value Then SlotCount = SheetIndx.Range("E28").Value
      If Cell.Value = SheetIndx.Range("B29").Value Then SlotCount = SheetIndx.Range("E29").Value
      If Cell.Value = SheetIndx.Range("B30").Value Then SlotCount = SheetIndx.Range("E30").Value
      If Cell.Value = SheetIndx.Range("B31").Value Then SlotCount = SheetIndx.Range("E31").Value
      If Cell.Value = SheetIndx.Range("B32").Value Then SlotCount = SheetIndx.Range("E32").Value
      If Cell.Value = SheetIndx.Range("B33").Value Then SlotCount = SheetIndx.Range("E33").Value
      If Cell.Value = SheetIndx.Range("B34").Value Then SlotCount = SheetIndx.Range("E34").Value
      If Cell.Value = SheetIndx.Range("B35").Value Then SlotCount = SheetIndx.Range("E35").Value
      If Cell.Value = SheetIndx.Range("B36").Value Then SlotCount = SheetIndx.Range("E36").Value
      If Cell.Value = SheetIndx.Range("B37").Value Then SlotCount = SheetIndx.Range("E37").Value
      If Cell.Value = SheetIndx.Range("B38").Value Then SlotCount = SheetIndx.Range("E38").Value
      If Cell.Value = SheetIndx.Range("B39").Value Then SlotCount = SheetIndx.Range("E39").Value
      If Cell.Value = SheetIndx.Range("B40").Value Then SlotCount = SheetIndx.Range("E40").Value
      If Cell.Value = SheetIndx.Range("B41").Value Then SlotCount = SheetIndx.Range("E41").Value
      If Cell.Value = SheetIndx.Range("B42").Value Then SlotCount = SheetIndx.Range("E42").Value
      If Cell.Value = SheetIndx.Range("B43").Value Then SlotCount = SheetIndx.Range("E43").Value
      If Cell.Value = SheetIndx.Range("B44").Value Then SlotCount = SheetIndx.Range("E44").Value
      If Cell.Value = SheetIndx.Range("B45").Value Then SlotCount = SheetIndx.Range("E45").Value
      If Cell.Value = SheetIndx.Range("B46").Value Then SlotCount = SheetIndx.Range("E46").Value
      If Cell.Value = SheetIndx.Range("B47").Value Then SlotCount = SheetIndx.Range("E47").Value
      If Cell.Value = SheetIndx.Range("B48").Value Then SlotCount = SheetIndx.Range("E48").Value
      If Cell.Value = SheetIndx.Range("B49").Value Then SlotCount = SheetIndx.Range("E49").Value
      If Cell.Value = SheetIndx.Range("B50").Value Then SlotCount = SheetIndx.Range("E50").Value
      If Cell.Value = SheetIndx.Range("B51").Value Then SlotCount = SheetIndx.Range("E51").Value
      If Cell.Value = SheetIndx.Range("B52").Value Then SlotCount = SheetIndx.Range("E52").Value
      If Cell.Value = SheetIndx.Range("B53").Value Then SlotCount = SheetIndx.Range("E53").Value
      If Cell.Value = SheetIndx.Range("B54").Value Then SlotCount = SheetIndx.Range("E54").Value
      If Cell.Value = SheetIndx.Range("B55").Value Then SlotCount = SheetIndx.Range("E55").Value
      If Cell.Value = SheetIndx.Range("B56").Value Then SlotCount = SheetIndx.Range("E56").Value
      If Cell.Value = SheetIndx.Range("B57").Value Then SlotCount = SheetIndx.Range("E57").Value
      If Cell.Value = SheetIndx.Range("B58").Value Then SlotCount = SheetIndx.Range("E58").Value
      If Cell.Value = SheetIndx.Range("B59").Value Then SlotCount = SheetIndx.Range("E59").Value
      If Cell.Value = SheetIndx.Range("B60").Value Then SlotCount = SheetIndx.Range("E60").Value
      If Cell.Value = SheetIndx.Range("B61").Value Then SlotCount = SheetIndx.Range("E61").Value
      If Cell.Value = SheetIndx.Range("B62").Value Then SlotCount = SheetIndx.Range("E62").Value
      If Cell.Value = SheetIndx.Range("B63").Value Then SlotCount = SheetIndx.Range("E63").Value
      If Cell.Value = SheetIndx.Range("B64").Value Then SlotCount = SheetIndx.Range("E64").Value
      If Cell.Value = SheetIndx.Range("B65").Value Then SlotCount = SheetIndx.Range("E65").Value
      If Cell.Value = SheetIndx.Range("B66").Value Then SlotCount = SheetIndx.Range("E66").Value
      If Cell.Value = SheetIndx.Range("B67").Value Then SlotCount = SheetIndx.Range("E67").Value
      If Cell.Value = SheetIndx.Range("B68").Value Then SlotCount = SheetIndx.Range("E68").Value
      If Cell.Value = SheetIndx.Range("B69").Value Then SlotCount = SheetIndx.Range("E69").Value
      If Cell.Value = SheetIndx.Range("B70").Value Then SlotCount = SheetIndx.Range("E70").Value
      If Cell.Value = SheetIndx.Range("B71").Value Then SlotCount = SheetIndx.Range("E71").Value
      If Cell.Value = SheetIndx.Range("B72").Value Then SlotCount = SheetIndx.Range("E72").Value
      If Cell.Value = SheetIndx.Range("B73").Value Then SlotCount = SheetIndx.Range("E73").Value
      If Cell.Value = SheetIndx.Range("B74").Value Then SlotCount = SheetIndx.Range("E74").Value
      If Cell.Value = SheetIndx.Range("B75").Value Then SlotCount = SheetIndx.Range("E75").Value
      If Cell.Value = SheetIndx.Range("B76").Value Then SlotCount = SheetIndx.Range("E76").Value
      If Cell.Value = SheetIndx.Range("B77").Value Then SlotCount = SheetIndx.Range("E77").Value
      If Cell.Value = SheetIndx.Range("B78").Value Then SlotCount = SheetIndx.Range("E78").Value
      If Cell.Value = SheetIndx.Range("B79").Value Then SlotCount = SheetIndx.Range("E79").Value
      If Cell.Value = SheetIndx.Range("B80").Value Then SlotCount = SheetIndx.Range("E80").Value
      If Cell.Value = SheetIndx.Range("B81").Value Then SlotCount = SheetIndx.Range("E81").Value
      If Cell.Value = SheetIndx.Range("B82").Value Then SlotCount = SheetIndx.Range("E82").Value
      If Cell.Value = SheetIndx.Range("B83").Value Then SlotCount = SheetIndx.Range("E83").Value
      If Cell.Value = SheetIndx.Range("B84").Value Then SlotCount = SheetIndx.Range("E84").Value
      If Cell.Value = SheetIndx.Range("B85").Value Then SlotCount = SheetIndx.Range("E85").Value
      If Cell.Value = SheetIndx.Range("B86").Value Then SlotCount = SheetIndx.Range("E86").Value
      If Cell.Value = SheetIndx.Range("B87").Value Then SlotCount = SheetIndx.Range("E87").Value
      If Cell.Value = SheetIndx.Range("B88").Value Then SlotCount = SheetIndx.Range("E88").Value
      If Cell.Value = SheetIndx.Range("B89").Value Then SlotCount = SheetIndx.Range("E89").Value
      If Cell.Value = SheetIndx.Range("B90").Value Then SlotCount = SheetIndx.Range("E90").Value
      If Cell.Value = SheetIndx.Range("B91").Value Then SlotCount = SheetIndx.Range("E91").Value
      If Cell.Value = SheetIndx.Range("B92").Value Then SlotCount = SheetIndx.Range("E92").Value
      If Cell.Value = SheetIndx.Range("B93").Value Then SlotCount = SheetIndx.Range("E93").Value
      If Cell.Value = SheetIndx.Range("B94").Value Then SlotCount = SheetIndx.Range("E94").Value
      If Cell.Value = SheetIndx.Range("B95").Value Then SlotCount = SheetIndx.Range("E95").Value
      If Cell.Value = SheetIndx.Range("B96").Value Then SlotCount = SheetIndx.Range("E96").Value
      If Cell.Value = SheetIndx.Range("B97").Value Then SlotCount = SheetIndx.Range("E97").Value
      If Cell.Value = SheetIndx.Range("B98").Value Then SlotCount = SheetIndx.Range("E98").Value
      If Cell.Value = SheetIndx.Range("B99").Value Then SlotCount = SheetIndx.Range("E99").Value
      If Cell.Value = SheetIndx.Range("B100").Value Then SlotCount = SheetIndx.Range("E100").Value
      If Cell.Value = SheetIndx.Range("B101").Value Then SlotCount = SheetIndx.Range("E101").Value
      If Cell.Value = SheetIndx.Range("B102").Value Then SlotCount = SheetIndx.Range("E102").Value
      If Cell.Value = SheetIndx.Range("B103").Value Then SlotCount = SheetIndx.Range("E103").Value
      If Cell.Value = SheetIndx.Range("B104").Value Then SlotCount = SheetIndx.Range("E104").Value
      If Cell.Value = SheetIndx.Range("B105").Value Then SlotCount = SheetIndx.Range("E105").Value
      If Cell.Value = SheetIndx.Range("B106").Value Then SlotCount = SheetIndx.Range("E106").Value
      If Cell.Value = SheetIndx.Range("B107").Value Then SlotCount = SheetIndx.Range("E107").Value
      If Cell.Value = SheetIndx.Range("B108").Value Then SlotCount = SheetIndx.Range("E108").Value
      If Cell.Value = SheetIndx.Range("B109").Value Then SlotCount = SheetIndx.Range("E109").Value
      If Cell.Value = SheetIndx.Range("B110").Value Then SlotCount = SheetIndx.Range("E110").Value
      If Cell.Value = SheetIndx.Range("B111").Value Then SlotCount = SheetIndx.Range("E111").Value
      If Cell.Value = SheetIndx.Range("B112").Value Then SlotCount = SheetIndx.Range("E112").Value
      If Cell.Value = SheetIndx.Range("B113").Value Then SlotCount = SheetIndx.Range("E113").Value
      If Cell.Value = SheetIndx.Range("B114").Value Then SlotCount = SheetIndx.Range("E114").Value
      If Cell.Value = SheetIndx.Range("B115").Value Then SlotCount = SheetIndx.Range("E115").Value
      If Cell.Value = SheetIndx.Range("B116").Value Then SlotCount = SheetIndx.Range("E116").Value
      If Cell.Value = SheetIndx.Range("B117").Value Then SlotCount = SheetIndx.Range("E117").Value
      If Cell.Value = SheetIndx.Range("B118").Value Then SlotCount = SheetIndx.Range("E118").Value
      If Cell.Value = SheetIndx.Range("B119").Value Then SlotCount = SheetIndx.Range("E119").Value
      If Cell.Value = SheetIndx.Range("B120").Value Then SlotCount = SheetIndx.Range("E120").Value
      If Cell.Value = SheetIndx.Range("B121").Value Then SlotCount = SheetIndx.Range("E121").Value
      If Cell.Value = SheetIndx.Range("B122").Value Then SlotCount = SheetIndx.Range("E122").Value
      If Cell.Value = SheetIndx.Range("B123").Value Then SlotCount = SheetIndx.Range("E123").Value
      If Cell.Value = SheetIndx.Range("B124").Value Then SlotCount = SheetIndx.Range("E124").Value
      If Cell.Value = SheetIndx.Range("B125").Value Then SlotCount = SheetIndx.Range("E125").Value
      If Cell.Value = SheetIndx.Range("B126").Value Then SlotCount = SheetIndx.Range("E126").Value
      If Cell.Value = SheetIndx.Range("B127").Value Then SlotCount = SheetIndx.Range("E127").Value
      If Cell.Value = SheetIndx.Range("B128").Value Then SlotCount = SheetIndx.Range("E128").Value
      If Cell.Value = SheetIndx.Range("B129").Value Then SlotCount = SheetIndx.Range("E129").Value
      If Cell.Value = SheetIndx.Range("B130").Value Then SlotCount = SheetIndx.Range("E130").Value
      If Cell.Value = SheetIndx.Range("B131").Value Then SlotCount = SheetIndx.Range("E131").Value
      If Cell.Value = SheetIndx.Range("B132").Value Then SlotCount = SheetIndx.Range("E132").Value
      If Cell.Value = SheetIndx.Range("B133").Value Then SlotCount = SheetIndx.Range("E133").Value
      If Cell.Value = SheetIndx.Range("B134").Value Then SlotCount = SheetIndx.Range("E134").Value
      If Cell.Value = SheetIndx.Range("B135").Value Then SlotCount = SheetIndx.Range("E135").Value
      If Cell.Value = SheetIndx.Range("B136").Value Then SlotCount = SheetIndx.Range("E136").Value
        
     'If SlotCount > Meanval Then DegreeOfFreedom = DegreeOfFreedom - 1
     
     'If SlotCount < Meanval Then DegreeOfFreedom = DegreeOfFreedom + 1
           
'__________________________________________________________________________________________________________________________
'[The Dynamic DegreeOfFreedom Algothrithm]
     
     If SlotCount = 0 Then
        Equalizer = True
        CountOfTrue = CountOfTrue + 1
     Else
        If SlotCount <= LowestFreq + DegreeOfFreedom Then
            Equalizer = True
            CountOfTrue = CountOfTrue + 1
     Else
        If SlotCount > LowestFreq + DegreeOfFreedom Then
            Equalizer = False
            CountOfFalse = CountOfFalse + 1
        End If
        End If
        End If
     
     
     'If DegreeOfFreedom > LowestDOF Then LowestDOF = DegreeOfFreedom
   
     
     'Adjusting the DegreeOfFreedom
     
     
     If CountOfFalse > CountOfTrue And DegreeOfFreedom < Difference Then
           DegreeOfFreedom = DegreeOfFreedom + 1
            Else
                If CountOfFalse < CountOfTrue And DegreeOfFreedom > 1 Then       'Condition for lowering DegreeOfFreedom
                    DegreeOfFreedom = DegreeOfFreedom - 1
                        Else
                            If CountOfFalse = CountOfTrue Then
                                DegreeOfFreedom = Difference
                End If
                End If
                End If
      
                

     'MsgBox CountOfTrue
     'MsgBox CountOfFalse
     'MsgBox DegreeOfFreedom
            
    
End Function