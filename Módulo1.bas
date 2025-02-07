Attribute VB_Name = "Módulo1"

Sub Macro3()

    ValoresAteAbril
    ValoresAteAgosto
    ValoresAteDezembro
    
End Sub
    
    
    ' MÊS JANEIRO
    Sub ValoresAteAbril()
    ' Define as variáveis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verificação
    Set wsOrigem = ThisWorkbook.Sheets("Lançamento TELECOM")
    Set wsDestino = ThisWorkbook.Sheets("TELECOM - BRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox2 As String
    Dim valorComboBoxMes As String
    
    valorComboBox2 = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBox2.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBoxMes.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim valorDinheiro As Variant
    Dim numRequisicao As Variant
    Dim valorEncargos As Variant
    
    Dim colunaValor As Variant
    Dim colunaReq As Variant
    Dim colunaEnc As Variant
    If valorComboBoxMes = wsVerificacao.Range("H19").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "I"
        colunaValor = "J"
        colunaEnc = "K"
        
        valorDinheiro = wsOrigem.Range("E12").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
            
    End If
    
    
        ' MÊS FEVEREIRO
    If valorComboBoxMes = wsVerificacao.Range("H20").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "P"
        colunaValor = "Q"
        colunaEnc = "R"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    ' MÊS MARÇO
    If valorComboBoxMes = wsVerificacao.Range("H21").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "W"
        colunaValor = "Y"
        colunaEnc = "Z"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    
    ' MÊS ABRIL
    If valorComboBoxMes = wsVerificacao.Range("H22").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "AD"
        colunaValor = "AE"
        colunaEnc = "AF"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    End Sub
    
    Sub ValoresAteAgosto()
    
    ' Define as variáveis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verificação
    Set wsOrigem = ThisWorkbook.Sheets("Lançamento TELECOM")
    Set wsDestino = ThisWorkbook.Sheets("TELECOM - BRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox2 As String
    Dim valorComboBoxMes As String
    
    valorComboBox2 = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBox2.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBoxMes.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim valorDinheiro As Variant
    Dim numRequisicao As Variant
    Dim valorEncargos As Variant
    
    Dim colunaValor As Variant
    Dim colunaReq As Variant
    Dim colunaEnc As Variant
    
    ' MÊS MAIO
    If valorComboBoxMes = wsVerificacao.Range("H23").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "AK"
        colunaValor = "AL"
        colunaEnc = "AM"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    ' MÊS JUNHO
    If valorComboBoxMes = wsVerificacao.Range("H24").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "AR"
        colunaValor = "AS"
        colunaEnc = "AT"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    ' MÊS JULHO
    If valorComboBoxMes = wsVerificacao.Range("H25").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "AY"
        colunaValor = "AZ"
        colunaEnc = "BA"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    'MÊS AGOSTO
    If valorComboBoxMes = wsVerificacao.Range("H26").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "BF"
        colunaValor = "BG"
        colunaEnc = "BH"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    End Sub
    
    
    Sub ValoresAteDezembro()
    
    ' Define as variáveis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verificação
    Set wsOrigem = ThisWorkbook.Sheets("Lançamento TELECOM")
    Set wsDestino = ThisWorkbook.Sheets("TELECOM - BRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox2 As String
    Dim valorComboBoxMes As String
    
    valorComboBox2 = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBox2.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBoxMes.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim valorDinheiro As Variant
    Dim numRequisicao As Variant
    Dim valorEncargos As Variant
    
    Dim colunaValor As Variant
    Dim colunaReq As Variant
    Dim colunaEnc As Variant
    
    'MÊS SETEMBRO
    If valorComboBoxMes = wsVerificacao.Range("H27").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "BM"
        colunaValor = "BN"
        colunaEnc = "BO"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    'MÊS OUTUBRO
    If valorComboBoxMes = wsVerificacao.Range("H28").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "BT"
        colunaValor = "BU"
        colunaEnc = "BV"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    'MÊS NOVEMBRO
    If valorComboBoxMes = wsVerificacao.Range("H29").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "CB"
        colunaValor = "CC"
        colunaEnc = "CD"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If
    
    'MÊS DEZEMBRO
    If valorComboBoxMes = wsVerificacao.Range("H30").Value Then
        ' Copia o valor da célula E12 da planilha de origem
        
        colunaReq = "CJ"
        colunaValor = "CK"
        colunaEnc = "CL"
        
        valorDinheiro = wsOrigem.Range("E12").Value
       
        valorEncargos = wsOrigem.Range("E18").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox2 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaReq & "5").Value = numRequisicao
            wsDestino.Range(colunaValor & "5").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "5").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "6").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaReq & "7").Value = numRequisicao
            wsDestino.Range(colunaValor & "7").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "7").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "8").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "9").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaReq & "10").Value = numRequisicao
            wsDestino.Range(colunaValor & "10").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "10").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "11").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "12").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "14").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "15").Value = valorEncargos
               
        ElseIf valorComboBox2 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaReq & "16").Value = numRequisicao
            wsDestino.Range(colunaValor & "16").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "16").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "17").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaReq & "18").Value = numRequisicao
            wsDestino.Range(colunaValor & "18").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "18").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("L5").Value Then
            'DDN CR01
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "19").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaReq & "20").Value = numRequisicao
            wsDestino.Range(colunaValor & "20").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "20").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaReq & "21").Value = numRequisicao
            wsDestino.Range(colunaValor & "21").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "21").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaReq & "22").Value = numRequisicao
            wsDestino.Range(colunaValor & "22").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "22").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaReq & "23").Value = numRequisicao
            wsDestino.Range(colunaValor & "23").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "23").Value = valorEncargos
              
        ElseIf valorComboBox2 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaReq & "24").Value = numRequisicao
            wsDestino.Range(colunaValor & "24").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "24").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaReq & "26").Value = numRequisicao
            wsDestino.Range(colunaValor & "26").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "26").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaReq & "27").Value = numRequisicao
            wsDestino.Range(colunaValor & "27").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "27").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaReq & "29").Value = numRequisicao
            wsDestino.Range(colunaValor & "29").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "29").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaReq & "30").Value = numRequisicao
            wsDestino.Range(colunaValor & "30").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "30").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaReq & "31").Value = numRequisicao
            wsDestino.Range(colunaValor & "31").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "31").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaReq & "33").Value = numRequisicao
            wsDestino.Range(colunaValor & "33").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "33").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaReq & "34").Value = numRequisicao
            wsDestino.Range(colunaValor & "34").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "34").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaReq & "35").Value = numRequisicao
            wsDestino.Range(colunaValor & "35").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "35").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaReq & "36").Value = numRequisicao
            wsDestino.Range(colunaValor & "36").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "36").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaReq & "37").Value = numRequisicao
            wsDestino.Range(colunaValor & "37").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "37").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaReq & "38").Value = numRequisicao
            wsDestino.Range(colunaValor & "38").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "38").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaReq & "39").Value = numRequisicao
            wsDestino.Range(colunaValor & "39").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "39").Value = valorEncargos
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaReq & "40").Value = numRequisicao
            wsDestino.Range(colunaValor & "40").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "40").Value = valorEncargos
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaReq & "41").Value = numRequisicao
            wsDestino.Range(colunaValor & "41").Value = valorDinheiro
            wsDestino.Range(colunaEnc & "41").Value = valorEncargos
            
        End If
        
    End If

    
End Sub


Sub LimparRc()
    ' Definir a planilha onde as células estão localizadas
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Lançamento TELECOM")

    ' Limpar o conteúdo das células específicas
    ws.Range("E12").ClearContents
    ws.Range("E18").ClearContents
    ws.Range("E25").ClearContents
End Sub

' -------------------------------------------------------------------------------------------------

Sub EnviarMigo()

    ValoresAteAbrilMigo
    ValoresAteAgostoMigo
    ValoresAteDezembroMigo
    
' -------------------------------------------------------------------------------------------------
End Sub



    
    
    ' MÊS JANEIRO
    Sub ValoresAteAbrilMigo()
    ' Define as variáveis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verificação
    Set wsOrigem = ThisWorkbook.Sheets("Lançamento TELECOM")
    Set wsDestino = ThisWorkbook.Sheets("TELECOM - BRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox4 As String
    Dim valorComboBoxMes As String
    
    valorComboBox4 = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBox4.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBox5.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim pedido As Variant
    Dim migo As Variant
    
    Dim colunaPedido As Variant
    Dim colunaMigo As Variant

    If valorComboBoxMes = wsVerificacao.Range("H19").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "L"
        colunaMigo = "M"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "32").Value = pedido
            wsDestino.Range(colunaMigo & "32").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
           wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
            
    End If
    
    
        ' MÊS FEVEREIRO
    If valorComboBoxMes = wsVerificacao.Range("H20").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "S"
        colunaMigo = "T"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    ' MÊS MARÇO
    If valorComboBoxMes = wsVerificacao.Range("H21").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "Z"
        colunaMigo = "AA"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    
    ' MÊS ABRIL
    If valorComboBoxMes = wsVerificacao.Range("H22").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "AG"
        colunaMigo = "AH"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    End Sub
    
    Sub ValoresAteAgostoMigo()
    
     ' Define as variáveis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verificação
    Set wsOrigem = ThisWorkbook.Sheets("Lançamento TELECOM")
    Set wsDestino = ThisWorkbook.Sheets("TELECOM - BRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox4 As String
    Dim valorComboBox5 As String
    
    valorComboBox4 = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBox4.Value
    valorComboBox5 = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBox5.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim pedido As Variant
    Dim migo As Variant
    
    Dim colunaPedido As Variant
    Dim colunaMigo As Variant
    
    ' MÊS MAIO
     If valorComboBox5 = wsVerificacao.Range("H23").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "AN"
        colunaMigo = "AO"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    ' MÊS JUNHO
    If valorComboBox5 = wsVerificacao.Range("H24").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "AU"
        colunaMigo = "AV"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    ' MÊS JULHO
   If valorComboBox5 = wsVerificacao.Range("H25").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "BB"
        colunaMigo = "BC"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    'MÊS AGOSTO
    If valorComboBox5 = wsVerificacao.Range("H26").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "BI"
        colunaMigo = "BJ"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    End Sub
    
    
    Sub ValoresAteDezembroMigo()
    
     ' Define as variáveis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verificação
    Set wsOrigem = ThisWorkbook.Sheets("Lançamento TELECOM")
    Set wsDestino = ThisWorkbook.Sheets("TELECOM - BRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox4 As String
    Dim valorComboBox5 As String
    
    valorComboBox4 = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBox4.Value
    valorComboBox5 = ThisWorkbook.Sheets("Lançamento TELECOM").ComboBox5.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim pedido As Variant
    Dim migo As Variant
    
    Dim colunaPedido As Variant
    Dim colunaMigo As Variant
    
    'MÊS SETEMBRO
    If valorComboBox5 = wsVerificacao.Range("H27").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "BP"
        colunaMigo = "BQ"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    'MÊS OUTUBRO
    If valorComboBox5 = wsVerificacao.Range("H28").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "BX"
        colunaMigo = "BY"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    'MÊS NOVEMBRO
    If valorComboBox5 = wsVerificacao.Range("H29").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "CF"
        colunaMigo = "CG"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
    'MÊS DEZEMBRO
    If valorComboBox5 = wsVerificacao.Range("H30").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "CN"
        colunaMigo = "CO"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
         'VIVO 0371942405 CR01
        If valorComboBox4 = wsVerificacao.Range("E4").Value Then
            wsDestino.Range(colunaPedido & "5").Value = pedido
            wsDestino.Range(colunaMigo & "5").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F4").Value Then
            'VIVO 0154012356 ES01
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G4").Value Then
            'VIVO 0373741211 ES01
            wsDestino.Range(colunaPedido & "7").Value = pedido
            wsDestino.Range(colunaMigo & "7").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H4").Value Then
            'VIVO 0372127129 ES02
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I4").Value Then
            'VIVO 0374041895 ES05
           wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J4").Value Then
            'VIVO 373923811 ES03
            wsDestino.Range(colunaPedido & "10").Value = pedido
            wsDestino.Range(colunaMigo & "10").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E6").Value Then
            'TIM
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E7").Value Then
            'Y3
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E8").Value Then
            'Manutençao Central Telefonica -  SP
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F8").Value Then
            'Locação de Central Telefonica - Matriz e Filiais
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
               
        ElseIf valorComboBox4 = wsVerificacao.Range("E9").Value Then
            'TMS
            wsDestino.Range(colunaPedido & "16").Value = pedido
            wsDestino.Range(colunaMigo & "16").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E5").Value Then
            'DDN ESAB
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("L5").Value Then
            'DDN ES03
            wsDestino.Range(colunaPedido & "18").Value = pedido
            wsDestino.Range(colunaMigo & "18").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("F5").Value Then
            'DDN CR01
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G5").Value Then
            'VPE ESAB
            wsDestino.Range(colunaPedido & "20").Value = pedido
            wsDestino.Range(colunaMigo & "20").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H5").Value Then
            'VPE CR01
            wsDestino.Range(colunaPedido & "21").Value = pedido
            wsDestino.Range(colunaMigo & "21").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E11").Value Then
            'Telefonia Matriz 7090 ES01
            wsDestino.Range(colunaPedido & "22").Value = pedido
            wsDestino.Range(colunaMigo & "22").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F11").Value Then
            'Telefonia Site II 9812 ES07
            wsDestino.Range(colunaPedido & "23").Value = pedido
            wsDestino.Range(colunaMigo & "23").Value = migo
              
        ElseIf valorComboBox4 = wsVerificacao.Range("K5").Value Then
            'BSP ESAB
            wsDestino.Range(colunaPedido & "24").Value = pedido
            wsDestino.Range(colunaMigo & "24").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E13").Value Then
            'OI POA 3774
            wsDestino.Range(colunaPedido & "26").Value = pedido
            wsDestino.Range(colunaMigo & "26").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("F13").Value Then
            'OI POA 3871
            wsDestino.Range(colunaPedido & "27").Value = pedido
            wsDestino.Range(colunaMigo & "27").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I5").Value Then
            'INN ESAB RIO/MATRIZ
            wsDestino.Range(colunaPedido & "29").Value = pedido
            wsDestino.Range(colunaMigo & "29").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("J5").Value Then
            'INN CR01
            wsDestino.Range(colunaPedido & "30").Value = pedido
            wsDestino.Range(colunaMigo & "30").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("M5").Value Then
            'INN ESAB ES05
            wsDestino.Range(colunaPedido & "31").Value = pedido
            wsDestino.Range(colunaMigo & "31").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("G11").Value Then
            'INTERNET MATRIZ E SP ALGAR
            wsDestino.Range(colunaPedido & "33").Value = pedido
            wsDestino.Range(colunaMigo & "33").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H11").Value Then
            'INTERNET UNIDADE II ALGAR
            wsDestino.Range(colunaPedido & "34").Value = pedido
            wsDestino.Range(colunaMigo & "34").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E12").Value Then
            'INTERNET SALVADOR ALGAR
            wsDestino.Range(colunaPedido & "35").Value = pedido
            wsDestino.Range(colunaMigo & "35").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F12").Value Then
            'LAN TO LAN SITE II
            wsDestino.Range(colunaPedido & "36").Value = pedido
            wsDestino.Range(colunaMigo & "36").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G12").Value Then
            'LAN TO LAN CONDOR
            wsDestino.Range(colunaPedido & "37").Value = pedido
            wsDestino.Range(colunaMigo & "37").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E14").Value Then
            'ÁVATO
            wsDestino.Range(colunaPedido & "38").Value = pedido
            wsDestino.Range(colunaMigo & "38").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E15").Value Then
            'CENTURY
            wsDestino.Range(colunaPedido & "39").Value = pedido
            wsDestino.Range(colunaMigo & "39").Value = migo
             
        ElseIf valorComboBox4 = wsVerificacao.Range("E10").Value Then
            'Link MPLS AT&T AVPN
            wsDestino.Range(colunaPedido & "40").Value = pedido
            wsDestino.Range(colunaMigo & "40").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F10").Value Then
            'Link NetBond AT&T With Azure
            wsDestino.Range(colunaPedido & "41").Value = pedido
            wsDestino.Range(colunaMigo & "41").Value = migo
            
        End If
        
    End If
    
End Sub

Sub LimparMigo()
    ' Definir a planilha onde as células estão localizadas
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Lançamento TELECOM")

    ' Limpar o conteúdo das células específicas
    ws.Range("V15").ClearContents
    ws.Range("V22").ClearContents
End Sub
