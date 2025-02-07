Attribute VB_Name = "M�dulo2"
Sub EnviarRcInfra()

    ValoresAteAbril
    ValoresAteAgosto
    ValoresAteDezembro
    
End Sub
    
    
    ' M�S JANEIRO
    Sub ValoresAteAbril()
    ' Define as vari�veis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verifica��o
    Set wsOrigem = ThisWorkbook.Sheets("Lan�amento INFRA")
    Set wsDestino = ThisWorkbook.Sheets("INFRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox2 As String
    Dim valorComboBoxMes As String
    
    valorComboBox2 = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBox2.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBoxMes.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim valorDinheiro As Variant
    Dim numRequisicao As Variant
    
    Dim colunaValor As Variant
    Dim colunaReq As Variant
    If valorComboBoxMes = wsVerificacao.Range("H19").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "I"
        colunaValor = "J"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
       
            'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
            
    End If
    
    
        ' M�S FEVEREIRO
    If valorComboBoxMes = wsVerificacao.Range("H20").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "P"
        colunaValor = "Q"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
       'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    ' M�S MAR�O
    If valorComboBoxMes = wsVerificacao.Range("H21").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "V"
        colunaValor = "W"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
        'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    
    ' M�S ABRIL
    If valorComboBoxMes = wsVerificacao.Range("H22").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "AB"
        colunaValor = "AC"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
        'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    End Sub
    
    Sub ValoresAteAgosto()
    
    ' Define as vari�veis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verifica��o
    Set wsOrigem = ThisWorkbook.Sheets("Lan�amento INFRA")
    Set wsDestino = ThisWorkbook.Sheets("INFRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox2 As String
    Dim valorComboBoxMes As String
    
    valorComboBox2 = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBox2.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBoxMes.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim valorDinheiro As Variant
    Dim numRequisicao As Variant
    Dim valorEncargos As Variant
    
    Dim colunaValor As Variant
    Dim colunaReq As Variant
    Dim colunaEnc As Variant
    
    ' M�S MAIO
    If valorComboBoxMes = wsVerificacao.Range("H23").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "AH"
        colunaValor = "AI"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
            'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    ' M�S JUNHO
    If valorComboBoxMes = wsVerificacao.Range("H24").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "AN"
        colunaValor = "AO"
        
        valorDinheiro = wsOrigem.Range("E15").Value
       
        numRequisicao = wsOrigem.Range("E25").Value
        
            'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    ' M�S JULHO
    If valorComboBoxMes = wsVerificacao.Range("H25").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "AT"
        colunaValor = "AU"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
            'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    'M�S AGOSTO
    If valorComboBoxMes = wsVerificacao.Range("H26").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "AZ"
        colunaValor = "BA"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
            'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    End Sub
    
    
    Sub ValoresAteDezembro()
    
    ' Define as vari�veis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verifica��o
    Set wsOrigem = ThisWorkbook.Sheets("Lan�amento INFRA")
    Set wsDestino = ThisWorkbook.Sheets("INFRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox2 As String
    Dim valorComboBoxMes As String
    
    valorComboBox2 = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBox2.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBoxMes.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim valorDinheiro As Variant
    Dim numRequisicao As Variant
    
    Dim colunaValor As Variant
    Dim colunaReq As Variant
    
    'M�S SETEMBRO
    If valorComboBoxMes = wsVerificacao.Range("H27").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "BF"
        colunaValor = "BG"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
            'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    'M�S OUTUBRO
    If valorComboBoxMes = wsVerificacao.Range("H28").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "BM"
        colunaValor = "BN"
        
        valorDinheiro = wsOrigem.Range("E12").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
            'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    'M�S NOVEMBRO
    If valorComboBoxMes = wsVerificacao.Range("H29").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "BT"
        colunaValor = "BU"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
            'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If
    
    'M�S DEZEMBRO
    If valorComboBoxMes = wsVerificacao.Range("H30").Value Then
        ' Copia o valor da c�lula E12 da planilha de origem
        
        colunaReq = "CA"
        colunaValor = "CB"
        
        valorDinheiro = wsOrigem.Range("E15").Value
        
        numRequisicao = wsOrigem.Range("E25").Value
        
            'SIMPRESS
        If valorComboBox2 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaReq & "6").Value = numRequisicao
            wsDestino.Range(colunaValor & "6").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaReq & "8").Value = numRequisicao
            wsDestino.Range(colunaValor & "8").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaReq & "9").Value = numRequisicao
            wsDestino.Range(colunaValor & "9").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaReq & "11").Value = numRequisicao
            wsDestino.Range(colunaValor & "11").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaReq & "12").Value = numRequisicao
            wsDestino.Range(colunaValor & "12").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaReq & "13").Value = numRequisicao
            wsDestino.Range(colunaValor & "13").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaReq & "14").Value = numRequisicao
            wsDestino.Range(colunaValor & "14").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaReq & "15").Value = numRequisicao
            wsDestino.Range(colunaValor & "15").Value = valorDinheiro
            
        ElseIf valorComboBox2 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaReq & "17").Value = numRequisicao
            wsDestino.Range(colunaValor & "17").Value = valorDinheiro
             
        ElseIf valorComboBox2 = wsVerificacao.Range("E36").Value Then
            'LENOVO
            wsDestino.Range(colunaReq & "19").Value = numRequisicao
            wsDestino.Range(colunaValor & "19").Value = valorDinheiro
            
        End If
        
    End If

    
End Sub


Sub LimparRCinfra()
    ' Definir a planilha onde as c�lulas est�o localizadas
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Lan�amento INFRA")

    ' Limpar o conte�do das c�lulas espec�ficas
    ws.Range("E15").ClearContents
    ws.Range("E25").ClearContents
End Sub

' -------------------------------------------------------------------------------------------------

Sub EnviarServico()

    ValoresAteAbrilServico
    ValoresAteAgostoServico
    ValoresAteDezembroServico
    
' -------------------------------------------------------------------------------------------------
End Sub



    Sub ValoresAteAbrilServico()
    ' Define as vari�veis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verifica��o
    Set wsOrigem = ThisWorkbook.Sheets("Lan�amento INFRA")
    Set wsDestino = ThisWorkbook.Sheets("INFRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox4 As String
    Dim valorComboBoxMes As String
    
    valorComboBox4 = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBox4.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBox5.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim pedido As Variant
    Dim migo As Variant
    
    Dim colunaPedido As Variant
    Dim colunaMigo As Variant

    'M�S JANEIRO
    If valorComboBoxMes = wsVerificacao.Range("H19").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "L"
        colunaMigo = "M"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
               
        End If
        
            
    End If
    
    'M�S FEVEREIRO
    If valorComboBoxMes = wsVerificacao.Range("H20").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "R"
        colunaMigo = "S"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
               
        End If
        
            
    End If
    
    'M�S MAR�O
     If valorComboBoxMes = wsVerificacao.Range("H21").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "X"
        colunaMigo = "Y"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
        End If
            
    End If

        'M�S ABRIL
    If valorComboBoxMes = wsVerificacao.Range("H22").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "AD"
        colunaMigo = "AE"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        End If
        
            
    End If
    
End Sub
    
    
    
    Sub ValoresAteAgostoServico()
    ' Define as vari�veis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verifica��o
    Set wsOrigem = ThisWorkbook.Sheets("Lan�amento INFRA")
    Set wsDestino = ThisWorkbook.Sheets("INFRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox4 As String
    Dim valorComboBoxMes As String
    
    valorComboBox4 = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBox4.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBox5.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim pedido As Variant
    Dim migo As Variant
    
    Dim colunaPedido As Variant
    Dim colunaMigo As Variant

    'M�S MAIO
    If valorComboBoxMes = wsVerificacao.Range("H23").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "AJ"
        colunaMigo = "AK"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
               
        End If
        
            
    End If
    
    'M�S JUNHO
    If valorComboBoxMes = wsVerificacao.Range("H24").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "AQ"
        colunaMigo = "AR"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
            
        End If
        
            
    End If
    
        'M�S JULHO
    If valorComboBoxMes = wsVerificacao.Range("H25").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "AV"
        colunaMigo = "AW"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
               
        End If
        
            
    End If

        'M�S AGOSTO
    If valorComboBoxMes = wsVerificacao.Range("H26").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "BB"
        colunaMigo = "BC"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
               
        End If
        
            
    End If
    
End Sub




Sub ValoresAteDezembroServico()
    ' Define as vari�veis para as planilhas
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsVerificacao As Worksheet
    
    ' Define as planilhas de origem, destino e verifica��o
    Set wsOrigem = ThisWorkbook.Sheets("Lan�amento INFRA")
    Set wsDestino = ThisWorkbook.Sheets("INFRA")
    Set wsVerificacao = ThisWorkbook.Sheets("Sheet2")
    
    ' Acessa os valores dos ComboBoxes
    Dim valorComboBox4 As String
    Dim valorComboBoxMes As String
    
    valorComboBox4 = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBox4.Value
    valorComboBoxMes = ThisWorkbook.Sheets("Lan�amento INFRA").ComboBox5.Value
    
   ' Declarando as variaveis que usaremos para receber os valores a sem preenchidos
    Dim pedido As Variant
    Dim migo As Variant
    
    Dim colunaPedido As Variant
    Dim colunaMigo As Variant

    'M�S SETEMBRO
    If valorComboBoxMes = wsVerificacao.Range("H27").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "BI"
        colunaMigo = "BJ"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
               
        End If
        
            
    End If
    
    'M�S OUTUBRO
    If valorComboBoxMes = wsVerificacao.Range("H28").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "BP"
        colunaMigo = "BQ"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
               
        End If
        
            
    End If
    
        'M�S NOVEMBRO
    If valorComboBoxMes = wsVerificacao.Range("H29").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "BW"
        colunaMigo = "BX"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
               
        End If
        
            
    End If

        'M�S DEZEMBRO
    If valorComboBoxMes = wsVerificacao.Range("H30").Value Then
        ' Copia o valor da planilha de origem
        
        colunaPedido = "CD"
        colunaMigo = "CE"
        
        pedido = wsOrigem.Range("V15").Value
        
        migo = wsOrigem.Range("V22").Value
        
        
         'SIMPRESS
        If valorComboBox4 = wsVerificacao.Range("E33").Value Then
            wsDestino.Range(colunaPedido & "6").Value = pedido
            wsDestino.Range(colunaMigo & "6").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E34").Value Then
            'AGASUS ES01
            wsDestino.Range(colunaPedido & "8").Value = pedido
            wsDestino.Range(colunaMigo & "8").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F34").Value Then
            'AGASUS CR01
            wsDestino.Range(colunaPedido & "9").Value = pedido
            wsDestino.Range(colunaMigo & "9").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("E35").Value Then
            'CABTEC ES01
            wsDestino.Range(colunaPedido & "11").Value = pedido
            wsDestino.Range(colunaMigo & "11").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("F35").Value Then
            'CABTEC ES07
            wsDestino.Range(colunaPedido & "12").Value = pedido
            wsDestino.Range(colunaMigo & "12").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("G35").Value Then
            'CABTEC CR01
            wsDestino.Range(colunaPedido & "13").Value = pedido
            wsDestino.Range(colunaMigo & "13").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("H35").Value Then
            'CABTEC ES03
            wsDestino.Range(colunaPedido & "14").Value = pedido
            wsDestino.Range(colunaMigo & "14").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("I35").Value Then
            'CABTEC ES05
            wsDestino.Range(colunaPedido & "15").Value = pedido
            wsDestino.Range(colunaMigo & "15").Value = migo
            
        ElseIf valorComboBox4 = wsVerificacao.Range("J35").Value Then
            'CABTEC GARANTIA ES01
            wsDestino.Range(colunaPedido & "17").Value = pedido
            wsDestino.Range(colunaMigo & "17").Value = migo
             
        ElseIf valorComboBo4 = wsVerificacao.Range("F36").Value Then
            'LENOVO
            wsDestino.Range(colunaPedido & "19").Value = pedido
            wsDestino.Range(colunaMigo & "19").Value = migo
               
        End If
        
            
    End If
    
    End Sub
    
    
    
    Sub LimparServico()
    ' Definir a planilha onde as c�lulas est�o localizadas
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Lan�amento INFRA")

    ' Limpar o conte�do das c�lulas espec�ficas
    ws.Range("V15").ClearContents
    ws.Range("V22").ClearContents
End Sub
