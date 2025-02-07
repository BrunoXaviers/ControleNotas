VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ComboBox1_Change()
    Dim wsOrigem As Worksheet
    Set wsOrigem = ThisWorkbook.Sheets("Sheet2")
    
    ' Adiciona itens ao ComboBox1 se eles ainda não existirem
    AdicionarItemAoComboBox ComboBox1, wsOrigem.Range("D33").Value
    AdicionarItemAoComboBox ComboBox1, wsOrigem.Range("D34").Value
    AdicionarItemAoComboBox ComboBox1, wsOrigem.Range("D35").Value
    AdicionarItemAoComboBox ComboBox1, wsOrigem.Range("D36").Value
    
    ' Verifica se o valor de ComboBox1 é igual ao conteúdo da célula D4 ou D5
    If ComboBox1.Value = wsOrigem.Range("D33").Value Then
        ComboBox2.Enabled = True
        PreencherComboBox2 wsOrigem, wsOrigem.Range("E33")
    ElseIf ComboBox1.Value = wsOrigem.Range("D34").Value Then
        ComboBox2.Enabled = True
        PreencherComboBox2 wsOrigem, wsOrigem.Range("E34:F34")
    ElseIf ComboBox1.Value = wsOrigem.Range("D35").Value Then
        ComboBox2.Enabled = True
        PreencherComboBox2 wsOrigem, wsOrigem.Range("E35:J35")
    ElseIf ComboBox1.Value = wsOrigem.Range("D36").Value Then
        ComboBox2.Enabled = True
        PreencherComboBox2 wsOrigem, wsOrigem.Range("E36")
    Else
        ComboBox2.Enabled = False
        ComboBox2.Clear
    End If
    
    ' Limpar ComboBoxMes antes de preencher novamente
    ComboBoxMes.Clear
    
    ' Preencher o array com os valores desejados
    Dim valores() As Variant
    ReDim valores(1 To 12)
    
    valores(1) = wsOrigem.Range("H19").Value
    valores(2) = wsOrigem.Range("H20").Value
    valores(3) = wsOrigem.Range("H21").Value
    valores(4) = wsOrigem.Range("H22").Value
    valores(5) = wsOrigem.Range("H23").Value
    valores(6) = wsOrigem.Range("H24").Value
    valores(7) = wsOrigem.Range("H25").Value
    valores(8) = wsOrigem.Range("H26").Value
    valores(9) = wsOrigem.Range("H27").Value
    valores(10) = wsOrigem.Range("H28").Value
    valores(11) = wsOrigem.Range("H29").Value
    valores(12) = wsOrigem.Range("H30").Value
    
    ' Adicionar os itens ao ComboBoxMes
    AdicionarItensAoComboBox ComboBoxMes, valores

End Sub
Sub AdicionarItensAoComboBox(comboBox As MSForms.comboBox, valores() As Variant)
    Dim i As Integer
    For i = LBound(valores) To UBound(valores)
        comboBox.AddItem valores(i)
    Next i
End Sub

' Função para adicionar item ao ComboBox somente se ele ainda não existir
Sub AdicionarItemAoComboBox(cb As comboBox, item As Variant)
    Dim i As Long
    Dim existe As Boolean
    existe = False

    ' Verifica se o item já existe no ComboBox
    For i = 1 To cb.ListCount
        If cb.List(i - 1) = item Then
            existe = True
            Exit For
        End If
    Next i

    ' Se o item não existir, adiciona-o ao ComboBox
    If Not existe Then
        cb.AddItem item
    End If
End Sub

' Função para preencher ComboBox2 com valores de um intervalo específico
Sub PreencherComboBox2(ws As Worksheet, rng As Range)
    Dim cel As Range
    
    ' Limpa ComboBox2 antes de adicionar novos itens
    ComboBox2.Clear
    
    ' Adiciona cada célula do intervalo ao ComboBox2
    For Each cel In rng
        ComboBox2.AddItem cel.Value
    Next cel
End Sub




Sub ComboBox3_Change()
    Dim wsOrigem As Worksheet
    Set wsOrigem = ThisWorkbook.Sheets("Sheet2")
    
    ' Adiciona itens ao ComboBox1 se eles ainda não existirem
    AdicionarItemAoComboBox ComboBox3, wsOrigem.Range("D33").Value
    AdicionarItemAoComboBox ComboBox3, wsOrigem.Range("D34").Value
    AdicionarItemAoComboBox ComboBox3, wsOrigem.Range("D35").Value
    AdicionarItemAoComboBox ComboBox3, wsOrigem.Range("D36").Value
    
    ' Verifica se o valor de ComboBox1 é igual ao conteúdo da célula D4 ou D5
    If ComboBox3.Value = wsOrigem.Range("D33").Value Then
        ComboBox4.Enabled = True
        PreencherComboBox4 wsOrigem, wsOrigem.Range("E33")
    ElseIf ComboBox3.Value = wsOrigem.Range("D34").Value Then
        ComboBox4.Enabled = True
        PreencherComboBox4 wsOrigem, wsOrigem.Range("E34:F34")
    ElseIf ComboBox3.Value = wsOrigem.Range("D35").Value Then
        ComboBox4.Enabled = True
        PreencherComboBox4 wsOrigem, wsOrigem.Range("E35:J35")
    ElseIf ComboBox3.Value = wsOrigem.Range("D36").Value Then
        ComboBox4.Enabled = True
        PreencherComboBox4 wsOrigem, wsOrigem.Range("E36")
    Else
        ComboBox4.Enabled = False
        ComboBox4.Clear
    End If
    
    ' Limpar ComboBoxMes antes de preencher novamente
    ComboBox5.Clear
    
    ' Preencher o array com os valores desejados
    Dim valores() As Variant
    ReDim valores(1 To 12)
    
    valores(1) = wsOrigem.Range("H19").Value
    valores(2) = wsOrigem.Range("H20").Value
    valores(3) = wsOrigem.Range("H21").Value
    valores(4) = wsOrigem.Range("H22").Value
    valores(5) = wsOrigem.Range("H23").Value
    valores(6) = wsOrigem.Range("H24").Value
    valores(7) = wsOrigem.Range("H25").Value
    valores(8) = wsOrigem.Range("H26").Value
    valores(9) = wsOrigem.Range("H27").Value
    valores(10) = wsOrigem.Range("H28").Value
    valores(11) = wsOrigem.Range("H29").Value
    valores(12) = wsOrigem.Range("H30").Value
    
    ' Adicionar os itens ao ComboBoxMes
    AdicionarItensAoComboBox ComboBox5, valores

End Sub
Sub AdicionarItensAoComboBox4(comboBox As MSForms.comboBox, valores() As Variant)
    Dim i As Integer
    For i = LBound(valores) To UBound(valores)
        comboBox.AddItem valores(i)
    Next i
End Sub

' Função para adicionar item ao ComboBox somente se ele ainda não existir
Sub AdicionarItemAoComboBox4(cb As comboBox, item As Variant)
    Dim i As Long
    Dim existe As Boolean
    existe = False

    ' Verifica se o item já existe no ComboBox
    For i = 1 To cb.ListCount
        If cb.List(i - 1) = item Then
            existe = True
            Exit For
        End If
    Next i

    ' Se o item não existir, adiciona-o ao ComboBox
    If Not existe Then
        cb.AddItem item
    End If
End Sub

' Função para preencher ComboBox2 com valores de um intervalo específico
Sub PreencherComboBox4(ws As Worksheet, rng As Range)
    Dim cel As Range
    
    ' Limpa ComboBox2 antes de adicionar novos itens
    ComboBox4.Clear
    
    ' Adiciona cada célula do intervalo ao ComboBox2
    For Each cel In rng
        ComboBox4.AddItem cel.Value
    Next cel
End Sub

Private Sub EnviarMigo_Click()

End Sub
