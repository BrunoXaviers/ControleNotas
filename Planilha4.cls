VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' VERIFICAR AS DATAS DE VENCIMENTO E CRIAR UM ALERTA
Function DataDia() As Integer
    Dim dataAtual As Date
    Dim dia As Integer
    
    dataAtual = Date
    dia = Day(dataAtual)
    
    DataDia = dia
    
End Function

Function DataMes() As String
    Dim dataAtual As Date
    Dim mes As Integer
    Dim mesCelulaAtual As String
    
    dataAtual = Date
    mes = Month(dataAtual)
    
    Select Case mes
        Case 1
            mesCelulaAtual = "I"
        Case 2
            mesCelulaAtual = "P"
        Case 3
            mesCelulaAtual = "W"
        Case 4
            mesCelulaAtual = "AD"
        Case 5
            mesCelulaAtual = "AK"
        Case 6
            mesCelulaAtual = "AR"
        Case 7
            mesCelulaAtual = "AY"
        Case 8
            mesCelulaAtual = "BF"
        Case 9
            mesCelulaAtual = "BM"
        Case 10
            mesCelulaAtual = "BT"
        Case 11
            mesCelulaAtual = "CB"
        Case 12
            mesCelulaAtual = "CJ"
    End Select
    
    DataMes = mesCelulaAtual
End Function


Sub PintarCelula()
    Dim diaCelula As Integer
    Dim diferenca As Integer
    Dim dia As Integer
    Dim mes As String
    Dim i As Integer
    Dim j As Integer
    Dim contador As Integer
    Dim wsOrigem As Worksheet
    Set wsOrigem = ThisWorkbook.Sheets("Sheet2")
    dia = DataDia()

    mes = DataMes()
    
    MsgBox (mes)
    
    j = 5
    
    For i = 8 To 10
    diaCelula = wsOrigem.Range("B" & i).Value
    
    diferenca = diaCelula - dia
    
    If diferenca < 4 Then
        Range("A" & j).Interior.Color = RGB(255, 0, 0) ' Vermelho
    ElseIf diferenca < 7 Then
        Range("A" & j).Interior.Color = RGB(255, 255, 0) ' Amarelo
    Else
        Range("A" & j).Interior.ColorIndex = xlNone ' Sem cor
    End If
    j = j + 1
    Next i
    
End Sub
