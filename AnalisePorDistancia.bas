Attribute VB_Name = "AnalisePorDistancia"
Sub FiltrarPorRaioProximidade()
    Dim wsMatriz As Worksheet ' Original: ws
    Dim ultimaLinhaOrigem As Long, i As Long, j As Long ' Original: ultimaLinha
    Dim limiteRaio As Double ' Original: raioMax
    Dim proximaLinhaDestino As Long ' Original: linhaDestino

    ' Definição da planilha de processamento
    Set wsMatriz = Worksheets("Distancias")
    
    ' Valor do filtro de distância definido pelo usuário
    limiteRaio = wsMatriz.Range("M2").Value
    
    ' Chamada do procedimento de ordenação (Mantendo a lógica original)
    Call OrdenarRegistrosPorProximidade
    
    ' Limpeza da área de saída de dados
    wsMatriz.Range("J6:P1000").ClearContents

    ' Configuração do Cabeçalho de Resultados
    wsMatriz.Range("J4:P4").Value = Array("Area_Logistica", "ID_Ponto", "Municipio", "Localizacao", "Coord_Lat", "Coord_Long", "Dist_Calculada_KM")

    proximaLinhaDestino = 6
    ' Localiza a última linha preenchida na coluna de métricas (Coluna H)
    ultimaLinhaOrigem = wsMatriz.Cells(wsMatriz.Rows.Count, "H").End(xlUp).Row

    ' Processamento de Filtragem: Percorre a matriz de dados original
    For i = 3 To ultimaLinhaOrigem
        ' Verifica se o registro está dentro do raio estipulado e não é nulo
        If wsMatriz.Cells(i, "H").Value <= limiteRaio And wsMatriz.Cells(i, "H").Value <> "" Then
            
            ' Transferência seletiva de dados (Colunas B:H para J:P)
            For j = 2 To 8
                wsMatriz.Cells(proximaLinhaDestino, j + 8).Value = wsMatriz.Cells(i, j).Value
            Next j
            
            proximaLinhaDestino = proximaLinhaDestino + 1
        End If
    Next i
    
End Sub

' Procedimento de apoio renomeado para manter consistência
Private Sub OrdenarRegistrosPorProximidade()
    ' A lógica original deste sub (GerarDistanciasOrdenadas) deve residir aqui
End Sub

