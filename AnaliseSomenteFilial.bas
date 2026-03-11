Attribute VB_Name = "AnaliseSomenteFilial"
Sub OrdenarRegistrosPorProximidade()
    Dim wsFonte As Worksheet, wsProcessamento As Worksheet
    Dim ultimaLinhaFonte As Long, i As Long, j As Long, k As Long
    Dim coordLatRef As Double, coordLonRef As Double
    Dim coordLatDest As Double, coordLonDest As Double
    Dim matrizDados As Variant
    Dim matrizResultados() As Variant
    Dim idReferencia As Long
    Dim contadorIdx As Long

    ' Definição das planilhas (Nomes anonimizados)
    Set wsFonte = Worksheets("Matriz")
    Set wsProcessamento = Worksheets("Distancias")

    ' Captura do ID do ponto de referência selecionado no painel
    idReferencia = wsProcessamento.Range("E2").Value

    ' Estrutura esperada: Setor (A), ID_Ponto (B), Municipio (C), Localizacao (D), Lat (E), Long (F)
    ultimaLinhaFonte = wsFonte.Cells(wsFonte.Rows.Count, "B").End(xlUp).Row
    matrizDados = wsFonte.Range("A2:F" & ultimaLinhaFonte).Value

    ' Localiza coordenadas do ponto de referência selecionado
    For i = 1 To UBound(matrizDados)
        If matrizDados(i, 2) = idReferencia Then
            coordLatRef = NormalizarCoordenada(matrizDados(i, 5))
            coordLonRef = NormalizarCoordenada(matrizDados(i, 6))
            Exit For
        End If
    Next i

    ' Dimensiona matriz para os resultados (excluindo o ponto de referência)
    ReDim matrizResultados(1 To UBound(matrizDados) - 1, 1 To 7)
    contadorIdx = 0

    ' Cálculo de métricas para os demais pontos da rede
    For i = 1 To UBound(matrizDados)
        If matrizDados(i, 2) <> idReferencia Then
            coordLatDest = NormalizarCoordenada(matrizDados(i, 5))
            coordLonDest = NormalizarCoordenada(matrizDados(i, 6))
            contadorIdx = contadorIdx + 1

            matrizResultados(contadorIdx, 1) = matrizDados(i, 1) ' Setor/Região
            matrizResultados(contadorIdx, 2) = matrizDados(i, 2) ' ID_Ponto
            matrizResultados(contadorIdx, 3) = matrizDados(i, 3) ' Municipio
            matrizResultados(contadorIdx, 4) = matrizDados(i, 4) ' Localizacao
            matrizResultados(contadorIdx, 5) = matrizDados(i, 5) ' Latitude
            matrizResultados(contadorIdx, 6) = matrizDados(i, 6) ' Longitude
            ' Cálculo de distância via fórmula esférica
            matrizResultados(contadorIdx, 7) = Round(AlgoritmoHaversine(coordLatRef, coordLonRef, coordLatDest, coordLonDest), 2)
        End If
    Next i

    ' Ordenação por proximidade (Bubble Sort)
    Dim temporario(1 To 7) As Variant
    For i = 1 To contadorIdx - 1
        For j = i + 1 To contadorIdx
            If matrizResultados(j, 7) < matrizResultados(i, 7) Then
                For k = 1 To 7
                    temporario(k) = matrizResultados(i, k)
                    matrizResultados(i, k) = matrizResultados(j, k)
                    matrizResultados(j, k) = temporario(k)
                Next k
            End If
        Next j
    Next i

    ' Saída dos dados estruturados
    With wsProcessamento
        .Range("B6:H1000").ClearContents
        .Range("B4").Resize(1, 7).Value = Array("Setor", "ID_Ponto", "Municipio", "Localizacao", "Coord_Lat", "Coord_Long", "Distancia_KM")
        
        ' Preenchimento da grade de análise
        For i = 1 To contadorIdx
            For j = 1 To 7
                .Cells(i + 5, j + 1).Value = matrizResultados(i, j)
            Next j
        Next i
    End With
    
    MsgBox "Processamento Concluído com Sucesso!"
End Sub
