Attribute VB_Name = "trazer_distancias"
Function AlgoritmoHaversine(latRef As Double, lonRef As Double, latDest As Double, lonDest As Double) As Double
    Const VALOR_PI As Double = 3.14159265358979
    Const RAIO_TERRA_KM As Double = 6371

    Dim deltaLat As Double, deltaLon As Double
    Dim varA As Double, varC As Double
    Dim radLatRef As Double, radLatDest As Double

    deltaLat = (latDest - latRef) * VALOR_PI / 180
    deltaLon = (lonDest - lonRef) * VALOR_PI / 180
    radLatRef = latRef * VALOR_PI / 180
    radLatDest = latDest * VALOR_PI / 180

    varA = Sin(deltaLat / 2) ^ 2 + Cos(radLatRef) * Cos(radLatDest) * Sin(deltaLon / 2) ^ 2
    varC = 2 * Atn(Sqr(varA) / Sqr(1 - varA))

    AlgoritmoHaversine = RAIO_TERRA_KM * varC
End Function

Function NormalizarCoordenada(entradaDados As Variant) As Double
    Dim textoProcessado As String
    textoProcessado = Trim(CStr(entradaDados))
    textoProcessado = Replace(textoProcessado, ".", "")
    textoProcessado = Replace(textoProcessado, ",", Application.DecimalSeparator)
    NormalizarCoordenada = CDbl(textoProcessado)
End Function

Sub ProcessarMatrizProximidade()
    Dim wsDados As Worksheet
    Set wsDados = ThisWorkbook.Sheets("Matriz")

    Dim limiteLinhas As Long, limiteColunas As Long
    limiteLinhas = wsDados.Cells(wsDados.Rows.Count, "B").End(xlUp).Row
    limiteColunas = wsDados.Cells(1, wsDados.Columns.Count).End(xlToLeft).Column

    Dim i As Long, j As Long, k As Long
    Dim latOrigem As Double, lonOrigem As Double
    Dim latAlvo As Double, lonAlvo As Double
    Dim calculoDistancia As Double
    Dim idPontoOrigem As String, idPontoAlvo As String

    For i = 2 To limiteLinhas
        idPontoOrigem = wsDados.Cells(i, "B").Value
        latOrigem = NormalizarCoordenada(wsDados.Cells(i, "E").Value)
        lonOrigem = NormalizarCoordenada(wsDados.Cells(i, "F").Value)

        For j = 7 To limiteColunas
            idPontoAlvo = wsDados.Cells(1, j).Value

            For k = 2 To limiteLinhas
                If wsDados.Cells(k, "B").Value = idPontoAlvo Then
                    latAlvo = NormalizarCoordenada(wsDados.Cells(k, "E").Value)
                    lonAlvo = NormalizarCoordenada(wsDados.Cells(k, "F").Value)
                    Exit For
                End If
            Next k

            calculoDistancia = AlgoritmoHaversine(latOrigem, lonOrigem, latAlvo, lonAlvo)
            wsDados.Cells(i, j).Value = Round(calculoDistancia, 2)
        Next j
    Next i
    
    MsgBox "Processamento da Matriz Geográfica Concluído!", vbInformation
End Sub

