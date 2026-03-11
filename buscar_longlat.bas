Attribute VB_Name = "buscar_longlat"
Sub CapturarMetadadosGeograficos()
    Dim clienteHTTP As Object
    Dim endpointURL As String
    Dim stringResposta As String
    Dim objetoJSON As Object
    Dim limiteLinhas As Long
    Dim i As Long
    Dim strLocalidade As String, strLogradouro As String
    Dim queryConsulta As String

    Set clienteHTTP = CreateObject("MSXML2.XMLHTTP")

    limiteLinhas = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To limiteLinhas
        strLocalidade = Cells(i, 1).Value
        strLogradouro = Cells(i, 2).Value
        queryConsulta = strLogradouro & ", " & strLocalidade

        queryConsulta = Replace(queryConsulta, " ", "+")
        endpointURL = "https://nominatim.openstreetmap.org/search?q=" & queryConsulta & "&format=json&limit=1"

        clienteHTTP.Open "GET", endpointURL, False
        clienteHTTP.setRequestHeader "User-Agent", "AppGestaoGeografica"
        clienteHTTP.Send

        stringResposta = clienteHTTP.responseText

        If InStr(stringResposta, "lat") > 0 Then
            Set objetoJSON = JsonConverter.ParseJson(stringResposta)
            If objetoJSON.Count > 0 Then
                Cells(i, 3).Value = objetoJSON(1)("lat")
                Cells(i, 4).Value = objetoJSON(1)("lon")
            Else
                Cells(i, 3).Value = "Registro nulo"
                Cells(i, 4).Value = "Registro nulo"
            End If
        Else
            Cells(i, 3).Value = "Falha_Ref"
            Cells(i, 4).Value = "Falha_Ref"
        End If

        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
    Next i

    MsgBox "Sincronização Geográfica Concluída!", vbInformation
End Sub
