Sub CotizarAutoMercadoLibre()

    Dim modelo As String
    Dim accessToken As String
    Dim url As String
    Dim http As Object
    Dim json As Object
    Dim result As Object
    Dim i As Integer

    ' Modelo de auto en A2
    modelo = Trim(Cells(2, 1).Value)
    
    ' Token de acceso de tu app
    accessToken = "APP_USR-5183773064841897-072412-2e137b60093173bd4810074534d37684-242507349"
    
    If modelo = "" Then
        MsgBox "Ingresá un modelo en la celda A2"
        Exit Sub
    End If

    ' Crear conexión HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Construir URL
    url = "https://api.mercadolibre.com/sites/MLA/search?q=" & Replace(modelo, " ", "%20") & "%200km&category=MLA1743&sort=price_asc&limit=3"
    
    ' Hacer GET
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Bearer " & accessToken
    http.Send
    
    ' Validar respuesta
    If http.Status <> 200 Then
        MsgBox "Error al obtener datos: " & http.Status
        Exit Sub
    End If
    
    ' Parsear JSON
    Set json = JsonConverter.ParseJson(http.ResponseText)

    ' Limpiar celdas anteriores
    Range("B2:D4").ClearContents

    ' Mostrar resultados
    i = 2
    For Each result In json("results")
        Cells(i, 2).Value = result("title")
        Cells(i, 3).Value = result("price")
        Cells(i, 4).Value = result("permalink")
        i = i + 1
        If i > 4 Then Exit For
    Next

    MsgBox "Cotizaciones cargadas correctamente."

End Sub
