Attribute VB_Name = "GetTrackingInfo"
' DHL API Connection Macro Version 1.0.1
' Created on 31/08/2023
' Updated on 04/09/2023
'
' API documentation available on: https://developer.dhl.com/api-reference/shipment-tracking#reference-docs-section/
'
' The addressable API base URL/URI environments are:
' https://api-eu.dhl.com/track/shipments
'
'----LIBRARIES---------------
' VBA-JSON-2.3.1
'----------------------------
'
'----REFERENCES--------------
' Microsoft XML, v6.0
' Microsoft Scripting Runtime
'----------------------------

Option Explicit

    ' Declaración de variables
    Dim strApiKey As String
    Dim variableName As String
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim objRequest As Object
    Dim json As Object


Function DhlApiRequest(ByVal trackingNumber As String) As String
    

    ' Crear un objeto XMLHTTP
    Set objRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configurar el API key y el URL del request
    variableName = "My-DHL-API-Key"
    strApiKey = Environ(variableName) 'Introducir la API Key guardada en variables de entorno
    strUrl = "https://api-eu.dhl.com/track/shipments" 'API Endpoint

    ' Construir la URL de la petición a la API con el tracking
    strUrl = strUrl & "?trackingNumber=" & trackingNumber
    
    ' Definir Asynchronus
    blnAsync = False
    
    With objRequest
        ' Realizar una petición (request) a la API de tipo GET
        .Open "GET", strUrl, blnAsync
        
        'Añadir el formato al header
        .setRequestHeader "Content-Type", "application/json"
        'Añadir el API Key al header
        .setRequestHeader "DHL-API-Key", strApiKey

        ' Enviar petición (request)
        .send
        

        ' Comprobar respuesta exitosa
        If .status = 200 Then
'            ' Parsear la respuesta que viene en formato JSON
'            Set json = JsonConverter.ParseJson(.responseText)

            ' Acceder a información específica en la respuesta JSON
'            Dim origin As String
'            Dim destination As String
'            Dim status As String
'            origin = json("shipments")(1)("origin")("address")("addressLocality")
'            destination = json("shipments")(1)("destination")("address")("addressLocality")
'            status = json("shipments")(1)("status")("description")


'            DhlApiRequest = "Tracking: " & trackingNumber & vbNewLine & "Origen: " & origin & vbNewLine & "Destino: " & destination & vbNewLine & "Estado: " & status
            DhlApiRequest = .responseText
        Else
            ' En caso de que no haya éxito en la respuesta
            DhlApiRequest = "Error: " & .status & " - " & .statusText

        End If
    End With

    ' Clean up the XMLHTTP object
    Set objRequest = Nothing

End Function

Sub GetTrackingInfo()
Attribute GetTrackingInfo.VB_Description = "y"
Attribute GetTrackingInfo.VB_ProcData.VB_Invoke_Func = "y\n14"
     
    ' Declaración de variables
    Dim trackingCell As Range
    Dim trackingNumber As String
    Dim result As String
    
    ' Introducir el tracking number del que se quiere realizar seguimiento
    Set trackingCell = Selection
    trackingNumber = trackingCell.Value
    
    If trackingNumber <> "" Then
        ' Obtener los datos del envío de la API
        result = DhlApiRequest(trackingNumber)
        ' Introducir los datos en Excel
        If result <> "Error: 404 - Not Found" Then ' Comprobar tracking erróneos
                
            If InputResultIntoSheet(result) Then
                Debug.Print "Datos introducidos"
            Else
                MsgBox ("El número de tracking no es un servicio express")
            End If
        
        Else
            MsgBox ("Lo sentimos, su intento de rastreo no se realizó correctamente. Compruebe su número de seguimiento.")
        End If
            
    Else
        Debug.Print "La celda está vacía"
    End If

End Sub

Function InputResultIntoSheet(ByVal result As String) As Boolean
    
' --- Extracción y ordenación de la información ---
    
    ' Declaración de variables
    Dim json As Object
    Dim service As String
    Dim origin As String
    Dim destination As String
    Dim status As String
    Dim deliveryDay As String
    Dim deliveryHour As String
    ' Parsear la respuesta que viene en formato JSON
    Set json = JsonConverter.ParseJson(result)

    ' Acceder a información específica en la respuesta JSON
    
    service = json("shipments")(1)("service")
    If service = "express" Then
    
        origin = json("shipments")(1)("origin")("address")("addressLocality")
        destination = json("shipments")(1)("destination")("address")("addressLocality")
        status = json("shipments")(1)("status")("status")
        deliveryDay = Left(json("shipments")(1)("status")("timestamp"), 10)
        deliveryHour = Mid(json("shipments")(1)("status")("timestamp"), 12)
        
    End If
        
    
' --- Introducción de la información en la hoja de Excel ---

    ' Declaración de variables
    Dim trackingCell As Range
    Dim statusCell As Range
    Dim delDayCell As Range
    Dim delTimeCell As Range
    
    ' Definición de las celdas
    Set trackingCell = Selection
    Set statusCell = ActiveCell.Offset(0, 2)
    Set delDayCell = ActiveCell.Offset(0, 3)
    Set delTimeCell = ActiveCell.Offset(0, 4)
    
    ' Introducción de datos
    If service = "express" Then ' Solo intrudocir datos en Excel para pedidos con servicio express
    
        If status = "delivered" Then
            statusCell.Value = "Entregado"
            delDayCell.Value = deliveryDay
            delTimeCell.Value = deliveryHour
        ElseIf status = "on hold" Then
            statusCell.Value = "Retraso"
        Else
            statusCell.Value = "En tránsito"
        End If
        
        InputResultIntoSheet = True
        
    Else
        
        InputResultIntoSheet = False
    
    End If
    

End Function
