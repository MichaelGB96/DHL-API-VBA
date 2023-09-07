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
'
' Author: Miguel G�mez
' email: mikegomezb96@gmail.com
' github: MichaelGB96

Option Explicit

    ' Declaraci�n de variables
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

    ' Construir la URL de la petici�n a la API con el tracking
    strUrl = strUrl & "?trackingNumber=" & trackingNumber
    
    ' Definir Asynchronus
    blnAsync = False
    
    With objRequest
        ' Realizar una petici�n (request) a la API de tipo GET
        .Open "GET", strUrl, blnAsync
        
        'A�adir el formato al header
        .setRequestHeader "Content-Type", "application/json"
        'A�adir el API Key al header
        .setRequestHeader "DHL-API-Key", strApiKey

        ' Enviar petici�n (request)
        .send
        

        ' Comprobar respuesta exitosa
        If .status = 200 Then
'            ' Parsear la respuesta que viene en formato JSON
'            Set json = JsonConverter.ParseJson(.responseText)

            ' Acceder a informaci�n espec�fica en la respuesta JSON
'            Dim origin As String
'            Dim destination As String
'            Dim status As String
'            origin = json("shipments")(1)("origin")("address")("addressLocality")
'            destination = json("shipments")(1)("destination")("address")("addressLocality")
'            status = json("shipments")(1)("status")("description")


'            DhlApiRequest = "Tracking: " & trackingNumber & vbNewLine & "Origen: " & origin & vbNewLine & "Destino: " & destination & vbNewLine & "Estado: " & status
            DhlApiRequest = .responseText
        Else
            ' En caso de que no haya �xito en la respuesta
            DhlApiRequest = "Error: " & .status & " - " & .statusText

        End If
    End With

    ' Clean up the XMLHTTP object
    Set objRequest = Nothing

End Function

Sub GetTrackingInfo()
Attribute GetTrackingInfo.VB_Description = "y"
Attribute GetTrackingInfo.VB_ProcData.VB_Invoke_Func = "y\n14"
     
    ' Declaraci�n de variables
    Dim trackingCell As Range
    Dim trackingNumber As String
    Dim result As String
    Dim msgbxTitle As String
    Dim msgbxResult As String
    
    ' Introducir el tracking number del que se quiere realizar seguimiento
    Set trackingCell = Selection
    trackingNumber = trackingCell.Value
    
    If Len(trackingNumber) = 10 Then
    
        msgbxTitle = "DHL Tracking"
        ' Obtener los datos del env�o de la API
        result = DhlApiRequest(trackingNumber)
        ' Introducir los datos en Excel
        If result <> "Error: 404 - Not Found" Then ' Comprobar tracking err�neos
                
            If InputResultIntoSheet(result) Then
                Debug.Print "Datos introducidos"
            Else
                msgbxResult = MsgBox("El n�mero de tracking no es un servicio express", 0, msgbxTitle)
            End If
        
        Else
            msgbxResult = MsgBox("Lo sentimos, su intento de rastreo no se realiz� correctamente. Compruebe su n�mero de seguimiento.", 0, msgbxTitle)
        End If
            
    Else
        msgbxTitle = "Tracking"
        msgbxResult = MsgBox("N�mero de seguimiento incorrecto.", 0, msgbxTitle)
    End If

End Sub

Function InputResultIntoSheet(ByVal result As String) As Boolean
    
' --- Extracci�n y ordenaci�n de la informaci�n ---
    
    ' Declaraci�n de variables
    Dim json As Object
    Dim service As String
    Dim origin As String
    Dim destination As String
    Dim status As String
    Dim deliveryDay As String
    Dim deliveryHour As String
    Dim deliveryETA As String
    ' Parsear la respuesta que viene en formato JSON
    Set json = JsonConverter.ParseJson(result)

    ' Acceder a informaci�n espec�fica en la respuesta JSON
    
    service = json("shipments")(1)("service")
    If service = "express" Then
    
        origin = json("shipments")(1)("origin")("address")("addressLocality")
        destination = json("shipments")(1)("destination")("address")("addressLocality")
        status = json("shipments")(1)("status")("status")
        deliveryDay = Left(json("shipments")(1)("status")("timestamp"), 10)
        deliveryHour = Mid(json("shipments")(1)("status")("timestamp"), 12)
        
        If status = "transit" Then ' Si est� en tr�nsito, informar del ETA
            
            deliveryETA = "ETA " & json("shipments")(1)("estimatedTimeOfDelivery") & " " & json("shipments")(1)("estimatedTimeOfDeliveryRemark")
        
        End If
        
    End If
        
    
' --- Introducci�n de la informaci�n en la hoja de Excel ---

    ' Declaraci�n de variables
    Dim trackingCell As Range
    Dim statusCell As Range
    Dim delDayCell As Range
    Dim delTimeCell As Range
    Dim commCell As Range
    
    ' Definici�n de las celdas
    Set trackingCell = Selection
    Set statusCell = ActiveCell.Offset(0, 2)
    Set delDayCell = ActiveCell.Offset(0, 3)
    Set delTimeCell = ActiveCell.Offset(0, 4)
    Set commCell = ActiveCell.Offset(0, 6)
    
    ' Introducci�n de datos
    If service = "express" Then ' Solo intrudocir datos en Excel para pedidos con servicio express
    
        If status = "delivered" Then
            statusCell.Value = "Entregado"
            delDayCell.Value = deliveryDay
            delTimeCell.Value = deliveryHour
        ElseIf status = "transit" Then
            statusCell.Value = "En tr�nsito"
            commCell.Value = deliveryETA
        Else
            statusCell.Value = "Retraso"
        End If
        
        InputResultIntoSheet = True
        
    Else
        
        InputResultIntoSheet = False
    
    End If
    

End Function
