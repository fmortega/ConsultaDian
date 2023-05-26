# ConsultaDian


    Public Sub consultaMasiva()

        Dim rngRUT As Range
       
        Range("c2:i500").ClearContents
       
        For Each rngRUT In Range("b2:b500")
       
            rngRUT.Offset(0, 1).Resize(1, 7) = ConsultaDIAN(rngRUT.Value)
       
        Next rngRUT

    End Sub


    Public Function ConsultaDIAN(strRut As String) As Variant

        ' se va a producir un error hasta que no esté disponible el elemento
        ' vamos a esperar y probamos en un rato a ver si ya está disponible
        
        On Error GoTo esperar
        
        Dim Nombres, consulta, estado, pp As Object
        Dim IE As Object
        Dim arrRespuesta(6) As String
       
        ' Sólo hacemos la consulta del RUT si tiene 9 dígitos
        If Len(strRut) < 2 Or Not IsNumeric(strRut) Then
            ConsultaDIAN = Array("RUT Inválido", "RUT Inválido")
            Exit Function
        End If
       
        Application.StatusBar = "Consultando RUT " & strRut & "."
        
        Set IE = CreateObject("InternetExplorer.Application")
       
        IE.Navigate "https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces"
       
        IE.Document.all.Item("vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit").Value = strRut
       
        IE.Document.getElementbyId("vistaConsultaEstadoRUT:formConsultaEstadoRUT:btnBuscar").Click
        

        
        arrRespuesta(0) = IE.Document.getElementbyId("vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv").innertext
        
        arrRespuesta(1) = IE.Document.getElementbyId("vistaConsultaEstadoRUT:formConsultaEstadoRUT:estado").innertext
        
        
        If IE.Document.getElementsByClassName("tipoFilaNormalVerde")(4).innertext = "REGISTRO ACTIVO" Or IE.Document.getElementsByClassName("tipoFilaNormalVerde")(4).innertext = "REGISTRO CANCELADO" Then
            arrRespuesta(2) = "N/A"
            arrRespuesta(3) = "N/A"
            arrRespuesta(4) = "N/A"
            arrRespuesta(5) = "N/A"
            arrRespuesta(6) = IE.Document.getElementsByClassName("tipoFilaNormalVerde")(2).innertext
        
        Else
            
            arrRespuesta(2) = IE.Document.getElementsByClassName("tipoFilaNormalVerde")(2).innertext
            arrRespuesta(3) = IE.Document.getElementsByClassName("tipoFilaNormalVerde")(3).innertext
            arrRespuesta(4) = IE.Document.getElementsByClassName("tipoFilaNormalVerde")(4).innertext
            arrRespuesta(5) = IE.Document.getElementsByClassName("tipoFilaNormalVerde")(5).innertext
            arrRespuesta(6) = "N/A"
        
        End If
        
        ConsultaDIAN = arrRespuesta
        
salir:

        Application.StatusBar = False
       
        IE.Quit

        Set IE = Nothing
       
        Exit Function
       
esperar:

        Select Case Err.Number
       
            Case 91, 424, -2147467259
       
                Application.StatusBar = Application.StatusBar & "."
               
                Application.Wait Now + TimeValue("0:00:00")
               
                Resume
               
            Case Else
       
                MsgBox "Error " & Err.Number & ": " & Err.Description
           
                Resume salir
           
        End Select
       
    End Function

    Private Sub Class_Initialize()
        MsgBox "Class_Initialize"
    End Sub
