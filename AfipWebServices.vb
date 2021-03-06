Option Strict Off

Namespace CardonerSistemas
    Module AfipWebServices

#Region "Declarations"
        Friend Const ServicioFacturacionElectronica As String = "wsfe"

        Friend Const SolicitudCaeResultadoAceptado As String = "A"
        Friend Const SolicitudCaeResultadoRechazado As String = "R"
        Friend Const SolicitudCaeResultadoParcial As String = "P"

        Friend Class ComprobanteAsociado
            Friend Property TipoComprobante As Short
            Friend Property PuntoVenta As Short
            Friend Property ComprobanteNumero As Integer
        End Class

        Friend Class Tributo
            Friend Property ID As Short
            Friend Property Descripcion As String
            Friend Property BaseImponible As Decimal
            Friend Property Alicuota As Decimal
            Friend Property Importe As Decimal
        End Class

        Friend Class IVA
            Friend Property ID As Short
            Friend Property BaseImponible As Decimal
            Friend Property Importe As Decimal
        End Class

        Friend Class Opcional
            Friend Property ID As String
            Friend Property Valor As String
        End Class

        Friend Class FacturaElectronicaCabecera
            Friend Property Concepto As Short
            Friend Property TipoDocumento As Short
            Friend Property DocumentoNumero As Long
            Friend Property TipoComprobante As Short
            Friend Property PuntoVenta As Short
            Friend Property ComprobanteDesde As Integer
            Friend Property ComprobanteHasta As Integer
            Friend Property ComprobanteFecha As Date
            Friend Property ImporteTotal As Decimal
            Friend Property ImporteTotalConc As Decimal            ' Importe neto no gravado - Para comprobantes "C", debe ser cero.
            Friend Property ImporteNeto As Decimal                 ' Importe neto gravado - Para comprobantes "C", debe ser igual al Subtotal.
            Friend Property ImporteOperacionesExentas As Decimal   ' Para comprobantes "C", debe ser cero.
            Friend Property ImporteTributos As Decimal
            Friend Property ImporteIVA As Decimal                  ' Para comprobantes "C", debe ser cero.
            Friend Property FechaServicioDesde As Date
            Friend Property FechaServicioHasta As Date
            Friend Property FechaVencimientoPago As Date
            Friend Property MonedaID As String
            Friend Property MonedaCotizacion As Decimal            ' Para pesos argentinos, debe ser 1.

            Friend Property ComprobantesAsociados As List(Of ComprobanteAsociado)
            Friend Property Tributos As List(Of Tributo)
            Friend Property IVAs As List(Of IVA)
            Friend Property Opcionales As List(Of Opcional)

            Friend Sub New()
                ComprobantesAsociados = New List(Of ComprobanteAsociado)
                Tributos = New List(Of Tributo)
                IVAs = New List(Of IVA)()
                Opcionales = New List(Of Opcional)
            End Sub
        End Class

        Friend Class ResultadoCAE
            Friend Property Resultado As Char
            Friend Property Numero As String
            Friend Property FechaVencimiento As Date
            Friend Property Observaciones As String
            Friend Property ErrorMessage As String
        End Class

        Friend Class ResultadoConsultaComprobante
            Friend Property Concepto As Short
            Friend Property TipoDocumento As Short
            Friend Property DocumentoNumero As Long
            Friend Property TipoComprobante As Short
            Friend Property PuntoVenta As Short
            Friend Property ComprobanteDesde As Integer
            Friend Property ComprobanteHasta As Integer
            Friend Property ComprobanteFecha As Date
            Friend Property ImporteTotal As Decimal
            Friend Property ImporteTotalConc As Decimal            ' Importe neto no gravado - Para comprobantes "C", debe ser cero.
            Friend Property ImporteNeto As Decimal                 ' Importe neto gravado - Para comprobantes "C", debe ser igual al Subtotal.
            Friend Property ImporteTributos As Decimal
            Friend Property ImporteIVA As Decimal                  ' Para comprobantes "C", debe ser cero.
            Friend Property FechaServicioDesde As Date
            Friend Property FechaServicioHasta As Date
            Friend Property FechaVencimientoPago As Date
            Friend Property MonedaID As String
            Friend Property MonedaCotizacion As Decimal            ' Para pesos argentinos, debe ser 1.
            Friend Property Resultado As Char
            Friend Property CodigoAutorizacion As String
            Friend Property EmisionTipo As String
            Friend Property FechaVencimiento As Date
            Friend Property FechaHoraProceso As Date
            Friend Property Observaciones As String
            Friend Property ErrorMessage As String
        End Class

#End Region

#Region "Clase Principal"

        Friend Class WebService
            Friend Property LogPath As String = ""
            Friend Property LogFileName As String = ""
            Friend Property Certificado As String
            Friend Property ClavePrivada As String
            Friend Property WSAA_URL As String
            Friend Property WSFEv1_URL As String
            Friend Property InternetProxy As String
            Friend Property CUIT_Emisor As String
            Friend Property MonedaLocal As Moneda
            Friend Property MonedaLocalCotizacion As MonedaCotizacion

            ' Propiedades de Resultado
            Friend Property TicketAcceso As String
            Friend Property WSFEv1 As Object
            Friend Property UltimoResultadoCAE As ResultadoCAE
            Friend Property UltimoComprobanteAutorizado As String
            Friend Property UltimoResultadoConsultaComprobante As ResultadoConsultaComprobante

            Friend Sub New()
                UltimoResultadoCAE = New ResultadoCAE
                UltimoResultadoConsultaComprobante = New ResultadoConsultaComprobante
            End Sub

            Friend Function Login(ByVal ServicioNombre As String) As Boolean
                Dim WSAA As Object
                Dim TicketRequerimientoAcceso As String
                Dim MensajeFirmado As String

                ' Crear objeto interface Web Service Autenticación y Autorización
                Try
                    WSAA = CreateObject("WSAA")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, StrDup(20, "="))
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se ha creado el objeto WSAA.")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Versión de WSAA: " & WSAA.Version)

                Catch ex As Exception
                    If ex.HResult = CardonerSistemas.ErrorHandler.ActiveXCreateError Then
                        MsgBox("La librería PyAfipWs para el Servicio WSAA (wsaa.exe), no está instalada correctamente." & vbCrLf & "Contáctese con el personal de Soporte Técnico.", MsgBoxStyle.Critical, My.Application.Info.Title)
                    Else
                        CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al crear el objeto WSAA de la librería PyAfipWs.")
                    End If
                    Return False
                End Try

                Try
                    If WSAA.Version < "2.07" Then
                        ' Es la versión antigua de WSAA, utilizo los métodos correspondientes

                        ' Generar un Ticket de Requerimiento de Acceso (TRA)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se generará un Ticket de Requerimiento de Acceso utilizando el método 'CreateTRA'.")
                        TicketRequerimientoAcceso = WSAA.CreateTRA(ServicioNombre, pAfipWebServicesConfig.TtlTicketRequerimientoAcceso)
                        If TicketRequerimientoAcceso <> "" Then
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Ticket de Requerimiento de Acceso: " & TicketRequerimientoAcceso)

                            ' Generar el mensaje firmado (CMS)
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se generará el Mensaje Firmado utilizando el método 'SignTRA'.")
                            MensajeFirmado = WSAA.SignTRA(TicketRequerimientoAcceso, Certificado, ClavePrivada)
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Mensaje Firmado: " & MensajeFirmado)

                            If MensajeFirmado <> "" Then
                                ' Conectar al webservice
                                CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se ejecutará el método 'Conectar'.")
                                WSAA.Conectar("", WSAA_URL, InternetProxy)
                                CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se conectó al WebService.")

                                CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se autenticará utilizando el método 'LoginCMS'.")
                                TicketAcceso = WSAA.LoginCMS(MensajeFirmado)
                                If TicketAcceso <> "" Then
                                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se autenticó correctamente.")
                                    Return True
                                Else
                                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Falló la autenticación.")
                                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "XmlRequest:  " & WSAA.XmlRequest)
                                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "XmlResponse: " & WSAA.XmlResponse)
                                    Return False
                                End If
                            Else
                                Return False
                            End If
                        Else
                            Return False
                        End If
                    Else
                        ' Es la versión nueva de WSAA, utilizo el método Autenticar
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se autenticará utilizando el método 'Autenticar'.")
                        TicketAcceso = WSAA.Autenticar(ServicioNombre, Certificado, ClavePrivada, WSAA_URL)
                        If TicketAcceso <> "" Then
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se autenticó correctamente.")
                            Return True
                        Else
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Falló la autenticación.")
                            WSAA.AnalizarXml("XmlResponse")
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Codigo de Fallo: " & WSAA.ObtenerTagXml("faultcode"))
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Mensaje de Fallo: " & WSAA.ObtenerTagXml("faultstring"))
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Certificado: " & Certificado)
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Clave Privada: " & ClavePrivada)
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "URL: " & WSAA_URL)
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "XmlRequest:  " & WSAA.XmlRequest)
                            CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "XmlResponse: " & WSAA.XmlResponse)
                            CardonerSistemas.ErrorHandler.ProcessError(New Exception("Excepción: " & WSAA.Excepcion), "Error al iniciar sesión en el Servidor de AFIP.")
                            Return False
                        End If
                    End If

                Catch
                    ' Muestro los errores
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Ocurrió un error.")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Excepción: " & WSAA.Excepcion)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Traza: " & WSAA.Traceback)
                    CardonerSistemas.ErrorHandler.ProcessError(New Exception("Excepción: " & WSAA.Excepcion), "Error al iniciar sesión en el Servidor de AFIP.")
                    Return False
                End Try

                'En versiones recientes (2.04a o superior), si no hubo excepción es posible revisar y obtener datos avanzados del ticket de acceso (útiles para depuración y solución de errores):

                '        Origen(Source) : WSAA.ObtenerTagXml("source")
                '        Destino(Destination) : WSAA.ObtenerTagXml("destination")
                '        ID Único : WSAA.ObtenerTagXml("uniqueId")
                'Fecha de Generación: WSAA.ObtenerTagXml("generationTime")
                'Fecha de Expiración: WSAA.ObtenerTagXml("expirationTime")
                'Si ha ocurrido error (llamar previamente a WSAA.AnalizarXml("XmlResponse") para analizar la respuesta):
            End Function

            Friend Function FacturaElectronica_Login() As Boolean
                Return Login(ServicioFacturacionElectronica)
            End Function

            Friend Function FacturaElectronica_Conectar() As Boolean

                ' Crear objeto interface Web Service Autenticación y Autorización
                Try
                    WSFEv1 = CreateObject("WSFEv1")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, StrDup(20, "="))
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se ha creado el objeto WSFEv1.")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Versión de WSFEv1: " & WSFEv1.Version)

                Catch ex As Exception
                    If ex.HResult = -2146233088 Then
                        MsgBox("La librería PyAfipWs para el Servicio WSFEv1 (wsfev1.exe), no está instalada correctamente." & vbCrLf & "Contáctese con el personal de Soporte Técnico.", MsgBoxStyle.Critical, My.Application.Info.Title)
                    Else
                        CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al crear el objeto WSFEv1 de la librería PyAfipWs.")
                    End If
                    WSFEv1 = Nothing
                    Return False
                End Try

                Try
                    ' Establecer Ticket de Acceso en un solo paso
                    WSFEv1.SetTicketAcceso(TicketAcceso.ToString)

                    ' CUIT del emisor (debe estar registrado en la AFIP)
                    WSFEv1.Cuit = CUIT_Emisor

                    ' Conectar al Servicio Web de Facturación. Proxy: usuario:clave@localhost:8000
                    WSFEv1.Conectar("", WSFEv1_URL.ToString, InternetProxy.ToString)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se ejecutó el método 'Conectar'")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "DebugLog: " & WSFEv1.DebugLog.Trim())

                    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
                    WSFEv1.Dummy()
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se ejecutó el método 'Dummy'")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "AppServer status:  " & WSFEv1.AppServerStatus)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "DbServer status:   " & WSFEv1.DbServerStatus)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "AuthServer status: " & WSFEv1.AuthServerStatus)

                    ' Fin
                    Return True

                Catch ex As Exception
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Ha ocurrido un error:")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "    Excepción: " & WSFEv1.Excepcion)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "    Traza:     " & WSFEv1.Traceback)
                    CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al conectar con el Servicio de Factura Electrónica.")
                    WSFEv1 = Nothing
                    Return False
                End Try
            End Function

            Friend Function FacturaElectronica_ObtenerCAE(ByRef FacturaAGenerar As FacturaElectronicaCabecera) As Boolean
                Dim CAE As String

                If WSFEv1 Is Nothing Then
                    Return False
                End If

                Try

                    With FacturaAGenerar
                        ' Creo la factura
                        WSFEv1.CrearFactura(.Concepto, .TipoDocumento, .DocumentoNumero, .TipoComprobante, .PuntoVenta, .ComprobanteDesde, .ComprobanteHasta,
                                        CS_ValueTranslation.FromDecimalToUString(.ImporteTotal), CS_ValueTranslation.FromDecimalToUString(.ImporteTotalConc), CS_ValueTranslation.FromDecimalToUString(.ImporteNeto), CS_ValueTranslation.FromDecimalToUString(.ImporteIVA), CS_ValueTranslation.FromDecimalToUString(.ImporteTributos), CS_ValueTranslation.FromDecimalToUString(.ImporteOperacionesExentas),
                                        CS_ValueTranslation.FromDateToUString(.ComprobanteFecha), CS_ValueTranslation.FromDateToUString(.FechaVencimientoPago), CS_ValueTranslation.FromDateToUString(.FechaServicioDesde), CS_ValueTranslation.FromDateToUString(.FechaServicioHasta), .MonedaID, CS_ValueTranslation.FromDecimalToUString(.MonedaCotizacion))

                        ' Agrego los comprobantes asociados:
                        For Each ComprobanteAsociadoActual In .ComprobantesAsociados
                            WSFEv1.AgregarCmpAsoc(ComprobanteAsociadoActual.TipoComprobante, ComprobanteAsociadoActual.PuntoVenta, ComprobanteAsociadoActual.ComprobanteNumero)
                        Next

                        ' Agrego los tributos (impuestos varios)
                        For Each TributoActual In .Tributos
                            WSFEv1.AgregarTributo(TributoActual.ID, TributoActual.Descripcion, CS_ValueTranslation.FromDecimalToUString(TributoActual.BaseImponible), CS_ValueTranslation.FromDecimalToUString(TributoActual.Alicuota), CS_ValueTranslation.FromDecimalToUString(TributoActual.Importe))
                        Next

                        ' Agrego tasas de IVA
                        For Each IVAActual In .IVAs
                            WSFEv1.AgregarIva(IVAActual.ID, CS_ValueTranslation.FromDecimalToUString(IVAActual.BaseImponible), CS_ValueTranslation.FromDecimalToUString(IVAActual.Importe))
                        Next
                    End With

                    ' Habilito reprocesamiento automático (predeterminado):
                    WSFEv1.Reprocesar = True

                    ' Solicito CAE:
                    CAE = WSFEv1.CAESolicitar.ToString
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se ejecutó el método 'CAESolicitar'")

                    UltimoResultadoCAE.Resultado = WSFEv1.Resultado
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Resultado: " & UltimoResultadoCAE.Resultado)
                    If WSFEv1.Resultado = SolicitudCaeResultadoAceptado Then
                        UltimoResultadoCAE.Numero = CAE
                        UltimoResultadoCAE.FechaVencimiento = Date.ParseExact(WSFEv1.Vencimiento, "yyyyMMdd", Nothing)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Tipo:  " & FacturaAGenerar.TipoComprobante)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Nro.:  " & FacturaAGenerar.PuntoVenta & "-" & FacturaAGenerar.ComprobanteDesde)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "CAE:               " & CAE)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Fecha Vencimiento: " & UltimoResultadoCAE.FechaVencimiento.ToShortDateString)
                    Else
                        'TODO: Arreglar el encoding de las Observaciones y el Mensaje de Error
                        UltimoResultadoCAE.Observaciones = WSFEv1.Obs
                        UltimoResultadoCAE.ErrorMessage = WSFEv1.ErrMsg

                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Tipo:  " & FacturaAGenerar.TipoComprobante)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Nro.:  " & FacturaAGenerar.PuntoVenta & "-" & FacturaAGenerar.ComprobanteDesde)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Observaciones: " & UltimoResultadoCAE.Observaciones)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Error:         " & UltimoResultadoCAE.ErrorMessage)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "XmlRequest:" & vbCrLf & WSFEv1.XmlRequest)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "XmlResponse:" & vbCrLf & WSFEv1.XmlResponse)
                    End If

                    Return True

                Catch ex As Exception
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Ocurrió un error.")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Excepción: " & WSFEv1.Excepcion)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Traza:     " & WSFEv1.Traceback)
                    CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al crear la Factura Electrónica.")
                    Return False
                End Try
            End Function

            Friend Function FacturaElectronica_ConectarYObtenerCAE(ByRef FacturaAGenerar As FacturaElectronicaCabecera) As Boolean
                If FacturaElectronica_Conectar() Then
                    Return FacturaElectronica_ObtenerCAE(FacturaAGenerar)
                Else
                    Return False
                End If
            End Function

            Friend Function FacturaElectronica_ObtenerUltimoNumeroComprobante(ByVal TipoComprobante As Short, ByVal PuntoVenta As Short) As Boolean
                If WSFEv1 Is Nothing Then
                    Return False
                End If

                Try
                    UltimoComprobanteAutorizado = WSFEv1.CompUltimoAutorizado(TipoComprobante, PuntoVenta)
                    Return True

                Catch ex As Exception
                    Return False
                End Try
            End Function

            Friend Function FacturaElectronica_ConectarYObtenerUltimoNumeroComprobante(ByVal TipoComprobante As Short, ByVal PuntoVenta As Short) As Boolean
                If FacturaElectronica_Conectar() Then
                    Return FacturaElectronica_ObtenerUltimoNumeroComprobante(TipoComprobante, PuntoVenta)
                Else
                    Return False
                End If
            End Function

            Friend Function FacturaElectronica_ConsultarComprobante(ByVal TipoComprobante As Short, ByVal PuntoVenta As Short, ByVal ComprobanteNumero As Integer) As Boolean
                If WSFEv1 Is Nothing Then
                    Return False
                End If

                Try
                    Dim CAE As String
                    CAE = WSFEv1.CompConsultar(TipoComprobante, PuntoVenta, ComprobanteNumero)

                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Tipo:  " & TipoComprobante)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Nro.:  " & PuntoVenta & "-" & ComprobanteNumero)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Observaciones: " & WSFEv1.Obs)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Error:         " & WSFEv1.ErrMsg)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "XmlRequest:" & vbCrLf & WSFEv1.XmlRequest)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "XmlResponse:" & vbCrLf & WSFEv1.XmlResponse)

                    With UltimoResultadoConsultaComprobante
                        .Resultado = WSFEv1.Resultado
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Resultado: " & .Resultado)
                        .CodigoAutorizacion = CAE
                        If WSFEv1.Resultado = SolicitudCaeResultadoAceptado Then
                            .Concepto = WSFEv1.ObtenerCampoFactura("concepto")
                            .TipoDocumento = WSFEv1.ObtenerCampoFactura("tipo_doc")
                            .DocumentoNumero = WSFEv1.ObtenerCampoFactura("nro_doc")
                            .TipoComprobante = WSFEv1.ObtenerCampoFactura("tipo_cbte")
                            .PuntoVenta = WSFEv1.ObtenerCampoFactura("punto_vta")
                            .ComprobanteDesde = WSFEv1.ObtenerCampoFactura("cbt_desde")
                            .ComprobanteHasta = WSFEv1.ObtenerCampoFactura("cbt_hasta")
                            .ComprobanteFecha = Date.ParseExact(WSFEv1.ObtenerCampoFactura("fecha_cbte"), "yyyyMMdd", Nothing)
                            .ImporteTotal = WSFEv1.ObtenerCampoFactura("imp_total")
                            .ImporteTotalConc = WSFEv1.ObtenerCampoFactura("imp_tot_conc")
                            .ImporteNeto = WSFEv1.ObtenerCampoFactura("imp_neto")
                            .ImporteTributos = WSFEv1.ObtenerCampoFactura("imp_trib")
                            .ImporteIVA = WSFEv1.ObtenerCampoFactura("imp_iva")
                            .FechaServicioDesde = Date.ParseExact(WSFEv1.ObtenerCampoFactura("fecha_serv_desde"), "yyyyMMdd", Nothing)
                            .FechaServicioHasta = Date.ParseExact(WSFEv1.ObtenerCampoFactura("fecha_serv_hasta"), "yyyyMMdd", Nothing)
                            .FechaVencimientoPago = Date.ParseExact(WSFEv1.ObtenerCampoFactura("fecha_venc_pago"), "yyyyMMdd", Nothing)
                            .MonedaID = WSFEv1.ObtenerCampoFactura("moneda_id")
                            .MonedaCotizacion = WSFEv1.ObtenerCampoFactura("moneda_ctz")
                            '.EmisionTipo = WSFEv1.ObtenerCampoFactura("emision_tipo")
                            .FechaVencimiento = Date.ParseExact(WSFEv1.Vencimiento, "yyyyMMdd", Nothing)
                            '.FechaHoraProceso = Date.ParseExact(WSFEv1.ObtenerCampoFactura("FchProceso"), "yyyyMMddhhnnss", Nothing)
                        End If
                        .Observaciones = WSFEv1.Obs
                        .ErrorMessage = WSFEv1.ErrMsg
                    End With

                Catch ex As Exception
                    UltimoResultadoConsultaComprobante.ErrorMessage = "Se produjo una excepción:" & vbCrLf & vbCrLf & ex.Message
                End Try

                Return True
            End Function

            Friend Function FacturaElectronica_ConectarYConsultarComprobante(ByVal TipoComprobante As Short, ByVal PuntoVenta As Short, ByVal ComprobanteNumero As Integer) As Boolean
                If FacturaElectronica_Conectar() Then
                    Return FacturaElectronica_ConsultarComprobante(TipoComprobante, PuntoVenta, ComprobanteNumero)
                Else
                    Return False
                End If
            End Function
        End Class

#End Region

    End Module
End Namespace