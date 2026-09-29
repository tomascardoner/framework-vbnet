Option Strict Off

Imports System.Linq

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
            Friend Property CondicionIVAReceptorId As Integer

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
            Friend Property ModoHomologacion As Boolean = True
            Friend Property MonedaLocal As Moneda
            Friend Property MonedaLocalCotizacion As MonedaCotizacion

            ' Propiedades de Resultado
            Friend Property TicketAcceso As String
            Friend Property WSFEv1 As Object
            Friend Property UltimoResultadoCAE As ResultadoCAE
            Friend Property UltimoComprobanteAutorizado As String
            Friend Property UltimoResultadoConsultaComprobante As ResultadoConsultaComprobante

            ' Credenciales de ARCA (certificado + clave privada ya cargados), obtenidas en 'Login'.
            ' Se reutilizan en cada llamada porque Armuna.Framework.Tax cachea internamente el Ticket de Acceso
            ' (WSAA) por CUIT/entorno/servicio, así que no hace falta administrar el ticket manualmente aquí.
            Private mCredenciales As Armuna.Framework.Tax.Arca.ArcaCredentials

            Friend Sub New()
                UltimoResultadoCAE = New ResultadoCAE
                UltimoResultadoConsultaComprobante = New ResultadoConsultaComprobante
            End Sub

            Friend Function Login(ByVal ServicioNombre As String) As Boolean
                CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, StrDup(20, "="))

                Dim Entorno As Armuna.Framework.Tax.Arca.ArcaEnvironment = If(ModoHomologacion, Armuna.Framework.Tax.Arca.ArcaEnvironment.Homologacion, Armuna.Framework.Tax.Arca.ArcaEnvironment.Produccion)
                Dim Credenciales As Armuna.Framework.Tax.Arca.ArcaCredentials = Nothing
                Dim CredencialesResultMessage As String = Nothing

                If Not Armuna.Framework.Tax.Arca.ArcaCredentialsLoader.TryLoadFromPemFiles(CUIT_Emisor, Certificado, ClavePrivada, Entorno, Credenciales, CredencialesResultMessage) Then
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "No se pudo cargar el certificado: " & CredencialesResultMessage)
                    CardonerSistemas.ErrorHandler.ProcessError(New Exception(CredencialesResultMessage), "Error al cargar el certificado para el Servicio de AFIP.")
                    Return False
                End If

                Try
                    Dim Resultado = Armuna.Framework.Tax.Arca.Wsaa.WsaaService.GetTicketAsync(Credenciales, ServicioNombre).GetAwaiter().GetResult()
                    If Resultado.success Then
                        mCredenciales = Credenciales
                        TicketAcceso = Resultado.ticket.Token
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se autenticó correctamente.")
                        Return True
                    Else
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Falló la autenticación.")
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, Resultado.resultMessage)
                        CardonerSistemas.ErrorHandler.ProcessError(New Exception(Resultado.resultMessage), "Error al iniciar sesión en el Servidor de AFIP.")
                        Return False
                    End If

                Catch ex As Exception
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Ocurrió un error.")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Excepción: " & ex.Message)
                    CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al iniciar sesión en el Servidor de AFIP.")
                    Return False
                End Try
            End Function

            Friend Function FacturaElectronica_Login() As Boolean
                Return Login(ServicioFacturacionElectronica)
            End Function

            Friend Function FacturaElectronica_Conectar() As Boolean
                If mCredenciales Is Nothing Then
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "No se ha iniciado sesión. Debe invocar 'Login' o 'FacturaElectronica_Login' antes de conectar.")
                    WSFEv1 = Nothing
                    Return False
                End If

                Try
                    Dim Resultado = Armuna.Framework.Tax.Arca.Wsaa.WsaaService.GetTicketAsync(mCredenciales, ServicioFacturacionElectronica).GetAwaiter().GetResult()
                    If Resultado.success Then
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Se conectó al Servicio de Factura Electrónica.")
                        WSFEv1 = mCredenciales
                        Return True
                    Else
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "No se pudo conectar: " & Resultado.resultMessage)
                        CardonerSistemas.ErrorHandler.ProcessError(New Exception(Resultado.resultMessage), "Error al conectar con el Servicio de Factura Electrónica.")
                        WSFEv1 = Nothing
                        Return False
                    End If

                Catch ex As Exception
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Ha ocurrido un error: " & ex.Message)
                    CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al conectar con el Servicio de Factura Electrónica.")
                    WSFEv1 = Nothing
                    Return False
                End Try
            End Function

            Friend Function FacturaElectronica_ObtenerCAE(ByRef FacturaAGenerar As FacturaElectronicaCabecera) As Boolean
                If WSFEv1 Is Nothing Then
                    Return False
                End If

                Try
                    Dim Detalle As New Armuna.Framework.Tax.Arca.Wsfe.FacturaElectronicaDetalle With {
                        .Concepto = CType(FacturaAGenerar.Concepto, Armuna.Framework.Tax.Arca.Comprobante.Concepto),
                        .DocTipo = FacturaAGenerar.TipoDocumento,
                        .DocNro = FacturaAGenerar.DocumentoNumero,
                        .CbteDesde = FacturaAGenerar.ComprobanteDesde,
                        .CbteHasta = FacturaAGenerar.ComprobanteHasta,
                        .CbteFecha = FacturaAGenerar.ComprobanteFecha,
                        .ImpTotal = FacturaAGenerar.ImporteTotal,
                        .ImpTotConc = FacturaAGenerar.ImporteTotalConc,
                        .ImpNeto = FacturaAGenerar.ImporteNeto,
                        .ImpOpEx = FacturaAGenerar.ImporteOperacionesExentas,
                        .ImpTributos = FacturaAGenerar.ImporteTributos,
                        .ImpIva = FacturaAGenerar.ImporteIVA,
                        .MonedaId = FacturaAGenerar.MonedaID,
                        .MonedaCotizacion = FacturaAGenerar.MonedaCotizacion,
                        .CondicionIvaReceptorId = FacturaAGenerar.CondicionIVAReceptorId,
                        .FechaServicioDesde = If(FacturaAGenerar.FechaServicioDesde = Date.MinValue, CType(Nothing, Date?), FacturaAGenerar.FechaServicioDesde),
                        .FechaServicioHasta = If(FacturaAGenerar.FechaServicioHasta = Date.MinValue, CType(Nothing, Date?), FacturaAGenerar.FechaServicioHasta),
                        .FechaVencimientoPago = If(FacturaAGenerar.FechaVencimientoPago = Date.MinValue, CType(Nothing, Date?), FacturaAGenerar.FechaVencimientoPago),
                        .ComprobantesAsociados = FacturaAGenerar.ComprobantesAsociados.Select(Function(c) New Armuna.Framework.Tax.Arca.Wsfe.FacturaElectronicaComprobanteAsociado(c.TipoComprobante, c.PuntoVenta, c.ComprobanteNumero)).ToList(),
                        .Tributos = FacturaAGenerar.Tributos.Select(Function(t) New Armuna.Framework.Tax.Arca.Wsfe.FacturaElectronicaTributo(t.ID, t.BaseImponible, t.Alicuota, t.Importe, t.Descripcion)).ToList(),
                        .AlicuotasIva = FacturaAGenerar.IVAs.Select(Function(i) New Armuna.Framework.Tax.Arca.Wsfe.FacturaElectronicaAlicuotaIva(i.ID, i.BaseImponible, i.Importe)).ToList(),
                        .Opcionales = FacturaAGenerar.Opcionales.Select(Function(o) New Armuna.Framework.Tax.Arca.Wsfe.FacturaElectronicaOpcional(o.ID, o.Valor)).ToList()
                    }

                    Dim Resultado = Armuna.Framework.Tax.Arca.Wsfe.WsfeService.SolicitarCaeAsync(mCredenciales, FacturaAGenerar.PuntoVenta, FacturaAGenerar.TipoComprobante, {Detalle}).GetAwaiter().GetResult()

                    If Not Resultado.success Then
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Ocurrió un error: " & Resultado.resultMessage)
                        UltimoResultadoCAE.ErrorMessage = Resultado.resultMessage
                        CardonerSistemas.ErrorHandler.ProcessError(New Exception(Resultado.resultMessage), "Error al crear la Factura Electrónica.")
                        Return False
                    End If

                    Dim DetalleResultado = Resultado.resultado.Detalles.FirstOrDefault()
                    Dim ResultadoTexto As String = If(DetalleResultado IsNot Nothing, DetalleResultado.Resultado, Resultado.resultado.Resultado)

                    UltimoResultadoCAE.Resultado = If(String.IsNullOrEmpty(ResultadoTexto), "R"c, ResultadoTexto.Chars(0))
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Resultado: " & UltimoResultadoCAE.Resultado)

                    If UltimoResultadoCAE.Resultado = SolicitudCaeResultadoAceptado AndAlso DetalleResultado IsNot Nothing Then
                        UltimoResultadoCAE.Numero = DetalleResultado.Cae
                        UltimoResultadoCAE.FechaVencimiento = DetalleResultado.CaeFechaVencimiento.GetValueOrDefault()
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Tipo:  " & FacturaAGenerar.TipoComprobante)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Nro.:  " & FacturaAGenerar.PuntoVenta & "-" & FacturaAGenerar.ComprobanteDesde)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "CAE:               " & UltimoResultadoCAE.Numero)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Fecha Vencimiento: " & UltimoResultadoCAE.FechaVencimiento.ToShortDateString)
                    Else
                        UltimoResultadoCAE.Observaciones = String.Join(vbCrLf, If(DetalleResultado IsNot Nothing, DetalleResultado.Observaciones, New List(Of String)))
                        UltimoResultadoCAE.ErrorMessage = String.Join(vbCrLf, Resultado.resultado.Errores)

                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Tipo:  " & FacturaAGenerar.TipoComprobante)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Nro.:  " & FacturaAGenerar.PuntoVenta & "-" & FacturaAGenerar.ComprobanteDesde)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Observaciones: " & UltimoResultadoCAE.Observaciones)
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Error:         " & UltimoResultadoCAE.ErrorMessage)
                    End If

                    Return True

                Catch ex As Exception
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Critical, "Ocurrió un error.")
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Excepción: " & ex.Message)
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
                    Dim Resultado = Armuna.Framework.Tax.Arca.Wsfe.WsfeService.ObtenerUltimoComprobanteAutorizadoAsync(mCredenciales, PuntoVenta, TipoComprobante).GetAwaiter().GetResult()
                    If Resultado.success Then
                        UltimoComprobanteAutorizado = Resultado.ultimoComprobante.ToString()
                        Return True
                    Else
                        Return False
                    End If

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
                    Dim Resultado = Armuna.Framework.Tax.Arca.Wsfe.WsfeService.ConsultarComprobanteAsync(mCredenciales, PuntoVenta, TipoComprobante, ComprobanteNumero).GetAwaiter().GetResult()

                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Tipo:  " & TipoComprobante)
                    CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Comprobante Nro.:  " & PuntoVenta & "-" & ComprobanteNumero)

                    If Not Resultado.success Then
                        UltimoResultadoConsultaComprobante.ErrorMessage = Resultado.resultMessage
                        Return True
                    End If

                    Dim Consulta = Resultado.resultado

                    With UltimoResultadoConsultaComprobante
                        .Resultado = If(String.IsNullOrEmpty(Consulta.Resultado), " "c, Consulta.Resultado.Chars(0))
                        CS_FileLog.WriteLine(LogPath, LogFileName, LogEntryType.Information, "Resultado: " & .Resultado)
                        .CodigoAutorizacion = Consulta.CodigoAutorizacion

                        If .Resultado = SolicitudCaeResultadoAceptado Then
                            .Concepto = Consulta.Concepto.GetValueOrDefault()
                            .TipoDocumento = Consulta.DocumentoTipo.GetValueOrDefault()
                            .DocumentoNumero = Consulta.DocumentoNro.GetValueOrDefault()
                            .TipoComprobante = Consulta.ComprobanteTipo.GetValueOrDefault()
                            .PuntoVenta = Consulta.PuntoVenta.GetValueOrDefault()
                            .ComprobanteDesde = Consulta.ComprobanteNumero.GetValueOrDefault()
                            .ComprobanteHasta = Consulta.ComprobanteHasta.GetValueOrDefault()
                            .ComprobanteFecha = Consulta.ComprobanteFecha.GetValueOrDefault()
                            .ImporteTotal = Consulta.ImporteTotal.GetValueOrDefault()
                            .ImporteTotalConc = Consulta.ImporteTotalConc.GetValueOrDefault()
                            .ImporteNeto = Consulta.ImporteNeto.GetValueOrDefault()
                            .ImporteTributos = Consulta.ImporteTributos.GetValueOrDefault()
                            .ImporteIVA = Consulta.ImporteIVA.GetValueOrDefault()
                            .FechaServicioDesde = Consulta.FechaServicioDesde.GetValueOrDefault()
                            .FechaServicioHasta = Consulta.FechaServicioHasta.GetValueOrDefault()
                            .FechaVencimientoPago = Consulta.FechaVencimientoPago.GetValueOrDefault()
                            .MonedaID = Consulta.MonedaId
                            .MonedaCotizacion = Consulta.MonedaCotizacion.GetValueOrDefault()
                            .EmisionTipo = Consulta.EmisionTipo
                            .FechaVencimiento = Consulta.CaeFechaVencimiento.GetValueOrDefault()
                            .FechaHoraProceso = Consulta.FechaProceso.GetValueOrDefault()
                        End If

                        .Observaciones = String.Join(vbCrLf, Consulta.Observaciones)
                        .ErrorMessage = String.Join(vbCrLf, Consulta.Errores)
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
