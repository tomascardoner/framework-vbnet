Imports System.Data.SqlClient
Imports System.IO

Namespace CardonerSistemas.Database.Ado

    Friend Class SqlServer

        Private Const ErrorLoginFailed As String = "Login failed for user "

#Region "Properties"

        Friend Property ApplicationName As String
        Friend Property AttachDBFilename As String
        Friend Property Datasource As String
        Friend Property IntegratedSecurity As Boolean
        Friend Property InitialCatalog As String
        Friend Property UserId As String
        Friend Property PasswordEncrypted As String
        Friend Property Password As String
        Friend Property TrustServerCertificate As Boolean
        Friend Property MultipleActiveResultsets As Boolean
        Friend Property WorkstationID As String
        Friend Property ConnectTimeout As Byte
        Friend Property ConnectRetryCount As Byte
        Friend Property ConnectRetryInterval As Byte
        Friend Property ConnectionString As String

        Friend Property Connection As SqlConnection

        Friend Function PasswordEncrypt() As Boolean
            Dim encryptedPassword As String = String.Empty
            If Encrypt.StringCipher.Encrypt(Password, Constants.PublicEncryptionPassword, encryptedPassword) Then
                PasswordEncrypted = encryptedPassword
                Return True
            Else
                Return False
            End If
        End Function

        Friend Function PasswordUnencrypt() As Boolean
            Dim unencryptedPassword As String = String.Empty
            If Encrypt.StringCipher.Decrypt(PasswordEncrypted, Constants.PublicEncryptionPassword, unencryptedPassword) Then
                Password = unencryptedPassword
                Return True
            Else
                Return False
            End If
        End Function

#End Region

#Region "Connection"

        Friend Function SetProperties(datasourceValue As String, initialCatalogValue As String, attachDBFilenameValue As String, integratedSecurityValue As Boolean, userIdValue As String, passwordEncryptedValue As String, trustServerCertificateValue As Boolean, connectTimeoutValue As Byte, connectRetryCountValue As Byte, connectRetryIntervalValue As Byte) As Boolean
            Dim selectedDatasourceIndex As Integer

            If datasourceValue.Contains(Constants.StringListSeparator) Then
                ' Muestro la ventana de selección del Datasource
                Using selectDatasource As New SelectDatasource()
                    selectDatasource.comboboxDataSource.Items.AddRange(datasourceValue.Split(Convert.ToChar(Constants.StringListSeparator)))
                    If selectDatasource.ShowDialog() <> DialogResult.OK Then
                        Return False
                    End If
                    selectedDatasourceIndex = selectDatasource.comboboxDataSource.SelectedIndex
                    selectDatasource.Close()
                End Using

                ' Asigno las propiedades
                Datasource = SelectProperty(datasourceValue, selectedDatasourceIndex)
                InitialCatalog = SelectProperty(initialCatalogValue, selectedDatasourceIndex)
                UserId = SelectProperty(userIdValue, selectedDatasourceIndex)
                PasswordEncrypted = SelectProperty(passwordEncryptedValue, selectedDatasourceIndex)
            Else
                Datasource = datasourceValue
                InitialCatalog = initialCatalogValue
                UserId = userIdValue
                PasswordEncrypted = passwordEncryptedValue
            End If
            AttachDBFilename = attachDBFilenameValue
            IntegratedSecurity = integratedSecurityValue
            If IntegratedSecurity Then
                UserId = String.Empty
                PasswordEncrypted = String.Empty
            End If
            TrustServerCertificate = trustServerCertificateValue
            ConnectTimeout = connectTimeoutValue
            ConnectRetryCount = connectRetryCountValue
            ConnectRetryInterval = connectRetryIntervalValue
            ApplicationName = My.Application.Info.Title
            MultipleActiveResultsets = True
            WorkstationID = Environment.MachineName

            Return True
        End Function

        Private Shared Function SelectProperty(value As String, selectedIndex As Integer) As String
            If value.Contains(Constants.StringListSeparator) Then
                Dim values As String()
                values = value.Split(Convert.ToChar(Constants.StringListSeparator))
                If (values.GetUpperBound(0) >= selectedIndex) Then
                    Return values(selectedIndex)
                Else
                    Return String.Empty
                End If
            Else
                Return value
            End If
        End Function

        Friend Sub CreateConnectionString()
            Dim scsb As New SqlConnectionStringBuilder()

            With scsb
                .ApplicationName = ApplicationName
                .DataSource = Datasource
                If AttachDBFilename IsNot Nothing AndAlso AttachDBFilename.Trim.Length > 0 Then
                    .AttachDBFilename = AttachDBFilename
                End If
                If InitialCatalog IsNot Nothing AndAlso InitialCatalog.Trim.Length > 0 Then
                    .InitialCatalog = InitialCatalog
                End If
                .IntegratedSecurity = IntegratedSecurity
                If (UserId IsNot Nothing) AndAlso UserId.Trim.Length > 0 Then
                    .UserID = UserId
                End If
                If (Password IsNot Nothing) AndAlso Password.Trim.Length > 0 Then
                    .Password = Password
                End If
                .TrustServerCertificate = TrustServerCertificate
                .ConnectTimeout = ConnectTimeout
                .ConnectRetryCount = ConnectRetryCount
                .ConnectRetryInterval = ConnectRetryInterval
                .MultipleActiveResultSets = MultipleActiveResultsets
                .WorkstationID = WorkstationID

                ConnectionString = .ConnectionString
            End With
        End Sub

        Friend Function Connect(ByVal errorMessage As String) As Boolean
            Try
                Connection = New SqlConnection(ConnectionString)
                Connection.Open()
                Return True

            Catch ex As Exception
                ErrorHandler.ProcessError(ex, errorMessage)
                Return False
            End Try
        End Function

        Friend Function Connect() As Boolean
            Return Connect("Error al conectarse a la Base de Datos.")
        End Function

        Friend Function Connect(ByRef databaseConfig As DatabaseConfig, ByRef newLoginData As Boolean) As Boolean
            newLoginData = False

            Do While True
                Try
                    Connection = New SqlConnection(ConnectionString)
                    Connection.Open()
                    Return True
                Catch ex As Exception
                    If ex.HResult = -2146232060 AndAlso ex.Message.Contains(ErrorLoginFailed) Then
                        ' Los datos de inicio de sesión en la base de datos son incorrectos.
                        MessageBox.Show("Los datos de inicio de sesión a la base de datos son incorrectos.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                        ' Pido datos nuevos.
                        Using loginInfo As New LoginInfo()
                            loginInfo.textboxUsuario.Text = UserId
                            loginInfo.textboxPassword.Text = Password
                            If loginInfo.ShowDialog() <> DialogResult.OK Then
                                Return False
                            End If
                            UserId = loginInfo.textboxUsuario.Text.TrimAndReduce()
                            databaseConfig.UserId = loginInfo.textboxUsuario.Text.TrimAndReduce()
                            Password = loginInfo.textboxPassword.Text.Trim()
                            If Password.Length > 0 Then
                                If PasswordEncrypt() Then
                                    databaseConfig.Password = PasswordEncrypted
                                End If
                            Else
                                databaseConfig.Password = String.Empty
                            End If
                            CreateConnectionString()
                            loginInfo.Close()
                        End Using
                        newLoginData = True
                        Return True
                    Else
                        ErrorHandler.ProcessError(ex, "Error al conectarse a la Base de Datos.")
                        Return False
                    End If
                End Try
            Loop
#Disable Warning BC42353 ' Function doesn't return a value on all code paths
        End Function
#Enable Warning BC42353 ' Function doesn't return a value on all code paths

        Friend Function IsConnected() As Boolean
            Return Not (Connection Is Nothing OrElse Connection.State = ConnectionState.Closed OrElse Connection.State = ConnectionState.Broken)
        End Function

        Friend Function Close() As Boolean
            If IsConnected() Then
                Try
                    Connection.Close()
                    Connection = Nothing
                    Return True
                Catch ex As Exception
                    ErrorHandler.ProcessError(ex, "Error al cerrar la conexión a la Base de Datos.")
                    Return False
                End Try
            Else
                Return True
            End If
        End Function

#End Region

#Region "Retrieve data"

        Friend Function OpenDataReader(ByRef dataReader As SqlDataReader, ByVal commandText As String, ByVal commandType As CommandType, ByVal commandBehavior As CommandBehavior, ByVal errorMessage As String) As Boolean

            Try
                Dim Command As New SqlCommand With {
                    .Connection = Connection,
                    .CommandText = commandText,
                    .CommandType = commandType
                }
                dataReader = Command.ExecuteReader(commandBehavior)

                Command = Nothing

                Return True

            Catch ex As Exception
                ErrorHandler.ProcessError(ex, errorMessage)
                Return False

            End Try
        End Function

        Friend Function OpenDataSet(ByRef dataAdapter As SqlDataAdapter, ByRef dataSet As DataSet, ByVal selectCommandText As String, ByVal sourceTable As String, ByVal errorMessage As String) As Boolean
            Try

                dataSet = New DataSet()
                dataAdapter = New SqlDataAdapter(selectCommandText, Connection) With {
                    .MissingSchemaAction = MissingSchemaAction.AddWithKey
                }
                dataAdapter.Fill(dataSet, sourceTable)
                Return True

            Catch ex As Exception

                ErrorHandler.ProcessError(ex, errorMessage)
                Return False

            End Try

        End Function

        Friend Function OpenDataTable(ByRef dataTable As DataTable, ByVal selectCommandText As String, ByVal sourceTable As String, ByVal errorMessage As String) As Boolean
            Try
                Dim dataAdapter As SqlDataAdapter = Nothing
                Dim dataSet As DataSet = Nothing

                If OpenDataSet(dataAdapter, dataSet, selectCommandText, sourceTable, errorMessage) Then
                    dataSet.Tables(0).TableName = sourceTable
                    dataTable = dataSet.Tables(sourceTable)
                    Return True
                Else
                    Return False
                End If

            Catch ex As Exception

                ErrorHandler.ProcessError(ex, errorMessage)
                Return False

            End Try
        End Function


        Friend Function Execute(ByVal commandText As String, ByVal commandType As CommandType, ByVal errorMessage As String) As Boolean
            Try
                Dim command As New SqlCommand With {
                    .Connection = Connection,
                    .CommandText = commandText,
                    .CommandType = commandType
                }
                command.ExecuteNonQuery()
                Return True

            Catch ex As Exception
                ErrorHandler.ProcessError(ex, errorMessage)
                Return False

            End Try

        End Function

        Friend Function Execute(ByVal commandText As String, ByVal commandType As CommandType, ByRef sqlParameterCollection As SqlParameterCollection, ByVal errorMessage As String) As Boolean
            Try
                Dim Command As New SqlCommand With {
                    .Connection = Connection,
                    .CommandText = commandText,
                    .CommandType = commandType
                }
                For Each parameter As SqlParameter In sqlParameterCollection
                    Command.Parameters.Add(parameter)
                Next
                Command.ExecuteNonQuery()
                Return True

            Catch ex As Exception
                ErrorHandler.ProcessError(ex, errorMessage)
                Return False

            End Try
        End Function

#End Region

#Region "Get values - data reader"

        Friend Shared Function GetOrdinalSafe(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Integer
            Try
                Return dataReader.GetOrdinal(columnName)

            Catch ex As Exception

                Return -1
            End Try
        End Function

        Friend Shared Function DataReaderGetString(ByRef dataReader As SqlDataReader, ByVal columnName As String) As String
            Return dataReader.GetString(dataReader.GetOrdinal(columnName))
        End Function

        Friend Shared Function DataReaderGetStringSafeAsEmpty(ByRef dataReader As SqlDataReader, ByVal columnName As String) As String
            Dim result As String = DataReaderGetStringSafeAsNull(dataReader, columnName)

            If result = Nothing Then
                Return String.Empty
            Else
                Return result
            End If
        End Function

        Friend Shared Function DataReaderGetStringSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As String
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetString(columnIndex)
            End If
        End Function

        Friend Shared Function DataReaderGetByte(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Byte
            Return dataReader.GetByte(dataReader.GetOrdinal(columnName))
        End Function

        Friend Shared Function DataReaderGetByteSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Byte
            Dim result As Byte? = DataReaderGetByteSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return Byte.MinValue
            End If
        End Function

        Friend Shared Function DataReaderGetByteSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Byte?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetByte(columnIndex)
            End If
        End Function

        Friend Shared Function DataReaderGetShort(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Short
            Return dataReader.GetInt16(dataReader.GetOrdinal(columnName))
        End Function

        Friend Shared Function DataReaderGetShortSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Short
            Dim result As Short? = DataReaderGetShortSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return Short.MinValue
            End If
        End Function

        Friend Shared Function DataReaderGetShortSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Short?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetInt16(columnIndex)
            End If
        End Function

        Friend Shared Function DataReaderGetInteger(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Integer
            Return dataReader.GetInt32(dataReader.GetOrdinal(columnName))
        End Function

        Friend Shared Function DataReaderGetIntegerSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Integer
            Dim result As Integer? = DataReaderGetIntegerSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return Integer.MinValue
            End If
        End Function

        Friend Shared Function DataReaderGetIntegerSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Integer?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetInt32(columnIndex)
            End If
        End Function

        Friend Shared Function DataReaderGetLong(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Long
            Return dataReader.GetInt64(dataReader.GetOrdinal(columnName))
        End Function

        Friend Shared Function DataReaderGetLongSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Long
            Dim result As Long? = DataReaderGetLongSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return Long.MinValue
            End If
        End Function

        Friend Shared Function DataReaderGetLongSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Long?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetInt64(columnIndex)
            End If
        End Function

        Friend Shared Function DataReaderGetDecimal(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Decimal
            Return dataReader.GetDecimal(dataReader.GetOrdinal(columnName))
        End Function

        Friend Shared Function DataReaderGetDecimalSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Decimal
            Dim result As Decimal? = DataReaderGetDecimalSafeAsNull(dataReader, columnName)
            If result.HasValue Then
                Return result.Value
            Else
                Return Decimal.MinValue
            End If
        End Function

        Friend Shared Function DataReaderGetDecimalSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Decimal?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetDecimal(columnIndex)
            End If
        End Function

        Friend Shared Function DataReaderGetBoolean(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Boolean
            Return dataReader.GetBoolean(dataReader.GetOrdinal(columnName))
        End Function

        Friend Shared Function DataReaderGetBooleanSafeAsByte(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Byte
            Dim result As Boolean? = DataReaderGetBooleanSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                If result.Value Then
                    Return 1
                Else
                    Return 0
                End If
            Else
                Return 2
            End If
        End Function

        Friend Shared Function DataReaderGetBooleanSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Boolean?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetBoolean(columnIndex)
            End If
        End Function

        Friend Shared Function DataReaderGetDateTime(ByRef dataReader As SqlDataReader, ByVal columnName As String) As System.DateTime
            Return dataReader.GetDateTime(dataReader.GetOrdinal(columnName))
        End Function

        Friend Shared Function DataReaderGetDateTimeSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As System.DateTime
            Dim result As System.DateTime? = DataReaderGetDateTimeSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return System.DateTime.MinValue
            End If
        End Function

        Friend Shared Function DataReaderGetDateTimeSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As System.DateTime?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetDateTime(columnIndex)
            End If
        End Function

#End Region

#Region "Data reader - Get varbinary"

        Friend Shared Function DataReaderGetStream(ByRef dataReader As SqlDataReader, ByVal columnName As String, Optional ByVal errorMessage As String = "") As Stream
            Try
                Return dataReader.GetStream(dataReader.GetOrdinal(columnName))
            Catch ex As Exception
                If errorMessage <> "" Then
                    CardonerSistemas.ErrorHandler.ProcessError(ex, errorMessage)
                End If
                Return Nothing
            End Try
        End Function

        Friend Shared Function DataReaderGetStreamAsImage(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Image
            Dim stream As Stream = DataReaderGetStream(dataReader, columnName)

            If stream Is Nothing Then
                Return Nothing
            Else
                Try
                    Return Bitmap.FromStream(stream)
                Catch ex As Exception
                    Return Nothing
                End Try
            End If
        End Function

#End Region

    End Class

End Namespace