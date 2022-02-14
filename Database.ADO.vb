Imports System.Data.SqlClient
Imports System.IO

Namespace CardonerSistemas.Database.Ado

    Friend Class SqlServer

        Friend Property ApplicationName As String
        Friend Property AttachDBFilename As String
        Friend Property Datasource As String
        Friend Property InitialCatalog As String
        Friend Property UserId As String
        Friend Property Password As String
        Friend Property MultipleActiveResultsets As Boolean
        Friend Property WorkstationID As String

        Friend Property ConnectionString As String

        Friend Property Connection As SqlConnection

#Region "Connection"

        Friend Sub CreateConnectionString()
            Dim scsb As SqlConnectionStringBuilder = New SqlConnectionStringBuilder()

            With scsb
                .ApplicationName = ApplicationName
                .DataSource = Datasource
                If (Not AttachDBFilename Is Nothing) AndAlso AttachDBFilename.Trim.Length > 0 Then
                    .AttachDBFilename = AttachDBFilename
                End If
                If (Not InitialCatalog Is Nothing) AndAlso InitialCatalog.Trim.Length > 0 Then
                    .InitialCatalog = InitialCatalog
                End If
                If (Not UserId Is Nothing) AndAlso UserId.Trim.Length > 0 Then
                    .UserID = UserId
                End If
                If (Not Password Is Nothing) AndAlso Password.Trim.Length > 0 Then
                    .Password = Password
                End If
                .MultipleActiveResultSets = MultipleActiveResultsets
                .WorkstationID = WorkstationID

                ConnectionString = .ConnectionString
            End With
        End Sub

        Friend Function Connect() As Boolean
            Try

                Connection = New SqlConnection(ConnectionString)
                Connection.Open()
                Return True

            Catch ex As Exception

                CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al conectarse a la Base de Datos.")
                Return False

            End Try
        End Function

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

                    CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al cerrar la conexión a la Base de Datos.")
                    Return False
                End Try
            Else
                Return True
            End If
        End Function

#End Region

#Region "Retrieve data"

        Friend Function Connect(ByVal errorMessage As String) As Boolean
            Try
                Connection = New SqlConnection(ConnectionString)
                Connection.Open()
                Return True

            Catch ex As Exception
                CardonerSistemas.ErrorHandler.ProcessError(ex, errorMessage)
                Return False

            End Try
        End Function

        Friend Function OpenDataReader(ByRef dataReader As SqlDataReader, ByVal commandText As String, ByVal commandType As CommandType, ByVal commandBehavior As CommandBehavior, ByVal errorMessage As String) As Boolean

            Try
                Dim Command As SqlCommand = New SqlCommand()
                Command.Connection = Connection
                Command.CommandText = commandText
                Command.CommandType = commandType

                dataReader = Command.ExecuteReader(commandBehavior)

                Command = Nothing

                Return True

            Catch ex As Exception
                CardonerSistemas.ErrorHandler.ProcessError(ex, errorMessage)
                Return False

            End Try
        End Function

        Friend Function OpenDataSet(ByRef dataAdapter As SqlDataAdapter, ByRef dataSet As DataSet, ByVal selectCommandText As String, ByVal sourceTable As String, ByVal errorMessage As String) As Boolean
            Try

                dataSet = New DataSet()
                dataAdapter = New SqlDataAdapter(selectCommandText, Connection)
                Dim commandBuilder As SqlCommandBuilder = New SqlCommandBuilder(dataAdapter)
                dataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                dataAdapter.Fill(dataSet, sourceTable)

                Return True

            Catch ex As Exception

                CardonerSistemas.ErrorHandler.ProcessError(ex, errorMessage)
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

                CardonerSistemas.ErrorHandler.ProcessError(ex, errorMessage)
                Return False

            End Try

        End Function


        Friend Function Execute(ByVal commandText As String, ByVal commandType As CommandType, ByVal errorMessage As String) As Boolean
            Try
                Dim command As SqlCommand = New SqlCommand()

                command.Connection = Connection
                command.CommandText = commandText
                command.CommandType = commandType
                command.ExecuteNonQuery()
                Return True

            Catch ex As Exception
                CardonerSistemas.ErrorHandler.ProcessError(ex, errorMessage)
                Return False

            End Try

        End Function

        Friend Function Execute(ByVal commandText As String, ByVal commandType As CommandType, ByRef sqlParameterCollection As SqlParameterCollection, ByVal errorMessage As String) As Boolean
            Try
                Dim Command As SqlCommand = New SqlCommand()
                Command.Connection = Connection
                Command.CommandText = commandText
                Command.CommandType = commandType
                For Each parameter As SqlParameter In sqlParameterCollection
                    Command.Parameters.Add(parameter)
                Next
                Command.ExecuteNonQuery()
                Return True

            Catch ex As Exception
                CardonerSistemas.ErrorHandler.ProcessError(ex, errorMessage)
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

        Friend Function DataReaderGetByte(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Byte
            Return dataReader.GetByte(dataReader.GetOrdinal(columnName))
        End Function

        Friend Function DataReaderGetByteSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Byte
            Dim result As Byte? = DataReaderGetByteSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return Byte.MinValue
            End If
        End Function

        Friend Function DataReaderGetByteSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Byte?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetByte(columnIndex)
            End If
        End Function

        Friend Function DataReaderGetShort(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Short
            Return dataReader.GetInt16(dataReader.GetOrdinal(columnName))
        End Function

        Friend Function DataReaderGetShortSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Short
            Dim result As Short? = DataReaderGetShortSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return Short.MinValue
            End If
        End Function

        Friend Function DataReaderGetShortSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Short?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetInt16(columnIndex)
            End If
        End Function

        Friend Function DataReaderGetInteger(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Integer
            Return dataReader.GetInt32(dataReader.GetOrdinal(columnName))
        End Function

        Friend Function DataReaderGetIntegerSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Integer
            Dim result As Integer? = DataReaderGetIntegerSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return Integer.MinValue
            End If
        End Function

        Friend Function DataReaderGetIntegerSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Integer?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetInt32(columnIndex)
            End If
        End Function

        Friend Function DataReaderGetLong(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Long
            Return dataReader.GetInt64(dataReader.GetOrdinal(columnName))
        End Function

        Friend Function DataReaderGetLongSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Long
            Dim result As Long? = DataReaderGetLongSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return Long.MinValue
            End If
        End Function

        Friend Function DataReaderGetLongSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Long?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetInt64(columnIndex)
            End If
        End Function

        Friend Function DataReaderGetDecimal(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Decimal
            Return dataReader.GetDecimal(dataReader.GetOrdinal(columnName))
        End Function

        Friend Function DataReaderGetDecimalSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Decimal
            Dim result As Decimal? = DataReaderGetDecimalSafeAsNull(dataReader, columnName)
            If result.HasValue Then
                Return result.Value
            Else
                Return Decimal.MinValue
            End If
        End Function

        Friend Function DataReaderGetDecimalSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Decimal?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetDecimal(columnIndex)
            End If
        End Function

        Friend Function DataReaderGetBoolean(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Boolean
            Return dataReader.GetBoolean(dataReader.GetOrdinal(columnName))
        End Function

        Friend Function DataReaderGetBooleanSafeAsByte(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Byte
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

        Friend Function DataReaderGetBooleanSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Boolean?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetBoolean(columnIndex)
            End If
        End Function

        Friend Function DataReaderGetDateTime(ByRef dataReader As SqlDataReader, ByVal columnName As String) As System.DateTime
            Return dataReader.GetDateTime(dataReader.GetOrdinal(columnName))
        End Function

        Friend Function DataReaderGetDateTimeSafeAsMinValue(ByRef dataReader As SqlDataReader, ByVal columnName As String) As System.DateTime
            Dim result As System.DateTime? = DataReaderGetDateTimeSafeAsNull(dataReader, columnName)

            If result.HasValue Then
                Return result.Value
            Else
                Return System.DateTime.MinValue
            End If
        End Function

        Friend Function DataReaderGetDateTimeSafeAsNull(ByRef dataReader As SqlDataReader, ByVal columnName As String) As System.DateTime?
            Dim columnIndex As Integer = dataReader.GetOrdinal(columnName)

            If dataReader.IsDBNull(columnIndex) Then
                Return Nothing
            Else
                Return dataReader.GetDateTime(columnIndex)
            End If
        End Function

#End Region

#Region "Data reader - Get varbinary"

        Friend Function DataReaderGetStream(ByRef dataReader As SqlDataReader, ByVal columnName As String, Optional ByVal errorMessage As String = "") As Stream
            Try
                Return dataReader.GetStream(dataReader.GetOrdinal(columnName))
            Catch ex As Exception
                If errorMessage <> "" Then
                    CardonerSistemas.ErrorHandler.ProcessError(ex, errorMessage)
                End If
                Return Nothing
            End Try
        End Function

        Friend Function DataReaderGetStreamAsImage(ByRef dataReader As SqlDataReader, ByVal columnName As String) As Image
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