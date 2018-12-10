Friend Class CS_Database_ADONET
    Private mConnectionCurrent As OleDb.OleDbConnection
    Friend Property DataSource As String
    Friend Property ConnectionString As String


    Friend Const JET_REPLICATION_MEMOTYPE_COLUMNNAMEPREFIX As String = "Gen_"
    Friend Shared JET_REPLICATION_COLUMNS() As String = {"s_ColLineage", "s_Generation", "s_GUID", "s_Lineage"}

    Friend Shared Function AskForConnectionString(ByVal connectionstringCurrent As String)
        Dim dataLinks As New MSDASC.DataLinks

        Dim connection As ADODB.Connection

        If connectionstringCurrent.Trim = String.Empty Then

            'get a new connection string
            Try
                connection = CType(dataLinks.PromptNew, ADODB.Connection)
                Return connection.ConnectionString.ToString()
            Catch ex As Exception
                CS_Error.ProcessError(ex, "Error constructing Connection String.")
                Return connectionstringCurrent
            End Try
        Else
            'edit connection string
            connection = New ADODB.Connection
            connection.ConnectionString = connectionstringCurrent
            'set local COM compatible data type
            Dim oConnection As Object = connection

            Try
                'prompt user to edit the given connect string
                If CBool(dataLinks.PromptEdit(oConnection)) Then
                    Return connection.ConnectionString
                Else
                    Return connectionstringCurrent
                End If

            Catch ex As Exception
                CS_Error.ProcessError(ex, "Error editing Connection String.")
                Return connectionstringCurrent
            End Try

            oConnection = Nothing
        End If

        dataLinks = Nothing
        connection = Nothing
    End Function

    Friend Sub CreateConnectionString(ByVal Provider As String)
        Dim odsb As System.Data.OleDb.OleDbConnectionStringBuilder = New System.Data.OleDb.OleDbConnectionStringBuilder()

        With odsb
            .DataSource = DataSource
            .Provider = Provider

            ConnectionString = .ConnectionString
        End With
    End Sub

    Friend Function Connect() As Boolean

        Cursor.Current = Cursors.WaitCursor

        Try
            mConnectionCurrent = New System.Data.OleDb.OleDbConnection
            mConnectionCurrent.ConnectionString = ConnectionString
            mConnectionCurrent.Open()

            Cursor.Current = Cursors.Default
            Return True

        Catch ex As Exception
            CS_Error.ProcessError(ex, "Error al crear la conexión a la Base de Datos." & ControlChars.CrLf & ControlChars.CrLf & "ConnectionString: " & ConnectionString)
            Return False
        End Try
    End Function

    Friend Function Disconnect() As Boolean
        Try
            If Not mConnectionCurrent Is Nothing Then
                If mConnectionCurrent.State = ConnectionState.Open Then
                    mConnectionCurrent.Close()
                    mConnectionCurrent = Nothing
                End If
            End If
            Return True

        Catch ex As Exception
            CS_Error.ProcessError(ex, "Error al cerrar la conexión a la Base de Datos.")

            Return False
        End Try
    End Function

    Friend Function OpenDataReader(ByVal CommandText As String, ByRef DataReader As Data.Common.DbDataReader, ByVal ErrorMessage As String) As Boolean
        Dim Command As System.Data.Common.DbCommand

        Try
            Cursor.Current = Cursors.WaitCursor

            Command = New Data.OleDb.OleDbCommand
            Command.Connection = mConnectionCurrent

            Command.CommandText = CommandText
            Command.CommandType = CommandType.Text

            DataReader = Command.ExecuteReader()
            Command = Nothing

            Cursor.Current = Cursors.Default

            Return True

        Catch ex As Exception
            CS_Error.ProcessError(ex, ErrorMessage)

            Return False
        End Try
    End Function

    Function OpenDataSet(ByRef DataAdapter As OleDb.OleDbDataAdapter, ByRef DataSet As System.Data.DataSet, ByVal SelectCommandText As String, ByVal SourceTable As String, ByVal ErrorMessage As String) As Boolean

        Try
            Cursor.Current = Cursors.WaitCursor

            DataSet = New System.Data.DataSet
            DataAdapter = New OleDb.OleDbDataAdapter(SelectCommandText, mConnectionCurrent)
            Dim CommandBuilder As New OleDb.OleDbCommandBuilder(DataAdapter)
            DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            DataAdapter.Fill(DataSet, SourceTable)

            Cursor.Current = Cursors.Default
            Return True

        Catch ex As Exception
            CS_Error.ProcessError(ex, ErrorMessage)
            Return False
        End Try
    End Function

    Function OpenDataTable(ByRef DataTable As System.Data.DataTable, ByVal SelectCommandText As String, ByVal SourceTable As String, ByVal ErrorMessage As String) As Boolean
        Dim DataAdapter As OleDb.OleDbDataAdapter = Nothing
        Dim DataSet As System.Data.DataSet = Nothing

        If OpenDataSet(DataAdapter, DataSet, SelectCommandText, SourceTable, ErrorMessage) Then
            Try
                Cursor.Current = Cursors.WaitCursor
                DataSet.Tables(0).TableName = SourceTable
                DataTable = DataSet.Tables(SourceTable)
                Cursor.Current = Cursors.Default
                Return True

            Catch ex As Exception
                CS_Error.ProcessError(ex, ErrorMessage)
                Return False

            End Try
        Else
            OpenDataTable = False
        End If

    End Function
End Class