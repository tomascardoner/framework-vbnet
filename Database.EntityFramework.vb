Imports System.Data.Entity.Core.EntityClient

Namespace CardonerSistemas.Database

    Friend Class EntityFramework
        Private Const ErrorRelatedEntityNumber As Integer = 547
        Private Const ErrorRelatedEntityMessage As String = "The DELETE statement conflicted with the REFERENCE constraint ""{CONSTRAINT_NAME}"". The conflict occurred in database ""{DATABASE_NAME}"", table ""{TABLE_NAME}"", column '{COLUMN_NAME}'."
        Private Const ErrorDuplicatedEntityNumber As Integer = 2601
        Private Const ErrorDuplicatedEntityMessage As String = "Cannot insert duplicate key row in object '{TABLE_NAME}' with unique index '{UNIQUE_INDEX}'. The duplicate key value is ({VALUE})."

        Friend Enum Errors As Integer
            NoDBError
            Unknown
            InvalidColumn = 207
            RelatedEntity = 547
            DuplicatedEntity = 2601
            PrimaryKeyViolation = 2627
            StringOrBinaryDataWillBeTruncated = 8152
            UserDefinedError = 50000
        End Enum

        Friend Shared Function CreateConnectionString(ByVal Provider As String, ByVal ProviderConnectionString As String, ByVal modelName As String) As String
            Dim ecb As EntityConnectionStringBuilder = New EntityConnectionStringBuilder()

            ecb.Metadata = String.Format("res://*/{0}.csdl|res://*/{0}.ssdl|res://*/{0}.msl", modelName)
            ecb.Provider = Provider
            ecb.ProviderConnectionString = ProviderConnectionString

            Return ecb.ConnectionString
        End Function

        Friend Shared Function TryDecodeDbUpdateException(ByRef ex As System.Data.Entity.Infrastructure.DbUpdateException) As Errors
            If (TypeOf (ex.InnerException) Is System.Data.Entity.Core.UpdateException) Or (TypeOf (ex.InnerException.InnerException) Is System.Data.SqlClient.SqlException) Then
                Dim SQLException As System.Data.SqlClient.SqlException

                SQLException = CType(ex.InnerException.InnerException, SqlClient.SqlException)

                For Each Err As System.Data.SqlClient.SqlError In SQLException.Errors
                    If [Enum].IsDefined(GetType(Errors), Err.Number) Then
                        Return CType([Enum].ToObject(GetType(Errors), Err.Number), Errors)
                    End If
                Next

                Return Errors.Unknown
            Else
                Return Errors.NoDBError
            End If
        End Function

    End Class

End Namespace