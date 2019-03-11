Module CS_Database_EF_SQL
    Private Const ERROR_RELATEDENTITY_NUMBER As Integer = 547
    Private Const ERROR_RELATEDENTITY_MESSAGE As String = "The DELETE statement conflicted with the REFERENCE constraint ""{CONSTRAINT_NAME}"". The conflict occurred in database ""{DATABASE_NAME}"", table ""{TABLE_NAME}"", column '{COLUMN_NAME}'."
    Private Const ERROR_DUPLICATEDENTITY_NUMBER As Integer = 2601
    Private Const ERROR_DUPLICATEDENTITY_MESSAGE As String = "Cannot insert duplicate key row in object '{TABLE_NAME}' with unique index '{UNIQUE_INDEX}'. The duplicate key value is ({VALUE})."

    Friend Enum Errors
        NoDBError
        Unknown
        InvalidColumn
        RelatedEntity
        DuplicatedEntity
        PrimaryKeyViolation
        UserDefinedError
    End Enum

    Friend Function TryDecodeDbUpdateException(ByRef ex As System.Data.Entity.Infrastructure.DbUpdateException) As Errors
        If (TypeOf (ex.InnerException) Is System.Data.Entity.Core.UpdateException) Or (TypeOf (ex.InnerException.InnerException) Is System.Data.SqlClient.SqlException) Then
            Dim SQLException As System.Data.SqlClient.SqlException

            SQLException = CType(ex.InnerException.InnerException, SqlClient.SqlException)

            For Each Err As System.Data.SqlClient.SqlError In SQLException.Errors
                Select Case Err.Number
                    Case 207
                        Return Errors.InvalidColumn
                    Case 547
                        Return Errors.RelatedEntity
                    Case 2601
                        Return Errors.DuplicatedEntity
                    Case 2627
                        Return Errors.PrimaryKeyViolation
                    Case 50000
                        Return Errors.UserDefinedError
                End Select
            Next

            Return Errors.Unknown
        Else
            Return Errors.NoDBError
        End If
    End Function

End Module