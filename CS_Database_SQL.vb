Friend Class CS_Database_SQL
    Friend Property ApplicationName As String
    Friend Property AttachDBFilename As String
    Friend Property DataSource As String
    Friend Property InitialCatalog As String
    Friend Property UserID As String
    Friend Property Password As String
    Friend Property MultipleActiveResultsets As Boolean
    Friend Property WorkstationID As String
    Friend Property ConnectionString As String

    Friend Sub CreateConnectionString()
        Dim scsb As System.Data.SqlClient.SqlConnectionStringBuilder = New System.Data.SqlClient.SqlConnectionStringBuilder()

        With scsb
            .ApplicationName = ApplicationName
            .DataSource = DataSource
            If (Not AttachDBFilename Is Nothing) AndAlso AttachDBFilename.Trim.Length > 0 Then
                .AttachDBFilename = AttachDBFilename
            End If
            If (Not InitialCatalog Is Nothing) AndAlso InitialCatalog.Trim.Length > 0 Then
                .InitialCatalog = InitialCatalog
            End If
            If (Not UserID Is Nothing) AndAlso UserID.Trim.Length > 0 Then
                .UserID = UserID
            End If
            If (Not Password Is Nothing) AndAlso Password.Trim.Length > 0 Then
                .Password = Password
            End If
            .MultipleActiveResultSets = MultipleActiveResultsets
            .WorkstationID = WorkstationID

            ConnectionString = .ConnectionString
        End With
    End Sub
End Class