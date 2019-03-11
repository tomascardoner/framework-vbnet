Imports System.Security.Cryptography

Friend Class CS_Encrypt_TripleDES
    Private TripleDes As New TripleDESCryptoServiceProvider

    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()
        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)

        Return hash
    End Function

    Public Sub New(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

    Friend Function Encrypt(ByVal plaintext As String) As String

        ' Convert the plaintext string to a byte array.
        Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        Return Convert.ToBase64String(ms.ToArray)
    End Function

    Friend Function Decrypt(ByVal encryptedText As String, ByRef decryptedText As String) As Boolean
        If encryptedText.Trim = "" Then
            decryptedText = ""
            Return True
        Else
            Try
                ' Convert the encrypted text string to a byte array.
                Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedText)

                ' Create the stream.
                Dim ms As New System.IO.MemoryStream
                ' Create the decoder to write to the stream.
                Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

                ' Use the crypto stream to write the byte array to the stream.
                decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
                decStream.FlushFinalBlock()

                ' Convert the plaintext stream to a string.
                decryptedText = System.Text.Encoding.Unicode.GetString(ms.ToArray)
                Return True
            Catch
                decryptedText = ""
                Return False
            End Try
        End If
    End Function

End Class