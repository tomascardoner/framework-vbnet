Imports System.IO
Imports System.Security.Cryptography
Imports System.Text.Encoding

Friend Class CS_Encrypt_AES256
    Public Function Decrypt(ByVal encryptedText As String, ByVal secretKey As String) As String
        Dim plainText As String = Nothing
        Using inputStream As MemoryStream = New MemoryStream(System.Convert.FromBase64String(encryptedText))
            Dim algorithm As RijndaelManaged = getAlgorithm(secretKey)
            Using cryptoStream As CryptoStream = New CryptoStream(inputStream, algorithm.CreateDecryptor(), CryptoStreamMode.Read)
                Dim outputBuffer(0 To CType(inputStream.Length - 1, Integer)) As Byte
                Dim readBytes As Integer = cryptoStream.Read(outputBuffer, 0, CType(inputStream.Length, Integer))
                plainText = Unicode.GetString(outputBuffer, 0, readBytes)
            End Using
        End Using
        Return plainText
    End Function


    Public Function Encrypt(ByVal plainText As String, ByVal secretKey As String) As String
        Dim encryptedPassword As String = Nothing
        Using outputStream As MemoryStream = New MemoryStream()
            Dim algorithm As RijndaelManaged = getAlgorithm(secretKey)
            Using cryptoStream As CryptoStream = New CryptoStream(outputStream, algorithm.CreateEncryptor(), CryptoStreamMode.Write)
                Dim inputBuffer() As Byte = Unicode.GetBytes(plainText)
                cryptoStream.Write(inputBuffer, 0, inputBuffer.Length)
                cryptoStream.FlushFinalBlock()
                encryptedPassword = System.Convert.ToBase64String(outputStream.ToArray())
            End Using
        End Using
        Return encryptedPassword
    End Function


    Private Function getAlgorithm(ByVal secretKey As String) As RijndaelManaged
        Const salt As String = "put a salt key here"
        Const keySize As Integer = 256

        Dim keyBuilder As Rfc2898DeriveBytes = New Rfc2898DeriveBytes(secretKey, Unicode.GetBytes(salt))
        Dim algorithm As RijndaelManaged = New RijndaelManaged()
        algorithm.KeySize = keySize
        algorithm.IV = keyBuilder.GetBytes(CType(algorithm.BlockSize / 8, Integer))
        algorithm.Key = keyBuilder.GetBytes(CType(algorithm.KeySize / 8, Integer))
        algorithm.Padding = PaddingMode.PKCS7
        Return algorithm
    End Function
End Class
