Imports System.Text
Imports System.Security.Cryptography
Imports System.IO

Namespace CardonerSistemas.Encrypt
    Friend Module StringCipher
        ' This constant is used to determine the keysize of the encryption algorithm in bits.
        ' We divide this by 8 within the code below to get the equivalent number of bytes.
        Private Const Keysize As Integer = 256

        ' This constant determines the number of iterations for the password bytes generation function.
        Private Const DerivationIterations As Integer = 1000

        Friend Function Encrypt(plainText As String, passPhrase As String, ByRef encryptedText As String) As Boolean
            If String.IsNullOrWhiteSpace(plainText) Then
                encryptedText = String.Empty
                Return True
            End If
            ' Salt and IV is randomly generated each time, but is preprended to encrypted cipher text
            ' so that the same Salt and IV values can be used when decrypting.  
            Dim saltStringBytes = Generate256BitsOfRandomEntropy()
            Dim ivStringBytes = Generate256BitsOfRandomEntropy()
            Dim plainTextBytes = Encoding.UTF8.GetBytes(plainText)

            Try
                Using password = New Rfc2898DeriveBytes(passPhrase, saltStringBytes, DerivationIterations)
                    Dim keyBytes = password.GetBytes(CInt(Keysize / 8))
                    Using symmetricKey = New RijndaelManaged()
                        symmetricKey.BlockSize = 256
                        symmetricKey.Mode = CipherMode.CBC
                        symmetricKey.Padding = PaddingMode.PKCS7
                        Using encryptor = symmetricKey.CreateEncryptor(keyBytes, ivStringBytes)
                            Using memoryStream = New MemoryStream()
                                Using cryptoStream = New CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write)
                                    cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length)
                                    cryptoStream.FlushFinalBlock()
                                    ' Create the final bytes as a concatenation of the random salt bytes, the random iv bytes and the cipher bytes.
                                    Dim cipherTextBytes = saltStringBytes
                                    cipherTextBytes = cipherTextBytes.Concat(ivStringBytes).ToArray()
                                    cipherTextBytes = cipherTextBytes.Concat(memoryStream.ToArray()).ToArray()
                                    memoryStream.Close()
                                    cryptoStream.Close()
                                    encryptedText = Convert.ToBase64String(cipherTextBytes)
                                    Return True
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                Return False
            End Try
        End Function

        Friend Function Decrypt(cipherText As String, passPhrase As String, ByRef decryptedText As String) As Boolean
            If String.IsNullOrWhiteSpace(cipherText) Then
                decryptedText = String.Empty
                Return True
            End If

            Try
                ' Get the complete stream of bytes that represent:
                ' [32 bytes of Salt] + [32 bytes of IV] + [n bytes of CipherText]
                Dim cipherTextBytesWithSaltAndIv = Convert.FromBase64String(cipherText)
                ' Get the saltbytes by extracting the first 32 bytes from the supplied cipherText bytes.
                Dim saltStringBytes = cipherTextBytesWithSaltAndIv.Take(CInt(Keysize / 8)).ToArray()
                ' Get the IV bytes by extracting the next 32 bytes from the supplied cipherText bytes.
                Dim ivStringBytes = cipherTextBytesWithSaltAndIv.Skip(CInt(Keysize / 8)).Take(CInt(Keysize / 8)).ToArray()
                ' Get the actual cipher text bytes by removing the first 64 bytes from the cipherText string.
                Dim cipherTextBytes = cipherTextBytesWithSaltAndIv.Skip(CInt((Keysize / 8) * 2)).Take(cipherTextBytesWithSaltAndIv.Length - (CInt((Keysize / 8) * 2))).ToArray()

                Using password = New Rfc2898DeriveBytes(passPhrase, saltStringBytes, DerivationIterations)
                    Dim keyBytes = password.GetBytes(CInt(Keysize / 8))
                    Using symmetricKey = New RijndaelManaged()
                        symmetricKey.BlockSize = 256
                        symmetricKey.Mode = CipherMode.CBC
                        symmetricKey.Padding = PaddingMode.PKCS7
                        Using decryptor = symmetricKey.CreateDecryptor(keyBytes, ivStringBytes)
                            Using memoryStream = New MemoryStream(cipherTextBytes)
                                Using cryptoStream = New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)
                                    Dim plainTextBytes(cipherTextBytes.Length) As Byte
                                    Dim decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)
                                    memoryStream.Close()
                                    cryptoStream.Close()
                                    decryptedText = Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount)
                                    Return True
                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                ErrorHandler.ProcessError(ex, "Error al desencriptar la cadena de texto.")
                Return False
            End Try
        End Function

        Private Function Generate256BitsOfRandomEntropy() As Byte()
            Dim randomBytes(31) As Byte ' 32 Bytes will give us 256 bits.
            Using rngCsp = New RNGCryptoServiceProvider()
                ' Fill the array with cryptographically secure random bytes.
                rngCsp.GetBytes(randomBytes)
            End Using
            Return randomBytes
        End Function
    End Module
End Namespace