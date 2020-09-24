Imports System.Security.Cryptography
Public NotInheritable Class Cipher
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Private TripleDes As New TripleDESCryptoServiceProvider
    Private Key As Byte() = {&H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &H10, &H11, &H12, &H13, &H14, &H15, &H16}
    Private IV As Byte() = {&H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &H10, &H11, &H12, &H13, &H14, &H15, &H16}
    Sub New()
    End Sub
    Sub CatHash(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)

    End Sub
    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()

        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim keyBytes() As Byte =
        System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function

    Public Function Encode(ByVal plainText As String) As String

        On Error GoTo Err

        If plainText = String.Empty Then
            Throw New ArgumentNullException(plainText, "The value specified for the plainText parameter cannot be null")
        End If

        'convert the plain text string to byte array
        Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plainText)

        'create the stream
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.

        CatHash(System.Text.Encoding.Unicode.GetString(Key))
        Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
        'use the cryp stream to write the bype array to the Stream
        encStream.Write(plaintextBytes, 0, plaintextBytes.Count)
        encStream.FlushFinalBlock()

        Return Convert.ToBase64String(ms.ToArray)
Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function


    Public Function Decode(ByVal encryptedText As String) As String



        If encryptedText = String.Empty Then
            Throw New ArgumentNullException(encryptedText, "The value specified for the plainText parameter cannot be null")
        End If

        'convert the plain text string to byte array
        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedText)

        'create the stream
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream.
        CatHash(System.Text.Encoding.Unicode.GetString(Key))
        Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

        'use the cryp stream to write the bype array to the Stream
        decStream.Write(encryptedBytes, 0, encryptedBytes.Count)
        decStream.FlushFinalBlock()

        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function

End Class

