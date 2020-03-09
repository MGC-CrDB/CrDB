Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

' http://aspsnippets.com/Articles/Encrypt-and-Decrypt-Username-or-Password-stored-in-database-in-ASPNet-using-C-and-VBNet.aspx
' http://msdn.microsoft.com/de-de/library/ms172831.aspx

Public Class CryptClass
    Protected CryptionKey As String
    Protected TripleDes As New TripleDESCryptoServiceProvider

    Public Sub New()
        Me.CryptionKey = "MAKV2SPBNI99212"
        Me.TripleDes.Key = TruncateHash(Me.CryptionKey, TripleDes.KeySize \ 8)
        Me.TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

    Sub New(ByVal CryptKey As String)
        ' Initialize the crypto provider.
        Me.CryptionKey = CryptKey
        Me.TripleDes.Key = TruncateHash(Me.CryptionKey, TripleDes.KeySize \ 8)
        Me.TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

    Public Function Encrypt(clearText As String) As String
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(Me.CryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, &H65, &H64, &H76, &H65, &H64, &H65, &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function

    Public Function Decrypt(cipherText As String) As String
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(Me.CryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function

    Public Function EncryptData(ByVal PlainText As String) As String
        ' Convert the plaintext string to a byte array.
        Dim PlainTextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(PlainText)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(PlainTextBytes, 0, PlainTextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        Return Convert.ToBase64String(ms.ToArray)
    End Function

    Public Function DecryptData(ByVal EncryptedText As String) As String
        ' Convert the encrypted text string to a byte array.
        Dim EncryptedBytes() As Byte = Convert.FromBase64String(EncryptedText)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream.
        Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        decStream.Write(EncryptedBytes, 0, EncryptedBytes.Length)
        decStream.FlushFinalBlock()

        ' Convert the plaintext stream to a string.
        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function

    Protected Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()
        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim KeyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
        Dim Hash() As Byte = sha1.ComputeHash(KeyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve Hash(length - 1)
        Return Hash
    End Function
End Class
