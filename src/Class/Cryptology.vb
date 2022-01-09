Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class Cryptology

#Region " Declare"
    Private Const Salt As String = "d5fg4df5sg4ds5fg45sdfg4"
    Private Const SizeOfBuffer As Integer = 1024 * 8

    Public Shared Encoded As Boolean = False
    Public Shared Decoded As Boolean = False
#End Region

#Region " Enc|File"
    Friend Shared Sub EncryptFile(inputPath As String, outputPath As String, password As String)
        Dim input = New FileStream(inputPath, FileMode.Open, FileAccess.Read)
        Dim output = New FileStream(outputPath, FileMode.OpenOrCreate, FileAccess.Write)

        ' Essentially, if you want to use RijndaelManaged as AES you need to make sure that:
        ' 1.The block size is set to 128 bits
        ' 2.You are not using CFB mode, or if you are the feedback size is also 128 bits

        Dim key = New Rfc2898DeriveBytes(password, Encoding.ASCII.GetBytes(Salt))

        Using algorithm = New RijndaelManaged()
            algorithm.KeySize = 256 'Can be 128, 192, or 256.
            algorithm.BlockSize = 128
            algorithm.Key = key.GetBytes(algorithm.KeySize \ 8)
            algorithm.IV = key.GetBytes(algorithm.BlockSize \ 8)
            Try
                Using encryptedStream = New CryptoStream(output, algorithm.CreateEncryptor(), CryptoStreamMode.Write)
                    CopyStream(input, encryptedStream)
                End Using
                Encoded = True
            Catch ex As Exception
                Encoded = False
                MsgBox(ex.Message)
                Exit Sub
            End Try
        End Using
    End Sub
#End Region

#Region " Dec|File"
    Friend Shared Sub DecryptFile(inputPath As String, outputPath As String, password As String)
        Dim input = New FileStream(inputPath, FileMode.Open, FileAccess.Read)
        Dim output = New FileStream(outputPath, FileMode.OpenOrCreate, FileAccess.Write)

        ' Essentially, if you want to use RijndaelManaged as AES you need to make sure that:
        ' 1.The block size is set to 128 bits
        ' 2.You are not using CFB mode, or if you are the feedback size is also 128 bits

        Dim key = New Rfc2898DeriveBytes(password, Encoding.ASCII.GetBytes(Salt))

        Using algorithm = New RijndaelManaged()
            algorithm.KeySize = 256 'Can be 128, 192, or 256.
            algorithm.BlockSize = 128
            algorithm.Key = key.GetBytes(algorithm.KeySize \ 8)
            algorithm.IV = key.GetBytes(algorithm.BlockSize \ 8)
            Try
                Using decryptedStream = New CryptoStream(output, algorithm.CreateDecryptor(), CryptoStreamMode.Write)
                    CopyStream(input, decryptedStream)
                End Using
                Decoded = True
            Catch generatedExceptionName As CryptographicException
                Decoded = False
                MsgBox("Please suppy a correct password", MsgBoxStyle.Critical)
                Exit Sub
            Catch ex As Exception
                Decoded = False
                MsgBox(ex.Message)
                Exit Sub
            End Try
        End Using
    End Sub
#End Region

#Region " Copy Stream"
    Private Shared Sub CopyStream(input As Stream, output As Stream)
        Using output
            Using input
                Dim buffer As Byte() = New Byte(SizeOfBuffer - 1) {}
                Dim read As Integer
                While (InlineAssignHelper(read, input.Read(buffer, 0, buffer.Length))) > 0
                    output.Write(buffer, 0, read)
                End While
            End Using
        End Using
    End Sub
    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function
#End Region

#Region " Enc|Text"
    Friend Shared Function RijndaelDecrypt(ByVal UDecryptU As String, ByVal UKeyU As String) As String
        Dim XoAesProviderX As New RijndaelManaged
        Dim XbtCipherX() As Byte
        Dim XbtSaltX() As Byte = New Byte() {1, 2, 3, 4, 5, 6, 7, 8}
        Dim XoKeyGeneratorX As New Rfc2898DeriveBytes(UKeyU, XbtSaltX)
        XoAesProviderX.Key = XoKeyGeneratorX.GetBytes(XoAesProviderX.Key.Length)
        XoAesProviderX.IV = XoKeyGeneratorX.GetBytes(XoAesProviderX.IV.Length)
        Dim XmsX As New IO.MemoryStream
        Dim XcsX As New CryptoStream(XmsX, XoAesProviderX.CreateDecryptor(),
          CryptoStreamMode.Write)
        Try
            XbtCipherX = Convert.FromBase64String(UDecryptU)
            XcsX.Write(XbtCipherX, 0, XbtCipherX.Length)
            XcsX.Close()
            UDecryptU = Encoding.UTF8.GetString(XmsX.ToArray)
        Catch
        End Try
        Return UDecryptU
    End Function
#End Region

#Region " Dec|Text"
    Friend Shared Function Rijndaelcrypt(ByVal File As String, ByVal Key As String) As String
        Dim oAesProvider As New RijndaelManaged
        Dim btClear() As Byte
        Dim btSalt() As Byte = New Byte() {1, 2, 3, 4, 5, 6, 7, 8}
        Dim oKeyGenerator As New Rfc2898DeriveBytes(Key, btSalt)
        oAesProvider.Key = oKeyGenerator.GetBytes(oAesProvider.Key.Length)
        oAesProvider.IV = oKeyGenerator.GetBytes(oAesProvider.IV.Length)
        Dim ms As New IO.MemoryStream
        Dim cs As New CryptoStream(ms,
          oAesProvider.CreateEncryptor(),
          CryptoStreamMode.Write)
        btClear = Encoding.UTF8.GetBytes(File)
        cs.Write(btClear, 0, btClear.Length)
        cs.Close()
        File = Convert.ToBase64String(ms.ToArray)
        Return File
    End Function
#End Region

End Class
