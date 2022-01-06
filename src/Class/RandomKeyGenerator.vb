Imports System.Security.Cryptography
Imports System.Text

Public Class RandomKeyGenerator

#Region " Random string number"
    Public Shared Function RandomStringNumber(ByVal length As Integer) As String
        Const valid As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
        Dim res As New StringBuilder()

        Using rng As New RNGCryptoServiceProvider()
            Dim uintBuffer As Byte() = New Byte(3) {}

            While Math.Max(Threading.Interlocked.Decrement(length), length + 1) > 0
                rng.GetBytes(uintBuffer)
                Dim num As UInteger = BitConverter.ToUInt32(uintBuffer, 0)
                res.Append(valid(CInt((num Mod CUInt(valid.Length)))))
            End While
        End Using

        Return res.ToString()
    End Function
#End Region

#Region " Random string number special char"
    Public Shared Function RandomStringNumberSpecialChar(ByVal length As Integer) As String
        Const valid As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!§$%&/@()=?*+~#'-_<>|^°"
        Dim res As New StringBuilder()

        Using rng As New RNGCryptoServiceProvider()
            Dim uintBuffer As Byte() = New Byte(3) {}

            While Math.Max(Threading.Interlocked.Decrement(length), length + 1) > 0
                rng.GetBytes(uintBuffer)
                Dim num As UInteger = BitConverter.ToUInt32(uintBuffer, 0)
                res.Append(valid(CInt((num Mod CUInt(valid.Length)))))
            End While
        End Using

        Return res.ToString()
    End Function
#End Region

#Region " Random"
    Public Shared Function RandomStringParams(ByVal length As Integer) As String
        Dim valid As String = Nothing

        If Main.cbLowercase.Checked Then
            valid &= "abcdefghijklmnopqrstuvwxyz"
        End If
        If Main.cbUppercase.Checked Then
            valid &= "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        End If
        If Main.cbNumbers.Checked Then
            valid &= "1234567890"
        End If
        If Main.cbSpecialChars.Checked Then
            valid &= "!#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
        End If

        If valid Is Nothing Then Return ""

        Dim res As New StringBuilder()
        Using rng As New RNGCryptoServiceProvider()
            Dim uintBuffer As Byte() = New Byte(3) {}

            While Math.Max(Threading.Interlocked.Decrement(length), length + 1) > 0
                rng.GetBytes(uintBuffer)
                Dim num As UInteger = BitConverter.ToUInt32(uintBuffer, 0)
                res.Append(valid(CInt((num Mod CUInt(valid.Length)))))
            End While
        End Using

        Return res.ToString()
    End Function
#End Region

End Class