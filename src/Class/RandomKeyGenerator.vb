Imports System.Security.Cryptography
Imports System.Text

Public Class RandomKeyGenerator

#Region " Declare"
    Private Const lowerLetters As String = "abcdefghijklmnopqrstuvwxyz"
    Private Const upperLetters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Private Const numbers As String = "1234567890"
    Private Const specialChars As String = "!#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
#End Region

#Region " Random string number"
    Public Shared Function RandomStringNumber(ByVal length As Integer) As String
        Const valid As String = lowerLetters & upperLetters
        Return RngGenerate(valid, length)
    End Function
#End Region

#Region " Random string number special char"
    Public Shared Function RandomStringNumberSpecialChar(ByVal length As Integer) As String
        Const valid As String = lowerLetters & upperLetters & numbers & specialChars
        Return RngGenerate(valid, length)
    End Function
#End Region

#Region " Random"
    Public Shared Function RandomStringParams(ByVal length As Integer) As String
        Dim valid As String = Nothing

        If Main.cbLowercase.Checked Then
            valid &= lowerLetters
        End If
        If Main.cbUppercase.Checked Then
            valid &= upperLetters
        End If
        If Main.cbNumbers.Checked Then
            valid &= numbers
        End If
        If Main.cbSpecialChars.Checked Then
            valid &= specialChars
        End If

        If valid Is Nothing Then Return ""

        Return RngGenerate(valid, length)

    End Function
#End Region

#Region " RNG Generate"
    Public Shared Function RngGenerate(ByVal valid As String, ByVal length As Integer) As String
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