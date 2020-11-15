Public Class RandomKeyGenerator

#Region " Random string number"
    Public Shared Function RandomStringNumber(ByVal length As Integer) As String
        Dim sb As New System.Text.StringBuilder
        Dim chars() As String = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z",
               "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "X",
               "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"}
        Dim upperBound As Integer = UBound(chars)

        For x As Integer = 1 To length
            sb.Append(chars(CInt(Int(Rnd() * upperBound))))
        Next

        Return sb.ToString
    End Function
#End Region

#Region " Random string number special char"
    Public Shared Function RandomStringNumberSpecialChar(ByVal length As Integer) As String
        Dim sb As New System.Text.StringBuilder
        Dim chars() As String = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z",
               "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "X",
               "1", "2", "3", "4", "5", "6", "7", "8", "9", "0",
               "!", "§", "$", "%", "&", "/", "@", "(", ")", "=", "?", "*", "+", "~", "#", "'", "-", "_", "<", ">", "|", "^", "°"}
        Dim upperBound As Integer = UBound(chars)

        For x As Integer = 1 To length
            sb.Append(chars(CInt(Int(Rnd() * upperBound))))
        Next

        Return sb.ToString
    End Function
#End Region

#Region " Random"
    Public Shared Function RandomStringParams(ByVal length As Integer) As String
        Try

            Dim sb As New System.Text.StringBuilder
            Dim chars() As Char = Nothing
            Dim charString As String = Nothing

            If Main.cbLowercase.Checked Then
                charString &= "a" & "b" & "c" & "d" & "e" & "f" & "g" & "h" & "i" & "j" & "k" & "l" & "m" & "n" & "o" & "p" & "q" & "r" & "s" & "t" & "u" & "v" & "w" & "x" & "y" & "z"
            End If
            If Main.cbUppercase.Checked Then
                charString &= "A" & "B" & "C" & "D" & "E" & "F" & "G" & "H" & "I" & "J" & "K" & "L" & "M" & "N" & "O" & "P" & "Q" & "R" & "S" & "T" & "U" & "V" & "W" & "X" & "Y" & "X" & "Y" & "Z"
            End If
            If Main.cbNumbers.Checked Then
                charString &= "1" & "2" & "3" & "4" & "5" & "6" & "7" & "8" & "9" & "0"
            End If
            If Main.cbSpecialChars.Checked Then
                charString &= "|" & "!" & "£" & "$" & "%" & "&" & "/" & "(" & ")" & "=" & "?" & "*" & "[" & "]" & "{" & "}" & ";" & ":" & "-" & "@" & "<" & ">" & "^" & "°" & "à" & "ò" & "é" & "§" & "ù" & "è" & "+"
            End If

            If charString Is Nothing Then Return ""

            chars = charString.ToCharArray

            Dim upperBound As Integer = UBound(chars)

            For x As Integer = 1 To length
                sb.Append(chars(CInt(Int(Rnd() * upperBound))))
            Next
            Return sb.ToString

        Catch ex As Exception
            Return Nothing
        End Try

    End Function
#End Region

End Class