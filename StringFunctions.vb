Public Class StringFunctions
    Private Function TruncateStringToLength(sStringToTrucate As String, nLength As Integer) As String
        If String.IsNullOrEmpty(sStringToTrucate) OrElse sStringToTrucate.Length <= nLength Then Return sStringToTrucate

        Dim sTruncatedString As String = Nothing
        Dim characterAtSpecifiedLength As Char = sStringToTrucate(nLength)

        If Char.IsWhiteSpace(characterAtSpecifiedLength) Then
            sTruncatedString = sStringToTrucate.Substring(0, nLength) + "..."

            Return sTruncatedString
        ElseIf characterAtSpecifiedLength = "."c Then
            Return sStringToTrucate
        Else
            Dim nIndexOfNextCharacterSpace As Integer = sStringToTrucate.IndexOf(" ", nLength - 1)
            sTruncatedString = sStringToTrucate.Substring(0, nIndexOfNextCharacterSpace) & "..."

            Return sTruncatedString
        End If
    End Function
End Class