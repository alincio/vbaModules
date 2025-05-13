' How to use: =ConvertCurrency("USD", "GBP", 10, "2025-01-01")

Function ConvertCurrency(fromCurrency As String, toCurrency As String, amount As Double, Optional dateValue As String = "") As Variant
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim apiKey As String
    Dim startPos As Long, endPos As Long, valueText As String

    ' enter your key
    apiKey = "YOUR KEY"

    ' Build URL
    url = "https://api.exchangerate.host/convert?access_key=" & apiKey & _
          "&from=" & fromCurrency & "&to=" & toCurrency & _
          "&amount=" & amount

    If dateValue <> "" Then
        url = url & "&date=" & dateValue
    End If

    ' Send request
    Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo ConnectionError
    http.Open "GET", url, False
    http.Send
    response = http.ResponseText

    ' export result
    startPos = InStr(response, """result"":") + 9
    endPos = InStr(startPos, response, "}")

    If startPos > 0 And endPos > startPos Then
        valueText = Mid(response, startPos, endPos - startPos)
        If IsNumeric(valueText) Then
            ConvertCurrency = CDbl(valueText)
        Else
            ConvertCurrency = "Invalid result"
        End If
    Else
        ConvertCurrency = "Cannot export the result"
    End If
    Exit Function

ConnectionError:
    ConvertCurrency = "Connection error"
End Function




