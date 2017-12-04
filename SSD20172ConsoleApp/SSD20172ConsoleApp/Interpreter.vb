Public Class Interpreter
    Public Shared Function GetValueFromModelRequestString(ByVal modelRequestString As String, ByVal attribute As String, Optional ByVal isString As Boolean = True) As String
        Dim props() As String = modelRequestString.Split(",")
        Dim prop As String = ""
        Dim quote As String = Chr(34)
        Dim found As Boolean = False
        Dim index As Integer = 0

        While Not found And index < props.Length
            If props(index).Contains(quote + attribute + quote) Then
                prop = props(index)
                found = True
            Else
                index += 1
            End If
        End While
        If found Then
            Dim value As String = prop.Split(":")(1)
            If isString Then
                Return value.Substring(1, value.Length - 2)
            Else
                Return value
            End If
        Else
            Return ""
        End If
    End Function

    Public Shared Function GetValueFromModelOutputString(ByVal modelOutputString As String, ByVal identifier As String, ByVal type As String) As Decimal
        Dim searchIndex As Integer = modelOutputString.IndexOf(identifier) + identifier.Length
        Dim valueLength As Integer = 0
        Dim columnNumber As Integer = 0
        If type = "Average" Then
            columnNumber = 1
        ElseIf type = "Minimum" Then
            columnNumber = 3
        ElseIf type = "Maximum" Then
            columnNumber = 4
        End If
        If columnNumber = 0 Then
            Return 0
        Else
            For columnIndex = 1 To columnNumber
                While modelOutputString.Chars(searchIndex) = " "
                    searchIndex += 1
                End While
                If columnIndex = columnNumber Then
                    While modelOutputString.Chars(searchIndex + valueLength) <> " "
                        valueLength += 1
                    End While
                Else
                    While modelOutputString.Chars(searchIndex) <> " "
                        searchIndex += 1
                    End While
                End If
            Next
            Return Decimal.Parse(modelOutputString.Substring(searchIndex, valueLength))
        End If
    End Function
End Class
