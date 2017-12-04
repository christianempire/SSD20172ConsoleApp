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

    Public Shared Function GetValueFromModelOutputString(ByVal identifier As String, ByVal type As String)
        Return ""
    End Function
End Class
