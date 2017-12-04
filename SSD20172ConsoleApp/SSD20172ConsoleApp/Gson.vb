Public Class Gson
    Private Attributes As Dictionary(Of String, String)

    Public Sub New()
        Me.Attributes = New Dictionary(Of String, String)
    End Sub

    Public Sub AddAttribute(ByVal Attribute As String, ByVal Value As String)
        Me.Attributes.Add(Attribute, Value)
    End Sub

    Public Function GetString() As String
        Dim quote As String = Chr(34)
        Dim gsonString As String = "{"
        For Each Kvp As KeyValuePair(Of String, String) In Me.Attributes
            gsonString += quote + Kvp.Key + quote + ":" + quote + Kvp.Value + quote + ","
        Next
        gsonString = gsonString.Substring(0, gsonString.Length - 1) + "}"
        Return gsonString
    End Function
End Class