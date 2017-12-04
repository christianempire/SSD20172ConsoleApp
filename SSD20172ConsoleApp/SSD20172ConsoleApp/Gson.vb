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
            gsonString += quote + Kvp.Key + quote + ":"
            If Kvp.Value.StartsWith("{") Or Kvp.Value.StartsWith("[") Then
                gsonString += Kvp.Value
            Else
                gsonString += quote + Kvp.Value + quote
            End If
            gsonString += ","
        Next
        Return gsonString.Substring(0, gsonString.Length - 1) + "}"
    End Function

    Public Shared Function GetArrayString(ByVal gsonArray() As Gson) As String
        Dim arrayString As String = "["
        For Each gson As Gson In gsonArray
            arrayString += gson.GetString() + ","
        Next
        Return arrayString.Substring(0, arrayString.Length - 1) + "]"
    End Function
End Class