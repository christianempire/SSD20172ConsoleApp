Public Class Proxy
    Const ApiMiddlewareRequestUrl As String = "http://localhost:60650/middleware/request"
    'Const ApiMiddlewareRequestUrl As String = "http://52.226.131.0/api-arena/middleware/request"

    Public Shared Function GetModelRequest() As String
        Dim webClient As New Net.WebClient
        Return webClient.DownloadString(ApiMiddlewareRequestUrl)
    End Function

    Public Shared Function PostModelRequest(ByVal payload As String) As String
        Dim webClient As New Net.WebClient
        webClient.Headers.Add("Content-Type", "application/json")
        Return webClient.UploadString(ApiMiddlewareRequestUrl, "POST", payload)
    End Function
End Class
