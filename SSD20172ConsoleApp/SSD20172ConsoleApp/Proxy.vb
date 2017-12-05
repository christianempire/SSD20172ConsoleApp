Public Class Proxy
    'Const ApiMiddlewareRequestUrl As String = "http://localhost:60650/middleware/request"
    Const ApiMiddlewareRequestUrl As String = "http://52.226.131.0/api-arena/middleware/request"

    Public Shared Function GetModelRequest() As String
        Try
            Dim webClient As New Net.WebClient
            Return webClient.DownloadString(ApiMiddlewareRequestUrl)
        Catch ex As Exception
            Return "Error downloading the request"
        End Try
    End Function

    Public Shared Function PostModelRequest(ByVal payload As String) As String
        Try
            Dim webClient As New Net.WebClient
            webClient.Headers.Add("Content-Type", "application/json")
            Return webClient.UploadString(ApiMiddlewareRequestUrl, "POST", payload)
        Catch ex As Exception
            Return "Error uploading the results"
        End Try
    End Function
End Class
