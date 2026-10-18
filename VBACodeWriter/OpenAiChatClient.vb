Imports System.Collections.Generic
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Security.Authentication
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web.Script.Serialization

Friend Class ChatMessage
    Public Property Role As String
    Public Property Content As String

    Public Sub New(role As String, content As String)
        Me.Role = role
        Me.Content = content
    End Sub
End Class

Friend NotInheritable Class OpenAiChatClient
    Implements IDisposable

    Private ReadOnly client As HttpClient

    Public Sub New(Optional handler As HttpMessageHandler = Nothing)
        If handler Is Nothing Then handler = New HttpClientHandler() With {.SslProtocols = SslProtocols.Tls12}
        client = New HttpClient(handler)
        client.Timeout = TimeSpan.FromSeconds(120)
        client.MaxResponseContentBufferSize = 4 * 1024 * 1024
    End Sub

    Public Async Function SendAsync(messages As IEnumerable(Of ChatMessage), apiKey As String, model As String,
                                    cancellationToken As CancellationToken) As Task(Of String)
        If String.IsNullOrWhiteSpace(apiKey) Then Throw New InvalidOperationException("Enter an OpenAI API key.")
        If String.IsNullOrWhiteSpace(model) Then Throw New InvalidOperationException("Enter a chat-compatible OpenAI model.")

        Dim serializer As New JavaScriptSerializer() With {.MaxJsonLength = 4 * 1024 * 1024}
        Dim payloadMessages As New List(Of Object)
        For Each message In messages
            payloadMessages.Add(New With {.role = message.Role, .content = message.Content})
        Next
        Dim payload = serializer.Serialize(New With {
            .model = model.Trim(), .messages = payloadMessages, .max_completion_tokens = 4096, .store = False})

        Using request As New HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions")
            request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey.Trim())
            request.Content = New StringContent(payload, Encoding.UTF8, "application/json")
            Using response = Await client.SendAsync(request, cancellationToken).ConfigureAwait(False)
                If Not response.IsSuccessStatusCode Then
                    Select Case CInt(response.StatusCode)
                        Case 401, 403
                            Throw New InvalidOperationException("OpenAI rejected the credentials or access. Check the API key and model permissions.")
                        Case 429
                            Throw New InvalidOperationException("OpenAI rate or quota limit reached. Check billing or try again later.")
                        Case Else
                            Throw New InvalidOperationException("OpenAI returned HTTP " & CInt(response.StatusCode).ToString() & ". Check the model and request size, or try again later.")
                    End Select
                End If

                Dim body = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                cancellationToken.ThrowIfCancellationRequested()
                Dim root = serializer.Deserialize(Of Dictionary(Of String, Object))(body)
                If root Is Nothing OrElse Not root.ContainsKey("choices") Then
                    Throw New InvalidOperationException("OpenAI returned an unexpected response.")
                End If
                Dim choices = TryCast(root("choices"), System.Collections.IList)
                If choices Is Nothing OrElse choices.Count = 0 Then Throw New InvalidOperationException("OpenAI returned no response.")
                Dim choice = DirectCast(choices(0), Dictionary(Of String, Object))
                Dim message = DirectCast(choice("message"), Dictionary(Of String, Object))
                Dim content As String = If(message.ContainsKey("content"), TryCast(message("content"), String), Nothing)
                If String.IsNullOrWhiteSpace(content) Then
                    Throw New InvalidOperationException("OpenAI returned no text. The request may have been refused; revise it and try again.")
                End If
                If choice.ContainsKey("finish_reason") AndAlso CStr(choice("finish_reason")) = "length" Then
                    content &= vbCrLf & vbCrLf & "**Response truncated by the output limit. Ask for a smaller section before using the code.**"
                End If
                Return content
            End Using
        End Using
    End Function

    Public Shared Function DescribeNetworkFailure(failure As Exception) As String
        Dim cause = failure
        Do While cause IsNot Nothing
            Dim networkError = TryCast(cause, WebException)
            If networkError IsNot Nothing Then
                Select Case networkError.Status
                    Case WebExceptionStatus.SecureChannelFailure
                        Return "OpenAI TLS handshake failed (SecureChannelFailure). Check TLS 1.2 support and HTTPS inspection."
                    Case WebExceptionStatus.TrustFailure
                        Return "OpenAI certificate validation failed (TrustFailure). Check Windows certificate trust and HTTPS inspection."
                    Case WebExceptionStatus.NameResolutionFailure
                        Return "Unable to resolve api.openai.com (NameResolutionFailure). Check DNS and network access."
                    Case WebExceptionStatus.ProxyNameResolutionFailure
                        Return "Unable to resolve the proxy (ProxyNameResolutionFailure). Check the system proxy settings."
                    Case WebExceptionStatus.ConnectFailure
                        Return "Unable to connect to OpenAI (ConnectFailure). Check firewall and proxy settings."
                    Case WebExceptionStatus.Timeout
                        Return "The OpenAI connection timed out. Check the network and try again."
                    Case Else
                        Return "OpenAI network error (" & networkError.Status.ToString() & "). Check the connection and proxy settings."
                End Select
            End If
            If TypeOf cause Is AuthenticationException Then
                Return "OpenAI TLS authentication failed. Check certificate trust and HTTPS inspection."
            End If
            cause = cause.InnerException
        Loop
        Return "Unable to reach OpenAI. Check the connection and proxy settings."
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        client.Dispose()
    End Sub
End Class