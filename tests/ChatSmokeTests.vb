Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security.Authentication
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web.Script.Serialization

Module Program
    Private checks As Integer

    <STAThread>
    Sub Main(args As String())
        Try
            If Array.IndexOf(args, "--check-network") >= 0 Then
                CheckNetwork()
                Return
            End If
            If Array.IndexOf(args, "--inspect-access") >= 0 Then InspectAccessWindows()
            TestTransport()
            TestMarkdown()
            Application.EnableVisualStyles()
            TestSidebarAndForm()
            Console.WriteLine("PASS: " & checks.ToString() & " checks; no API calls or database edits.")
        Catch ex As Exception
            Console.Error.WriteLine(ex.ToString())
            Environment.ExitCode = 1
        End Try
    End Sub

    Private Sub CheckNetwork()
        Console.WriteLine("Credential-free GET /v1/models; no procedure text is transmitted.")
        Console.WriteLine("Process TLS defaults: " & ServicePointManager.SecurityProtocol.ToString())
        For Each protocol In New SslProtocols() {SslProtocols.None, SslProtocols.Tls, SslProtocols.Tls12}
            Using handler As New HttpClientHandler() With {.SslProtocols = protocol}
                Using client As New HttpClient(handler) With {.Timeout = TimeSpan.FromSeconds(20)}
                    Try
                        Using response = client.GetAsync("https://api.openai.com/v1/models").GetAwaiter().GetResult()
                            Console.WriteLine(protocol.ToString() & ": HTTP " & CInt(response.StatusCode).ToString())
                        End Using
                    Catch ex As Exception
                        Dim cause As Exception = ex
                        Do While cause IsNot Nothing
                            Dim networkError = TryCast(cause, WebException)
                            Console.WriteLine(protocol.ToString() & ": " & cause.GetType().Name &
                                              If(networkError Is Nothing, "", " / " & networkError.Status.ToString()))
                            cause = cause.InnerException
                        Loop
                    End Try
                End Using
            End Using
        Next
        Dim originalProtocols = ServicePointManager.SecurityProtocol
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
            Using transport As New OpenAiChatClient()
                Dim client = Field(Of HttpClient)(transport, "client")
                client.Timeout = TimeSpan.FromSeconds(20)
                Using response = client.GetAsync("https://api.openai.com/v1/models").GetAwaiter().GetResult()
                    Check(response.StatusCode = HttpStatusCode.Unauthorized, "production client reaches OpenAI despite legacy host TLS defaults")
                End Using
                Check(ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls, "production client leaves host TLS settings unchanged")
            End Using
        Finally
            ServicePointManager.SecurityProtocol = originalProtocols
        End Try
    End Sub

    Private Sub Check(condition As Boolean, name As String)
        If Not condition Then Throw New Exception("FAIL: " & name)
        checks += 1
        Console.WriteLine("PASS: " & name)
    End Sub

    Private Function Reply(content As String, Optional finishReason As String = "stop") As String
        Return New JavaScriptSerializer().Serialize(New With {.choices = New Object() {
            New With {.message = New With {.content = content}, .finish_reason = finishReason}}})
    End Function

    Private Sub TestTransport()
        For Each status In New WebExceptionStatus() {WebExceptionStatus.SecureChannelFailure, WebExceptionStatus.TrustFailure,
            WebExceptionStatus.NameResolutionFailure, WebExceptionStatus.ProxyNameResolutionFailure, WebExceptionStatus.ConnectFailure}
            Dim failure As New HttpRequestException("sensitive request", New WebException("fake-session-key", status))
            Dim explanation = OpenAiChatClient.DescribeNetworkFailure(failure)
            Check(explanation.Contains(status.ToString()) AndAlso Not explanation.Contains("fake-session-key") AndAlso
                  Not explanation.Contains("sensitive"), "safe network diagnosis: " & status.ToString())
        Next
        Dim expected = "## Result" & vbLf & "```vba" & vbLf & "MsgBox ""Hello""" & vbLf & "```"
        Dim handler As New FakeHandler With {.Body = Reply(expected)}
        Dim messages As New List(Of ChatMessage) From {
            New ChatMessage("system", "Return Markdown."),
            New ChatMessage("user", "Sub Example()" & vbCrLf & "End Sub" & vbCrLf & "Explain ""this"".")}
        Using client As New OpenAiChatClient(handler)
            Dim result = client.SendAsync(messages, "fake-session-key", "gpt-4.1-mini", CancellationToken.None).GetAwaiter().GetResult()
            Check(result = expected, "response text preserved")
            Check(handler.Endpoint = "https://api.openai.com/v1/chat/completions", "fixed HTTPS endpoint")
            Check(handler.Authorization = "Bearer fake-session-key", "per-request authorization")
            Dim payload = New JavaScriptSerializer().Deserialize(Of Dictionary(Of String, Object))(handler.RequestBody)
            Check(CStr(payload("model")) = "gpt-4.1-mini" AndAlso Not CBool(payload("store")), "model and no-storage request")
            Dim sentMessages = DirectCast(payload("messages"), System.Collections.IList)
            Check(CStr(DirectCast(sentMessages(1), Dictionary(Of String, Object))("content")) = messages(1).Content, "procedure and prompt JSON escaping")
            handler.Body = Reply("Partial", "length")
            Check(client.SendAsync(messages, "fake-session-key", "test-model", CancellationToken.None).GetAwaiter().GetResult().Contains("truncated"), "truncation warning")
            For Each status In New HttpStatusCode() {HttpStatusCode.Unauthorized, CType(429, HttpStatusCode), HttpStatusCode.InternalServerError}
                handler.Status = status
                handler.Body = "sensitive error body"
                Try
                    client.SendAsync(messages, "fake-session-key", "test-model", CancellationToken.None).GetAwaiter().GetResult()
                    Throw New Exception("Expected HTTP error")
                Catch ex As InvalidOperationException
                    Check(Not ex.Message.Contains("sensitive") AndAlso Not ex.Message.Contains("fake-session-key"), "safe HTTP " & CInt(status).ToString() & " error")
                End Try
            Next
            handler.Status = HttpStatusCode.OK
            handler.Body = Reply(Nothing)
            Try
                client.SendAsync(messages, "fake-session-key", "test-model", CancellationToken.None).GetAwaiter().GetResult()
                Throw New Exception("Expected empty response error")
            Catch ex As InvalidOperationException
                Check(ex.Message.Contains("no text"), "empty or refused response handled")
            End Try
            handler.WaitForCancellation = True
            Using cancellation As New CancellationTokenSource()
                Dim request = client.SendAsync(messages, "fake-session-key", "test-model", cancellation.Token)
                cancellation.Cancel()
                Try
                    request.GetAwaiter().GetResult()
                    Throw New Exception("Expected cancellation")
                Catch ex As OperationCanceledException
                    Check(True, "in-flight cancellation")
                End Try
            End Using
        End Using
    End Sub

    Private Sub TestMarkdown()
        Dim markdown = "# Heading" & vbLf & "**Bold** and *italic* and `inline`." & vbLf & vbLf &
            "```vba" & vbLf & "Sub Example()" & vbLf & "    MsgBox ""Hi""" & vbLf & "End Sub" & vbLf & "```" & vbLf & vbLf &
            "~~~sql" & vbLf & "SELECT * FROM Example;" & vbLf & "~~~" & vbLf & vbLf & "    indented code" & vbLf
        Dim snippets = ChatMarkdown.GetCodeBlocks(markdown)
        Check(snippets.Count = 3, "backtick, tilde and indented code blocks")
        Check(snippets(0).Contains("    MsgBox ""Hi""") AndAlso Not snippets(0).Contains("```"), "snippet indentation without fences")
        Check(ChatMarkdown.GetCodeBlocks("A plain response").Count = 0, "response without snippets")
        Check(ChatMarkdown.GetCodeBlocks("> ```vba" & vbLf & "> Debug.Print 1" & vbLf & "> ```").Count = 1, "nested code block")
        Using view As New RichTextBox()
            ChatMarkdown.AppendResponse(view, markdown & vbLf & "<script>alert('test')</script>")
            Check(view.Text.Contains("Heading") AndAlso view.Text.Contains("SELECT *"), "native Markdown rendering")
            view.Select(view.Text.IndexOf("Bold", StringComparison.Ordinal), 4)
            Check(view.SelectionFont.Bold, "bold formatting")
            view.Select(view.Text.IndexOf("Sub Example", StringComparison.Ordinal), 3)
            Check(view.SelectionFont.Name = "Consolas", "monospaced code formatting")
            Check(view.Text.Contains("<script>"), "HTML displayed as inert text")
        End Using
    End Sub

    Private Function Field(Of T)(instance As Object, name As String) As T
        Return DirectCast(instance.GetType().GetField(name, BindingFlags.Instance Or BindingFlags.NonPublic).GetValue(instance), T)
    End Function

    Private Sub InvokePrivate(instance As Object, name As String, ParamArray arguments As Object())
        instance.GetType().GetMethod(name, BindingFlags.Instance Or BindingFlags.NonPublic).Invoke(instance, arguments)
    End Sub

    Private Sub TestSidebarAndForm()
        Dim previousKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY")
        Try
            Environment.SetEnvironmentVariable("OPENAI_API_KEY", Nothing)
            Dim formType = GetType(VBACodeWriter.ProcedureChatForm)
            Using panel = DirectCast(Activator.CreateInstance(formType, New Object() {Nothing}), Form)
                Check(Field(Of TableLayoutPanel)(panel, "chatLayout").RowStyles(2).Height > 0, "missing key prompts inline")
                Dim input = Field(Of TextBox)(panel, "keyInput")
                Check(input.UseSystemPasswordChar, "masked key input")
                input.Text = "fake-session-key"
                InvokePrivate(panel, "UseKey", Nothing, EventArgs.Empty)
                Check(input.Text.Length = 0 AndAlso Field(Of String)(panel, "apiKey") = "fake-session-key", "session key retained and input cleared")
                Using editor As New Form() With {.IsMdiContainer = True, .Size = New Size(1200, 850), .StartPosition = FormStartPosition.Manual, .Location = New Point(10, 10)}
                    editor.Show()
                    Application.DoEvents()
                    Dim workspace As MdiClient = Nothing
                    For Each control As Control In editor.Controls
                        If TypeOf control Is MdiClient Then workspace = DirectCast(control, MdiClient)
                    Next
                    workspace.Dock = DockStyle.None
                    workspace.Bounds = editor.ClientRectangle
                    AddHandler editor.ClientSizeChanged, Sub() workspace.Bounds = editor.ClientRectangle
                    Dim original = workspace.Bounds
                    Using host As New VbeChatSidebar(panel)
                        host.Attach(editor.Handle)
                        Application.DoEvents()
                        Check(panel.Visible AndAlso Not panel.TopLevel, "modeless child sidebar")
                        Check(workspace.Width < original.Width AndAlso panel.Left >= workspace.Right,
                            "sidebar reserves editor space: original=" & original.ToString() & ", workspace=" & workspace.Bounds.ToString() & ", panel=" & panel.Bounds.ToString())
                        InvokePrivate(host, "UpdateLayout", Nothing, EventArgs.Empty)
                        Check(panel.Left = workspace.Right, "stable layout on repeated tick")
                        host.Detach()
                        Check(workspace.Bounds = original, "editor bounds restored on hide")
                        host.Attach(editor.Handle)
                        Check(Field(Of String)(panel, "apiKey") = "fake-session-key", "session key survives reopening")
                        editor.Size = New Size(1100, 750)
                        Application.DoEvents()
                        InvokePrivate(host, "UpdateLayout", Nothing, EventArgs.Empty)
                        Check(panel.Left = workspace.Right, "sidebar follows editor resize")
                        Dim view = Field(Of RichTextBox)(panel, "transcript")
                        ChatMarkdown.AppendText(view, "You - Example.UpdateRecord" & vbCrLf, FontStyle.Bold)
                        ChatMarkdown.AppendText(view, "How can I improve this procedure?" & vbCrLf & vbCrLf)
                        ChatMarkdown.AppendText(view, "OpenAI" & vbCrLf, FontStyle.Bold)
                        ChatMarkdown.AppendResponse(view, "## Review" & vbLf & "Use **explicit declarations** and close the recordset." & vbLf & vbLf & "```vba" & vbLf & "Dim records As DAO.Recordset" & vbLf & "Set records = CurrentDb.OpenRecordset(""Example"")" & vbLf & "```")
                        Using screenshot As New Bitmap(panel.Width, panel.Height)
                            panel.DrawToBitmap(screenshot, New Rectangle(Point.Empty, panel.Size))
                            screenshot.Save(IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "chat-sidebar.png"))
                        End Using
                        Check(Field(Of TextBox)(panel, "prompt").Height >= 60 AndAlso view.Height > 100, "usable chat layout")
                        Dim longResponse As New StringBuilder()
                        For paragraphNumber As Integer = 1 To 80
                            longResponse.AppendLine("Paragraph " & paragraphNumber.ToString() & ": Read this response from the beginning.")
                            longResponse.AppendLine()
                        Next
                        For replyNumber As Integer = 1 To 2
                            ChatMarkdown.AppendText(view, "OpenAI" & vbCrLf, FontStyle.Bold)
                            Dim responseStart = view.TextLength
                            ChatMarkdown.AppendResponse(view, longResponse.ToString())
                            Application.DoEvents()
                            Check(view.SelectionStart = responseStart AndAlso view.SelectionLength = 0,
                                  "reply " & replyNumber.ToString() & " anchors at its own beginning")
                            Dim startPosition = view.GetPositionFromCharIndex(responseStart)
                            Check(startPosition.Y >= 0 AndAlso startPosition.Y < view.Font.Height * 2,
                                  "reply " & replyNumber.ToString() & " begins at the top of the viewport")
                        Next
                        host.Detach()
                        editor.Size = New Size(450, 350)
                        Application.DoEvents()
                        host.Attach(editor.Handle)
                        Check(Not panel.Visible, "small editor collapses sidebar")
                        editor.Size = New Size(1200, 850)
                        Application.DoEvents()
                        InvokePrivate(host, "UpdateLayout", Nothing, EventArgs.Empty)
                        Check(panel.Visible, "sidebar returns when space permits")
                    End Using
                    Check(workspace.Width = editor.ClientSize.Width, "editor restored on disposal")
                    editor.Close()
                End Using
            End Using
            Environment.SetEnvironmentVariable("OPENAI_API_KEY", "fake-environment-key")
            Using panel = DirectCast(Activator.CreateInstance(formType, New Object() {Nothing}), Form)
                Check(Field(Of String)(panel, "apiKey") = "fake-environment-key" AndAlso Field(Of TableLayoutPanel)(panel, "chatLayout").RowStyles(2).Height = 0, "environment key default")
            End Using
        Finally
            Environment.SetEnvironmentVariable("OPENAI_API_KEY", previousKey)
        End Try
    End Sub

    Private Sub InspectAccessWindows()
        For Each accessProcess As Process In Process.GetProcessesByName("MSACCESS")
            Dim foundEditor As Boolean
            EnumWindows(
                Function(windowHandle, state)
                    Dim processId As Integer
                    GetWindowThreadProcessId(windowHandle, processId)
                    If processId <> accessProcess.Id Then Return True
                    Dim className As New StringBuilder(256)
                    GetClassName(windowHandle, className, className.Capacity)
                    If className.ToString().IndexOf("wndclass_desked", StringComparison.OrdinalIgnoreCase) < 0 Then Return True
                    foundEditor = True
                    Dim foundWorkspace As Boolean
                    EnumChildWindows(windowHandle,
                        Function(childHandle, childState)
                            Dim childClass As New StringBuilder(256)
                            GetClassName(childHandle, childClass, childClass.Capacity)
                            If childClass.ToString() = "MDIClient" Then foundWorkspace = True
                            Return True
                        End Function, IntPtr.Zero)
                    Console.WriteLine("Access VBE window found; MDI workspace: " & foundWorkspace.ToString())
                    Return True
                End Function, IntPtr.Zero)
            If Not foundEditor Then Console.WriteLine("Access is running; open its VBA editor for a live workspace check.")
        Next
    End Sub

    Private Delegate Function WindowCallback(windowHandle As IntPtr, state As IntPtr) As Boolean
    <DllImport("user32.dll")>
    Private Function EnumWindows(callback As WindowCallback, state As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")>
    Private Function EnumChildWindows(parent As IntPtr, callback As WindowCallback, state As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Function GetClassName(windowHandle As IntPtr, className As StringBuilder, count As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Private Function GetWindowThreadProcessId(windowHandle As IntPtr, ByRef processId As Integer) As Integer
    End Function
End Module

Friend Class FakeHandler
    Inherits HttpMessageHandler

    Public Property Body As String
    Public Property Status As HttpStatusCode = HttpStatusCode.OK
    Public Property RequestBody As String
    Public Property Endpoint As String
    Public Property Authorization As String
    Public Property WaitForCancellation As Boolean

    Protected Overrides Async Function SendAsync(request As HttpRequestMessage, cancellationToken As CancellationToken) As Task(Of HttpResponseMessage)
        RequestBody = Await request.Content.ReadAsStringAsync().ConfigureAwait(False)
        Endpoint = request.RequestUri.AbsoluteUri
        Authorization = request.Headers.Authorization.ToString()
        If WaitForCancellation Then Await Task.Delay(Timeout.Infinite, cancellationToken).ConfigureAwait(False)
        Return New HttpResponseMessage(Status) With {.Content = New StringContent(Body)}
    End Function
End Class