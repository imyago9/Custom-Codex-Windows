Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks

Namespace CodexNativeAgent.AppServer
    Public NotInheritable Class RpcServerRequest
        Public Sub New(id As JsonNode, methodName As String, paramsNode As JsonNode)
            Me.Id = id
            Me.MethodName = methodName
            Me.ParamsNode = paramsNode
        End Sub

        Public ReadOnly Property Id As JsonNode
        Public ReadOnly Property MethodName As String
        Public ReadOnly Property ParamsNode As JsonNode
    End Class

    Public NotInheritable Class RpcErrorException
        Inherits Exception

        Public Sub New(code As Integer, message As String, dataNode As JsonNode)
            MyBase.New(message)
            Me.Code = code
            Me.DataNode = dataNode
        End Sub

        Public ReadOnly Property Code As Integer
        Public ReadOnly Property DataNode As JsonNode
    End Class

    Public NotInheritable Class CodexAppServerClient
        Implements IAsyncDisposable

        Private ReadOnly _pendingRequests As New ConcurrentDictionary(Of String, TaskCompletionSource(Of JsonNode))(StringComparer.Ordinal)
        Private ReadOnly _writeLock As New SemaphoreSlim(1, 1)
        Private ReadOnly _jsonWriteOptions As New JsonSerializerOptions With {
            .WriteIndented = False
        }

        Private _process As Process
        Private _stdin As StreamWriter
        Private _stdoutTask As Task
        Private _stderrTask As Task
        Private _cts As CancellationTokenSource
        Private _nextRequestId As Long
        Private _stopGate As Integer
        Private _defaultRequestTimeout As TimeSpan = TimeSpan.FromMinutes(2)

        Public Event RawMessage(direction As String, payload As String)
        Public Event NotificationReceived(methodName As String, paramsNode As JsonNode)
        Public Event ServerRequestReceived(request As RpcServerRequest)
        Public Event Disconnected(reason As String)

        Public ReadOnly Property IsRunning As Boolean
            Get
                Return _process IsNot Nothing AndAlso Not _process.HasExited
            End Get
        End Property

        Public Property DefaultRequestTimeout As TimeSpan
            Get
                Return _defaultRequestTimeout
            End Get
            Set(value As TimeSpan)
                If value <= TimeSpan.Zero Then
                    Throw New ArgumentOutOfRangeException(NameOf(value), "Default request timeout must be positive.")
                End If

                _defaultRequestTimeout = value
            End Set
        End Property

        Public ReadOnly Property ProcessId As Integer
            Get
                If _process Is Nothing Then
                    Return 0
                End If

                Try
                    Return _process.Id
                Catch
                    Return 0
                End Try
            End Get
        End Property

        Public Async Function StartAsync(executablePath As String,
                                         arguments As String,
                                         workingDirectory As String,
                                         cancellationToken As CancellationToken,
                                         Optional environmentVariables As IDictionary(Of String, String) = Nothing) As Task
            If IsRunning Then
                Throw New InvalidOperationException("Codex App Server is already running.")
            End If

            Dim startInfo As New ProcessStartInfo() With {
                .FileName = executablePath,
                .Arguments = arguments,
                .UseShellExecute = False,
                .RedirectStandardInput = True,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True,
                .StandardOutputEncoding = Encoding.UTF8,
                .StandardErrorEncoding = Encoding.UTF8
            }

            If Not String.IsNullOrWhiteSpace(workingDirectory) Then
                startInfo.WorkingDirectory = workingDirectory
            End If

            If environmentVariables IsNot Nothing Then
                For Each kvp In environmentVariables
                    If String.IsNullOrWhiteSpace(kvp.Key) Then
                        Continue For
                    End If

                    startInfo.EnvironmentVariables(kvp.Key) = If(kvp.Value, String.Empty)
                Next
            End If

            _process = New Process() With {
                .StartInfo = startInfo,
                .EnableRaisingEvents = True
            }

            If Not _process.Start() Then
                Throw New InvalidOperationException("Failed to start Codex App Server process.")
            End If

            _stdin = _process.StandardInput
            _stdin.AutoFlush = True
            _cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
            _stopGate = 0

            Dim token = _cts.Token
            _stdoutTask = ReadStdoutLoopAsync(token)
            _stderrTask = ReadStderrLoopAsync(token)

            Await Task.CompletedTask.ConfigureAwait(False)
        End Function

        Public Async Function InitializeAsync(clientName As String, clientVersion As String, cancellationToken As CancellationToken) As Task
            Dim paramsNode As New JsonObject()
            Dim clientInfo As New JsonObject()
            clientInfo("name") = clientName
            clientInfo("version") = clientVersion
            paramsNode("clientInfo") = clientInfo

            Dim capabilities As New JsonObject()
            capabilities("experimentalApi") = True
            paramsNode("capabilities") = capabilities

            Await SendRequestAsync("initialize", paramsNode, TimeSpan.FromSeconds(30), cancellationToken).ConfigureAwait(False)
            Await SendNotificationAsync("initialized", Nothing, cancellationToken).ConfigureAwait(False)
        End Function

        Public Async Function SendRequestAsync(methodName As String,
                                               Optional paramsNode As JsonNode = Nothing,
                                               Optional timeout As TimeSpan? = Nothing,
                                               Optional cancellationToken As CancellationToken = Nothing) As Task(Of JsonNode)
            If String.IsNullOrWhiteSpace(methodName) Then
                Throw New ArgumentException("Method name is required.", NameOf(methodName))
            End If

            EnsureRunning()

            Dim requestId As Long = Interlocked.Increment(_nextRequestId)
            Dim requestKey As String = requestId.ToString(CultureInfo.InvariantCulture)
            Dim tcs As New TaskCompletionSource(Of JsonNode)(TaskCreationOptions.RunContinuationsAsynchronously)

            If Not _pendingRequests.TryAdd(requestKey, tcs) Then
                Throw New InvalidOperationException($"Could not register request id {requestKey}.")
            End If

            Dim requestMessage As New JsonObject()
            requestMessage("id") = requestId
            requestMessage("method") = methodName

            If paramsNode IsNot Nothing Then
                requestMessage("params") = CloneJson(paramsNode)
            End If

            Dim cancellationRegistration As CancellationTokenRegistration = Nothing
            Try
                If cancellationToken.CanBeCanceled Then
                    cancellationRegistration = cancellationToken.Register(
                        Sub()
                            Dim removed As TaskCompletionSource(Of JsonNode) = Nothing
                            If _pendingRequests.TryRemove(requestKey, removed) Then
                                removed.TrySetCanceled(cancellationToken)
                            End If
                        End Sub)
                End If

                Await SendMessageAsync(requestMessage, cancellationToken).ConfigureAwait(False)

                Dim timeoutValue As TimeSpan = If(timeout.HasValue, timeout.Value, _defaultRequestTimeout)
                Dim timeoutTask = Task.Delay(timeoutValue, cancellationToken)
                Dim completedTask = Await Task.WhenAny(tcs.Task, timeoutTask).ConfigureAwait(False)

                If completedTask Is timeoutTask Then
                    Dim removed As TaskCompletionSource(Of JsonNode) = Nothing
                    _pendingRequests.TryRemove(requestKey, removed)
                    Throw New TimeoutException($"Timed out waiting for response to '{methodName}'.")
                End If

                Return Await tcs.Task.ConfigureAwait(False)
            Catch
                Dim removed As TaskCompletionSource(Of JsonNode) = Nothing
                _pendingRequests.TryRemove(requestKey, removed)
                Throw
            Finally
                cancellationRegistration.Dispose()
            End Try
        End Function

        Public Async Function SendNotificationAsync(methodName As String,
                                                    Optional paramsNode As JsonNode = Nothing,
                                                    Optional cancellationToken As CancellationToken = Nothing) As Task
            If String.IsNullOrWhiteSpace(methodName) Then
                Throw New ArgumentException("Method name is required.", NameOf(methodName))
            End If

            EnsureRunning()

            Dim notification As New JsonObject()
            notification("method") = methodName

            If paramsNode IsNot Nothing Then
                notification("params") = CloneJson(paramsNode)
            End If

            Await SendMessageAsync(notification, cancellationToken).ConfigureAwait(False)
        End Function

        Public Async Function SendResultAsync(requestId As JsonNode,
                                              resultNode As JsonNode,
                                              Optional cancellationToken As CancellationToken = Nothing) As Task
            EnsureRunning()

            Dim response As New JsonObject()
            response("id") = CloneJson(requestId)
            response("result") = If(resultNode Is Nothing, New JsonObject(), CloneJson(resultNode))

            Await SendMessageAsync(response, cancellationToken).ConfigureAwait(False)
        End Function

        Public Async Function SendErrorAsync(requestId As JsonNode,
                                             code As Integer,
                                             message As String,
                                             Optional dataNode As JsonNode = Nothing,
                                             Optional cancellationToken As CancellationToken = Nothing) As Task
            EnsureRunning()

            Dim response As New JsonObject()
            response("id") = CloneJson(requestId)

            Dim errorObject As New JsonObject()
            errorObject("code") = code
            errorObject("message") = message
            If dataNode IsNot Nothing Then
                errorObject("data") = CloneJson(dataNode)
            End If

            response("error") = errorObject
            Await SendMessageAsync(response, cancellationToken).ConfigureAwait(False)
        End Function

        Public Async Function StopAsync(Optional reason As String = "Disconnected.") As Task
            If Interlocked.Exchange(_stopGate, 1) = 1 Then
                Return
            End If

            Try
                If _cts IsNot Nothing Then
                    _cts.Cancel()
                End If

                If _stdin IsNot Nothing Then
                    Try
                        _stdin.Close()
                    Catch
                    End Try
                End If

                If _process IsNot Nothing AndAlso Not _process.HasExited Then
                    Try
                        _process.Kill(entireProcessTree:=True)
                    Catch
                    End Try
                End If

                If _stdoutTask IsNot Nothing Then
                    Await Task.WhenAny(_stdoutTask, Task.Delay(250)).ConfigureAwait(False)
                End If

                If _stderrTask IsNot Nothing Then
                    Await Task.WhenAny(_stderrTask, Task.Delay(250)).ConfigureAwait(False)
                End If
            Finally
                FailAllPendingRequests(New IOException(reason))

                If _stdin IsNot Nothing Then
                    _stdin.Dispose()
                    _stdin = Nothing
                End If

                If _process IsNot Nothing Then
                    _process.Dispose()
                    _process = Nothing
                End If

                If _cts IsNot Nothing Then
                    _cts.Dispose()
                    _cts = Nothing
                End If

                _stdoutTask = Nothing
                _stderrTask = Nothing
                RaiseEvent Disconnected(reason)

                Interlocked.Exchange(_stopGate, 0)
            End Try
        End Function

        Public Function DisposeAsync() As ValueTask Implements IAsyncDisposable.DisposeAsync
            Return New ValueTask(DisposeInternalAsync())
        End Function

        Private Async Function DisposeInternalAsync() As Task
            Await StopAsync("Disposed.").ConfigureAwait(False)
            _writeLock.Dispose()
        End Function

        Private Async Function ReadStdoutLoopAsync(token As CancellationToken) As Task
            Try
                While Not token.IsCancellationRequested
                    Dim line = Await _process.StandardOutput.ReadLineAsync().ConfigureAwait(False)
                    If line Is Nothing Then
                        Exit While
                    End If

                    RaiseEvent RawMessage("recv", line)
                    Await HandleIncomingLineAsync(line).ConfigureAwait(False)
                End While
            Catch ex As Exception
                RaiseEvent RawMessage("transport-error", ex.Message)
            End Try

            If _cts IsNot Nothing AndAlso Not _cts.IsCancellationRequested Then
                QueueUnexpectedDisconnect("Codex App Server disconnected.")
            End If
        End Function

        Private Async Function ReadStderrLoopAsync(token As CancellationToken) As Task
            Try
                While Not token.IsCancellationRequested
                    Dim line = Await _process.StandardError.ReadLineAsync().ConfigureAwait(False)
                    If line Is Nothing Then
                        Exit While
                    End If

                    RaiseEvent RawMessage("stderr", line)
                End While
            Catch ex As Exception
                RaiseEvent RawMessage("transport-error", ex.Message)
            End Try
        End Function

        Private Async Function HandleIncomingLineAsync(line As String) As Task
            Dim root As JsonNode
            Try
                root = JsonNode.Parse(line)
            Catch ex As JsonException
                RaiseEvent RawMessage("parse-error", ex.Message)
                Return
            End Try

            Dim payloadObject = AsObject(root)
            If payloadObject Is Nothing Then
                Return
            End If

            Dim methodNode As JsonNode = Nothing
            Dim hasMethodNode = payloadObject.TryGetPropertyValue("method", methodNode)
            Dim methodName As String = String.Empty
            Dim hasMethodName = hasMethodNode AndAlso TryGetStringValue(methodNode, methodName)

            Dim idNode As JsonNode = Nothing
            Dim hasId = payloadObject.TryGetPropertyValue("id", idNode) AndAlso idNode IsNot Nothing

            If hasMethodName AndAlso hasId Then
                Dim paramsNode As JsonNode = Nothing
                payloadObject.TryGetPropertyValue("params", paramsNode)
                Dim request As New RpcServerRequest(CloneJson(idNode), methodName, CloneJson(paramsNode))
                RaiseEvent ServerRequestReceived(request)
                Return
            End If

            If hasMethodName Then
                Dim paramsNode As JsonNode = Nothing
                payloadObject.TryGetPropertyValue("params", paramsNode)
                RaiseEvent NotificationReceived(methodName, CloneJson(paramsNode))
                Return
            End If

            If hasId Then
                HandleResponse(payloadObject, idNode)
            End If

            Await Task.CompletedTask.ConfigureAwait(False)
        End Function

        Private Sub HandleResponse(payloadObject As JsonObject, idNode As JsonNode)
            Dim requestKey = RequestIdToKey(idNode)
            If String.IsNullOrEmpty(requestKey) Then
                Return
            End If

            Dim tcs As TaskCompletionSource(Of JsonNode) = Nothing
            If Not _pendingRequests.TryRemove(requestKey, tcs) Then
                Return
            End If

            Dim errorNode As JsonNode = Nothing
            If payloadObject.TryGetPropertyValue("error", errorNode) AndAlso errorNode IsNot Nothing Then
                Dim errorObject = AsObject(errorNode)
                Dim code As Integer = -32000
                Dim message As String = "Unknown RPC error."
                Dim dataNode As JsonNode = Nothing

                If errorObject IsNot Nothing Then
                    Dim codeNode As JsonNode = Nothing
                    If errorObject.TryGetPropertyValue("code", codeNode) AndAlso codeNode IsNot Nothing Then
                        Dim codeLong As Long
                        If TryGetInt64Value(codeNode, codeLong) Then
                            code = CInt(Math.Max(Math.Min(codeLong, Integer.MaxValue), Integer.MinValue))
                        End If
                    End If

                    message = GetPropertyString(errorObject, "message", message)
                    errorObject.TryGetPropertyValue("data", dataNode)
                End If

                tcs.TrySetException(New RpcErrorException(code, message, CloneJson(dataNode)))
                Return
            End If

            Dim resultNode As JsonNode = Nothing
            If payloadObject.TryGetPropertyValue("result", resultNode) Then
                tcs.TrySetResult(CloneJson(resultNode))
            Else
                tcs.TrySetResult(New JsonObject())
            End If
        End Sub

        Private Async Function SendMessageAsync(message As JsonObject, cancellationToken As CancellationToken) As Task
            EnsureRunning()

            Dim payload = message.ToJsonString(_jsonWriteOptions)
            RaiseEvent RawMessage("send", payload)

            Await _writeLock.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                Await _stdin.WriteLineAsync(payload).ConfigureAwait(False)
                Await _stdin.FlushAsync().ConfigureAwait(False)
            Finally
                _writeLock.Release()
            End Try
        End Function

        Private Sub EnsureRunning()
            If Not IsRunning OrElse _stdin Is Nothing Then
                Throw New InvalidOperationException("Codex App Server is not running.")
            End If
        End Sub

        Private Sub FailAllPendingRequests(ex As Exception)
            For Each kvp In _pendingRequests
                Dim removed As TaskCompletionSource(Of JsonNode) = Nothing
                If _pendingRequests.TryRemove(kvp.Key, removed) Then
                    removed.TrySetException(ex)
                End If
            Next
        End Sub

        Private Sub QueueUnexpectedDisconnect(reason As String)
            Task.Run(
                Async Function()
                    Try
                        Await StopAsync(reason).ConfigureAwait(False)
                    Catch
                    End Try
                End Function)
        End Sub
    End Class
End Namespace
