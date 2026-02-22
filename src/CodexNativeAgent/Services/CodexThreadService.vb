Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Services
    Public NotInheritable Class CodexThreadService
        Implements IThreadService

        Private ReadOnly _clientProvider As Func(Of CodexAppServerClient)

        Public Sub New(clientProvider As Func(Of CodexAppServerClient))
            If clientProvider Is Nothing Then
                Throw New ArgumentNullException(NameOf(clientProvider))
            End If

            _clientProvider = clientProvider
        End Sub

        Public Async Function ListModelsAsync(cancellationToken As CancellationToken) As Task(Of IReadOnlyList(Of ModelSummary)) Implements IThreadService.ListModelsAsync
            Dim models As New List(Of ModelSummary)()
            Dim cursor As String = String.Empty

            Do
                Dim paramsNode As New JsonObject()
                paramsNode("limit") = 100
                paramsNode("includeHidden") = False

                If Not String.IsNullOrWhiteSpace(cursor) Then
                    paramsNode("cursor") = cursor
                End If

                Dim responseNode = Await CurrentClient().SendRequestAsync("model/list",
                                                                          paramsNode,
                                                                          cancellationToken:=cancellationToken)
                Dim responseObject = AsObject(responseNode)
                If responseObject Is Nothing Then
                    Exit Do
                End If

                Dim data = GetPropertyArray(responseObject, "data")
                If data IsNot Nothing Then
                    For Each modelNode In data
                        Dim modelObject = AsObject(modelNode)
                        If modelObject Is Nothing Then
                            Continue For
                        End If

                        Dim modelId = GetPropertyString(modelObject, "id")
                        If String.IsNullOrWhiteSpace(modelId) Then
                            Continue For
                        End If

                        models.Add(New ModelSummary() With {
                            .Id = modelId,
                            .DisplayName = GetPropertyString(modelObject, "displayName"),
                            .IsDefault = GetPropertyBoolean(modelObject, "isDefault", False)
                        })
                    Next
                End If

                cursor = GetPropertyString(responseObject, "nextCursor")
            Loop While Not String.IsNullOrWhiteSpace(cursor)

            Return models
        End Function

        Public Async Function StartThreadAsync(options As ThreadRequestOptions,
                                               cancellationToken As CancellationToken) As Task(Of JsonObject) Implements IThreadService.StartThreadAsync
            Dim paramsNode = BuildThreadParams(options, includeModel:=True)
            Dim responseNode = Await CurrentClient().SendRequestAsync("thread/start",
                                                                      paramsNode,
                                                                      cancellationToken:=cancellationToken)
            Dim responseObject = AsObject(responseNode)
            Dim threadObject = GetPropertyObject(responseObject, "thread")

            If threadObject Is Nothing Then
                Throw New InvalidOperationException("Thread start response did not include a thread payload.")
            End If

            Return threadObject
        End Function

        Public Async Function ListThreadsAsync(includeArchived As Boolean,
                                               cwd As String,
                                               cancellationToken As CancellationToken) As Task(Of IReadOnlyList(Of ThreadSummary)) Implements IThreadService.ListThreadsAsync
            Dim entries As New List(Of ThreadSummary)()
            Dim seenThreadIds As New HashSet(Of String)(StringComparer.Ordinal)
            Dim cursor As String = String.Empty
            Dim seenCursors As New HashSet(Of String)(StringComparer.Ordinal)

            Do
                Dim paramsNode As New JsonObject()
                paramsNode("limit") = 100
                paramsNode("sortKey") = "updated_at"
                paramsNode("archived") = includeArchived

                If Not String.IsNullOrWhiteSpace(cursor) Then
                    paramsNode("cursor") = cursor
                End If

                If Not String.IsNullOrWhiteSpace(cwd) Then
                    paramsNode("cwd") = cwd
                End If

                Dim responseNode = Await CurrentClient().SendRequestAsync("thread/list",
                                                                          paramsNode,
                                                                          cancellationToken:=cancellationToken)
                Dim responseObject = AsObject(responseNode)
                Dim data = GetPropertyArray(responseObject, "data")

                If data IsNot Nothing Then
                    For Each threadNode In data
                        Dim threadObject = AsObject(threadNode)
                        If threadObject Is Nothing Then
                            Continue For
                        End If

                        Dim threadId = GetPropertyString(threadObject, "id")
                        If String.IsNullOrWhiteSpace(threadId) OrElse Not seenThreadIds.Add(threadId) Then
                            Continue For
                        End If

                        Dim lastActiveNode = ExtractLastActiveNode(threadObject)
                        Dim lastActiveText = NormalizeTimestampText(lastActiveNode)
                        Dim lastActiveSortValue = ParseSortTimestamp(lastActiveNode)

                        entries.Add(New ThreadSummary() With {
                            .Id = threadId,
                            .Preview = GetPropertyString(threadObject, "preview"),
                            .UpdatedAtText = lastActiveText,
                            .UpdatedSortValue = lastActiveSortValue,
                            .LastActiveText = lastActiveText,
                            .LastActiveSortValue = lastActiveSortValue,
                            .Cwd = ExtractThreadWorkingDirectory(threadObject)
                        })
                    Next
                End If

                Dim nextCursor = GetPropertyString(responseObject, "nextCursor")
                If String.IsNullOrWhiteSpace(nextCursor) OrElse seenCursors.Contains(nextCursor) Then
                    Exit Do
                End If

                seenCursors.Add(nextCursor)
                cursor = nextCursor
            Loop

            Return entries
        End Function

        Public Async Function ResumeThreadAsync(threadId As String,
                                                options As ThreadRequestOptions,
                                                cancellationToken As CancellationToken) As Task(Of JsonObject) Implements IThreadService.ResumeThreadAsync
            Dim paramsNode = BuildThreadParams(options, includeModel:=True)
            paramsNode("threadId") = threadId

            Dim responseNode = Await CurrentClient().SendRequestAsync("thread/resume",
                                                                      paramsNode,
                                                                      cancellationToken:=cancellationToken)
            Dim responseObject = AsObject(responseNode)
            Dim threadObject = GetPropertyObject(responseObject, "thread")

            If threadObject Is Nothing Then
                Throw New InvalidOperationException("Thread resume response did not include a thread payload.")
            End If

            Return threadObject
        End Function

        Public Async Function ReadThreadAsync(threadId As String,
                                              includeTurns As Boolean,
                                              cancellationToken As CancellationToken) As Task(Of JsonObject) Implements IThreadService.ReadThreadAsync
            Dim paramsNode As New JsonObject()
            paramsNode("threadId") = threadId
            paramsNode("includeTurns") = includeTurns

            Dim responseNode = Await CurrentClient().SendRequestAsync("thread/read",
                                                                      paramsNode,
                                                                      cancellationToken:=cancellationToken)
            Dim responseObject = AsObject(responseNode)
            Dim threadObject = GetPropertyObject(responseObject, "thread")

            If threadObject Is Nothing Then
                Throw New InvalidOperationException("Thread read response did not include a thread payload.")
            End If

            Return threadObject
        End Function

        Public Async Function ForkThreadAsync(threadId As String,
                                              options As ThreadRequestOptions,
                                              cancellationToken As CancellationToken) As Task(Of JsonObject) Implements IThreadService.ForkThreadAsync
            Dim paramsNode = BuildThreadParams(options, includeModel:=False)
            paramsNode("threadId") = threadId

            Dim responseNode = Await CurrentClient().SendRequestAsync("thread/fork",
                                                                      paramsNode,
                                                                      cancellationToken:=cancellationToken)
            Dim responseObject = AsObject(responseNode)
            Dim threadObject = GetPropertyObject(responseObject, "thread")

            If threadObject Is Nothing Then
                Throw New InvalidOperationException("Thread fork response did not include a thread payload.")
            End If

            Return threadObject
        End Function

        Public Async Function ArchiveThreadAsync(threadId As String,
                                                 cancellationToken As CancellationToken) As Task Implements IThreadService.ArchiveThreadAsync
            Dim paramsNode As New JsonObject()
            paramsNode("threadId") = threadId

            Await CurrentClient().SendRequestAsync("thread/archive",
                                                   paramsNode,
                                                   cancellationToken:=cancellationToken)
        End Function

        Public Async Function UnarchiveThreadAsync(threadId As String,
                                                   cancellationToken As CancellationToken) As Task Implements IThreadService.UnarchiveThreadAsync
            Dim paramsNode As New JsonObject()
            paramsNode("threadId") = threadId

            Await CurrentClient().SendRequestAsync("thread/unarchive",
                                                   paramsNode,
                                                   cancellationToken:=cancellationToken)
        End Function

        Private Function CurrentClient() As CodexAppServerClient
            Dim client = _clientProvider()
            If client Is Nothing OrElse Not client.IsRunning Then
                Throw New InvalidOperationException("Not connected to Codex App Server.")
            End If

            Return client
        End Function

        Private Shared Function BuildThreadParams(options As ThreadRequestOptions,
                                                  includeModel As Boolean) As JsonObject
            Dim paramsNode As New JsonObject()
            Dim effectiveOptions = If(options, New ThreadRequestOptions())

            If includeModel AndAlso Not String.IsNullOrWhiteSpace(effectiveOptions.Model) Then
                paramsNode("model") = effectiveOptions.Model
            End If

            If Not String.IsNullOrWhiteSpace(effectiveOptions.ApprovalPolicy) Then
                paramsNode("approvalPolicy") = effectiveOptions.ApprovalPolicy
            End If

            If Not String.IsNullOrWhiteSpace(effectiveOptions.Sandbox) Then
                paramsNode("sandbox") = effectiveOptions.Sandbox
            End If

            If Not String.IsNullOrWhiteSpace(effectiveOptions.Cwd) Then
                paramsNode("cwd") = effectiveOptions.Cwd
            End If

            Return paramsNode
        End Function

        Private Shared Function ExtractLastActiveNode(threadObject As JsonObject) As JsonNode
            If threadObject Is Nothing Then
                Return Nothing
            End If

            Dim node = GetNestedProperty(threadObject, "updatedAt")
            If node IsNot Nothing Then
                Return node
            End If

            node = GetNestedProperty(threadObject, "updated_at")
            If node IsNot Nothing Then
                Return node
            End If

            node = GetNestedProperty(threadObject, "lastActiveAt")
            If node IsNot Nothing Then
                Return node
            End If

            node = GetNestedProperty(threadObject, "last_active_at")
            If node IsNot Nothing Then
                Return node
            End If

            node = GetNestedProperty(threadObject, "lastActivityAt")
            If node IsNot Nothing Then
                Return node
            End If

            node = GetNestedProperty(threadObject, "last_activity_at")
            If node IsNot Nothing Then
                Return node
            End If

            Return Nothing
        End Function

        Private Shared Function ExtractThreadWorkingDirectory(threadObject As JsonObject) As String
            If threadObject Is Nothing Then
                Return String.Empty
            End If

            Dim directKeys = {
                "cwd",
                "workingDirectory",
                "workingDir"
            }

            For Each key In directKeys
                Dim value = GetPropertyString(threadObject, key)
                If Not String.IsNullOrWhiteSpace(value) Then
                    Return value.Trim()
                End If
            Next

            Dim nestedCandidates = {
                GetNestedProperty(threadObject, "context", "cwd"),
                GetNestedProperty(threadObject, "context", "workingDirectory"),
                GetNestedProperty(threadObject, "workspace", "cwd"),
                GetNestedProperty(threadObject, "workspace", "workingDirectory"),
                GetNestedProperty(threadObject, "project", "cwd"),
                GetNestedProperty(threadObject, "project", "workingDirectory")
            }

            For Each node In nestedCandidates
                Dim value As String = String.Empty
                If TryGetStringValue(node, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then
                    Return value.Trim()
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Function NormalizeTimestampText(node As JsonNode) As String
            If node Is Nothing Then
                Return String.Empty
            End If

            Dim stringValue As String = String.Empty
            If TryGetStringValue(node, stringValue) Then
                Dim unixMs As Long
                If Long.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, unixMs) Then
                    Return UnixToLocalTime(unixMs)
                End If

                Return stringValue
            End If

            Dim unixValue As Long
            If TryGetInt64Value(node, unixValue) Then
                Return UnixToLocalTime(unixValue)
            End If

            Return node.ToJsonString()
        End Function

        Private Shared Function ParseSortTimestamp(node As JsonNode) As Long
            If node Is Nothing Then
                Return Long.MinValue
            End If

            Dim stringValue As String = String.Empty
            If TryGetStringValue(node, stringValue) Then
                Dim parsed As Long
                If Long.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                    Return NormalizeUnixForSort(parsed)
                End If

                Return Long.MinValue
            End If

            Dim unixValue As Long
            If TryGetInt64Value(node, unixValue) Then
                Return NormalizeUnixForSort(unixValue)
            End If

            Return Long.MinValue
        End Function

        Private Shared Function NormalizeUnixForSort(value As Long) As Long
            If value <= 0 Then
                Return value
            End If

            If value > 1000000000000L Then
                Return value
            End If

            Return value * 1000L
        End Function

        Private Shared Function UnixToLocalTime(unixValue As Long) As String
            Try
                If unixValue > 1000000000000L Then
                    Return DateTimeOffset.FromUnixTimeMilliseconds(unixValue).LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss")
                End If

                Return DateTimeOffset.FromUnixTimeSeconds(unixValue).LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss")
            Catch
                Return unixValue.ToString(CultureInfo.InvariantCulture)
            End Try
        End Function
    End Class
End Namespace
