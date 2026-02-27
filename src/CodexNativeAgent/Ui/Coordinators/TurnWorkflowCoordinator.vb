Imports System.Collections.Generic
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports CodexNativeAgent.AppServer
Imports CodexNativeAgent.Services
Imports CodexNativeAgent.Ui.ViewModels

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class TurnWorkflowCoordinator
        Public NotInheritable Class ApprovalQueuedEventArgs
            Inherits EventArgs

            Public Property Approval As PendingApprovalInfo
        End Class

        Public NotInheritable Class ApprovalResolvedEventArgs
            Inherits EventArgs

            Public Property RequestId As JsonNode
            Public Property MethodName As String = String.Empty
            Public Property Decision As String = String.Empty
            Public Property ThreadId As String = String.Empty
            Public Property TurnId As String = String.Empty
            Public Property ItemId As String = String.Empty
        End Class

        Private NotInheritable Class ApprovalThreadState
            Public Property ThreadKey As String = String.Empty
            Public Property ActiveApproval As PendingApprovalInfo
            Public ReadOnly Property PendingQueue As New Queue(Of PendingApprovalInfo)()
        End Class

        Private Const UnscopedApprovalThreadKey As String = "__unscoped_approval__"

        Private ReadOnly _viewModel As MainWindowViewModel
        Private ReadOnly _turnService As ITurnService
        Private ReadOnly _approvalService As IApprovalService
        Private ReadOnly _runUiActionAsync As Func(Of Func(Of Task), Task)
        Private ReadOnly _approvalStatesByThreadKey As New Dictionary(Of String, ApprovalThreadState)(StringComparer.Ordinal)

        Public Event ApprovalQueued As EventHandler(Of ApprovalQueuedEventArgs)
        Public Event ApprovalActivated As EventHandler(Of ApprovalQueuedEventArgs)
        Public Event ApprovalResolved As EventHandler(Of ApprovalResolvedEventArgs)

        Public Sub New(viewModel As MainWindowViewModel,
                       turnService As ITurnService,
                       approvalService As IApprovalService,
                       runUiActionAsync As Func(Of Func(Of Task), Task))
            If viewModel Is Nothing Then Throw New ArgumentNullException(NameOf(viewModel))
            If turnService Is Nothing Then Throw New ArgumentNullException(NameOf(turnService))
            If approvalService Is Nothing Then Throw New ArgumentNullException(NameOf(approvalService))
            If runUiActionAsync Is Nothing Then Throw New ArgumentNullException(NameOf(runUiActionAsync))

            _viewModel = viewModel
            _turnService = turnService
            _approvalService = approvalService
            _runUiActionAsync = runUiActionAsync
        End Sub

        Public ReadOnly Property HasActiveApproval As Boolean
            Get
                For Each state In _approvalStatesByThreadKey.Values
                    If state IsNot Nothing AndAlso state.ActiveApproval IsNot Nothing Then
                        Return True
                    End If
                Next

                Return False
            End Get
        End Property

        Public Function HasActiveApprovalForThread(threadId As String) As Boolean
            Dim targetThreadKey = ResolveApprovalThreadKeyForSelection(threadId)
            Dim threadState As ApprovalThreadState = Nothing
            If String.IsNullOrWhiteSpace(targetThreadKey) Then
                Return False
            End If

            Return _approvalStatesByThreadKey.TryGetValue(targetThreadKey, threadState) AndAlso
                   threadState IsNot Nothing AndAlso
                   threadState.ActiveApproval IsNot Nothing
        End Function

        Public Function HasPendingApprovalForThread(threadId As String) As Boolean
            Dim targetThreadKey = ResolveApprovalThreadKeyForSelection(threadId)
            Dim threadState As ApprovalThreadState = Nothing
            If String.IsNullOrWhiteSpace(targetThreadKey) Then
                Return False
            End If

            If Not _approvalStatesByThreadKey.TryGetValue(targetThreadKey, threadState) OrElse threadState Is Nothing Then
                Return False
            End If

            Return threadState.ActiveApproval IsNot Nothing OrElse threadState.PendingQueue.Count > 0
        End Function

        Public Sub RefreshApprovalPanelForThread(threadId As String, isAuthenticated As Boolean)
            Dim targetThreadKey = ResolveApprovalThreadKeyForSelection(threadId)
            Dim threadState As ApprovalThreadState = Nothing
            If Not String.IsNullOrWhiteSpace(targetThreadKey) Then
                _approvalStatesByThreadKey.TryGetValue(targetThreadKey, threadState)
            End If

            Dim activeApproval = threadState?.ActiveApproval
            Dim hasActiveApproval = activeApproval IsNot Nothing
            Dim pendingQueueCount = If(threadState Is Nothing, 0, threadState.PendingQueue.Count)
            Dim summaryText = If(hasActiveApproval, BuildApprovalHintText(activeApproval), String.Empty)
            Dim activeMethodName = If(hasActiveApproval, activeApproval.MethodName, String.Empty)
            Dim supportsExecpolicyAmendment = hasActiveApproval AndAlso activeApproval.SupportsExecpolicyAmendment

            _viewModel.ApprovalPanel.SetThreadScopedState(summaryText,
                                                          activeMethodName,
                                                          supportsExecpolicyAmendment,
                                                          pendingQueueCount)
            _viewModel.ApprovalPanel.UpdateAvailability(isAuthenticated, hasActiveApproval)
        End Sub

        Public Sub BindCommands(startTurnAsync As Func(Of Task),
                                steerTurnAsync As Func(Of Task),
                                interruptTurnAsync As Func(Of Task),
                                resolveApprovalAsync As Func(Of String, Task))
            _viewModel.TurnComposer.ConfigureCommands(
                Function()
                    Return _runUiActionAsync(startTurnAsync)
                End Function,
                Function()
                    Return _runUiActionAsync(steerTurnAsync)
                End Function,
                Function()
                    Return _runUiActionAsync(interruptTurnAsync)
                End Function)

            _viewModel.ApprovalPanel.ConfigureCommands(
                Function()
                    Return _runUiActionAsync(Function() resolveApprovalAsync("accept"))
                End Function,
                Function()
                    Return _runUiActionAsync(Function() resolveApprovalAsync("accept_session"))
                End Function,
                Function()
                    Return _runUiActionAsync(Function() resolveApprovalAsync("accept_amended"))
                End Function,
                Function()
                    Return _runUiActionAsync(Function() resolveApprovalAsync("decline"))
                End Function,
                Function()
                    Return _runUiActionAsync(Function() resolveApprovalAsync("cancel"))
                End Function)
        End Sub

        Public Sub ResetApprovalState()
            _approvalStatesByThreadKey.Clear()
            _viewModel.ApprovalPanel.ResetLifecycleState()
        End Sub

        Public Async Function RunStartTurnAsync(currentThreadId As String,
                                                rawInputText As String,
                                                modelId As String,
                                                effort As String,
                                                approvalPolicy As String,
                                                ensureThreadSelected As Action,
                                                beforeSubmitInput As Action(Of String),
                                                afterTurnStarted As Action(Of String)) As Task
            If ensureThreadSelected Is Nothing Then Throw New ArgumentNullException(NameOf(ensureThreadSelected))
            If beforeSubmitInput Is Nothing Then Throw New ArgumentNullException(NameOf(beforeSubmitInput))
            If afterTurnStarted Is Nothing Then Throw New ArgumentNullException(NameOf(afterTurnStarted))

            ensureThreadSelected()

            Dim inputText = If(rawInputText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(inputText) Then
                Throw New InvalidOperationException("Enter turn input before sending.")
            End If

            beforeSubmitInput(inputText)

            Dim result = Await _turnService.StartTurnAsync(currentThreadId,
                                                           inputText,
                                                           modelId,
                                                           effort,
                                                           approvalPolicy,
                                                           CancellationToken.None).ConfigureAwait(True)

            afterTurnStarted(If(result?.TurnId, String.Empty))
        End Function

        Public Async Function RunSteerTurnAsync(currentThreadId As String,
                                                currentTurnId As String,
                                                rawInputText As String,
                                                ensureThreadSelected As Action,
                                                beforeSubmitInput As Action(Of String),
                                                afterTurnSteered As Action(Of String)) As Task
            If ensureThreadSelected Is Nothing Then Throw New ArgumentNullException(NameOf(ensureThreadSelected))
            If beforeSubmitInput Is Nothing Then Throw New ArgumentNullException(NameOf(beforeSubmitInput))
            If afterTurnSteered Is Nothing Then Throw New ArgumentNullException(NameOf(afterTurnSteered))

            ensureThreadSelected()

            If String.IsNullOrWhiteSpace(currentTurnId) Then
                Throw New InvalidOperationException("No active turn to steer.")
            End If

            Dim steerText = If(rawInputText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(steerText) Then
                Throw New InvalidOperationException("Enter steer input before sending.")
            End If

            beforeSubmitInput(steerText)

            Dim returnedTurnId = Await _turnService.SteerTurnAsync(currentThreadId,
                                                                   currentTurnId,
                                                                   steerText,
                                                                   CancellationToken.None).ConfigureAwait(True)

            afterTurnSteered(If(returnedTurnId, String.Empty))
        End Function

        Public Async Function RunInterruptTurnAsync(currentThreadId As String,
                                                    currentTurnId As String,
                                                    ensureThreadSelected As Action,
                                                    afterInterruptRequested As Action(Of String)) As Task
            If ensureThreadSelected Is Nothing Then Throw New ArgumentNullException(NameOf(ensureThreadSelected))
            If afterInterruptRequested Is Nothing Then Throw New ArgumentNullException(NameOf(afterInterruptRequested))

            ensureThreadSelected()

            If String.IsNullOrWhiteSpace(currentTurnId) Then
                Throw New InvalidOperationException("No active turn to interrupt.")
            End If

            Await _turnService.InterruptTurnAsync(currentThreadId,
                                                  currentTurnId,
                                                  CancellationToken.None).ConfigureAwait(True)

            afterInterruptRequested(currentTurnId)
        End Function

        Public Async Function HandleServerRequestAsync(request As RpcServerRequest,
                                                       handleToolRequestUserInputAsync As Func(Of RpcServerRequest, Task),
                                                       handleUnsupportedToolCallAsync As Func(Of RpcServerRequest, Task),
                                                       handleChatgptTokenRefreshAsync As Func(Of RpcServerRequest, Task),
                                                       sendUnsupportedServerRequestErrorAsync As Func(Of RpcServerRequest, Integer, String, Task),
                                                       currentThreadResolver As Func(Of String),
                                                       refreshControlStates As Action,
                                                       refreshThreadRuntimeIndicators As Action,
                                                       appendSystemMessage As Action(Of String),
                                                       showStatus As Action(Of String, Boolean, Boolean)) As Task
            If request Is Nothing Then Throw New ArgumentNullException(NameOf(request))
            If handleToolRequestUserInputAsync Is Nothing Then Throw New ArgumentNullException(NameOf(handleToolRequestUserInputAsync))
            If handleUnsupportedToolCallAsync Is Nothing Then Throw New ArgumentNullException(NameOf(handleUnsupportedToolCallAsync))
            If handleChatgptTokenRefreshAsync Is Nothing Then Throw New ArgumentNullException(NameOf(handleChatgptTokenRefreshAsync))
            If sendUnsupportedServerRequestErrorAsync Is Nothing Then Throw New ArgumentNullException(NameOf(sendUnsupportedServerRequestErrorAsync))
            If refreshControlStates Is Nothing Then Throw New ArgumentNullException(NameOf(refreshControlStates))
            If refreshThreadRuntimeIndicators Is Nothing Then Throw New ArgumentNullException(NameOf(refreshThreadRuntimeIndicators))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If showStatus Is Nothing Then Throw New ArgumentNullException(NameOf(showStatus))

            Dim approvalInfo As PendingApprovalInfo = Nothing
            If _approvalService.TryCreateApproval(request, approvalInfo) Then
                QueueApproval(approvalInfo,
                              ResolveCurrentThreadId(currentThreadResolver),
                              refreshControlStates,
                              refreshThreadRuntimeIndicators,
                              appendSystemMessage,
                              showStatus)
                Return
            End If

            Select Case request.MethodName
                Case ToolRequestUserInputMethod
                    Await handleToolRequestUserInputAsync(request).ConfigureAwait(True)

                Case "item/tool/call"
                    Await handleUnsupportedToolCallAsync(request).ConfigureAwait(True)

                Case "account/chatgptAuthTokens/refresh"
                    Await handleChatgptTokenRefreshAsync(request).ConfigureAwait(True)

                Case Else
                    Await sendUnsupportedServerRequestErrorAsync(request,
                                                                 -32601,
                                                                 $"Unsupported server request method: {request.MethodName}").ConfigureAwait(True)
            End Select
        End Function

        Public Async Function ResolveApprovalAsync(action As String,
                                                   selectedThreadId As String,
                                                   sendResultAsync As Func(Of JsonNode, JsonObject, Task),
                                                   refreshControlStates As Action,
                                                   refreshThreadRuntimeIndicators As Action,
                                                   appendSystemMessage As Action(Of String),
                                                   showStatus As Action(Of String, Boolean, Boolean)) As Task
            If sendResultAsync Is Nothing Then Throw New ArgumentNullException(NameOf(sendResultAsync))
            If refreshControlStates Is Nothing Then Throw New ArgumentNullException(NameOf(refreshControlStates))
            If refreshThreadRuntimeIndicators Is Nothing Then Throw New ArgumentNullException(NameOf(refreshThreadRuntimeIndicators))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If showStatus Is Nothing Then Throw New ArgumentNullException(NameOf(showStatus))

            Dim threadKey = ResolveApprovalThreadKeyForSelection(selectedThreadId)
            Dim threadState As ApprovalThreadState = Nothing
            If String.IsNullOrWhiteSpace(threadKey) OrElse
               Not _approvalStatesByThreadKey.TryGetValue(threadKey, threadState) OrElse
               threadState Is Nothing OrElse
               threadState.ActiveApproval Is Nothing Then
                _viewModel.ApprovalPanel.RecordError("No active approval for the selected thread.")
                Return
            End If

            Dim activeApproval = threadState.ActiveApproval
            Dim decisionLabel As String = Nothing
            Dim decisionPayload = _approvalService.ResolveDecisionPayload(activeApproval, action, decisionLabel)
            If decisionPayload Is Nothing OrElse String.IsNullOrWhiteSpace(decisionLabel) Then
                _viewModel.ApprovalPanel.RecordError("No decision mapping is available for this approval type.")
                Throw New InvalidOperationException("No decision mapping is available for this approval type.")
            End If

            Try
                Dim resultNode As New JsonObject()
                resultNode("decision") = decisionPayload

                Await sendResultAsync(activeApproval.RequestId, resultNode).ConfigureAwait(True)
                appendSystemMessage($"Approval sent: {decisionLabel}")
                showStatus($"Approval sent: {decisionLabel}", False, True)
                _viewModel.ApprovalPanel.OnApprovalResolved(action, decisionLabel, threadState.PendingQueue.Count)

                RaiseEvent ApprovalResolved(Me,
                                            New ApprovalResolvedEventArgs() With {
                                                .RequestId = CloneJson(activeApproval.RequestId),
                                                .MethodName = activeApproval.MethodName,
                                                .Decision = decisionLabel,
                                                .ThreadId = activeApproval.ThreadId,
                                                .TurnId = activeApproval.TurnId,
                                                .ItemId = activeApproval.ItemId
                                            })

                threadState.ActiveApproval = Nothing
                ShowNextApprovalIfNeeded(threadState)
                RemoveThreadStateIfEmpty(threadKey)
                refreshControlStates()
                refreshThreadRuntimeIndicators()
            Catch ex As Exception
                _viewModel.ApprovalPanel.RecordError(ex.Message)
                Throw
            End Try
        End Function

        Private Sub QueueApproval(approvalInfo As PendingApprovalInfo,
                                  fallbackCurrentThreadId As String,
                                  refreshControlStates As Action,
                                  refreshThreadRuntimeIndicators As Action,
                                  appendSystemMessage As Action(Of String),
                                  showStatus As Action(Of String, Boolean, Boolean))
            If approvalInfo Is Nothing Then
                Return
            End If

            EnsureApprovalThreadScope(approvalInfo, fallbackCurrentThreadId)
            Dim threadKey = ResolveApprovalThreadKeyFromApproval(approvalInfo)
            Dim threadState = GetOrCreateThreadState(threadKey)
            threadState.PendingQueue.Enqueue(approvalInfo)

            RaiseEvent ApprovalQueued(Me,
                                      New ApprovalQueuedEventArgs() With {
                                          .Approval = approvalInfo
                                      })
            ShowNextApprovalIfNeeded(threadState)
            appendSystemMessage($"Approval queued: {approvalInfo.MethodName}")
            showStatus($"Approval queued: {approvalInfo.MethodName}", False, True)
            refreshControlStates()
            refreshThreadRuntimeIndicators()
        End Sub

        Private Sub ShowNextApprovalIfNeeded(threadState As ApprovalThreadState)
            If threadState Is Nothing Then
                Return
            End If

            If threadState.ActiveApproval IsNot Nothing Then
                Return
            End If

            If threadState.PendingQueue.Count = 0 Then
                Return
            End If

            threadState.ActiveApproval = threadState.PendingQueue.Dequeue()
            RaiseEvent ApprovalActivated(Me,
                                         New ApprovalQueuedEventArgs() With {
                                             .Approval = threadState.ActiveApproval
                                         })
        End Sub

        Private Function GetOrCreateThreadState(threadKey As String) As ApprovalThreadState
            Dim normalizedThreadKey = If(threadKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadKey) Then
                normalizedThreadKey = UnscopedApprovalThreadKey
            End If

            Dim threadState As ApprovalThreadState = Nothing
            If _approvalStatesByThreadKey.TryGetValue(normalizedThreadKey, threadState) AndAlso threadState IsNot Nothing Then
                Return threadState
            End If

            threadState = New ApprovalThreadState() With {
                .ThreadKey = normalizedThreadKey
            }
            _approvalStatesByThreadKey(normalizedThreadKey) = threadState
            Return threadState
        End Function

        Private Function ResolveApprovalThreadKeyFromApproval(approval As PendingApprovalInfo) As String
            Dim approvalThreadId = If(approval?.ThreadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(approvalThreadId) Then
                Return UnscopedApprovalThreadKey
            End If

            Return approvalThreadId
        End Function

        Private Shared Function ResolveApprovalThreadKeyForSelection(threadId As String) As String
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return UnscopedApprovalThreadKey
            End If

            Return normalizedThreadId
        End Function

        Private Shared Function ResolveCurrentThreadId(currentThreadResolver As Func(Of String)) As String
            If currentThreadResolver Is Nothing Then
                Return String.Empty
            End If

            Try
                Return If(currentThreadResolver.Invoke(), String.Empty).Trim()
            Catch
                Return String.Empty
            End Try
        End Function

        Private Shared Sub EnsureApprovalThreadScope(approvalInfo As PendingApprovalInfo,
                                                     fallbackCurrentThreadId As String)
            If approvalInfo Is Nothing Then
                Return
            End If

            Dim normalizedApprovalThreadId = If(approvalInfo.ThreadId, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(normalizedApprovalThreadId) Then
                approvalInfo.ThreadId = normalizedApprovalThreadId
                Return
            End If

            Dim normalizedFallbackThreadId = If(fallbackCurrentThreadId, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(normalizedFallbackThreadId) Then
                approvalInfo.ThreadId = normalizedFallbackThreadId
            End If
        End Sub

        Private Sub RemoveThreadStateIfEmpty(threadKey As String)
            Dim normalizedThreadKey = If(threadKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadKey) Then
                Return
            End If

            Dim threadState As ApprovalThreadState = Nothing
            If Not _approvalStatesByThreadKey.TryGetValue(normalizedThreadKey, threadState) OrElse threadState Is Nothing Then
                Return
            End If

            If threadState.ActiveApproval IsNot Nothing OrElse threadState.PendingQueue.Count > 0 Then
                Return
            End If

            _approvalStatesByThreadKey.Remove(normalizedThreadKey)
        End Sub

        Private Shared Function BuildApprovalHintText(approval As PendingApprovalInfo) As String
            If approval Is Nothing Then
                Return String.Empty
            End If

            Dim commandText = ExtractApprovalSummaryField(approval.Summary, "Command")
            If Not String.IsNullOrWhiteSpace(commandText) Then
                Return $"Command: {TruncateWithEllipsis(commandText, 140)}"
            End If

            Dim reasonText = ExtractApprovalSummaryField(approval.Summary, "Reason")
            If Not String.IsNullOrWhiteSpace(reasonText) Then
                Return TruncateWithEllipsis(reasonText, 140)
            End If

            Dim methodName = If(approval.MethodName, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(methodName) Then
                Return $"Approval required: {methodName}"
            End If

            Return "Approval required for current thread."
        End Function

        Private Shared Function ExtractApprovalSummaryField(summaryText As String,
                                                            fieldName As String) As String
            Dim normalizedSummary = If(summaryText, String.Empty)
            Dim normalizedFieldName = If(fieldName, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedSummary) OrElse String.IsNullOrWhiteSpace(normalizedFieldName) Then
                Return String.Empty
            End If

            Dim lines = normalizedSummary.Replace(ControlChars.CrLf, ControlChars.Lf).
                                          Replace(ControlChars.Cr, ControlChars.Lf).
                                          Split({ControlChars.Lf}, StringSplitOptions.None)
            Dim prefix = normalizedFieldName & ":"
            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty).Trim()
                If line.Length = 0 Then
                    Continue For
                End If

                If line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                    Return line.Substring(prefix.Length).Trim()
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Function TruncateWithEllipsis(value As String, maxChars As Integer) As String
            Dim normalized = If(value, String.Empty).Trim()
            Dim clampedMaxChars = Math.Max(1, maxChars)
            If normalized.Length <= clampedMaxChars Then
                Return normalized
            End If

            Return normalized.Substring(0, clampedMaxChars).TrimEnd() & "..."
        End Function
    End Class
End Namespace
