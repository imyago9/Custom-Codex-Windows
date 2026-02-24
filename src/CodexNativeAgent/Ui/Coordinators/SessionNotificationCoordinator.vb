Imports System.Globalization
Imports System.Collections.Generic
Imports System.Text.Json.Nodes
Imports CodexNativeAgent.AppServer
Imports CodexNativeAgent.Services

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class SessionNotificationCoordinator
        Private ReadOnly _runtimeStore As New TurnFlowRuntimeStore()
        Private ReadOnly _reportedUnknownMethods As New HashSet(Of String)(StringComparer.Ordinal)

        Public ReadOnly Property RuntimeStore As TurnFlowRuntimeStore
            Get
                Return _runtimeStore
            End Get
        End Property

        Public Sub ResetStreamingAgentItems()
            _runtimeStore.Reset()
            _reportedUnknownMethods.Clear()
        End Sub

        Public NotInheritable Class ProtocolDispatchMessage
            Public Property Direction As String = String.Empty
            Public Property Payload As String = String.Empty
        End Class

        Public NotInheritable Class TurnLifecycleDispatchMessage
            Public Property ThreadId As String = String.Empty
            Public Property TurnId As String = String.Empty
            Public Property Status As String = String.Empty
        End Class

        Public NotInheritable Class TurnMetadataDispatchMessage
            Public Property ThreadId As String = String.Empty
            Public Property TurnId As String = String.Empty
            Public Property Kind As String = String.Empty
            Public Property SummaryText As String = String.Empty
        End Class

        Public NotInheritable Class TokenUsageDispatchMessage
            Public Property ThreadId As String = String.Empty
            Public Property TurnId As String = String.Empty
            Public Property TokenUsage As JsonObject
        End Class

        Public NotInheritable Class NotificationDispatchResult
            Public Sub New(methodName As String)
                Me.MethodName = NormalizeDispatchValue(methodName)
            End Sub

            Public ReadOnly Property MethodName As String
            Public Property CurrentThreadId As String = String.Empty
            Public Property CurrentTurnId As String = String.Empty
            Public Property CurrentThreadChanged As Boolean
            Public Property CurrentTurnChanged As Boolean
            Public Property ThreadObject As JsonObject
            Public Property ShouldScrollTranscriptToBottom As Boolean
            Public Property ShouldRefreshAuthentication As Boolean
            Public ReadOnly Property ProtocolMessages As New List(Of ProtocolDispatchMessage)()
            Public ReadOnly Property SystemMessages As New List(Of String)()
            Public ReadOnly Property Diagnostics As New List(Of String)()
            Public ReadOnly Property RuntimeItems As New List(Of TurnItemRuntimeState)()
            Public ReadOnly Property TurnLifecycleMessages As New List(Of TurnLifecycleDispatchMessage)()
            Public ReadOnly Property TurnMetadataMessages As New List(Of TurnMetadataDispatchMessage)()
            Public ReadOnly Property TokenUsageMessages As New List(Of TokenUsageDispatchMessage)()
            Public ReadOnly Property ThreadIdsToMarkLastActive As New List(Of String)()
            Public ReadOnly Property LoginIdsToClear As New List(Of String)()
            Public ReadOnly Property RateLimitPayloads As New List(Of JsonObject)()
        End Class

        Public NotInheritable Class ServerRequestDispatchResult
            Public Sub New(methodName As String)
                Me.MethodName = NormalizeDispatchValue(methodName)
            End Sub

            Public ReadOnly Property MethodName As String
            Public ReadOnly Property ProtocolMessages As New List(Of ProtocolDispatchMessage)()
            Public ReadOnly Property Diagnostics As New List(Of String)()
            Public ReadOnly Property RuntimeItems As New List(Of TurnItemRuntimeState)()
        End Class

        Public NotInheritable Class ApprovalResolutionDispatchResult
            Public ReadOnly Property ProtocolMessages As New List(Of ProtocolDispatchMessage)()
            Public ReadOnly Property Diagnostics As New List(Of String)()
            Public ReadOnly Property RuntimeItems As New List(Of TurnItemRuntimeState)()
        End Class

        Public Function DispatchNotification(methodName As String,
                                             paramsNode As JsonNode,
                                             currentThreadId As String,
                                             currentTurnId As String) As NotificationDispatchResult
            Dim result As New NotificationDispatchResult(methodName)
            Dim nextThreadId = NormalizeDispatchValue(currentThreadId)
            Dim nextTurnId = NormalizeDispatchValue(currentTurnId)
            Dim currentThreadChanged = False
            Dim currentTurnChanged = False

            HandleNotification(
                methodName,
                paramsNode,
                Sub(threadObject)
                    result.ThreadObject = CloneJsonObject(threadObject)
                End Sub,
                Function() nextThreadId,
                Sub(value)
                    nextThreadId = NormalizeDispatchValue(value)
                    currentThreadChanged = True
                End Sub,
                Function() nextTurnId,
                Sub(value)
                    nextTurnId = NormalizeDispatchValue(value)
                    currentTurnChanged = True
                End Sub,
                Sub(threadId)
                    AddValue(result.ThreadIdsToMarkLastActive, threadId)
                End Sub,
                Sub(message)
                    AddValue(result.SystemMessages, message)
                End Sub,
                Sub(role, text)
                End Sub,
                Sub(itemId, renderAsReasoningChainStep)
                End Sub,
                Sub(itemId, delta)
                End Sub,
                Sub(itemId, finalText, renderAsReasoningChainStep)
                End Sub,
                Sub(itemId)
                End Sub,
                Sub(itemId, delta)
                End Sub,
                Sub(itemId, finalText)
                End Sub,
                Sub()
                    result.ShouldScrollTranscriptToBottom = True
                End Sub,
                Sub(direction, payload)
                    result.ProtocolMessages.Add(New ProtocolDispatchMessage() With {
                        .Direction = NormalizeDispatchValue(direction),
                        .Payload = If(payload, String.Empty)
                    })
                End Sub,
                Sub(itemObject)
                End Sub,
                Sub(loginId)
                    AddValue(result.LoginIdsToClear, loginId)
                End Sub,
                Sub()
                    result.ShouldRefreshAuthentication = True
                End Sub,
                Sub(rateLimitsPayload)
                    If rateLimitsPayload IsNot Nothing Then
                        result.RateLimitPayloads.Add(CloneJsonObject(rateLimitsPayload))
                    End If
                End Sub,
                Sub(item)
                    If item IsNot Nothing Then
                        result.RuntimeItems.Add(item)
                    End If
                End Sub,
                Sub(threadId, turnId, status)
                    result.TurnLifecycleMessages.Add(New TurnLifecycleDispatchMessage() With {
                        .ThreadId = NormalizeDispatchValue(threadId),
                        .TurnId = NormalizeDispatchValue(turnId),
                        .Status = NormalizeDispatchValue(status)
                    })
                End Sub,
                Sub(threadId, turnId, kind, summaryText)
                    result.TurnMetadataMessages.Add(New TurnMetadataDispatchMessage() With {
                        .ThreadId = NormalizeDispatchValue(threadId),
                        .TurnId = NormalizeDispatchValue(turnId),
                        .Kind = NormalizeDispatchValue(kind),
                        .SummaryText = If(summaryText, String.Empty)
                    })
                End Sub,
                Sub(threadId, turnId, tokenUsage)
                    result.TokenUsageMessages.Add(New TokenUsageDispatchMessage() With {
                        .ThreadId = NormalizeDispatchValue(threadId),
                        .TurnId = NormalizeDispatchValue(turnId),
                        .TokenUsage = CloneJsonObject(tokenUsage)
                    })
                End Sub,
                Sub(message)
                    AddValue(result.Diagnostics, message)
                End Sub)

            result.CurrentThreadId = nextThreadId
            result.CurrentTurnId = nextTurnId
            result.CurrentThreadChanged = currentThreadChanged
            result.CurrentTurnChanged = currentTurnChanged
            Return result
        End Function

        Public Function DispatchServerRequest(request As RpcServerRequest) As ServerRequestDispatchResult
            Dim result As New ServerRequestDispatchResult(If(request?.MethodName, String.Empty))

            HandleServerRequest(
                request,
                Sub(direction, payload)
                    result.ProtocolMessages.Add(New ProtocolDispatchMessage() With {
                        .Direction = NormalizeDispatchValue(direction),
                        .Payload = If(payload, String.Empty)
                    })
                End Sub,
                Sub(item)
                    If item IsNot Nothing Then
                        result.RuntimeItems.Add(item)
                    End If
                End Sub,
                Sub(message)
                    AddValue(result.Diagnostics, message)
                End Sub)

            Return result
        End Function

        Public Function DispatchApprovalResolution(requestId As JsonNode,
                                                   decision As String) As ApprovalResolutionDispatchResult
            Dim result As New ApprovalResolutionDispatchResult()

            HandleApprovalResolved(
                requestId,
                decision,
                Sub(direction, payload)
                    result.ProtocolMessages.Add(New ProtocolDispatchMessage() With {
                        .Direction = NormalizeDispatchValue(direction),
                        .Payload = If(payload, String.Empty)
                    })
                End Sub,
                Sub(item)
                    If item IsNot Nothing Then
                        result.RuntimeItems.Add(item)
                    End If
                End Sub,
                Sub(message)
                    AddValue(result.Diagnostics, message)
                End Sub)

            Return result
        End Function

        Public Sub HandleNotification(methodName As String,
                                      paramsNode As JsonNode,
                                      applyCurrentThreadFromThreadObject As Action(Of JsonObject),
                                      getCurrentThreadId As Func(Of String),
                                      setCurrentThreadId As Action(Of String),
                                      getCurrentTurnId As Func(Of String),
                                      setCurrentTurnId As Action(Of String),
                                      markThreadLastActive As Action(Of String),
                                      appendSystemMessage As Action(Of String),
                                      appendTranscript As Action(Of String, String),
                                      beginAssistantStream As Action(Of String, Boolean),
                                      appendAssistantStreamDelta As Action(Of String, String),
                                      completeAssistantStream As Action(Of String, String, Boolean),
                                      beginReasoningStream As Action(Of String),
                                      appendReasoningStreamDelta As Action(Of String, String),
                                      completeReasoningStream As Action(Of String, String),
                                      scrollTranscriptToBottom As Action,
                                      appendProtocol As Action(Of String, String),
                                      renderItem As Action(Of JsonObject),
                                      clearSessionLoginIfMatches As Action(Of String),
                                      requestAuthenticationRefresh As Action,
                                      notifyRateLimitsUpdatedUi As Action(Of JsonObject),
                                      Optional upsertRuntimeItem As Action(Of TurnItemRuntimeState) = Nothing,
                                      Optional appendTurnLifecycleMarker As Action(Of String, String, String) = Nothing,
                                      Optional upsertTurnMetadata As Action(Of String, String, String, String) = Nothing,
                                      Optional updateTokenUsageWidget As Action(Of String, String, JsonObject) = Nothing,
                                      Optional appendDiagnosticEvent As Action(Of String) = Nothing)
            If applyCurrentThreadFromThreadObject Is Nothing Then Throw New ArgumentNullException(NameOf(applyCurrentThreadFromThreadObject))
            If getCurrentThreadId Is Nothing Then Throw New ArgumentNullException(NameOf(getCurrentThreadId))
            If setCurrentThreadId Is Nothing Then Throw New ArgumentNullException(NameOf(setCurrentThreadId))
            If getCurrentTurnId Is Nothing Then Throw New ArgumentNullException(NameOf(getCurrentTurnId))
            If setCurrentTurnId Is Nothing Then Throw New ArgumentNullException(NameOf(setCurrentTurnId))
            If markThreadLastActive Is Nothing Then Throw New ArgumentNullException(NameOf(markThreadLastActive))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If scrollTranscriptToBottom Is Nothing Then Throw New ArgumentNullException(NameOf(scrollTranscriptToBottom))
            If appendProtocol Is Nothing Then Throw New ArgumentNullException(NameOf(appendProtocol))
            If renderItem Is Nothing Then Throw New ArgumentNullException(NameOf(renderItem))
            If clearSessionLoginIfMatches Is Nothing Then Throw New ArgumentNullException(NameOf(clearSessionLoginIfMatches))
            If requestAuthenticationRefresh Is Nothing Then Throw New ArgumentNullException(NameOf(requestAuthenticationRefresh))
            If notifyRateLimitsUpdatedUi Is Nothing Then Throw New ArgumentNullException(NameOf(notifyRateLimitsUpdatedUi))

            Dim paramsObject = TryCast(paramsNode, JsonObject)
            If paramsObject Is Nothing Then
                paramsObject = New JsonObject()
            End If

            Dim parsedEvent = TurnFlowEventParser.ParseNotification(methodName, paramsObject)
            Dim reduceResult = _runtimeStore.Reduce(parsedEvent)

            If reduceResult.IsSkipped Then
                LogDebug(appendProtocol, $"event_skipped method={methodName} reason={reduceResult.SkipReason}")
                If appendDiagnosticEvent IsNot Nothing Then
                    appendDiagnosticEvent($"{methodName}: {reduceResult.SkipReason}")
                End If
                Return
            End If

            If TypeOf parsedEvent Is ThreadStartedNotificationEvent Then
                Dim threadStarted = DirectCast(parsedEvent, ThreadStartedNotificationEvent)
                If threadStarted.ThreadObject IsNot Nothing Then
                    applyCurrentThreadFromThreadObject(threadStarted.ThreadObject)
                ElseIf Not String.IsNullOrWhiteSpace(threadStarted.ThreadId) Then
                    setCurrentThreadId(threadStarted.ThreadId)
                End If

                If Not String.IsNullOrWhiteSpace(getCurrentThreadId()) Then
                    markThreadLastActive(getCurrentThreadId())
                End If

                appendSystemMessage($"Thread started: {If(String.IsNullOrWhiteSpace(getCurrentThreadId()), threadStarted.ThreadId, getCurrentThreadId())}")
                Return
            End If

            If TypeOf parsedEvent Is ThreadArchivedNotificationEvent Then
                Dim threadArchived = DirectCast(parsedEvent, ThreadArchivedNotificationEvent)
                appendSystemMessage($"Thread archived: {threadArchived.ThreadId}")
                Return
            End If

            If TypeOf parsedEvent Is ThreadUnarchivedNotificationEvent Then
                Dim threadUnarchived = DirectCast(parsedEvent, ThreadUnarchivedNotificationEvent)
                appendSystemMessage($"Thread unarchived: {threadUnarchived.ThreadId}")
                Return
            End If

            If TypeOf parsedEvent Is ThreadStatusChangedNotificationEvent Then
                Dim statusChanged = DirectCast(parsedEvent, ThreadStatusChangedNotificationEvent)
                If Not String.IsNullOrWhiteSpace(statusChanged.ThreadId) Then
                    markThreadLastActive(statusChanged.ThreadId)
                End If

                If appendDiagnosticEvent IsNot Nothing Then
                    appendDiagnosticEvent($"thread/status/changed thread={statusChanged.ThreadId} status={statusChanged.StatusSummary}")
                End If
                Return
            End If

            If TypeOf parsedEvent Is AccountLoginCompletedNotificationEvent Then
                Dim accountLoginCompleted = DirectCast(parsedEvent, AccountLoginCompletedNotificationEvent)
                clearSessionLoginIfMatches(accountLoginCompleted.LoginId)
                If accountLoginCompleted.Success Then
                    appendSystemMessage("Account login completed.")
                Else
                    appendSystemMessage($"Account login failed: {accountLoginCompleted.ErrorMessage}")
                End If

                requestAuthenticationRefresh()
                Return
            End If

            If TypeOf parsedEvent Is AccountUpdatedNotificationEvent Then
                requestAuthenticationRefresh()
                Return
            End If

            If TypeOf parsedEvent Is AccountRateLimitsUpdatedNotificationEvent Then
                Dim rateLimitsUpdated = DirectCast(parsedEvent, AccountRateLimitsUpdatedNotificationEvent)
                notifyRateLimitsUpdatedUi(rateLimitsUpdated.Payload)
                Return
            End If

            If TypeOf parsedEvent Is ModelReroutedNotificationEvent Then
                Dim modelRerouted = DirectCast(parsedEvent, ModelReroutedNotificationEvent)
                appendSystemMessage($"Model rerouted: {modelRerouted.FromModel} -> {modelRerouted.ToModel} ({modelRerouted.Reason})")
                Return
            End If

            If TypeOf parsedEvent Is AppListUpdatedNotificationEvent Then
                Dim appListUpdated = DirectCast(parsedEvent, AppListUpdatedNotificationEvent)
                If appendDiagnosticEvent IsNot Nothing Then
                    appendDiagnosticEvent($"app/list/updated apps={appListUpdated.AppCount.ToString(CultureInfo.InvariantCulture)}")
                End If
                Return
            End If

            If TypeOf parsedEvent Is McpOauthLoginCompletedNotificationEvent Then
                Dim oauthCompleted = DirectCast(parsedEvent, McpOauthLoginCompletedNotificationEvent)
                If oauthCompleted.Success Then
                    appendSystemMessage($"MCP OAuth completed: {oauthCompleted.Name}")
                Else
                    appendSystemMessage($"MCP OAuth failed: {oauthCompleted.Name} ({oauthCompleted.ErrorMessage})")
                End If
                Return
            End If

            If TypeOf parsedEvent Is WindowsSandboxSetupCompletedNotificationEvent Then
                Dim sandboxCompleted = DirectCast(parsedEvent, WindowsSandboxSetupCompletedNotificationEvent)
                If appendDiagnosticEvent IsNot Nothing Then
                    appendDiagnosticEvent($"windowsSandbox/setupCompleted mode={sandboxCompleted.Mode} success={sandboxCompleted.Success.ToString(CultureInfo.InvariantCulture)} error={sandboxCompleted.ErrorMessage}")
                End If
                Return
            End If

            If TypeOf parsedEvent Is FuzzyFileSearchSessionUpdatedNotificationEvent Then
                Dim fuzzyUpdated = DirectCast(parsedEvent, FuzzyFileSearchSessionUpdatedNotificationEvent)
                If appendDiagnosticEvent IsNot Nothing Then
                    appendDiagnosticEvent($"fuzzyFileSearch/sessionUpdated session={fuzzyUpdated.SessionId} files={fuzzyUpdated.FileCount.ToString(CultureInfo.InvariantCulture)} query={fuzzyUpdated.Query}")
                End If
                Return
            End If

            If TypeOf parsedEvent Is FuzzyFileSearchSessionCompletedNotificationEvent Then
                Dim fuzzyCompleted = DirectCast(parsedEvent, FuzzyFileSearchSessionCompletedNotificationEvent)
                If appendDiagnosticEvent IsNot Nothing Then
                    appendDiagnosticEvent($"fuzzyFileSearch/sessionCompleted session={fuzzyCompleted.SessionId}")
                End If
                Return
            End If

            If TypeOf parsedEvent Is GenericServerNotificationEvent Then
                LogDebug(appendProtocol, $"event_generic method={methodName}")
                If appendDiagnosticEvent IsNot Nothing Then
                    Dim normalizedMethod = If(methodName, String.Empty).Trim()
                    If _reportedUnknownMethods.Add(normalizedMethod) Then
                        appendDiagnosticEvent($"Unhandled notification method: {methodName}")
                    End If
                End If
                Return
            End If

            If TypeOf parsedEvent Is TurnStartedEvent Then
                Dim turnStarted = DirectCast(parsedEvent, TurnStartedEvent)
                Dim effectiveThreadId = If(reduceResult.UpdatedTurn?.ThreadId, turnStarted.ThreadId)
                Dim effectiveTurnId = If(reduceResult.UpdatedTurn?.TurnId, turnStarted.TurnId)

                If Not String.IsNullOrWhiteSpace(effectiveThreadId) Then
                    setCurrentThreadId(effectiveThreadId)
                    markThreadLastActive(effectiveThreadId)
                End If

                If Not String.IsNullOrWhiteSpace(effectiveTurnId) Then
                    setCurrentTurnId(effectiveTurnId)
                End If

                appendSystemMessage($"Turn started: {getCurrentTurnId()}")
                If appendTurnLifecycleMarker IsNot Nothing Then
                    appendTurnLifecycleMarker(effectiveThreadId, effectiveTurnId, "started")
                End If
                Return
            End If

            If TypeOf parsedEvent Is TurnCompletedEvent Then
                Dim turnCompleted = DirectCast(parsedEvent, TurnCompletedEvent)
                Dim effectiveThreadId = If(reduceResult.UpdatedTurn?.ThreadId, turnCompleted.ThreadId)
                Dim effectiveTurnId = If(reduceResult.UpdatedTurn?.TurnId, turnCompleted.TurnId)

                If Not String.IsNullOrWhiteSpace(effectiveThreadId) Then
                    markThreadLastActive(effectiveThreadId)
                ElseIf Not String.IsNullOrWhiteSpace(getCurrentThreadId()) Then
                    markThreadLastActive(getCurrentThreadId())
                End If

                If StringComparer.Ordinal.Equals(effectiveTurnId, getCurrentTurnId()) Then
                    setCurrentTurnId(String.Empty)
                End If

                Dim status = NormalizeStatusForDisplay(turnCompleted.Status)
                appendSystemMessage($"Turn completed: {effectiveTurnId} ({status})")
                If appendTurnLifecycleMarker IsNot Nothing Then
                    appendTurnLifecycleMarker(effectiveThreadId, effectiveTurnId, status)
                End If

                If StringComparer.OrdinalIgnoreCase.Equals(status, "failed") AndAlso
                   appendDiagnosticEvent IsNot Nothing Then
                    Dim failureMessage = If(String.IsNullOrWhiteSpace(turnCompleted.ErrorMessage),
                                            "Turn failed.",
                                            $"Turn failed: {turnCompleted.ErrorMessage}")
                    appendDiagnosticEvent(failureMessage)
                End If

                Return
            End If

            If TypeOf parsedEvent Is TurnDiffUpdatedEvent Then
                Dim turnDiff = DirectCast(parsedEvent, TurnDiffUpdatedEvent)
                Dim effectiveThreadId = If(reduceResult.UpdatedTurn?.ThreadId, turnDiff.ThreadId)
                Dim effectiveTurnId = If(reduceResult.UpdatedTurn?.TurnId, turnDiff.TurnId)
                If upsertTurnMetadata IsNot Nothing Then
                    upsertTurnMetadata(effectiveThreadId, effectiveTurnId, "diff", turnDiff.SummaryText)
                End If
                Return
            End If

            If TypeOf parsedEvent Is TurnPlanUpdatedEvent Then
                Dim turnPlan = DirectCast(parsedEvent, TurnPlanUpdatedEvent)
                Dim effectiveThreadId = If(reduceResult.UpdatedTurn?.ThreadId, turnPlan.ThreadId)
                Dim effectiveTurnId = If(reduceResult.UpdatedTurn?.TurnId, turnPlan.TurnId)
                If upsertTurnMetadata IsNot Nothing Then
                    upsertTurnMetadata(effectiveThreadId, effectiveTurnId, "plan", turnPlan.SummaryText)
                End If
                Return
            End If

            If TypeOf parsedEvent Is ThreadTokenUsageUpdatedEvent Then
                Dim tokenUsage = DirectCast(parsedEvent, ThreadTokenUsageUpdatedEvent)
                Dim effectiveThreadId = If(reduceResult.UpdatedThread?.ThreadId, tokenUsage.ThreadId)
                Dim effectiveTurnId = If(reduceResult.UpdatedTurn?.TurnId, tokenUsage.TurnId)
                If updateTokenUsageWidget IsNot Nothing Then
                    updateTokenUsageWidget(effectiveThreadId, effectiveTurnId, tokenUsage.TokenUsage)
                End If
                Return
            End If

            If TypeOf parsedEvent Is ErrorEvent Then
                Dim turnError = DirectCast(parsedEvent, ErrorEvent)
                Dim effectiveTurnId = If(reduceResult.UpdatedTurn?.TurnId, turnError.TurnId)
                If Not String.IsNullOrWhiteSpace(turnError.Message) Then
                    appendSystemMessage($"Turn error: {turnError.Message}")
                    If appendDiagnosticEvent IsNot Nothing Then
                        appendDiagnosticEvent($"Turn error ({effectiveTurnId}): {turnError.Message}")
                    End If
                End If
                Return
            End If

            Dim updatedItem = reduceResult.UpdatedItem
            If updatedItem IsNot Nothing Then
                If reduceResult.IsFirstSeenItem Then
                    LogDebug(appendProtocol,
                             $"item_first_seen threadId={updatedItem.ThreadId} turnId={updatedItem.TurnId} itemId={updatedItem.ItemId} type={updatedItem.ItemType} phase={updatedItem.AgentMessagePhase}")
                End If

                If TypeOf parsedEvent Is ItemDeltaEvent Then
                    Dim deltaEvent = DirectCast(parsedEvent, ItemDeltaEvent)
                    LogDebug(appendProtocol,
                             $"item_delta method={deltaEvent.MethodName} threadId={updatedItem.ThreadId} turnId={updatedItem.TurnId} itemId={updatedItem.ItemId} delta_len={reduceResult.DeltaAppendedLength.ToString(CultureInfo.InvariantCulture)} preview={reduceResult.DeltaPreview}")
                End If

                If TypeOf parsedEvent Is ItemLifecycleEvent Then
                    Dim lifecycleEvent = DirectCast(parsedEvent, ItemLifecycleEvent)
                    If lifecycleEvent.IsCompleted Then
                        LogDebug(appendProtocol,
                                 $"item_completion itemId={updatedItem.ItemId} type={updatedItem.ItemType} replaced={reduceResult.CompletionReplaced.ToString(CultureInfo.InvariantCulture)}")
                    End If
                End If

                If upsertRuntimeItem IsNot Nothing Then
                    upsertRuntimeItem(updatedItem)
                ElseIf updatedItem.RawItemPayload IsNot Nothing Then
                    renderItem(updatedItem.RawItemPayload)
                End If

                scrollTranscriptToBottom()
            End If
        End Sub

        Public Sub HandleServerRequest(request As RpcServerRequest,
                                       appendProtocol As Action(Of String, String),
                                       Optional upsertRuntimeItem As Action(Of TurnItemRuntimeState) = Nothing,
                                       Optional appendDiagnosticEvent As Action(Of String) = Nothing)
            If request Is Nothing Then
                Return
            End If

            If appendProtocol Is Nothing Then
                Throw New ArgumentNullException(NameOf(appendProtocol))
            End If

            Dim parsedEvent = TurnFlowEventParser.ParseServerRequest(request)
            Dim reduceResult = _runtimeStore.Reduce(parsedEvent)

            If reduceResult.IsSkipped Then
                LogDebug(appendProtocol, $"approval_skipped method={request.MethodName} reason={reduceResult.SkipReason}")
                If appendDiagnosticEvent IsNot Nothing Then
                    appendDiagnosticEvent(reduceResult.SkipReason)
                End If
                Return
            End If

            If TypeOf parsedEvent Is GenericServerRequestEvent Then
                LogDebug(appendProtocol, $"server_request_generic method={request.MethodName}")
                Dim normalizedMethod = If(request.MethodName, String.Empty).Trim()
                If StringComparer.Ordinal.Equals(normalizedMethod, ToolRequestUserInputMethod) OrElse
                   StringComparer.Ordinal.Equals(normalizedMethod, "item/tool/call") OrElse
                   StringComparer.Ordinal.Equals(normalizedMethod, "account/chatgptAuthTokens/refresh") Then
                    Return
                End If

                If appendDiagnosticEvent IsNot Nothing Then
                    If _reportedUnknownMethods.Add($"request:{normalizedMethod}") Then
                        appendDiagnosticEvent($"Unhandled server request method: {request.MethodName}")
                    End If
                End If
                Return
            End If

            If reduceResult.PendingApproval IsNot Nothing Then
                Dim pending = reduceResult.PendingApproval
                LogDebug(appendProtocol,
                         $"approval_requested method={pending.MethodName} threadId={pending.ThreadId} turnId={pending.TurnId} itemId={pending.ItemId} requestId={pending.RequestIdKey}")
            End If

            If reduceResult.UpdatedItem IsNot Nothing AndAlso upsertRuntimeItem IsNot Nothing Then
                upsertRuntimeItem(reduceResult.UpdatedItem)
            End If
        End Sub

        Public Sub HandleApprovalResolved(requestId As JsonNode,
                                          decision As String,
                                          appendProtocol As Action(Of String, String),
                                          Optional upsertRuntimeItem As Action(Of TurnItemRuntimeState) = Nothing,
                                          Optional appendDiagnosticEvent As Action(Of String) = Nothing)
            If appendProtocol Is Nothing Then
                Throw New ArgumentNullException(NameOf(appendProtocol))
            End If

            Dim reduceResult = _runtimeStore.ResolveApprovalDecision(requestId, decision)
            If reduceResult.IsSkipped Then
                LogDebug(appendProtocol, $"approval_resolve_skipped reason={reduceResult.SkipReason}")
                If appendDiagnosticEvent IsNot Nothing Then
                    appendDiagnosticEvent(reduceResult.SkipReason)
                End If
                Return
            End If

            Dim resolved = reduceResult.ResolvedApproval
            If resolved IsNot Nothing Then
                LogDebug(appendProtocol,
                         $"approval_resolved method={resolved.MethodName} threadId={resolved.ThreadId} turnId={resolved.TurnId} itemId={resolved.ItemId} requestId={resolved.RequestIdKey} decision={resolved.Decision}")
            End If

            If reduceResult.UpdatedItem IsNot Nothing AndAlso upsertRuntimeItem IsNot Nothing Then
                upsertRuntimeItem(reduceResult.UpdatedItem)
            End If
        End Sub

        Private Shared Sub LogDebug(appendProtocol As Action(Of String, String), message As String)
            If appendProtocol Is Nothing Then
                Return
            End If

            appendProtocol("debug", message)
        End Sub

        Private Shared Function NormalizeStatusForDisplay(status As String) As String
            If String.IsNullOrWhiteSpace(status) Then
                Return "completed"
            End If

            Return status.Trim()
        End Function

        Private Shared Function ReadObject(obj As JsonObject, key As String) As JsonObject
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonObject)
        End Function

        Private Shared Function ReadStringFirst(obj As JsonObject, ParamArray keys() As String) As String
            If obj Is Nothing OrElse keys Is Nothing Then
                Return String.Empty
            End If

            For Each key In keys
                If String.IsNullOrWhiteSpace(key) Then
                    Continue For
                End If

                Dim node As JsonNode = Nothing
                If obj.TryGetPropertyValue(key, node) AndAlso node IsNot Nothing Then
                    Dim jsonValue = TryCast(node, JsonValue)
                    If jsonValue IsNot Nothing Then
                        Dim stringValue As String = Nothing
                        If jsonValue.TryGetValue(Of String)(stringValue) Then
                            Return If(stringValue, String.Empty).Trim()
                        End If
                    End If
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Function ReadBoolean(obj As JsonObject,
                                            key As String,
                                            fallback As Boolean) As Boolean
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return fallback
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) OrElse node Is Nothing Then
                Return fallback
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue Is Nothing Then
                Return fallback
            End If

            Dim boolValue As Boolean
            If jsonValue.TryGetValue(Of Boolean)(boolValue) Then
                Return boolValue
            End If

            Dim stringValue As String = Nothing
            If jsonValue.TryGetValue(Of String)(stringValue) AndAlso
               Not String.IsNullOrWhiteSpace(stringValue) Then
                Dim parsed As Boolean
                If Boolean.TryParse(stringValue, parsed) Then
                    Return parsed
                End If
            End If

            Return fallback
        End Function

        Private Shared Function NormalizeDispatchValue(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Return value.Trim()
        End Function

        Private Shared Sub AddValue(target As ICollection(Of String), value As String)
            If target Is Nothing Then
                Return
            End If

            Dim normalized = NormalizeDispatchValue(value)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return
            End If

            target.Add(normalized)
        End Sub
    End Class
End Namespace
