Imports System.Collections.Generic
Imports System.Text.Json.Nodes

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class SessionNotificationCoordinator
        Private ReadOnly _streamingAgentItemIds As New HashSet(Of String)(StringComparer.Ordinal)
        Private ReadOnly _streamingCommentaryAgentItemIds As New HashSet(Of String)(StringComparer.Ordinal)

        Public Sub ResetStreamingAgentItems()
            _streamingAgentItemIds.Clear()
            _streamingCommentaryAgentItemIds.Clear()
        End Sub

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
                                      notifyRateLimitsUpdatedUi As Action(Of JsonObject)) 
            If applyCurrentThreadFromThreadObject Is Nothing Then Throw New ArgumentNullException(NameOf(applyCurrentThreadFromThreadObject))
            If getCurrentThreadId Is Nothing Then Throw New ArgumentNullException(NameOf(getCurrentThreadId))
            If setCurrentThreadId Is Nothing Then Throw New ArgumentNullException(NameOf(setCurrentThreadId))
            If getCurrentTurnId Is Nothing Then Throw New ArgumentNullException(NameOf(getCurrentTurnId))
            If setCurrentTurnId Is Nothing Then Throw New ArgumentNullException(NameOf(setCurrentTurnId))
            If markThreadLastActive Is Nothing Then Throw New ArgumentNullException(NameOf(markThreadLastActive))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If appendTranscript Is Nothing Then Throw New ArgumentNullException(NameOf(appendTranscript))
            If beginAssistantStream Is Nothing Then Throw New ArgumentNullException(NameOf(beginAssistantStream))
            If appendAssistantStreamDelta Is Nothing Then Throw New ArgumentNullException(NameOf(appendAssistantStreamDelta))
            If completeAssistantStream Is Nothing Then Throw New ArgumentNullException(NameOf(completeAssistantStream))
            If beginReasoningStream Is Nothing Then Throw New ArgumentNullException(NameOf(beginReasoningStream))
            If appendReasoningStreamDelta Is Nothing Then Throw New ArgumentNullException(NameOf(appendReasoningStreamDelta))
            If completeReasoningStream Is Nothing Then Throw New ArgumentNullException(NameOf(completeReasoningStream))
            If scrollTranscriptToBottom Is Nothing Then Throw New ArgumentNullException(NameOf(scrollTranscriptToBottom))
            If appendProtocol Is Nothing Then Throw New ArgumentNullException(NameOf(appendProtocol))
            If renderItem Is Nothing Then Throw New ArgumentNullException(NameOf(renderItem))
            If clearSessionLoginIfMatches Is Nothing Then Throw New ArgumentNullException(NameOf(clearSessionLoginIfMatches))
            If requestAuthenticationRefresh Is Nothing Then Throw New ArgumentNullException(NameOf(requestAuthenticationRefresh))
            If notifyRateLimitsUpdatedUi Is Nothing Then Throw New ArgumentNullException(NameOf(notifyRateLimitsUpdatedUi))

            Dim paramsObject = AsObject(paramsNode)

            Select Case methodName
                Case "thread/started"
                    Dim threadObject = GetPropertyObject(paramsObject, "thread")
                    If threadObject IsNot Nothing Then
                        applyCurrentThreadFromThreadObject(threadObject)
                        appendSystemMessage($"Thread started: {getCurrentThreadId()}")
                    End If

                Case "turn/started"
                    Dim turnObject = GetPropertyObject(paramsObject, "turn")
                    Dim threadId = GetPropertyString(paramsObject, "threadId")
                    If Not String.IsNullOrWhiteSpace(threadId) Then
                        setCurrentThreadId(threadId)
                        markThreadLastActive(threadId)
                    End If

                    If turnObject IsNot Nothing Then
                        Dim turnId = GetPropertyString(turnObject, "id")
                        If Not String.IsNullOrWhiteSpace(turnId) Then
                            setCurrentTurnId(turnId)
                        End If
                    End If

                    appendSystemMessage($"Turn started: {getCurrentTurnId()}")

                Case "turn/completed"
                    Dim turnObject = GetPropertyObject(paramsObject, "turn")
                    Dim threadId = GetPropertyString(paramsObject, "threadId")
                    If Not String.IsNullOrWhiteSpace(threadId) Then
                        markThreadLastActive(threadId)
                    ElseIf Not String.IsNullOrWhiteSpace(getCurrentThreadId()) Then
                        markThreadLastActive(getCurrentThreadId())
                    End If

                    Dim completedTurnId = GetPropertyString(turnObject, "id")
                    Dim status = GetPropertyString(turnObject, "status")
                    If StringComparer.Ordinal.Equals(completedTurnId, getCurrentTurnId()) Then
                        setCurrentTurnId(String.Empty)
                    End If

                    _streamingAgentItemIds.Clear()
                    _streamingCommentaryAgentItemIds.Clear()
                    appendSystemMessage($"Turn completed: {completedTurnId} ({status})")

                Case "item/started"
                    Dim itemObject = GetPropertyObject(paramsObject, "item")
                    If itemObject IsNot Nothing Then
                        Dim itemType = GetPropertyString(itemObject, "type")
                        If StringComparer.Ordinal.Equals(itemType, "agentMessage") Then
                            Dim itemId = GetPropertyString(itemObject, "id")
                            Dim isCommentary = IsCommentaryAgentMessage(itemObject)
                            If Not String.IsNullOrWhiteSpace(itemId) Then
                                _streamingAgentItemIds.Add(itemId)
                                If isCommentary Then
                                    _streamingCommentaryAgentItemIds.Add(itemId)
                                Else
                                    _streamingCommentaryAgentItemIds.Remove(itemId)
                                End If
                                beginAssistantStream(itemId, isCommentary)
                            End If
                        ElseIf StringComparer.Ordinal.Equals(itemType, "reasoning") Then
                            beginReasoningStream(GetPropertyString(itemObject, "id"))
                        End If
                    End If

                Case "item/agentMessage/delta"
                    Dim delta = GetPropertyString(paramsObject, "delta")
                    Dim itemId = GetPropertyString(paramsObject, "itemId")
                    Dim streamItemId = itemId

                    If String.IsNullOrWhiteSpace(streamItemId) Then
                        Dim singleStreamingId As String = Nothing
                        If TryGetSingleStreamingAgentItemId(singleStreamingId) Then
                            streamItemId = singleStreamingId
                        Else
                            streamItemId = "live"
                        End If
                    ElseIf Not _streamingAgentItemIds.Contains(streamItemId) Then
                        Dim singleStreamingId As String = Nothing
                        If TryGetSingleStreamingAgentItemId(singleStreamingId) AndAlso
                           StringComparer.Ordinal.Equals(singleStreamingId, "live") Then
                            ' Keep using the original fallback stream instead of creating a duplicate bubble
                            ' when early deltas arrive without an itemId and later deltas include it.
                            streamItemId = singleStreamingId
                        End If
                    End If

                    If Not _streamingAgentItemIds.Contains(streamItemId) Then
                        _streamingAgentItemIds.Add(streamItemId)
                        beginAssistantStream(streamItemId, _streamingCommentaryAgentItemIds.Contains(streamItemId))
                    End If

                    appendAssistantStreamDelta(streamItemId, delta)
                    scrollTranscriptToBottom()

                Case "item/completed"
                    Dim itemObject = GetPropertyObject(paramsObject, "item")
                    If itemObject IsNot Nothing Then
                        Dim itemType = GetPropertyString(itemObject, "type")
                        Dim itemId = GetPropertyString(itemObject, "id")

                        If StringComparer.Ordinal.Equals(itemType, "agentMessage") Then
                            Dim text = GetPropertyString(itemObject, "text")
                            Dim isCommentary = IsCommentaryAgentMessage(itemObject)
                            Dim completionStreamId = itemId
                            If String.IsNullOrWhiteSpace(completionStreamId) OrElse
                               Not _streamingAgentItemIds.Contains(completionStreamId) Then
                                Dim singleStreamingId As String = Nothing
                                If TryGetSingleStreamingAgentItemId(singleStreamingId) Then
                                    completionStreamId = singleStreamingId
                                End If
                            End If

                            If Not String.IsNullOrWhiteSpace(completionStreamId) AndAlso
                               _streamingAgentItemIds.Contains(completionStreamId) Then
                                _streamingAgentItemIds.Remove(completionStreamId)
                                If isCommentary Then
                                    _streamingCommentaryAgentItemIds.Add(completionStreamId)
                                End If
                                completeAssistantStream(completionStreamId, text, isCommentary)
                                _streamingCommentaryAgentItemIds.Remove(completionStreamId)
                            ElseIf Not String.IsNullOrWhiteSpace(text) Then
                                If isCommentary Then
                                    completeAssistantStream(String.Empty, text, True)
                                Else
                                    appendTranscript("assistant", text)
                                End If
                            End If
                        ElseIf StringComparer.Ordinal.Equals(itemType, "reasoning") Then
                            completeReasoningStream(itemId, GetPropertyString(itemObject, "text"))
                        Else
                            renderItem(itemObject)
                        End If
                    End If

                Case "item/commandExecution/outputDelta"
                    appendProtocol("cmd", GetPropertyString(paramsObject, "delta"))

                Case "item/fileChange/outputDelta"
                    appendProtocol("file", GetPropertyString(paramsObject, "delta"))

                Case "item/reasoning/textDelta"
                    Dim reasoningDelta = GetPropertyString(paramsObject, "delta")
                    Dim reasoningItemId = GetPropertyString(paramsObject, "itemId")
                    beginReasoningStream(reasoningItemId)
                    appendReasoningStreamDelta(reasoningItemId, reasoningDelta)
                    scrollTranscriptToBottom()
                    appendProtocol("reason", reasoningDelta)

                Case "error"
                    Dim errorObject = GetPropertyObject(paramsObject, "error")
                    Dim message = GetPropertyString(errorObject, "message", "Unknown error")
                    appendSystemMessage($"Turn error: {message}")

                Case "account/login/completed"
                    Dim success = GetPropertyBoolean(paramsObject, "success", False)
                    Dim loginId = GetPropertyString(paramsObject, "loginId")
                    Dim [error] = GetPropertyString(paramsObject, "error")

                    clearSessionLoginIfMatches(loginId)

                    If success Then
                        appendSystemMessage("Account login completed.")
                    Else
                        appendSystemMessage($"Account login failed: {[error]}")
                    End If

                    requestAuthenticationRefresh()

                Case "account/updated"
                    requestAuthenticationRefresh()

                Case "account/rateLimits/updated"
                    If paramsObject IsNot Nothing Then
                        notifyRateLimitsUpdatedUi(paramsObject)
                    End If

                Case "model/rerouted"
                    Dim fromModel = GetPropertyString(paramsObject, "fromModel")
                    Dim toModel = GetPropertyString(paramsObject, "toModel")
                    Dim reason = GetPropertyString(paramsObject, "reason")
                    appendSystemMessage($"Model rerouted: {fromModel} -> {toModel} ({reason})")

                Case Else
                    ' Keep unsupported notifications in protocol log only.
            End Select
        End Sub

        Private Function TryGetSingleStreamingAgentItemId(ByRef itemId As String) As Boolean
            itemId = Nothing
            If _streamingAgentItemIds.Count <> 1 Then
                Return False
            End If

            For Each existingId In _streamingAgentItemIds
                itemId = existingId
                Return True
            Next

            Return False
        End Function
    End Class
End Namespace
