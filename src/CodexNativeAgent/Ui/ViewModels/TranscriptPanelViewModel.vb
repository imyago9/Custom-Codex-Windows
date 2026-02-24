Imports System.Collections.Generic
Imports System.Globalization
Imports System.Collections.ObjectModel
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Text.Json.Nodes
Imports System.Windows
Imports System.Windows.Media
Imports CodexNativeAgent.Ui.Coordinators
Imports CodexNativeAgent.Ui.Mvvm
Imports CodexNativeAgent.Ui.ViewModels.Transcript

Namespace CodexNativeAgent.Ui.ViewModels
    Public NotInheritable Class TranscriptPanelViewModel
        Inherits ViewModelBase

        Private Const MaxLogChars As Integer = 500000
        Private Const MaxTranscriptEntries As Integer = 2000

        Private _transcriptText As String = String.Empty
        Private _protocolText As String = String.Empty
        Private ReadOnly _fullTranscriptBuilder As New StringBuilder()
        Private ReadOnly _fullProtocolBuilder As New StringBuilder()
        Private _loadingText As String = "Loading thread..."
        Private _loadingOverlayVisibility As Visibility = Visibility.Collapsed
        Private _collapseCommandDetailsByDefault As Boolean
        Private _showEventDotsInTranscript As Boolean
        Private _showSystemDotsInTranscript As Boolean
        Private _tokenUsageText As String = String.Empty
        Private _tokenUsageVisibility As Visibility = Visibility.Collapsed
        Private ReadOnly _items As New ObservableCollection(Of TranscriptEntryViewModel)()
        Private ReadOnly _runtimeEntriesByKey As New Dictionary(Of String, TranscriptEntryViewModel)(StringComparer.Ordinal)
        Private ReadOnly _activeAssistantStreams As New Dictionary(Of String, TranscriptEntryViewModel)(StringComparer.Ordinal)
        Private ReadOnly _activeAssistantStreamBuffers As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Private ReadOnly _activeAssistantRawPrefixes As New HashSet(Of String)(StringComparer.Ordinal)
        Private ReadOnly _assistantReasoningChainStreamIds As New HashSet(Of String)(StringComparer.Ordinal)
        Private ReadOnly _activeReasoningStreams As New Dictionary(Of String, TranscriptEntryViewModel)(StringComparer.Ordinal)
        Private ReadOnly _activeReasoningStreamBuffers As New Dictionary(Of String, String)(StringComparer.Ordinal)

        Public Property TranscriptText As String
            Get
                Return _transcriptText
            End Get
            Set(value As String)
                SetProperty(_transcriptText, If(value, String.Empty))
            End Set
        End Property

        Public Property ProtocolText As String
            Get
                Return _protocolText
            End Get
            Set(value As String)
                SetProperty(_protocolText, If(value, String.Empty))
            End Set
        End Property

        Public ReadOnly Property FullTranscriptText As String
            Get
                Return _fullTranscriptBuilder.ToString()
            End Get
        End Property

        Public ReadOnly Property FullProtocolText As String
            Get
                Return _fullProtocolBuilder.ToString()
            End Get
        End Property

        Public Property LoadingText As String
            Get
                Return _loadingText
            End Get
            Set(value As String)
                SetProperty(_loadingText, If(value, "Loading thread..."))
            End Set
        End Property

        Public Property LoadingOverlayVisibility As Visibility
            Get
                Return _loadingOverlayVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_loadingOverlayVisibility, value)
            End Set
        End Property

        Public ReadOnly Property Items As ObservableCollection(Of TranscriptEntryViewModel)
            Get
                Return _items
            End Get
        End Property

        Public Property CollapseCommandDetailsByDefault As Boolean
            Get
                Return _collapseCommandDetailsByDefault
            End Get
            Set(value As Boolean)
                SetProperty(_collapseCommandDetailsByDefault, value)
            End Set
        End Property

        Public Property ShowEventDotsInTranscript As Boolean
            Get
                Return _showEventDotsInTranscript
            End Get
            Set(value As Boolean)
                If SetProperty(_showEventDotsInTranscript, value) Then
                    RefreshTimelineDotRowVisibility()
                End If
            End Set
        End Property

        Public Property ShowSystemDotsInTranscript As Boolean
            Get
                Return _showSystemDotsInTranscript
            End Get
            Set(value As Boolean)
                If SetProperty(_showSystemDotsInTranscript, value) Then
                    RefreshTimelineDotRowVisibility()
                End If
            End Set
        End Property

        Public Property TokenUsageText As String
            Get
                Return _tokenUsageText
            End Get
            Set(value As String)
                SetProperty(_tokenUsageText, If(value, String.Empty))
                TokenUsageVisibility = If(String.IsNullOrWhiteSpace(_tokenUsageText), Visibility.Collapsed, Visibility.Visible)
            End Set
        End Property

        Public Property TokenUsageVisibility As Visibility
            Get
                Return _tokenUsageVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_tokenUsageVisibility, value)
            End Set
        End Property

        Public Sub ClearTranscript()
            TranscriptText = String.Empty
            _fullTranscriptBuilder.Clear()
            _items.Clear()
            _runtimeEntriesByKey.Clear()
            _activeAssistantStreams.Clear()
            _activeAssistantStreamBuffers.Clear()
            _activeAssistantRawPrefixes.Clear()
            _assistantReasoningChainStreamIds.Clear()
            _activeReasoningStreams.Clear()
            _activeReasoningStreamBuffers.Clear()
            TokenUsageText = String.Empty
        End Sub

        Public Sub SetTranscriptSnapshot(value As String)
            Dim text = If(value, String.Empty)
            TranscriptText = TrimLogText(text)
            _fullTranscriptBuilder.Clear()
            If text.Length > 0 Then
                _fullTranscriptBuilder.Append(text)
            End If
        End Sub

        Public Sub SetTranscriptDisplaySnapshot(entries As IEnumerable(Of TranscriptEntryDescriptor))
            _items.Clear()
            _runtimeEntriesByKey.Clear()
            _activeAssistantStreams.Clear()
            _activeAssistantStreamBuffers.Clear()
            _activeAssistantRawPrefixes.Clear()
            _assistantReasoningChainStreamIds.Clear()
            _activeReasoningStreams.Clear()
            _activeReasoningStreamBuffers.Clear()

            If entries Is Nothing Then
                Return
            End If

            For Each descriptor In entries
                AppendDescriptor(descriptor, appendToRaw:=False)
            Next
        End Sub

        Public Sub AppendTranscriptChunk(value As String)
            If String.IsNullOrEmpty(value) Then
                Return
            End If

            AppendRawTranscriptChunk(value)
        End Sub

        Public Sub AppendRoleMessage(role As String, text As String)
            Dim normalizedRole = If(role, String.Empty).Trim().ToLowerInvariant()
            Dim normalizedText = If(text, String.Empty)
            If String.IsNullOrWhiteSpace(normalizedText) Then
                Return
            End If

            AppendRawRoleLine(normalizedRole, normalizedText)

            Select Case normalizedRole
                Case "user"
                    AppendDescriptor(New TranscriptEntryDescriptor() With {
                        .Kind = "user",
                        .TimestampText = FormatLiveTimestamp(),
                        .RoleText = "You",
                        .BodyText = normalizedText
                    }, appendToRaw:=False)

                Case "user (steer)"
                    AppendDescriptor(New TranscriptEntryDescriptor() With {
                        .Kind = "user",
                        .TimestampText = FormatLiveTimestamp(),
                        .RoleText = "Steer",
                        .BodyText = normalizedText
                    }, appendToRaw:=False)

                Case "assistant"
                    AppendAssistantMessage(normalizedText, appendRaw:=False)

                Case "plan"
                    AppendPlan(normalizedText)

                Case "reasoning"
                    AppendReasoning(normalizedText)

                Case "command"
                    AppendDescriptor(New TranscriptEntryDescriptor() With {
                        .Kind = "command",
                        .TimestampText = FormatLiveTimestamp(),
                        .RoleText = "Command",
                        .BodyText = normalizedText,
                        .IsCommandLike = True
                    }, appendToRaw:=False)

                Case "filechange"
                    AppendDescriptor(New TranscriptEntryDescriptor() With {
                        .Kind = "fileChange",
                        .TimestampText = FormatLiveTimestamp(),
                        .RoleText = "Files",
                        .BodyText = normalizedText
                    }, appendToRaw:=False)

                Case "system"
                    AppendSystemMessage(normalizedText, appendRaw:=False)

                Case "error"
                    AppendErrorMessage(normalizedText, appendRaw:=False)

                Case Else
                    AppendDescriptor(New TranscriptEntryDescriptor() With {
                        .Kind = "event",
                        .TimestampText = FormatLiveTimestamp(),
                        .RoleText = normalizedRole,
                        .BodyText = normalizedText,
                        .IsMuted = True
                    }, appendToRaw:=False)
            End Select
        End Sub

        Public Sub AppendSystemMessage(message As String, Optional appendRaw As Boolean = True)
            If String.IsNullOrWhiteSpace(message) Then
                Return
            End If

            If appendRaw Then
                AppendRawRoleLine("system", message)
            End If

            If ShouldSuppressSystemMessage(message) Then
                Return
            End If

            AppendDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "system",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "System",
                .BodyText = message,
                .IsMuted = True
            }, appendToRaw:=False)
        End Sub

        Public Sub AppendErrorMessage(message As String, Optional appendRaw As Boolean = True)
            If String.IsNullOrWhiteSpace(message) Then
                Return
            End If

            If appendRaw Then
                AppendRawRoleLine("error", message)
            End If

            AppendDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "error",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Error",
                .BodyText = message,
                .IsError = True
            }, appendToRaw:=False)
        End Sub

        Public Sub AppendAssistantMessage(text As String, Optional appendRaw As Boolean = True)
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            If appendRaw Then
                AppendRawRoleLine("assistant", text)
            End If

            AppendDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "assistant",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Codex",
                .BodyText = text
            }, appendToRaw:=False)
        End Sub

        Public Sub AppendAssistantCommentaryReasoningStep(text As String, Optional appendRaw As Boolean = True)
            Dim normalizedText = If(text, String.Empty)
            If String.IsNullOrWhiteSpace(normalizedText) Then
                Return
            End If

            If appendRaw Then
                AppendRawRoleLine("assistant", normalizedText)
            End If

            Dim descriptor As New TranscriptEntryDescriptor() With {
                .Kind = "reasoning",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Reasoning",
                .BodyText = normalizedText,
                .DetailsText = String.Empty,
                .IsReasoning = True,
                .IsMuted = True
            }

            If Not ShouldDisplayDescriptor(descriptor) Then
                Return
            End If

            Dim entry = CreateEntryFromDescriptor(descriptor)
            ApplyReasoningChainStepFormatting(entry, normalizedText)
            _items.Add(entry)
            TrimEntriesIfNeeded()
        End Sub

        Public Sub BeginAssistantStream(itemId As String, Optional renderAsReasoningChainStep As Boolean = False)
            Dim streamId = NormalizeStreamId(itemId)
            If _activeAssistantStreams.ContainsKey(streamId) Then
                If renderAsReasoningChainStep Then
                    PromoteAssistantStreamToReasoningChainStep(streamId)
                End If
                Return
            End If

            Dim entry = CreateEntryFromDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = If(renderAsReasoningChainStep, "reasoning", "assistant"),
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = If(renderAsReasoningChainStep, "Reasoning", "Codex"),
                .BodyText = String.Empty,
                .IsMuted = renderAsReasoningChainStep,
                .IsReasoning = renderAsReasoningChainStep,
                .IsStreaming = True
            })
            entry.StreamingIndicatorVisibility = Visibility.Visible
            entry.StreamingIndicatorText = If(renderAsReasoningChainStep, "updating", "streaming")

            _activeAssistantStreams(streamId) = entry
            _activeAssistantStreamBuffers(streamId) = String.Empty
            If renderAsReasoningChainStep Then
                _assistantReasoningChainStreamIds.Add(streamId)
            Else
                _assistantReasoningChainStreamIds.Remove(streamId)
            End If
            _items.Add(entry)
            TrimEntriesIfNeeded()

            If _activeAssistantRawPrefixes.Add(streamId) Then
                AppendRawTranscriptChunk($"[{Date.Now:HH:mm:ss}] assistant: ")
            End If
        End Sub

        Public Sub AppendAssistantStreamDelta(itemId As String, delta As String)
            Dim streamId = NormalizeStreamId(itemId)
            If String.IsNullOrEmpty(delta) Then
                Return
            End If

            If Not _activeAssistantStreams.ContainsKey(streamId) Then
                BeginAssistantStream(streamId)
            End If

            Dim rawBuffer = String.Empty
            If _activeAssistantStreamBuffers.TryGetValue(streamId, rawBuffer) Then
                rawBuffer &= delta
            Else
                rawBuffer = delta
            End If
            _activeAssistantStreamBuffers(streamId) = rawBuffer

            ApplyAssistantStreamFormatting(_activeAssistantStreams(streamId),
                                           rawBuffer,
                                           _assistantReasoningChainStreamIds.Contains(streamId))
            AppendRawTranscriptChunk(delta)
        End Sub

        Public Sub CompleteAssistantStream(itemId As String,
                                           finalText As String,
                                           Optional renderAsReasoningChainStep As Boolean = False)
            Dim streamId = NormalizeStreamId(itemId)
            Dim existing As TranscriptEntryViewModel = Nothing
            If Not _activeAssistantStreams.TryGetValue(streamId, existing) AndAlso _activeAssistantStreams.Count = 1 Then
                For Each pair In _activeAssistantStreams
                    streamId = pair.Key
                    existing = pair.Value
                    Exit For
                Next
            End If

            If existing IsNot Nothing Then
                If renderAsReasoningChainStep Then
                    PromoteAssistantStreamToReasoningChainStep(streamId)
                End If

                Dim renderInReasoningChain = _assistantReasoningChainStreamIds.Contains(streamId)
                Dim bufferedText As String = String.Empty
                _activeAssistantStreamBuffers.TryGetValue(streamId, bufferedText)

                If existing IsNot Nothing AndAlso
                   String.IsNullOrWhiteSpace(existing.BodyText) AndAlso
                   String.IsNullOrWhiteSpace(bufferedText) AndAlso
                   Not String.IsNullOrWhiteSpace(finalText) Then
                    existing.BodyText = finalText
                End If

                If existing IsNot Nothing Then
                    Dim completedText = finalText
                    If String.IsNullOrWhiteSpace(completedText) Then
                        completedText = If(String.IsNullOrWhiteSpace(bufferedText), existing.BodyText, bufferedText)
                    End If
                    ApplyAssistantStreamFormatting(existing, completedText, renderInReasoningChain)
                    existing.StreamingIndicatorVisibility = Visibility.Collapsed
                End If

                _activeAssistantStreams.Remove(streamId)
                _activeAssistantStreamBuffers.Remove(streamId)
                _assistantReasoningChainStreamIds.Remove(streamId)
                If _activeAssistantRawPrefixes.Remove(streamId) Then
                    AppendRawTranscriptChunk(Environment.NewLine & Environment.NewLine)
                End If
                Return
            End If

            If Not String.IsNullOrWhiteSpace(finalText) Then
                If renderAsReasoningChainStep Then
                    AppendAssistantCommentaryReasoningStep(finalText)
                Else
                    AppendAssistantMessage(finalText)
                End If
            End If
        End Sub

        Public Sub AppendPlan(text As String)
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            AppendDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "plan",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Plan",
                .BodyText = text
            }, appendToRaw:=True)
        End Sub

        Public Sub AppendReasoning(text As String)
            Dim normalizedText = NormalizeReasoningPayloadText(text)
            If String.IsNullOrWhiteSpace(normalizedText) Then
                Return
            End If

            If StringComparer.Ordinal.Equals(normalizedText.Trim(), "[reasoning item]") Then
                Return
            End If

            AppendDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "reasoning",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Reasoning",
                .BodyText = normalizedText,
                .DetailsText = String.Empty,
                .IsReasoning = True,
                .IsMuted = True
            }, appendToRaw:=True, rawRole:="reasoning", rawBody:=normalizedText)
        End Sub

        Public Sub BeginReasoningStream(itemId As String)
            Dim streamId = NormalizeStreamId(itemId)
            If _activeReasoningStreams.ContainsKey(streamId) Then
                Return
            End If

            Dim entry = CreateEntryFromDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "reasoning",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Reasoning",
                .BodyText = String.Empty,
                .DetailsText = String.Empty,
                .IsReasoning = True,
                .IsMuted = True,
                .IsStreaming = True
            })
            entry.StreamingIndicatorVisibility = Visibility.Visible
            entry.StreamingIndicatorText = "thinking"

            _activeReasoningStreams(streamId) = entry
            _activeReasoningStreamBuffers(streamId) = String.Empty
            _items.Add(entry)
            TrimEntriesIfNeeded()
        End Sub

        Public Sub AppendReasoningStreamDelta(itemId As String, delta As String)
            Dim streamId = NormalizeStreamId(itemId)
            Dim normalizedDelta = NormalizeReasoningPayloadFragment(delta)
            If String.IsNullOrEmpty(normalizedDelta) Then
                Return
            End If

            If Not _activeReasoningStreams.ContainsKey(streamId) Then
                BeginReasoningStream(streamId)
            End If

            Dim entry = _activeReasoningStreams(streamId)
            Dim rawBuffer = String.Empty
            If _activeReasoningStreamBuffers.TryGetValue(streamId, rawBuffer) Then
                rawBuffer &= normalizedDelta
            Else
                rawBuffer = normalizedDelta
            End If
            _activeReasoningStreamBuffers(streamId) = rawBuffer

            ApplyReasoningMarkdownFormatting(entry, rawBuffer)
        End Sub

        Public Sub CompleteReasoningStream(itemId As String, finalText As String)
            Dim streamId = NormalizeStreamId(itemId)
            Dim entry As TranscriptEntryViewModel = Nothing
            If Not _activeReasoningStreams.TryGetValue(streamId, entry) AndAlso _activeReasoningStreams.Count = 1 Then
                For Each pair In _activeReasoningStreams
                    streamId = pair.Key
                    entry = pair.Value
                    Exit For
                Next
            End If

            If entry IsNot Nothing Then
                Dim bufferedText As String = String.Empty
                _activeReasoningStreamBuffers.TryGetValue(streamId, bufferedText)

                Dim completedText = If(String.IsNullOrWhiteSpace(finalText), bufferedText, NormalizeReasoningPayloadText(finalText))
                If Not String.IsNullOrWhiteSpace(completedText) Then
                    ApplyReasoningMarkdownFormatting(entry, completedText)
                    AppendRawRoleLine("reasoning", completedText)
                End If

                entry.StreamingIndicatorVisibility = Visibility.Collapsed

                ' Remove empty placeholder reasoning rows.
                If String.IsNullOrWhiteSpace(entry.BodyText) AndAlso String.IsNullOrWhiteSpace(entry.DetailsText) Then
                    _items.Remove(entry)
                End If

                _activeReasoningStreams.Remove(streamId)
                _activeReasoningStreamBuffers.Remove(streamId)
                Return
            End If

            If Not String.IsNullOrWhiteSpace(finalText) Then
                AppendReasoning(finalText)
            End If
        End Sub

        Public Sub AppendCommandExecution(commandText As String, status As String, output As String)
            Dim cleanCommand = If(commandText, String.Empty).Trim()
            Dim cleanStatus = If(status, String.Empty).Trim()
            Dim cleanOutput = If(output, String.Empty)

            Dim rawSummary As String = cleanCommand
            If Not String.IsNullOrWhiteSpace(cleanStatus) Then
                rawSummary = $"Command ({cleanStatus}): {cleanCommand}"
            End If
            If Not String.IsNullOrWhiteSpace(cleanOutput) Then
                rawSummary &= Environment.NewLine & cleanOutput.TrimEnd()
            End If

            AppendDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "command",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Command",
                .BodyText = If(String.IsNullOrWhiteSpace(cleanCommand), "(command)", cleanCommand),
                .SecondaryText = If(String.IsNullOrWhiteSpace(cleanStatus), String.Empty, $"status: {cleanStatus}"),
                .DetailsText = cleanOutput.Trim(),
                .IsCommandLike = True
            }, appendToRaw:=True, rawRole:="command", rawBody:=rawSummary)
        End Sub

        Public Sub AppendFileChangeSummary(status As String,
                                           count As Integer,
                                           details As String,
                                           Optional addedLineCount As Integer? = Nothing,
                                           Optional removedLineCount As Integer? = Nothing)
            Dim cleanStatus = If(status, String.Empty).Trim()
            Dim summary = $"{Math.Max(0, count)} change(s)"
            If Not String.IsNullOrWhiteSpace(cleanStatus) Then
                summary &= $" ({cleanStatus})"
            End If

            AppendDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "fileChange",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Files",
                .BodyText = summary,
                .DetailsText = If(details, String.Empty).Trim(),
                .AddedLineCount = SanitizeOptionalLineCount(addedLineCount),
                .RemovedLineCount = SanitizeOptionalLineCount(removedLineCount)
            }, appendToRaw:=True, rawRole:="fileChange", rawBody:=If(String.IsNullOrWhiteSpace(cleanStatus),
                                                                      $"{Math.Max(0, count)} change(s)",
                                                                      $"{Math.Max(0, count)} change(s), status={cleanStatus}"))
        End Sub

        Public Sub AppendUnknownItem(itemType As String, itemId As String)
            Dim body = $"{If(itemType, String.Empty)} ({If(itemId, String.Empty)})".Trim()
            If String.IsNullOrWhiteSpace(body) Then
                Return
            End If

            AppendDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "event",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Item",
                .BodyText = body,
                .IsMuted = True
            }, appendToRaw:=True, rawRole:="item", rawBody:=body)
        End Sub

        Public Sub AppendSnapshotDescriptor(descriptor As TranscriptEntryDescriptor)
            AppendDescriptor(descriptor, appendToRaw:=False)
        End Sub

        Public Sub UpsertRuntimeItem(itemState As TurnItemRuntimeState)
            If itemState Is Nothing Then
                Return
            End If

            Dim scopedKey = If(itemState.ScopedItemKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(scopedKey) Then
                Dim threadId = If(itemState.ThreadId, String.Empty).Trim()
                Dim turnId = If(itemState.TurnId, String.Empty).Trim()
                Dim itemId = If(itemState.ItemId, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(threadId) OrElse
                   String.IsNullOrWhiteSpace(turnId) OrElse
                   String.IsNullOrWhiteSpace(itemId) Then
                    Return
                End If

                scopedKey = $"{threadId}:{turnId}:{itemId}"
            End If

            If String.IsNullOrWhiteSpace(scopedKey) Then
                Return
            End If

            Dim descriptor = BuildRuntimeItemDescriptor(itemState)
            If descriptor Is Nothing Then
                Return
            End If

            UpsertRuntimeDescriptor($"item:{scopedKey}", descriptor)
        End Sub

        Public Sub UpsertTurnLifecycleMarker(threadId As String, turnId As String, status As String)
            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return
            End If

            Dim normalizedStatus = If(status, String.Empty).Trim()
            Dim descriptor = BuildTurnLifecycleDescriptorForSnapshot(normalizedTurnId,
                                                                     normalizedStatus,
                                                                     FormatLiveTimestamp())

            Dim lifecycleSlot = If(StringComparer.OrdinalIgnoreCase.Equals(normalizedStatus, "started"),
                                   "start",
                                   "end")

            ' Use turnId as the stable key so duplicate lifecycle notifications coalesce even
            ' when one notification is missing thread_id and another includes it.
            UpsertRuntimeDescriptor($"turn:lifecycle:{lifecycleSlot}:{normalizedTurnId}", descriptor)
        End Sub

        Public Sub UpsertTurnMetadataMarker(threadId As String,
                                            turnId As String,
                                            kind As String,
                                            summaryText As String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            Dim normalizedKind = If(kind, String.Empty).Trim().ToLowerInvariant()
            Dim normalizedSummary = If(summaryText, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedTurnId) OrElse
               String.IsNullOrWhiteSpace(normalizedKind) OrElse
               String.IsNullOrWhiteSpace(normalizedSummary) Then
                Return
            End If

            Dim descriptor = BuildTurnMetadataDescriptorForSnapshot(normalizedKind,
                                                                    normalizedSummary,
                                                                    FormatLiveTimestamp())

            UpsertRuntimeDescriptor($"turn:meta:{normalizedKind}:{normalizedThreadId}:{normalizedTurnId}", descriptor)
        End Sub

        Public Sub SetTokenUsageSummary(threadId As String, turnId As String, tokenUsage As JsonObject)
            If tokenUsage Is Nothing Then
                TokenUsageText = String.Empty
                Return
            End If

            Dim inputTokens = ReadLong(tokenUsage, "inputTokens", "input_tokens")
            Dim outputTokens = ReadLong(tokenUsage, "outputTokens", "output_tokens")
            Dim reasoningTokens = ReadLong(tokenUsage, "reasoningTokens", "reasoning_tokens")
            Dim cachedInputTokens = ReadLong(tokenUsage, "cachedInputTokens", "cached_input_tokens")
            Dim totalTokens = ReadLong(tokenUsage, "totalTokens", "total_tokens", "total")

            Dim parts As New List(Of String)()
            If inputTokens.HasValue Then parts.Add($"in {inputTokens.Value.ToString(CultureInfo.InvariantCulture)}")
            If outputTokens.HasValue Then parts.Add($"out {outputTokens.Value.ToString(CultureInfo.InvariantCulture)}")
            If reasoningTokens.HasValue Then parts.Add($"reasoning {reasoningTokens.Value.ToString(CultureInfo.InvariantCulture)}")
            If cachedInputTokens.HasValue Then parts.Add($"cached {cachedInputTokens.Value.ToString(CultureInfo.InvariantCulture)}")
            If totalTokens.HasValue Then parts.Add($"total {totalTokens.Value.ToString(CultureInfo.InvariantCulture)}")

            If parts.Count = 0 Then
                TokenUsageText = String.Empty
                Return
            End If

            Dim prefix = If(String.IsNullOrWhiteSpace(turnId), "Tokens", $"Tokens ({turnId})")
            TokenUsageText = $"{prefix}: {String.Join(" | ", parts)}"
        End Sub

        Public Sub AppendRuntimeDiagnosticEvent(message As String)
            Dim text = If(message, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            AppendDescriptor(New TranscriptEntryDescriptor() With {
                .Kind = "event",
                .TimestampText = FormatLiveTimestamp(),
                .RoleText = "Event",
                .BodyText = text,
                .IsMuted = True
            }, appendToRaw:=False)
        End Sub

        Public Sub ClearProtocol()
            ProtocolText = String.Empty
            _fullProtocolBuilder.Clear()
        End Sub

        Public Sub SetProtocolSnapshot(value As String)
            Dim text = If(value, String.Empty)
            ProtocolText = TrimLogText(text)
            _fullProtocolBuilder.Clear()
            If text.Length > 0 Then
                _fullProtocolBuilder.Append(text)
            End If
        End Sub

        Public Sub AppendProtocolChunk(value As String)
            If String.IsNullOrEmpty(value) Then
                Return
            End If

            _fullProtocolBuilder.Append(value)
            ProtocolText = TrimLogText(_protocolText & value)
        End Sub

        Private Sub UpsertRuntimeDescriptor(runtimeKey As String, descriptor As TranscriptEntryDescriptor)
            Dim normalizedKey = If(runtimeKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedKey) OrElse descriptor Is Nothing Then
                Return
            End If

            If Not ShouldDisplayDescriptor(descriptor) Then
                Return
            End If

            Dim replacement = CreateEntryFromDescriptor(descriptor)

            Dim existing As TranscriptEntryViewModel = Nothing
            If _runtimeEntriesByKey.TryGetValue(normalizedKey, existing) AndAlso existing IsNot Nothing Then
                Dim index = _items.IndexOf(existing)
                If index >= 0 Then
                    _items(index) = replacement
                Else
                    _items.Add(replacement)
                End If

                _runtimeEntriesByKey(normalizedKey) = replacement
                TrimEntriesIfNeeded()
                Return
            End If

            _runtimeEntriesByKey(normalizedKey) = replacement
            _items.Add(replacement)
            TrimEntriesIfNeeded()
        End Sub

        Private Function BuildRuntimeItemDescriptor(itemState As TurnItemRuntimeState) As TranscriptEntryDescriptor
            Return BuildRuntimeItemDescriptorForSnapshot(itemState, FormatLiveTimestamp())
        End Function

        Friend Shared Function BuildRuntimeItemDescriptorForSnapshot(itemState As TurnItemRuntimeState,
                                                                     Optional timestampText As String = "") As TranscriptEntryDescriptor
            If itemState Is Nothing Then
                Return Nothing
            End If

            Dim itemType = If(itemState.ItemType, String.Empty).Trim().ToLowerInvariant()
            Dim descriptor As New TranscriptEntryDescriptor() With {
                .TimestampText = If(timestampText, String.Empty)
            }

            Select Case itemType
                Case "usermessage"
                    descriptor.Kind = "user"
                    descriptor.RoleText = "You"
                    descriptor.BodyText = If(itemState.GenericText, String.Empty).Trim()

                Case "agentmessage"
                    Dim phase = If(itemState.AgentMessagePhase, String.Empty).Trim().ToLowerInvariant()
                    If StringComparer.Ordinal.Equals(phase, "commentary") Then
                        descriptor.Kind = "assistantCommentary"
                        descriptor.RoleText = "Commentary"
                        descriptor.IsMuted = True
                    ElseIf StringComparer.Ordinal.Equals(phase, "final_answer") Then
                        descriptor.Kind = "assistantFinal"
                        descriptor.RoleText = "Codex"
                    Else
                        descriptor.Kind = "assistant"
                        descriptor.RoleText = "Assistant"
                    End If
                    descriptor.BodyText = If(itemState.AgentMessageText, String.Empty)

                Case "plan"
                    descriptor.Kind = "plan"
                    descriptor.RoleText = "Plan"
                    descriptor.BodyText = If(itemState.IsCompleted AndAlso
                                             Not String.IsNullOrWhiteSpace(itemState.PlanFinalText),
                                             itemState.PlanFinalText,
                                             itemState.PlanStreamText)

                Case "reasoning"
                    descriptor.Kind = "reasoningCard"
                    descriptor.RoleText = "Reasoning"
                    descriptor.IsMuted = True
                    descriptor.IsReasoning = True
                    descriptor.UseRawReasoningLayout = True
                    descriptor.BodyText = If(String.IsNullOrWhiteSpace(itemState.ReasoningSummaryText),
                                             "Summary pending...",
                                             itemState.ReasoningSummaryText)
                    descriptor.SecondaryText = If(itemState.ReasoningSummaryParts.Count > 0,
                                                  $"{itemState.ReasoningSummaryParts.Count.ToString(CultureInfo.InvariantCulture)} summary part(s)",
                                                  String.Empty)
                    descriptor.DetailsText = If(itemState.ReasoningContentText, String.Empty)

                Case "commandexecution"
                    descriptor.Kind = "command"
                    descriptor.RoleText = "Command"
                    descriptor.IsCommandLike = True
                    descriptor.BodyText = If(String.IsNullOrWhiteSpace(itemState.CommandText),
                                             "(command)",
                                             itemState.CommandText)
                    descriptor.SecondaryText = BuildCommandSecondaryText(itemState)
                    descriptor.DetailsText = If(itemState.CommandOutputText, String.Empty).Trim()

                Case "filechange"
                    descriptor.Kind = "fileChange"
                    descriptor.RoleText = "Files"
                    descriptor.BodyText = BuildFileChangeBody(itemState)
                    descriptor.SecondaryText = BuildFileChangeSecondaryText(itemState)
                    descriptor.DetailsText = BuildFileChangeDetails(itemState)

                Case "mcptoolcall"
                    descriptor.Kind = "toolMcp"
                    descriptor.RoleText = "MCP"
                    descriptor.BodyText = BuildToolBody(itemState)
                    descriptor.SecondaryText = BuildToolSecondary(itemState)
                    descriptor.DetailsText = BuildToolDetails(itemState)

                Case "collabtoolcall"
                    descriptor.Kind = "toolCollab"
                    descriptor.RoleText = "Collaboration"
                    descriptor.BodyText = BuildToolBody(itemState)
                    descriptor.SecondaryText = BuildToolSecondary(itemState)
                    descriptor.DetailsText = BuildToolDetails(itemState)

                Case "websearch"
                    descriptor.Kind = "toolWebSearch"
                    descriptor.RoleText = "Web Search"
                    descriptor.BodyText = BuildToolBody(itemState)
                    descriptor.SecondaryText = BuildToolSecondary(itemState)
                    descriptor.DetailsText = BuildToolDetails(itemState)

                Case "imageview"
                    descriptor.Kind = "toolImageView"
                    descriptor.RoleText = "Image View"
                    descriptor.BodyText = BuildToolBody(itemState)
                    descriptor.SecondaryText = BuildToolSecondary(itemState)
                    descriptor.DetailsText = BuildToolDetails(itemState)

                Case "enteredreviewmode"
                    descriptor.Kind = "systemMarker"
                    descriptor.RoleText = "System"
                    descriptor.BodyText = "Entered review mode."
                    descriptor.IsMuted = True

                Case "exitedreviewmode"
                    descriptor.Kind = "systemMarker"
                    descriptor.RoleText = "System"
                    descriptor.BodyText = "Exited review mode."
                    descriptor.IsMuted = True

                Case "contextcompaction"
                    descriptor.Kind = "systemMarker"
                    descriptor.RoleText = "System"
                    descriptor.BodyText = "Context compaction completed."
                    descriptor.IsMuted = True

                Case Else
                    descriptor.Kind = "unknownItem"
                    descriptor.RoleText = "Item"
                    descriptor.IsMuted = True
                    descriptor.BodyText = $"{If(itemState.ItemType, "unknown")} ({itemState.ItemId})"
                    If itemState.RawItemPayload IsNot Nothing Then
                        descriptor.DetailsText = itemState.RawItemPayload.ToJsonString()
                    End If
            End Select

            descriptor.IsStreaming = Not itemState.IsCompleted AndAlso IsStreamingRuntimeKind(descriptor.Kind)
            If descriptor.IsStreaming AndAlso String.IsNullOrWhiteSpace(descriptor.BodyText) AndAlso
               String.IsNullOrWhiteSpace(descriptor.DetailsText) Then
                descriptor.BodyText = "..."
            End If

            If itemState.PendingApprovalCount > 0 Then
                If String.IsNullOrWhiteSpace(descriptor.SecondaryText) Then
                    descriptor.SecondaryText = itemState.PendingApprovalSummary
                Else
                    descriptor.SecondaryText &= $" | {itemState.PendingApprovalSummary}"
                End If
            End If

            Return descriptor
        End Function

        Friend Shared Function BuildTurnLifecycleDescriptorForSnapshot(turnId As String,
                                                                       status As String,
                                                                       Optional timestampText As String = "") As TranscriptEntryDescriptor
            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return Nothing
            End If

            Dim normalizedStatus = If(status, String.Empty).Trim()
            Dim compactStatus = normalizedStatus.Replace("-", String.Empty, StringComparison.Ordinal).
                                                 Replace("_", String.Empty, StringComparison.Ordinal).
                                                 Replace(" ", String.Empty, StringComparison.Ordinal).
                                                 ToLowerInvariant()
            Dim body As String

            Select Case compactStatus
                Case "", "completed"
                    body = $"Turn {normalizedTurnId} completed."
                Case "started"
                    body = $"Turn {normalizedTurnId} started."
                Case "interrupted"
                    body = $"Turn {normalizedTurnId} interrupted."
                Case "failed"
                    body = $"Turn {normalizedTurnId} failed."
                Case "cancelled", "canceled"
                    body = $"Turn {normalizedTurnId} canceled."
                Case "aborted"
                    body = $"Turn {normalizedTurnId} aborted."
                Case Else
                    body = $"Turn {normalizedTurnId} ended ({If(String.IsNullOrWhiteSpace(normalizedStatus), "completed", normalizedStatus)})."
            End Select

            Return New TranscriptEntryDescriptor() With {
                .Kind = "turnMarker",
                .TimestampText = If(timestampText, String.Empty),
                .RoleText = "Turn",
                .BodyText = body,
                .IsMuted = True
            }
        End Function

        Friend Shared Function BuildTurnMetadataDescriptorForSnapshot(kind As String,
                                                                      summaryText As String,
                                                                      Optional timestampText As String = "") As TranscriptEntryDescriptor
            Dim normalizedKind = If(kind, String.Empty).Trim().ToLowerInvariant()
            Dim normalizedSummary = If(summaryText, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedKind) OrElse String.IsNullOrWhiteSpace(normalizedSummary) Then
                Return Nothing
            End If

            Return New TranscriptEntryDescriptor() With {
                .Kind = If(StringComparer.Ordinal.Equals(normalizedKind, "diff"), "turnDiff", "turnPlan"),
                .TimestampText = If(timestampText, String.Empty),
                .RoleText = If(StringComparer.Ordinal.Equals(normalizedKind, "diff"), "Turn diff", "Turn plan"),
                .BodyText = normalizedSummary,
                .IsMuted = True
            }
        End Function

        Private Shared Function IsStreamingRuntimeKind(kind As String) As Boolean
            Select Case If(kind, String.Empty).Trim().ToLowerInvariant()
                Case "assistant", "assistantcommentary", "assistantfinal",
                     "plan", "reasoningcard", "command", "filechange",
                     "toolmcp", "toolcollab", "toolwebsearch", "toolimageview"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function BuildCommandSecondaryText(itemState As TurnItemRuntimeState) As String
            Dim parts As New List(Of String)()
            If Not String.IsNullOrWhiteSpace(itemState.CommandCwd) Then
                parts.Add($"cwd: {itemState.CommandCwd}")
            End If

            If Not String.IsNullOrWhiteSpace(itemState.Status) Then
                parts.Add($"status: {itemState.Status}")
            End If

            If itemState.CommandExitCode.HasValue Then
                parts.Add($"exit: {itemState.CommandExitCode.Value.ToString(CultureInfo.InvariantCulture)}")
            End If

            If itemState.CommandDurationMs.HasValue Then
                parts.Add($"duration: {itemState.CommandDurationMs.Value.ToString(CultureInfo.InvariantCulture)} ms")
            End If

            Return String.Join(" | ", parts)
        End Function

        Private Shared Function BuildFileChangeBody(itemState As TurnItemRuntimeState) As String
            Dim changeCount = If(itemState.FileChangeChanges Is Nothing, 0, itemState.FileChangeChanges.Count)
            Dim summary = $"{changeCount.ToString(CultureInfo.InvariantCulture)} change(s)"
            If Not String.IsNullOrWhiteSpace(itemState.FileChangeStatus) Then
                summary &= $" ({itemState.FileChangeStatus})"
            End If

            Return summary
        End Function

        Private Shared Function BuildFileChangeSecondaryText(itemState As TurnItemRuntimeState) As String
            If String.IsNullOrWhiteSpace(itemState.Status) Then
                Return String.Empty
            End If

            Return $"status: {itemState.Status}"
        End Function

        Private Shared Function BuildFileChangeDetails(itemState As TurnItemRuntimeState) As String
            Dim lines As New List(Of String)()
            Dim changes = itemState.FileChangeChanges
            If changes IsNot Nothing Then
                Dim shown = 0
                For Each changeNode In changes
                    If shown >= 24 Then
                        Exit For
                    End If

                    Dim changeObject = TryCast(changeNode, JsonObject)
                    If changeObject Is Nothing Then
                        Continue For
                    End If

                    Dim path = ReadString(changeObject, "path", "file")
                    If String.IsNullOrWhiteSpace(path) Then
                        Continue For
                    End If

                    Dim kind = ReadString(changeObject, "status", "kind", "type")
                    If String.IsNullOrWhiteSpace(kind) Then
                        lines.Add(path)
                    Else
                        lines.Add($"{kind}: {path}")
                    End If

                    shown += 1
                Next

                If changes.Count > shown Then
                    lines.Add($"... +{(changes.Count - shown).ToString(CultureInfo.InvariantCulture)} more")
                End If
            End If

            Dim output = If(itemState.FileChangeOutputText, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(output) Then
                If lines.Count > 0 Then
                    lines.Add(String.Empty)
                End If
                lines.Add(output)
            End If

            Return String.Join(Environment.NewLine, lines)
        End Function

        Private Shared Function BuildToolBody(itemState As TurnItemRuntimeState) As String
            If Not String.IsNullOrWhiteSpace(itemState.GenericText) Then
                Return itemState.GenericText.Trim()
            End If

            If itemState.RawItemPayload Is Nothing Then
                Return If(itemState.ItemType, "tool")
            End If

            Dim body = ReadString(itemState.RawItemPayload, "text", "title", "query", "prompt", "name")
            If Not String.IsNullOrWhiteSpace(body) Then
                Return body
            End If

            Return If(itemState.ItemType, "tool")
        End Function

        Private Shared Function BuildToolSecondary(itemState As TurnItemRuntimeState) As String
            If String.IsNullOrWhiteSpace(itemState.Status) Then
                Return String.Empty
            End If

            Return $"status: {itemState.Status}"
        End Function

        Private Shared Function BuildToolDetails(itemState As TurnItemRuntimeState) As String
            If itemState.RawItemPayload Is Nothing Then
                Return String.Empty
            End If

            Return itemState.RawItemPayload.ToJsonString()
        End Function

        Private Shared Function ReadLong(obj As JsonObject, ParamArray keys() As String) As Long?
            If obj Is Nothing OrElse keys Is Nothing Then
                Return Nothing
            End If

            For Each key In keys
                If String.IsNullOrWhiteSpace(key) Then
                    Continue For
                End If

                Dim value = ReadString(obj, key)
                If String.IsNullOrWhiteSpace(value) Then
                    Continue For
                End If

                Dim parsed As Long
                If Long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                    Return parsed
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function ReadString(obj As JsonObject, ParamArray keys() As String) As String
            If obj Is Nothing OrElse keys Is Nothing Then
                Return String.Empty
            End If

            For Each key In keys
                Dim node As JsonNode = Nothing
                If obj.TryGetPropertyValue(key, node) AndAlso node IsNot Nothing Then
                    Dim jsonValue = TryCast(node, JsonValue)
                    If jsonValue IsNot Nothing Then
                        Dim stringValue As String = Nothing
                        If jsonValue.TryGetValue(Of String)(stringValue) Then
                            If Not String.IsNullOrWhiteSpace(stringValue) Then
                                Return stringValue.Trim()
                            End If
                        End If

                        Dim longValue As Long
                        If jsonValue.TryGetValue(Of Long)(longValue) Then
                            Return longValue.ToString(CultureInfo.InvariantCulture)
                        End If
                    End If
                End If
            Next

            Return String.Empty
        End Function

        Private Sub AppendDescriptor(descriptor As TranscriptEntryDescriptor,
                                     appendToRaw As Boolean,
                                     Optional rawRole As String = Nothing,
                                     Optional rawBody As String = Nothing)
            If descriptor Is Nothing Then
                Return
            End If

            If appendToRaw Then
                Dim role = If(rawRole, descriptor.RoleText)
                Dim body = If(rawBody, descriptor.BodyText)
                If Not String.IsNullOrWhiteSpace(role) AndAlso Not String.IsNullOrWhiteSpace(body) Then
                    AppendRawRoleLine(role, body)
                End If
            End If

            If Not ShouldDisplayDescriptor(descriptor) Then
                Return
            End If

            _items.Add(CreateEntryFromDescriptor(descriptor))
            TrimEntriesIfNeeded()
        End Sub

        Private Sub PromoteAssistantStreamToReasoningChainStep(streamId As String)
            If String.IsNullOrWhiteSpace(streamId) Then
                Return
            End If

            Dim entry As TranscriptEntryViewModel = Nothing
            If Not _activeAssistantStreams.TryGetValue(streamId, entry) OrElse entry Is Nothing Then
                Return
            End If

            _assistantReasoningChainStreamIds.Add(streamId)
            entry.Kind = "reasoning"
            entry.RoleText = "Reasoning"
            entry.RowOpacity = 0.82R
            If entry.StreamingIndicatorVisibility = Visibility.Visible Then
                entry.StreamingIndicatorText = "updating"
            End If
        End Sub

        Private Shared Sub ApplyAssistantStreamFormatting(entry As TranscriptEntryViewModel,
                                                          markdownText As String,
                                                          renderAsReasoningChainStep As Boolean)
            If entry Is Nothing Then
                Return
            End If

            If renderAsReasoningChainStep Then
                ApplyReasoningChainStepFormatting(entry, markdownText)
                Return
            End If

            ApplyAssistantMarkdownFormatting(entry, markdownText)
        End Sub

        Private Shared Sub ApplyReasoningChainStepFormatting(entry As TranscriptEntryViewModel, markdownText As String)
            If entry Is Nothing Then
                Return
            End If

            ApplyAssistantMarkdownFormatting(entry, markdownText)

            If Not String.IsNullOrWhiteSpace(entry.BodyText) Then
                entry.BodyText = PrefixReasoningChainStepBody(entry.BodyText)
            ElseIf String.IsNullOrWhiteSpace(entry.DetailsText) Then
                entry.BodyText = "* Updating..."
            Else
                entry.BodyText = "* Codex update"
            End If
        End Sub

        Private Shared Function PrefixReasoningChainStepBody(text As String) As String
            Dim source = If(text, String.Empty)
            If String.IsNullOrWhiteSpace(source) Then
                Return "*"
            End If

            Dim trimmedStart = source.TrimStart()
            If trimmedStart.StartsWith("* ", StringComparison.Ordinal) Then
                Return trimmedStart
            End If

            Return "* " & trimmedStart
        End Function

        Private Function CreateEntryFromDescriptor(descriptor As TranscriptEntryDescriptor) As TranscriptEntryViewModel
            Dim entry As New TranscriptEntryViewModel() With {
                .Kind = If(descriptor.Kind, String.Empty),
                .TimestampText = If(descriptor.TimestampText, String.Empty),
                .RoleText = If(descriptor.RoleText, String.Empty),
                .BodyText = If(descriptor.BodyText, String.Empty),
                .SecondaryText = If(descriptor.SecondaryText, String.Empty),
                .DetailsText = If(descriptor.DetailsText, String.Empty)
            }
            ApplyTimelineDotRowVisibility(entry)

            If descriptor.AddedLineCount.HasValue AndAlso descriptor.AddedLineCount.Value > 0 Then
                entry.AddedLinesText = $"+{descriptor.AddedLineCount.Value}"
            End If

            If descriptor.RemovedLineCount.HasValue AndAlso descriptor.RemovedLineCount.Value > 0 Then
                entry.RemovedLinesText = $"-{descriptor.RemovedLineCount.Value}"
            End If

            entry.RowOpacity = If(descriptor.IsMuted, 0.82R, 1.0R)
            entry.BodyFontFamily = If(descriptor.IsCommandLike OrElse descriptor.IsMonospaceBody,
                                      New FontFamily("Cascadia Code"),
                                      New FontFamily("Segoe UI"))
            entry.DetailsFontFamily = New FontFamily("Cascadia Code")

            If StringComparer.OrdinalIgnoreCase.Equals(entry.Kind, "assistant") OrElse
               StringComparer.OrdinalIgnoreCase.Equals(entry.Kind, "assistantFinal") OrElse
               StringComparer.OrdinalIgnoreCase.Equals(entry.Kind, "assistantCommentary") Then
                ApplyAssistantMarkdownFormatting(entry, entry.BodyText)
            ElseIf StringComparer.OrdinalIgnoreCase.Equals(entry.Kind, "reasoning") AndAlso
                   Not descriptor.UseRawReasoningLayout Then
                Dim reasoningSource = If(String.IsNullOrWhiteSpace(descriptor.DetailsText), descriptor.BodyText, descriptor.DetailsText)
                ApplyReasoningMarkdownFormatting(entry, reasoningSource)
            End If

            entry.AllowDetailsCollapse = ShouldAllowDetailsCollapse(entry.Kind)
            entry.IsDetailsExpanded = Not (entry.AllowDetailsCollapse AndAlso ShouldCollapseDetailsByDefault(entry.Kind))
            entry.StreamingIndicatorVisibility = If(descriptor.IsStreaming, Visibility.Visible, Visibility.Collapsed)
            entry.StreamingIndicatorText = ResolveStreamingIndicatorText(entry.Kind)

            Select Case entry.Kind
                Case "assistant"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(222, 232, 238))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(48, 66, 74))
                Case "assistantCommentary"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(236, 236, 236))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(92, 92, 92))
                    entry.RowOpacity = 0.86R
                Case "assistantFinal"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(217, 230, 243))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(30, 60, 92))
                Case "user"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(226, 238, 228))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(42, 78, 52))
                Case "command"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(234, 233, 230))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(72, 69, 62))
                Case "fileChange"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(233, 236, 229))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(78, 83, 60))
                Case "plan"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(231, 235, 239))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(60, 72, 88))
                Case "reasoning"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(237, 237, 237))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(92, 92, 92))
                Case "reasoningCard"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(237, 237, 237))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(92, 92, 92))
                    entry.RowOpacity = 0.82R
                Case "turnMarker", "turnDiff", "turnPlan", "systemMarker"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(238, 238, 238))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(88, 88, 88))
                    entry.RowOpacity = 0.78R
                Case "toolMcp"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(230, 237, 246))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(45, 75, 112))
                Case "toolCollab"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(232, 243, 238))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(45, 86, 67))
                Case "toolWebSearch"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(246, 237, 228))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(107, 76, 44))
                Case "toolImageView"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(236, 232, 246))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(78, 62, 112))
                Case "unknownItem"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(238, 238, 238))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(88, 88, 88))
                    entry.RowOpacity = 0.84R
                Case "error"
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(250, 229, 225))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(128, 43, 26))
                Case Else
                    entry.RoleBadgeBackground = New SolidColorBrush(Color.FromRgb(238, 238, 238))
                    entry.RoleBadgeForeground = New SolidColorBrush(Color.FromRgb(88, 88, 88))
            End Select

            Return entry
        End Function

        Private Sub RefreshTimelineDotRowVisibility()
            For Each entry In _items
                ApplyTimelineDotRowVisibility(entry)
            Next
        End Sub

        Private Sub ApplyTimelineDotRowVisibility(entry As TranscriptEntryViewModel)
            If entry Is Nothing Then
                Return
            End If

            entry.RowVisibility = If(ShouldShowTimelineRowForKind(entry.Kind),
                                     Visibility.Visible,
                                     Visibility.Collapsed)
        End Sub

        Private Function ShouldShowTimelineRowForKind(kind As String) As Boolean
            Select Case If(kind, String.Empty).Trim().ToLowerInvariant()
                Case "event"
                    Return _showEventDotsInTranscript
                Case "system", "systemmarker"
                    Return _showSystemDotsInTranscript
                Case Else
                    Return True
            End Select
        End Function

        Private Function ShouldCollapseDetailsByDefault(kind As String) As Boolean
            If StringComparer.OrdinalIgnoreCase.Equals(kind, "reasoningCard") Then
                Return True
            End If

            Return _collapseCommandDetailsByDefault
        End Function

        Private Shared Function ShouldAllowDetailsCollapse(kind As String) As Boolean
            Select Case If(kind, String.Empty).Trim().ToLowerInvariant()
                Case "command", "filechange", "reasoningcard",
                     "toolmcp", "toolcollab", "toolwebsearch", "toolimageview"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function ResolveStreamingIndicatorText(kind As String) As String
            Select Case If(kind, String.Empty).Trim().ToLowerInvariant()
                Case "reasoning", "reasoningcard"
                    Return "thinking"
                Case "assistantcommentary"
                    Return "updating"
                Case Else
                    Return "streaming"
            End Select
        End Function

        Private Shared Sub ApplyAssistantMarkdownFormatting(entry As TranscriptEntryViewModel, markdownText As String)
            If entry Is Nothing Then
                Return
            End If

            Dim body As String = String.Empty
            Dim codeBlocks As String = String.Empty
            SplitAssistantMarkdown(markdownText, body, codeBlocks)

            entry.BodyText = body
            entry.DetailsText = codeBlocks
        End Sub

        Private Shared Sub ApplyReasoningMarkdownFormatting(entry As TranscriptEntryViewModel, markdownText As String)
            If entry Is Nothing Then
                Return
            End If

            Dim normalizedText = NormalizeReasoningPayloadText(markdownText)
            Dim body As String = String.Empty
            Dim codeBlocks As String = String.Empty
            SplitAssistantMarkdown(normalizedText, body, codeBlocks)

            If String.IsNullOrWhiteSpace(body) AndAlso String.IsNullOrWhiteSpace(codeBlocks) Then
                body = BuildReasoningPreview(normalizedText)
            End If

            ' Reasoning should read as a single clean block by default.
            ' Only keep a details block when there is actual fenced code content.
            entry.BodyText = If(String.IsNullOrWhiteSpace(body), "Thinking...", body)
            entry.DetailsText = codeBlocks
        End Sub

        Private Shared Function NormalizeReasoningPayloadFragment(value As String) As String
            Dim normalized = NormalizeReasoningPayloadText(value)
            Return If(normalized, String.Empty)
        End Function

        Private Shared Function NormalizeReasoningPayloadText(value As String) As String
            Dim source = If(value, String.Empty)
            If String.IsNullOrWhiteSpace(source) Then
                Return String.Empty
            End If

            Dim trimmed = source.Trim()
            Dim flattenedJson = TryFlattenReasoningJsonPayload(trimmed)
            If Not String.IsNullOrWhiteSpace(flattenedJson) Then
                Return flattenedJson
            End If

            Return source
        End Function

        Private Shared Function TryFlattenReasoningJsonPayload(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Dim firstChar = value(0)
            If firstChar <> "{"c AndAlso firstChar <> "["c AndAlso firstChar <> """"c Then
                Return String.Empty
            End If

            Try
                Dim node = JsonNode.Parse(value)
                If node Is Nothing Then
                    Return String.Empty
                End If

                Dim builder As New StringBuilder()
                CollectReasoningJsonText(node, builder)
                Dim flattened = NormalizeMarkdownProse(builder.ToString())
                If Not String.IsNullOrWhiteSpace(flattened) Then
                    Return flattened
                End If
            Catch
                ' Not valid JSON-shaped reasoning payload; leave as-is.
            End Try

            Return String.Empty
        End Function

        Private Shared Sub CollectReasoningJsonText(node As JsonNode, builder As StringBuilder)
            If node Is Nothing OrElse builder Is Nothing Then
                Return
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue IsNot Nothing Then
                Dim stringValue As String = Nothing
                If jsonValue.TryGetValue(Of String)(stringValue) Then
                    AppendReasoningJsonTextPart(builder, stringValue)
                End If
                Return
            End If

            Dim arrayNode = TryCast(node, JsonArray)
            If arrayNode IsNot Nothing Then
                For Each child In arrayNode
                    CollectReasoningJsonText(child, builder)
                Next
                Return
            End If

            Dim objectNode = TryCast(node, JsonObject)
            If objectNode Is Nothing Then
                Return
            End If

            Dim preferredKeys = {"text", "delta", "summary", "content", "parts", "items", "message"}
            Dim foundPreferred = False

            For Each key In preferredKeys
                Dim child As JsonNode = Nothing
                If objectNode.TryGetPropertyValue(key, child) Then
                    Dim before = builder.Length
                    CollectReasoningJsonText(child, builder)
                    If builder.Length > before Then
                        foundPreferred = True
                    End If
                End If
            Next

            If foundPreferred Then
                Return
            End If

            For Each pair In objectNode
                If ShouldIgnoreReasoningJsonKey(pair.Key) Then
                    Continue For
                End If

                CollectReasoningJsonText(pair.Value, builder)
            Next
        End Sub

        Private Shared Function ShouldIgnoreReasoningJsonKey(key As String) As Boolean
            Select Case If(key, String.Empty).Trim().ToLowerInvariant()
                Case "type", "id", "index", "kind", "role", "status"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Sub AppendReasoningJsonTextPart(builder As StringBuilder, text As String)
            Dim value = If(text, String.Empty)
            If String.IsNullOrWhiteSpace(value) Then
                Return
            End If

            If StringComparer.Ordinal.Equals(value.Trim(), "[reasoning item]") Then
                Return
            End If

            If builder.Length > 0 Then
                Dim lastChar = builder(builder.Length - 1)
                If Not Char.IsWhiteSpace(lastChar) AndAlso
                   lastChar <> "."c AndAlso lastChar <> ":"c AndAlso lastChar <> ";"c AndAlso
                   lastChar <> "!"c AndAlso lastChar <> "?"c Then
                    builder.Append(" "c)
                End If
            End If

            builder.Append(value.Trim())
        End Sub

        Private Shared Sub SplitAssistantMarkdown(markdownText As String,
                                                  ByRef body As String,
                                                  ByRef codeBlocks As String)
            body = String.Empty
            codeBlocks = String.Empty

            Dim source = If(markdownText, String.Empty)
            If String.IsNullOrWhiteSpace(source) Then
                Return
            End If

            Dim proseBuilder As New StringBuilder()
            Dim codeBuilder As New StringBuilder()
            Dim inFence = False
            Dim fenceLang = String.Empty

            Dim normalized = source.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)

            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty)
                Dim trimmed = line.Trim()

                If trimmed.StartsWith("```", StringComparison.Ordinal) Then
                    If inFence Then
                        inFence = False
                        fenceLang = String.Empty
                        codeBuilder.AppendLine()
                    Else
                        inFence = True
                        fenceLang = trimmed.Substring(3).Trim()
                        If codeBuilder.Length > 0 Then
                            codeBuilder.AppendLine()
                        End If
                        If Not String.IsNullOrWhiteSpace(fenceLang) Then
                            codeBuilder.AppendLine($"[{fenceLang}]")
                        End If
                    End If

                    Continue For
                End If

                If inFence Then
                    codeBuilder.AppendLine(line)
                Else
                    proseBuilder.AppendLine(CleanMarkdownProseLine(line))
                End If
            Next

            body = NormalizeMarkdownProse(proseBuilder.ToString())
            codeBlocks = codeBuilder.ToString().Trim()

            If String.IsNullOrWhiteSpace(body) AndAlso Not String.IsNullOrWhiteSpace(codeBlocks) Then
                body = "Code snippet"
            End If
        End Sub

        Private Shared Function CleanMarkdownProseLine(line As String) As String
            Dim text = If(line, String.Empty)
            Dim trimmedStart = text.TrimStart()

            If trimmedStart.StartsWith("#", StringComparison.Ordinal) Then
                text = trimmedStart.TrimStart("#"c).TrimStart()
            ElseIf trimmedStart.StartsWith("> ", StringComparison.Ordinal) Then
                text = trimmedStart.Substring(2)
            ElseIf trimmedStart.StartsWith("- ", StringComparison.Ordinal) OrElse
                   trimmedStart.StartsWith("* ", StringComparison.Ordinal) Then
                text = "• " & trimmedStart.Substring(2)
            Else
                text = trimmedStart
            End If

            ' Convert markdown links/images into plain visible labels.
            text = Regex.Replace(text, "!\[([^\]]*)\]\([^)]+\)", "$1")
            text = Regex.Replace(text, "\[([^\]]+)\]\([^)]+\)", "$1")

            text = text.Replace("**", String.Empty, StringComparison.Ordinal).
                        Replace("__", String.Empty, StringComparison.Ordinal).
                        Replace("`", String.Empty, StringComparison.Ordinal)

            Return text.TrimEnd()
        End Function

        Private Shared Function NormalizeMarkdownProse(text As String) As String
            Dim value = If(text, String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = value.Split({vbLf}, StringSplitOptions.None)
            Dim builder As New StringBuilder()
            Dim blankCount = 0

            For Each line In lines
                Dim current = If(line, String.Empty).TrimEnd()
                If String.IsNullOrWhiteSpace(current) Then
                    blankCount += 1
                    If blankCount > 1 Then
                        Continue For
                    End If
                Else
                    blankCount = 0
                End If

                If builder.Length > 0 Then
                    builder.AppendLine()
                End If
                builder.Append(current)
            Next

            Return builder.ToString().Trim()
        End Function

        Private Function ShouldDisplayDescriptor(descriptor As TranscriptEntryDescriptor) As Boolean
            If descriptor Is Nothing Then
                Return False
            End If

            If StringComparer.OrdinalIgnoreCase.Equals(descriptor.Kind, "system") Then
                Return Not ShouldSuppressSystemMessage(descriptor.BodyText)
            End If

            Return Not String.IsNullOrWhiteSpace(descriptor.BodyText) OrElse
                   Not String.IsNullOrWhiteSpace(descriptor.DetailsText)
        End Function

        Private Shared Function ShouldSuppressSystemMessage(message As String) As Boolean
            Dim text = If(message, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return True
            End If

            Return False
        End Function

        Private Shared Function BuildReasoningPreview(text As String) As String
            Dim normalized = If(text, String.Empty).Replace(ControlChars.Cr, " "c).
                                               Replace(ControlChars.Lf, " "c).
                                               Replace(ControlChars.Tab, " "c).Trim()
            Do While normalized.Contains("  ", StringComparison.Ordinal)
                normalized = normalized.Replace("  ", " ", StringComparison.Ordinal)
            Loop

            If normalized.Length > 180 Then
                Return normalized.Substring(0, 177) & "..."
            End If

            Return normalized
        End Function

        Private Shared Function NormalizeStreamId(itemId As String) As String
            If String.IsNullOrWhiteSpace(itemId) Then
                Return "live"
            End If

            Return itemId.Trim()
        End Function

        Private Shared Function FormatLiveTimestamp() As String
            Return Date.Now.ToString("HH:mm")
        End Function

        Private Shared Function SanitizeOptionalLineCount(value As Integer?) As Integer?
            If Not value.HasValue OrElse value.Value <= 0 Then
                Return Nothing
            End If

            Return value.Value
        End Function

        Private Sub AppendRawRoleLine(role As String, text As String)
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            Dim normalizedRole = If(role, String.Empty).Trim()
            Dim line = $"[{Date.Now:HH:mm:ss}] {normalizedRole}: {text}{Environment.NewLine}{Environment.NewLine}"
            AppendRawTranscriptChunk(line)
        End Sub

        Private Sub AppendRawTranscriptChunk(value As String)
            If String.IsNullOrEmpty(value) Then
                Return
            End If

            _fullTranscriptBuilder.Append(value)
            TranscriptText = TrimLogText(_transcriptText & value)
        End Sub

        Private Sub TrimEntriesIfNeeded()
            While _items.Count > MaxTranscriptEntries
                Dim removed = _items(0)
                _items.RemoveAt(0)
                RemoveRuntimeEntryKeyForEntry(removed)
            End While
        End Sub

        Private Sub RemoveRuntimeEntryKeyForEntry(entry As TranscriptEntryViewModel)
            If entry Is Nothing OrElse _runtimeEntriesByKey.Count = 0 Then
                Return
            End If

            Dim keyToRemove As String = Nothing
            For Each pair In _runtimeEntriesByKey
                If ReferenceEquals(pair.Value, entry) Then
                    keyToRemove = pair.Key
                    Exit For
                End If
            Next

            If Not String.IsNullOrWhiteSpace(keyToRemove) Then
                _runtimeEntriesByKey.Remove(keyToRemove)
            End If
        End Sub

        Private Shared Function TrimLogText(value As String) As String
            Dim text = If(value, String.Empty)
            If text.Length <= MaxLogChars Then
                Return text
            End If

            Return text.Substring(text.Length - (MaxLogChars \ 2))
        End Function
    End Class
End Namespace
