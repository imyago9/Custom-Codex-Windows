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
        Private Const EnableFileChangeTurnDiffMergeProtocolDebug As Boolean = True
        Private Const EnableTranscriptOrderingProtocolDebug As Boolean = True
        Private Shared ReadOnly TurnDiffAddedCountRegex As New Regex("(?<!\w)\+(?<n>\d+)\b", RegexOptions.Compiled Or RegexOptions.CultureInvariant)
        Private Shared ReadOnly TurnDiffRemovedCountRegex As New Regex("(?<!\w)-(?<n>\d+)\b", RegexOptions.Compiled Or RegexOptions.CultureInvariant)
        Private Shared ReadOnly TurnDiffInsertionsRegex As New Regex("(?<n>\d+)\s+insertions?\(\+\)", RegexOptions.Compiled Or RegexOptions.CultureInvariant Or RegexOptions.IgnoreCase)
        Private Shared ReadOnly TurnDiffDeletionsRegex As New Regex("(?<n>\d+)\s+deletions?\(-\)", RegexOptions.Compiled Or RegexOptions.CultureInvariant Or RegexOptions.IgnoreCase)

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
        Private ReadOnly _turnDiffSummaryByTurnId As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Private ReadOnly _fileChangeRuntimeKeyByTurnId As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Private ReadOnly _fileChangeTurnIdByRuntimeKey As New Dictionary(Of String, String)(StringComparer.Ordinal)
        Private ReadOnly _fileChangeItemStateByRuntimeKey As New Dictionary(Of String, TurnItemRuntimeState)(StringComparer.Ordinal)
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
            _turnDiffSummaryByTurnId.Clear()
            _fileChangeRuntimeKeyByTurnId.Clear()
            _fileChangeTurnIdByRuntimeKey.Clear()
            _fileChangeItemStateByRuntimeKey.Clear()
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
            _turnDiffSummaryByTurnId.Clear()
            _fileChangeRuntimeKeyByTurnId.Clear()
            _fileChangeTurnIdByRuntimeKey.Clear()
            _fileChangeItemStateByRuntimeKey.Clear()
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
                Dim entry = AppendDescriptor(descriptor, appendToRaw:=False)
                If descriptor Is Nothing OrElse entry Is Nothing Then
                    Continue For
                End If

                Dim runtimeKey = If(descriptor.RuntimeKey, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(runtimeKey) Then
                    Continue For
                End If

                _runtimeEntriesByKey(runtimeKey) = entry
            Next

            AppendTranscriptOrderingSnapshotDump("set_snapshot")
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
                .StatusText = NormalizeStatusToken(cleanStatus),
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

            Dim runtimeKey = $"item:{scopedKey}"
            Dim descriptor = BuildRuntimeItemDescriptor(itemState)
            If descriptor Is Nothing Then
                Return
            End If

            AppendTranscriptOrderingProtocolDebug(
                "runtime_item_prepare",
                $"runtimeKey={FormatProtocolDebugToken(runtimeKey)} threadId={FormatProtocolDebugToken(itemState.ThreadId)} turnId={FormatProtocolDebugToken(itemState.TurnId)} itemType={FormatProtocolDebugToken(itemState.ItemType)} seq={FormatOrderingDebugNullableLong(itemState.TurnItemStreamSequence)} ord={FormatOrderingDebugNullableInt(itemState.TurnItemOrderIndex)} ts={FormatOrderingDebugTimestamp(If(itemState.StartedAt, itemState.CompletedAt))}")

            RegisterRuntimeItemAssociations(itemState, runtimeKey)
            UpsertRuntimeDescriptor(runtimeKey, descriptor)

            If StringComparer.OrdinalIgnoreCase.Equals(If(descriptor.Kind, String.Empty), "fileChange") Then
                Dim cachedDiffSummary As String = Nothing
                _turnDiffSummaryByTurnId.TryGetValue(NormalizeTurnId(itemState.TurnId), cachedDiffSummary)
                AppendFileChangeTurnDiffMergeProtocolDebug("item_upsert",
                                                           itemState.TurnId,
                                                           runtimeKey,
                                                           cachedDiffSummary,
                                                           descriptor)
                RemoveStandaloneTurnDiffEntriesForTurn(itemState.TurnId)
            End If
        End Sub

        Public Sub UpsertTurnLifecycleMarker(threadId As String,
                                            turnId As String,
                                            status As String,
                                            Optional timestampUtc As DateTimeOffset? = Nothing)
            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return
            End If

            Dim normalizedStatus = If(status, String.Empty).Trim()
            Dim lifecycleTimestampText = If(timestampUtc.HasValue,
                                            timestampUtc.Value.ToLocalTime().ToString("HH:mm"),
                                            FormatLiveTimestamp())
            Dim descriptor = BuildTurnLifecycleDescriptorForSnapshot(normalizedTurnId,
                                                                     normalizedStatus,
                                                                     lifecycleTimestampText)
            descriptor.ThreadId = If(threadId, String.Empty).Trim()

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

            If StringComparer.Ordinal.Equals(normalizedKind, "diff") Then
                _turnDiffSummaryByTurnId(normalizedTurnId) = normalizedSummary
                If TryRefreshMergedFileChangeForTurn(normalizedTurnId) Then
                    AppendFileChangeTurnDiffMergeProtocolDebug("turn_diff_updated",
                                                               normalizedTurnId,
                                                               String.Empty,
                                                               normalizedSummary,
                                                               Nothing,
                                                               note:="refreshed_filechange=True")
                    RemoveStandaloneTurnDiffEntriesForTurn(normalizedTurnId)
                    Return
                End If

                AppendFileChangeTurnDiffMergeProtocolDebug("turn_diff_updated",
                                                           normalizedTurnId,
                                                           String.Empty,
                                                           normalizedSummary,
                                                           Nothing,
                                                           note:="refreshed_filechange=False")
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
            Dim existingIndexBefore = -1

            Dim existing As TranscriptEntryViewModel = Nothing
            If _runtimeEntriesByKey.TryGetValue(normalizedKey, existing) AndAlso existing IsNot Nothing Then
                Dim index = _items.IndexOf(existing)
                existingIndexBefore = index
                Dim shouldReposition = normalizedKey.StartsWith("item:", StringComparison.Ordinal) AndAlso
                                       (descriptor.TurnItemStreamSequence.HasValue OrElse
                                        descriptor.TurnItemSortTimestampUtc.HasValue OrElse
                                        descriptor.TurnItemOrderIndex.HasValue)
                If index >= 0 AndAlso shouldReposition Then
                    _items.RemoveAt(index)
                    Dim repositionInsertIndex = ResolveRuntimeInsertIndex(normalizedKey, descriptor)
                    If repositionInsertIndex >= 0 AndAlso repositionInsertIndex <= _items.Count Then
                        _items.Insert(repositionInsertIndex, replacement)
                    Else
                        _items.Add(replacement)
                    End If

                    AppendTranscriptOrderingProtocolDebug(
                        "upsert_reposition",
                        BuildOrderingUpsertDebugMessage(normalizedKey, descriptor, existingIndexBefore, _items.IndexOf(replacement), "reposition"))
                ElseIf index >= 0 Then
                    _items(index) = replacement
                    AppendTranscriptOrderingProtocolDebug(
                        "upsert_replace_in_place",
                        BuildOrderingUpsertDebugMessage(normalizedKey, descriptor, existingIndexBefore, index, "replace_in_place"))
                Else
                    _items.Add(replacement)
                    AppendTranscriptOrderingProtocolDebug(
                        "upsert_replace_missing_index",
                        BuildOrderingUpsertDebugMessage(normalizedKey, descriptor, existingIndexBefore, _items.Count - 1, "replace_missing_index"))
                End If

                _runtimeEntriesByKey(normalizedKey) = replacement
                TrimEntriesIfNeeded()
                Return
            End If

            _runtimeEntriesByKey(normalizedKey) = replacement
            Dim runtimeInsertIndex = ResolveRuntimeInsertIndex(normalizedKey, descriptor)
            If runtimeInsertIndex < 0 Then
                runtimeInsertIndex = ResolveLifecycleInsertIndex(normalizedKey, descriptor)
            End If
            If runtimeInsertIndex >= 0 AndAlso runtimeInsertIndex <= _items.Count Then
                _items.Insert(runtimeInsertIndex, replacement)
            Else
                _items.Add(replacement)
            End If
            AppendTranscriptOrderingProtocolDebug(
                "upsert_insert",
                BuildOrderingUpsertDebugMessage(normalizedKey, descriptor, existingIndexBefore, _items.IndexOf(replacement), "insert"))
            TrimEntriesIfNeeded()
        End Sub

        Private Function ResolveRuntimeInsertIndex(runtimeKey As String, descriptor As TranscriptEntryDescriptor) As Integer
            Dim normalizedRuntimeKey = If(runtimeKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedRuntimeKey) OrElse descriptor Is Nothing Then
                Return -1
            End If

            If Not normalizedRuntimeKey.StartsWith("item:", StringComparison.Ordinal) Then
                Return -1
            End If

            Dim normalizedTurnId = NormalizeTurnId(descriptor.TurnId)
            If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return -1
            End If

            If descriptor.TurnItemStreamSequence.HasValue Then
                Dim targetStreamSequence = descriptor.TurnItemStreamSequence.Value
                Dim targetOrderIndex = descriptor.TurnItemOrderIndex
                For i = 0 To _items.Count - 1
                    Dim candidate = _items(i)
                    If candidate Is Nothing Then
                        Continue For
                    End If

                    If Not StringComparer.Ordinal.Equals(NormalizeTurnId(candidate.TurnId), normalizedTurnId) Then
                        Continue For
                    End If

                    If Not candidate.TurnItemStreamSequence.HasValue Then
                        Continue For
                    End If

                    Dim candidateStreamSequence = candidate.TurnItemStreamSequence.Value
                    If candidateStreamSequence > targetStreamSequence Then
                        AppendTranscriptOrderingResolveDecisionDebug(
                            "stream_seq_gt",
                            normalizedRuntimeKey,
                            descriptor,
                            i,
                            candidate)
                        Return i
                    End If

                    If candidateStreamSequence = targetStreamSequence AndAlso
                       targetOrderIndex.HasValue AndAlso
                       candidate.TurnItemOrderIndex.HasValue AndAlso
                       candidate.TurnItemOrderIndex.Value > targetOrderIndex.Value Then
                        AppendTranscriptOrderingResolveDecisionDebug(
                            "stream_seq_eq_order_gt",
                            normalizedRuntimeKey,
                            descriptor,
                            i,
                            candidate)
                        Return i
                    End If
                Next
            End If

            If descriptor.TurnItemSortTimestampUtc.HasValue Then
                Dim targetTimestamp = descriptor.TurnItemSortTimestampUtc.Value
                Dim targetOrderIndex = descriptor.TurnItemOrderIndex
                For i = 0 To _items.Count - 1
                    Dim candidate = _items(i)
                    If candidate Is Nothing Then
                        Continue For
                    End If

                    If Not StringComparer.Ordinal.Equals(NormalizeTurnId(candidate.TurnId), normalizedTurnId) Then
                        Continue For
                    End If

                    If Not candidate.TurnItemSortTimestampUtc.HasValue Then
                        Continue For
                    End If

                    Dim candidateTimestamp = candidate.TurnItemSortTimestampUtc.Value
                    If candidateTimestamp > targetTimestamp Then
                        AppendTranscriptOrderingResolveDecisionDebug(
                            "timestamp_gt",
                            normalizedRuntimeKey,
                            descriptor,
                            i,
                            candidate)
                        Return i
                    End If

                    If candidateTimestamp = targetTimestamp AndAlso
                       targetOrderIndex.HasValue AndAlso
                       candidate.TurnItemOrderIndex.HasValue AndAlso
                       candidate.TurnItemOrderIndex.Value > targetOrderIndex.Value Then
                        AppendTranscriptOrderingResolveDecisionDebug(
                            "timestamp_eq_order_gt",
                            normalizedRuntimeKey,
                            descriptor,
                            i,
                            candidate)
                        Return i
                    End If
                Next
            End If

            If descriptor.TurnItemOrderIndex.HasValue Then
                Dim targetOrderIndex = descriptor.TurnItemOrderIndex.Value
                For i = 0 To _items.Count - 1
                    Dim candidate = _items(i)
                    If candidate Is Nothing Then
                        Continue For
                    End If

                    If Not StringComparer.Ordinal.Equals(NormalizeTurnId(candidate.TurnId), normalizedTurnId) Then
                        Continue For
                    End If

                    If Not candidate.TurnItemOrderIndex.HasValue Then
                        Continue For
                    End If

                    If candidate.TurnItemOrderIndex.Value > targetOrderIndex Then
                        AppendTranscriptOrderingResolveDecisionDebug(
                            "order_gt",
                            normalizedRuntimeKey,
                            descriptor,
                            i,
                            candidate)
                        Return i
                    End If
                Next
            End If

            For i = _items.Count - 1 To 0 Step -1
                Dim candidate = _items(i)
                If candidate Is Nothing Then
                    Continue For
                End If

                If Not StringComparer.OrdinalIgnoreCase.Equals(candidate.Kind, "turnMarker") Then
                    Continue For
                End If

                If Not StringComparer.Ordinal.Equals(NormalizeTurnId(candidate.TurnId), normalizedTurnId) Then
                    Continue For
                End If

                If Not IsTurnCompletionMarkerBody(candidate.BodyText) Then
                    Continue For
                End If

                AppendTranscriptOrderingResolveDecisionDebug(
                    "before_turn_completion",
                    normalizedRuntimeKey,
                    descriptor,
                    i,
                    candidate)
                Return i
            Next

            AppendTranscriptOrderingResolveDecisionDebug(
                "append_end",
                normalizedRuntimeKey,
                descriptor,
                -1,
                Nothing)
            Return -1
        End Function

        Private Function ResolveLifecycleInsertIndex(runtimeKey As String, descriptor As TranscriptEntryDescriptor) As Integer
            Dim normalizedRuntimeKey = If(runtimeKey, String.Empty).Trim()
            Dim normalizedTurnId = NormalizeTurnId(descriptor?.TurnId)
            If String.IsNullOrWhiteSpace(normalizedRuntimeKey) OrElse
               String.IsNullOrWhiteSpace(normalizedTurnId) OrElse
               Not normalizedRuntimeKey.StartsWith("turn:lifecycle:", StringComparison.Ordinal) Then
                Return -1
            End If

            If normalizedRuntimeKey.StartsWith("turn:lifecycle:start:", StringComparison.Ordinal) Then
                Dim lastUserIndex = -1
                For i = 0 To _items.Count - 1
                    Dim candidate = _items(i)
                    If candidate Is Nothing Then
                        Continue For
                    End If

                    If Not StringComparer.Ordinal.Equals(NormalizeTurnId(candidate.TurnId), normalizedTurnId) Then
                        Continue For
                    End If

                    If StringComparer.OrdinalIgnoreCase.Equals(candidate.Kind, "user") Then
                        lastUserIndex = i
                    End If
                Next

                If lastUserIndex >= 0 Then
                    Return lastUserIndex + 1
                End If

                For i = 0 To _items.Count - 1
                    Dim candidate = _items(i)
                    If candidate Is Nothing Then
                        Continue For
                    End If

                    If StringComparer.Ordinal.Equals(NormalizeTurnId(candidate.TurnId), normalizedTurnId) Then
                        Return i
                    End If
                Next

                Return -1
            End If

            If normalizedRuntimeKey.StartsWith("turn:lifecycle:end:", StringComparison.Ordinal) Then
                For i = _items.Count - 1 To 0 Step -1
                    Dim candidate = _items(i)
                    If candidate Is Nothing Then
                        Continue For
                    End If

                    If StringComparer.Ordinal.Equals(NormalizeTurnId(candidate.TurnId), normalizedTurnId) Then
                        Return i + 1
                    End If
                Next
            End If

            Return -1
        End Function

        Private Shared Function IsTurnCompletionMarkerBody(bodyText As String) As Boolean
            Dim normalized = If(bodyText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            If StringComparer.Ordinal.Equals(normalized, "Turn started.") Then
                Return False
            End If

            Return normalized.StartsWith("Turn ", StringComparison.Ordinal)
        End Function

        Private Sub RegisterRuntimeItemAssociations(itemState As TurnItemRuntimeState, runtimeKey As String)
            Dim normalizedRuntimeKey = If(runtimeKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedRuntimeKey) Then
                Return
            End If

            Dim itemType = If(itemState?.ItemType, String.Empty).Trim()
            Dim isFileChange = StringComparer.OrdinalIgnoreCase.Equals(itemType, "filechange")

            If Not isFileChange Then
                RemoveFileChangeRuntimeAssociation(normalizedRuntimeKey)
                Return
            End If

            Dim normalizedTurnId = NormalizeTurnId(itemState.TurnId)
            If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                RemoveFileChangeRuntimeAssociation(normalizedRuntimeKey)
                Return
            End If

            Dim previousRuntimeKey As String = Nothing
            If _fileChangeRuntimeKeyByTurnId.TryGetValue(normalizedTurnId, previousRuntimeKey) AndAlso
               Not StringComparer.Ordinal.Equals(previousRuntimeKey, normalizedRuntimeKey) Then
                _fileChangeTurnIdByRuntimeKey.Remove(previousRuntimeKey)
                _fileChangeItemStateByRuntimeKey.Remove(previousRuntimeKey)
            End If

            _fileChangeRuntimeKeyByTurnId(normalizedTurnId) = normalizedRuntimeKey
            _fileChangeTurnIdByRuntimeKey(normalizedRuntimeKey) = normalizedTurnId
            _fileChangeItemStateByRuntimeKey(normalizedRuntimeKey) = itemState
        End Sub

        Private Function TryRefreshMergedFileChangeForTurn(turnId As String) As Boolean
            Dim normalizedTurnId = NormalizeTurnId(turnId)
            If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return False
            End If

            Dim runtimeKey As String = Nothing
            If Not _fileChangeRuntimeKeyByTurnId.TryGetValue(normalizedTurnId, runtimeKey) OrElse
               String.IsNullOrWhiteSpace(runtimeKey) Then
                Return False
            End If

            Dim itemState As TurnItemRuntimeState = Nothing
            If Not _fileChangeItemStateByRuntimeKey.TryGetValue(runtimeKey, itemState) OrElse itemState Is Nothing Then
                RemoveFileChangeRuntimeAssociation(runtimeKey)
                Return False
            End If

            If Not _runtimeEntriesByKey.ContainsKey(runtimeKey) Then
                RemoveFileChangeRuntimeAssociation(runtimeKey)
                Return False
            End If

            Dim replacementDescriptor = BuildRuntimeItemDescriptor(itemState)
            If replacementDescriptor Is Nothing Then
                Return False
            End If

            Dim diffSummary As String = Nothing
            _turnDiffSummaryByTurnId.TryGetValue(normalizedTurnId, diffSummary)
            AppendFileChangeTurnDiffMergeProtocolDebug("refresh_from_turn_diff",
                                                       normalizedTurnId,
                                                       runtimeKey,
                                                       diffSummary,
                                                       replacementDescriptor)
            UpsertRuntimeDescriptor(runtimeKey, replacementDescriptor)
            Return True
        End Function

        Private Sub RemoveStandaloneTurnDiffEntriesForTurn(turnId As String)
            Dim normalizedTurnId = NormalizeTurnId(turnId)
            If String.IsNullOrWhiteSpace(normalizedTurnId) OrElse _runtimeEntriesByKey.Count = 0 Then
                Return
            End If

            Dim suffix = ":" & normalizedTurnId
            Dim keysToRemove As New List(Of String)()
            For Each pair In _runtimeEntriesByKey
                Dim runtimeKey = If(pair.Key, String.Empty)
                If runtimeKey.StartsWith("turn:meta:diff:", StringComparison.Ordinal) AndAlso
                   runtimeKey.EndsWith(suffix, StringComparison.Ordinal) Then
                    keysToRemove.Add(runtimeKey)
                End If
            Next

            For Each runtimeKey In keysToRemove
                RemoveRuntimeEntryByKey(runtimeKey)
            Next
        End Sub

        Private Sub RemoveRuntimeEntryByKey(runtimeKey As String)
            Dim normalizedKey = If(runtimeKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedKey) Then
                Return
            End If

            Dim existing As TranscriptEntryViewModel = Nothing
            If _runtimeEntriesByKey.TryGetValue(normalizedKey, existing) Then
                _runtimeEntriesByKey.Remove(normalizedKey)
                If existing IsNot Nothing Then
                    _items.Remove(existing)
                End If
            End If

            RemoveRuntimeKeyAssociations(normalizedKey)
        End Sub

        Private Sub RemoveRuntimeKeyAssociations(runtimeKey As String)
            Dim normalizedKey = If(runtimeKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedKey) Then
                Return
            End If

            RemoveFileChangeRuntimeAssociation(normalizedKey)
        End Sub

        Private Sub RemoveFileChangeRuntimeAssociation(runtimeKey As String)
            Dim normalizedKey = If(runtimeKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedKey) Then
                Return
            End If

            Dim normalizedTurnId As String = Nothing
            If _fileChangeTurnIdByRuntimeKey.TryGetValue(normalizedKey, normalizedTurnId) Then
                _fileChangeTurnIdByRuntimeKey.Remove(normalizedKey)
                Dim currentTurnRuntimeKey As String = Nothing
                If Not String.IsNullOrWhiteSpace(normalizedTurnId) AndAlso
                   _fileChangeRuntimeKeyByTurnId.TryGetValue(normalizedTurnId, currentTurnRuntimeKey) AndAlso
                   StringComparer.Ordinal.Equals(currentTurnRuntimeKey, normalizedKey) Then
                    _fileChangeRuntimeKeyByTurnId.Remove(normalizedTurnId)
                End If
            End If

            _fileChangeItemStateByRuntimeKey.Remove(normalizedKey)
        End Sub

        Private Shared Function NormalizeTurnId(turnId As String) As String
            Return If(turnId, String.Empty).Trim()
        End Function

        Private Function BuildRuntimeItemDescriptor(itemState As TurnItemRuntimeState) As TranscriptEntryDescriptor
            Dim descriptor = BuildRuntimeItemDescriptorForSnapshot(itemState, FormatRuntimeItemTimestamp(itemState))
            If descriptor Is Nothing OrElse itemState Is Nothing Then
                Return descriptor
            End If

            Dim normalizedTurnId = NormalizeTurnId(itemState.TurnId)
            If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return descriptor
            End If

            If StringComparer.OrdinalIgnoreCase.Equals(If(descriptor.Kind, String.Empty), "fileChange") Then
                Dim cachedDiffSummary As String = Nothing
                If _turnDiffSummaryByTurnId.TryGetValue(normalizedTurnId, cachedDiffSummary) Then
                    MergeTurnDiffIntoFileChangeDescriptor(descriptor, cachedDiffSummary)
                End If
            End If

            Return descriptor
        End Function

        Friend Shared Function BuildRuntimeItemDescriptorForSnapshot(itemState As TurnItemRuntimeState,
                                                                     Optional timestampText As String = "") As TranscriptEntryDescriptor
            If itemState Is Nothing Then
                Return Nothing
            End If

            Dim itemType = If(itemState.ItemType, String.Empty).Trim().ToLowerInvariant()
            Dim descriptor As New TranscriptEntryDescriptor() With {
                .ThreadId = If(itemState.ThreadId, String.Empty).Trim(),
                .TurnId = If(itemState.TurnId, String.Empty).Trim(),
                .TurnItemStreamSequence = itemState.TurnItemStreamSequence,
                .TurnItemOrderIndex = itemState.TurnItemOrderIndex,
                .TurnItemSortTimestampUtc = If(itemState.StartedAt, itemState.CompletedAt),
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
                    Dim reasoningSummaryText = SanitizeReasoningCardDisplayText(itemState.ReasoningSummaryText)
                    Dim reasoningContentText = SanitizeReasoningCardDisplayText(itemState.ReasoningContentText)
                    Dim hasReasoningSummary = Not String.IsNullOrWhiteSpace(reasoningSummaryText)
                    Dim hasReasoningContent = Not String.IsNullOrWhiteSpace(reasoningContentText)

                    descriptor.Kind = "reasoningCard"
                    descriptor.RoleText = "Reasoning"
                    descriptor.IsMuted = True
                    descriptor.IsReasoning = True
                    descriptor.UseRawReasoningLayout = True
                    If hasReasoningSummary Then
                        descriptor.BodyText = reasoningSummaryText
                    ElseIf hasReasoningContent Then
                        ' Show a placeholder only after reasoning content has actually started streaming.
                        descriptor.BodyText = "Summary pending..."
                    Else
                        descriptor.BodyText = String.Empty
                    End If
                    descriptor.SecondaryText = String.Empty
                    descriptor.DetailsText = reasoningContentText

                Case "commandexecution"
                    descriptor.Kind = "command"
                    descriptor.RoleText = "Command"
                    descriptor.IsCommandLike = True
                    descriptor.BodyText = If(String.IsNullOrWhiteSpace(itemState.CommandText),
                                             "(command)",
                                             itemState.CommandText)
                    descriptor.StatusText = NormalizeStatusToken(itemState.Status)
                    descriptor.SecondaryText = BuildCommandSecondaryText(itemState)
                    descriptor.DetailsText = If(itemState.CommandOutputText, String.Empty).Trim()

                Case "filechange"
                    descriptor.Kind = "fileChange"
                    descriptor.RoleText = "Files"
                    descriptor.BodyText = BuildFileChangeBody(itemState)
                    descriptor.SecondaryText = BuildFileChangeSecondaryText(itemState)
                    descriptor.FileChangeItems = BuildFileChangeInlineItems(itemState)
                    descriptor.DetailsText = BuildFileChangeExtraDetails(itemState)

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
               String.IsNullOrWhiteSpace(descriptor.DetailsText) AndAlso
               Not StringComparer.OrdinalIgnoreCase.Equals(descriptor.Kind, "reasoningCard") Then
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
                    body = "Turn completed."
                Case "started"
                    body = "Turn started."
                Case "interrupted"
                    body = "Turn interrupted."
                Case "failed"
                    body = "Turn failed."
                Case "cancelled", "canceled"
                    body = "Turn canceled."
                Case "aborted"
                    body = "Turn aborted."
                Case Else
                    body = $"Turn ended ({If(String.IsNullOrWhiteSpace(normalizedStatus), "completed", normalizedStatus)})."
            End Select

            Return New TranscriptEntryDescriptor() With {
                .Kind = "turnMarker",
                .TurnId = normalizedTurnId,
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

        Private Shared Function BuildFileChangeInlineItems(itemState As TurnItemRuntimeState) As List(Of TranscriptFileChangeListItemViewModel)
            Dim items As New List(Of TranscriptFileChangeListItemViewModel)()
            If itemState Is Nothing Then
                Return items
            End If

            Dim changes = itemState.FileChangeChanges
            If changes Is Nothing OrElse changes.Count = 0 Then
                Return items
            End If

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

                Dim added = ReadLong(changeObject,
                                     "addedLineCount", "added_line_count",
                                     "addedLines", "added_lines",
                                     "additions", "added")
                Dim removed = ReadLong(changeObject,
                                       "removedLineCount", "removed_line_count",
                                       "removedLines", "removed_lines",
                                       "deletions", "removed")

                items.Add(TranscriptFileChangeListItemViewModel.CreatePathItem(
                    path,
                    SanitizeOptionalLineCountLong(added),
                    SanitizeOptionalLineCountLong(removed)))
                shown += 1
            Next

            If changes.Count > shown Then
                items.Add(TranscriptFileChangeListItemViewModel.CreateOverflowItem(
                    $"... +{(changes.Count - shown).ToString(CultureInfo.InvariantCulture)} more"))
            End If

            Return items
        End Function

        Private Shared Function BuildFileChangeExtraDetails(itemState As TurnItemRuntimeState) As String
            If itemState Is Nothing Then
                Return String.Empty
            End If

            Return If(itemState.FileChangeOutputText, String.Empty).Trim()
        End Function

        Private NotInheritable Class TurnDiffFileCounts
            Public Property Added As Integer?
            Public Property Removed As Integer?
        End Class

        Friend Shared Sub MergeTurnDiffIntoFileChangeDescriptor(descriptor As TranscriptEntryDescriptor,
                                                                turnDiffSummary As String)
            If descriptor Is Nothing OrElse
               Not StringComparer.OrdinalIgnoreCase.Equals(If(descriptor.Kind, String.Empty), "fileChange") Then
                Return
            End If

            Dim normalizedDiffSummary = If(turnDiffSummary, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedDiffSummary) Then
                Return
            End If

            descriptor.DetailsText = MergeFileChangeDetailsWithTurnDiff(descriptor.DetailsText, normalizedDiffSummary)

            If descriptor.FileChangeItems Is Nothing OrElse descriptor.FileChangeItems.Count = 0 Then
                Return
            End If

            descriptor.FileChangeItems = MergeFileChangeInlineItemsWithTurnDiff(descriptor.FileChangeItems, normalizedDiffSummary)
        End Sub

        Private Shared Function MergeFileChangeDetailsWithTurnDiff(existingDetails As String,
                                                                   turnDiffSummary As String) As String
            Dim normalizedDiffSummary = If(turnDiffSummary, String.Empty).Trim()
            Dim normalizedExisting = If(existingDetails, String.Empty).Trim()

            If String.IsNullOrWhiteSpace(normalizedDiffSummary) Then
                Return normalizedExisting
            End If

            If String.IsNullOrWhiteSpace(normalizedExisting) Then
                Return normalizedDiffSummary
            End If

            If StringComparer.Ordinal.Equals(normalizedExisting, normalizedDiffSummary) Then
                Return normalizedExisting
            End If

            If normalizedExisting.IndexOf(normalizedDiffSummary, StringComparison.Ordinal) >= 0 Then
                Return normalizedExisting
            End If

            If normalizedDiffSummary.IndexOf(normalizedExisting, StringComparison.Ordinal) >= 0 Then
                Return normalizedDiffSummary
            End If

            Return normalizedDiffSummary & Environment.NewLine & Environment.NewLine & normalizedExisting
        End Function

        Private Shared Function MergeFileChangeInlineItemsWithTurnDiff(items As IList(Of TranscriptFileChangeListItemViewModel),
                                                                       turnDiffSummary As String) As List(Of TranscriptFileChangeListItemViewModel)
            Dim mergedItems As New List(Of TranscriptFileChangeListItemViewModel)()
            If items Is Nothing OrElse items.Count = 0 Then
                Return mergedItems
            End If

            Dim countsByPath = BuildTurnDiffCountsByPath(turnDiffSummary, items)
            For Each item In items
                If item Is Nothing Then
                    Continue For
                End If

                If item.IsOverflow Then
                    mergedItems.Add(TranscriptFileChangeListItemViewModel.CreateOverflowItem(item.OverflowText))
                    Continue For
                End If

                Dim addedCount = ParseSignedLineBadgeCount(item.AddedLinesText)
                Dim removedCount = ParseSignedLineBadgeCount(item.RemovedLinesText)

                Dim diffCounts As TurnDiffFileCounts = Nothing
                If TryResolveTurnDiffCountsForFilePath(countsByPath, item.FullPathText, diffCounts) AndAlso diffCounts IsNot Nothing Then
                    If diffCounts.Added.HasValue Then
                        addedCount = diffCounts.Added
                    End If
                    If diffCounts.Removed.HasValue Then
                        removedCount = diffCounts.Removed
                    End If
                End If

                mergedItems.Add(TranscriptFileChangeListItemViewModel.CreatePathItem(item.FullPathText, addedCount, removedCount))
            Next

            Return mergedItems
        End Function

        Private Shared Function TryResolveTurnDiffCountsForFilePath(countsByPath As IDictionary(Of String, TurnDiffFileCounts),
                                                                    filePath As String,
                                                                    ByRef resolvedCounts As TurnDiffFileCounts) As Boolean
            resolvedCounts = Nothing
            If countsByPath Is Nothing OrElse countsByPath.Count = 0 Then
                Return False
            End If

            Dim rawFilePath = If(filePath, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(rawFilePath) Then
                Return False
            End If

            Dim direct As TurnDiffFileCounts = Nothing
            If countsByPath.TryGetValue(rawFilePath, direct) AndAlso direct IsNot Nothing Then
                resolvedCounts = direct
                Return True
            End If

            Dim normalizedFilePath = NormalizeDiffPathForMatching(rawFilePath)
            If String.IsNullOrWhiteSpace(normalizedFilePath) Then
                Return False
            End If

            For Each candidate In BuildTurnDiffPathCandidates(rawFilePath)
                Dim candidateCounts As TurnDiffFileCounts = Nothing
                If countsByPath.TryGetValue(candidate, candidateCounts) AndAlso candidateCounts IsNot Nothing Then
                    resolvedCounts = candidateCounts
                    Return True
                End If

                Dim normalizedCandidate = NormalizeDiffPathForMatching(candidate)
                If String.IsNullOrWhiteSpace(normalizedCandidate) Then
                    Continue For
                End If

                For Each pair In countsByPath
                    Dim keyText = If(pair.Key, String.Empty)
                    Dim keyCounts = pair.Value
                    If keyCounts Is Nothing Then
                        Continue For
                    End If

                    Dim normalizedKey = NormalizeDiffPathForMatching(keyText)
                    If String.IsNullOrWhiteSpace(normalizedKey) Then
                        Continue For
                    End If

                    If StringComparer.OrdinalIgnoreCase.Equals(normalizedCandidate, normalizedKey) OrElse
                       normalizedCandidate.EndsWith("/" & normalizedKey, StringComparison.OrdinalIgnoreCase) OrElse
                       normalizedFilePath.EndsWith("/" & normalizedKey, StringComparison.OrdinalIgnoreCase) Then
                        resolvedCounts = keyCounts
                        Return True
                    End If
                Next
            Next

            Return False
        End Function

        Private Shared Function BuildTurnDiffCountsByPath(turnDiffSummary As String,
                                                          items As IEnumerable(Of TranscriptFileChangeListItemViewModel)) As Dictionary(Of String, TurnDiffFileCounts)
            Dim countsByPath As New Dictionary(Of String, TurnDiffFileCounts)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(turnDiffSummary) Then
                Return countsByPath
            End If

            TryPopulateTurnDiffCountsByPathFromJson(turnDiffSummary, countsByPath)
            TryPopulateTurnDiffCountsByUnifiedDiff(turnDiffSummary, countsByPath)

            Dim visiblePathItems As New List(Of TranscriptFileChangeListItemViewModel)()
            Dim lines = If(turnDiffSummary, String.Empty).Replace(ControlChars.CrLf, ControlChars.Lf).
                                               Replace(ControlChars.Cr, ControlChars.Lf).
                                               Split(ControlChars.Lf)
            If lines.Length = 0 OrElse items Is Nothing Then
                Return countsByPath
            End If

            For Each item In items
                If item Is Nothing OrElse item.IsOverflow Then
                    Continue For
                End If

                visiblePathItems.Add(item)

                Dim fullPath = If(item.FullPathText, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(fullPath) OrElse countsByPath.ContainsKey(fullPath) Then
                    Continue For
                End If

                Dim candidates = BuildTurnDiffPathCandidates(fullPath)
                If candidates.Count = 0 Then
                    Continue For
                End If

                For Each rawLine In lines
                    Dim lineText = If(rawLine, String.Empty)
                    If String.IsNullOrWhiteSpace(lineText) OrElse Not LineContainsAnyTurnDiffPathCandidate(lineText, candidates) Then
                        Continue For
                    End If

                    Dim parsedCounts As TurnDiffFileCounts = Nothing
                    If Not TryReadTurnDiffCountsFromText(lineText, parsedCounts) Then
                        Continue For
                    End If

                    countsByPath(fullPath) = parsedCounts
                    Exit For
                Next
            Next

            If visiblePathItems.Count = 1 Then
                Dim onlyItem = visiblePathItems(0)
                Dim onlyPath = If(onlyItem.FullPathText, String.Empty).Trim()
                If Not String.IsNullOrWhiteSpace(onlyPath) Then
                    Dim existing As TurnDiffFileCounts = Nothing
                    If Not countsByPath.TryGetValue(onlyPath, existing) OrElse existing Is Nothing OrElse
                       (Not existing.Added.HasValue AndAlso Not existing.Removed.HasValue) Then
                        Dim overallCounts As TurnDiffFileCounts = Nothing
                        If TryReadTurnDiffOverallCounts(turnDiffSummary, overallCounts) AndAlso overallCounts IsNot Nothing Then
                            countsByPath(onlyPath) = overallCounts
                        End If
                    End If
                End If
            End If

            Return countsByPath
        End Function

        Private Shared Function NormalizeDiffPathForMatching(pathText As String) As String
            Dim text = If(pathText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If

            text = text.Replace("\"c, "/"c)
            Do While text.Contains("//", StringComparison.Ordinal)
                text = text.Replace("//", "/", StringComparison.Ordinal)
            Loop

            Dim renameArrowIndex = text.IndexOf(" -> ", StringComparison.Ordinal)
            If renameArrowIndex >= 0 AndAlso renameArrowIndex + 4 < text.Length Then
                ' Prefer the destination path when matching rename strings against diff paths.
                text = text.Substring(renameArrowIndex + 4).Trim()
            End If

            If text.Length >= 2 AndAlso Char.IsLetter(text(0)) AndAlso text(1) = ":"c Then
                text = text.Substring(2)
            End If

            If text.StartsWith("./", StringComparison.Ordinal) Then
                text = text.Substring(2)
            End If

            text = text.Trim("/"c)
            Return text
        End Function

        Private Shared Sub TryPopulateTurnDiffCountsByUnifiedDiff(turnDiffSummary As String,
                                                                  countsByPath As IDictionary(Of String, TurnDiffFileCounts))
            If String.IsNullOrWhiteSpace(turnDiffSummary) OrElse countsByPath Is Nothing Then
                Return
            End If

            Dim lines = turnDiffSummary.Replace(ControlChars.CrLf, ControlChars.Lf).
                                        Replace(ControlChars.Cr, ControlChars.Lf).
                                        Split(ControlChars.Lf)
            If lines Is Nothing OrElse lines.Length = 0 Then
                Return
            End If

            Dim currentOldPath As String = String.Empty
            Dim currentNewPath As String = String.Empty
            Dim currentAdded As Integer = 0
            Dim currentRemoved As Integer = 0
            Dim hasCurrentFile As Boolean = False
            Dim inHunk As Boolean = False

            For Each rawLine In lines
                Dim lineText = If(rawLine, String.Empty)

                If lineText.StartsWith("diff --git ", StringComparison.Ordinal) Then
                    FlushUnifiedDiffFileCounts(countsByPath, currentOldPath, currentNewPath, currentAdded, currentRemoved)
                    hasCurrentFile = True
                    currentOldPath = String.Empty
                    currentNewPath = String.Empty
                    currentAdded = 0
                    currentRemoved = 0
                    inHunk = False
                    Continue For
                End If

                If Not inHunk AndAlso lineText.StartsWith("--- ", StringComparison.Ordinal) Then
                    If Not hasCurrentFile Then
                        hasCurrentFile = True
                        currentOldPath = String.Empty
                        currentNewPath = String.Empty
                        currentAdded = 0
                        currentRemoved = 0
                        inHunk = False
                    End If

                    Dim parsedOldPath = ParseUnifiedDiffPathMarker(lineText, "--- ")
                    If Not String.IsNullOrWhiteSpace(parsedOldPath) Then
                        currentOldPath = parsedOldPath
                    End If
                    Continue For
                End If

                If Not inHunk AndAlso lineText.StartsWith("+++ ", StringComparison.Ordinal) Then
                    If Not hasCurrentFile Then
                        hasCurrentFile = True
                        currentOldPath = String.Empty
                        currentNewPath = String.Empty
                        currentAdded = 0
                        currentRemoved = 0
                        inHunk = False
                    End If

                    Dim parsedNewPath = ParseUnifiedDiffPathMarker(lineText, "+++ ")
                    If Not String.IsNullOrWhiteSpace(parsedNewPath) Then
                        currentNewPath = parsedNewPath
                    End If
                    Continue For
                End If

                If Not hasCurrentFile Then
                    Continue For
                End If

                If lineText.StartsWith("@@ ", StringComparison.Ordinal) OrElse
                   lineText.StartsWith("@@", StringComparison.Ordinal) Then
                    inHunk = True
                    Continue For
                End If

                If lineText.StartsWith("+", StringComparison.Ordinal) Then
                    currentAdded += 1
                    Continue For
                End If

                If lineText.StartsWith("-", StringComparison.Ordinal) Then
                    currentRemoved += 1
                    Continue For
                End If
            Next

            FlushUnifiedDiffFileCounts(countsByPath, currentOldPath, currentNewPath, currentAdded, currentRemoved)
        End Sub

        Private Shared Function ParseUnifiedDiffPathMarker(lineText As String, markerPrefix As String) As String
            If String.IsNullOrWhiteSpace(lineText) OrElse String.IsNullOrWhiteSpace(markerPrefix) OrElse
               Not lineText.StartsWith(markerPrefix, StringComparison.Ordinal) Then
                Return String.Empty
            End If

            Dim token = lineText.Substring(markerPrefix.Length).Trim()
            If String.IsNullOrWhiteSpace(token) Then
                Return String.Empty
            End If

            Dim tabIndex = token.IndexOf(ControlChars.Tab)
            If tabIndex > 0 Then
                token = token.Substring(0, tabIndex)
            End If

            token = token.Trim()
            If String.IsNullOrWhiteSpace(token) OrElse StringComparer.Ordinal.Equals(token, "/dev/null") Then
                Return String.Empty
            End If

            If token.StartsWith("""", StringComparison.Ordinal) AndAlso token.EndsWith("""", StringComparison.Ordinal) AndAlso token.Length >= 2 Then
                token = token.Substring(1, token.Length - 2)
            End If

            If token.StartsWith("a/", StringComparison.Ordinal) OrElse token.StartsWith("b/", StringComparison.Ordinal) Then
                token = token.Substring(2)
            End If

            Return token.Trim()
        End Function

        Private Shared Sub FlushUnifiedDiffFileCounts(countsByPath As IDictionary(Of String, TurnDiffFileCounts),
                                                      oldPath As String,
                                                      newPath As String,
                                                      addedCount As Integer,
                                                      removedCount As Integer)
            If countsByPath Is Nothing Then
                Return
            End If

            Dim sanitizedAdded = SanitizeOptionalLineCount(If(addedCount > 0, CType(addedCount, Integer?), Nothing))
            Dim sanitizedRemoved = SanitizeOptionalLineCount(If(removedCount > 0, CType(removedCount, Integer?), Nothing))
            If Not sanitizedAdded.HasValue AndAlso Not sanitizedRemoved.HasValue Then
                Return
            End If

            Dim cleanOldPath = If(oldPath, String.Empty).Trim()
            Dim cleanNewPath = If(newPath, String.Empty).Trim()

            If Not String.IsNullOrWhiteSpace(cleanNewPath) Then
                SetTurnDiffCountsForPath(countsByPath, cleanNewPath, sanitizedAdded, sanitizedRemoved)
            End If

            If Not String.IsNullOrWhiteSpace(cleanOldPath) Then
                SetTurnDiffCountsForPath(countsByPath, cleanOldPath, sanitizedAdded, sanitizedRemoved)
            End If

            If Not String.IsNullOrWhiteSpace(cleanOldPath) AndAlso
               Not String.IsNullOrWhiteSpace(cleanNewPath) AndAlso
               Not StringComparer.OrdinalIgnoreCase.Equals(cleanOldPath, cleanNewPath) Then
                SetTurnDiffCountsForPath(countsByPath,
                                         $"{cleanOldPath} -> {cleanNewPath}",
                                         sanitizedAdded,
                                         sanitizedRemoved)
            End If
        End Sub

        Private Shared Sub SetTurnDiffCountsForPath(countsByPath As IDictionary(Of String, TurnDiffFileCounts),
                                                    pathText As String,
                                                    addedCount As Integer?,
                                                    removedCount As Integer?)
            If countsByPath Is Nothing Then
                Return
            End If

            Dim normalizedPath = If(pathText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedPath) Then
                Return
            End If

            Dim existing As TurnDiffFileCounts = Nothing
            If Not countsByPath.TryGetValue(normalizedPath, existing) OrElse existing Is Nothing Then
                existing = New TurnDiffFileCounts()
                countsByPath(normalizedPath) = existing
            End If

            If addedCount.HasValue Then
                existing.Added = addedCount
            End If
            If removedCount.HasValue Then
                existing.Removed = removedCount
            End If
        End Sub

        Private Shared Sub TryPopulateTurnDiffCountsByPathFromJson(turnDiffSummary As String,
                                                                   countsByPath As IDictionary(Of String, TurnDiffFileCounts))
            If String.IsNullOrWhiteSpace(turnDiffSummary) OrElse countsByPath Is Nothing Then
                Return
            End If

            Try
                Dim root = JsonNode.Parse(turnDiffSummary)
                If root Is Nothing Then
                    Return
                End If

                CollectTurnDiffCountsFromJsonNode(root, countsByPath, depth:=0)
            Catch
                ' Plain-text summaries are common; JSON parse is best-effort only.
            End Try
        End Sub

        Private Shared Sub CollectTurnDiffCountsFromJsonNode(node As JsonNode,
                                                             countsByPath As IDictionary(Of String, TurnDiffFileCounts),
                                                             depth As Integer)
            If node Is Nothing OrElse countsByPath Is Nothing OrElse depth > 4 Then
                Return
            End If

            Dim obj = TryCast(node, JsonObject)
            If obj IsNot Nothing Then
                TryAddTurnDiffCountsFromJsonObject(obj, countsByPath)

                Dim nestedDiff = TryCast(obj("diff"), JsonNode)
                If nestedDiff IsNot Nothing Then
                    CollectTurnDiffCountsFromJsonNode(nestedDiff, countsByPath, depth + 1)
                End If

                For Each arrayKey In New String() {"files", "changes", "items"}
                    Dim nestedArray = TryCast(obj(arrayKey), JsonNode)
                    If nestedArray IsNot Nothing Then
                        CollectTurnDiffCountsFromJsonNode(nestedArray, countsByPath, depth + 1)
                    End If
                Next

                Return
            End If

            Dim arr = TryCast(node, JsonArray)
            If arr Is Nothing Then
                Return
            End If

            For Each child In arr
                CollectTurnDiffCountsFromJsonNode(child, countsByPath, depth + 1)
            Next
        End Sub

        Private Shared Sub TryAddTurnDiffCountsFromJsonObject(obj As JsonObject,
                                                              countsByPath As IDictionary(Of String, TurnDiffFileCounts))
            If obj Is Nothing OrElse countsByPath Is Nothing Then
                Return
            End If

            Dim path = ReadString(obj, "path", "file", "filename", "name")
            If String.IsNullOrWhiteSpace(path) Then
                Return
            End If

            Dim addedCount = SanitizeOptionalLineCountLong(ReadLong(obj,
                                                                    "addedLineCount", "added_line_count",
                                                                    "addedLines", "added_lines",
                                                                    "additions", "added", "plus"))
            Dim removedCount = SanitizeOptionalLineCountLong(ReadLong(obj,
                                                                      "removedLineCount", "removed_line_count",
                                                                      "removedLines", "removed_lines",
                                                                      "deletions", "removed", "minus"))
            If Not addedCount.HasValue AndAlso Not removedCount.HasValue Then
                Return
            End If

            Dim normalizedPath = path.Trim()
            Dim existing As TurnDiffFileCounts = Nothing
            If Not countsByPath.TryGetValue(normalizedPath, existing) OrElse existing Is Nothing Then
                existing = New TurnDiffFileCounts()
                countsByPath(normalizedPath) = existing
            End If

            If addedCount.HasValue Then
                existing.Added = addedCount
            End If
            If removedCount.HasValue Then
                existing.Removed = removedCount
            End If
        End Sub

        Private Shared Function BuildTurnDiffPathCandidates(fullPathText As String) As List(Of String)
            Dim candidates As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each candidate In ExpandTurnDiffPathCandidates(fullPathText)
                Dim normalized = If(candidate, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalized) Then
                    Continue For
                End If

                If seen.Add(normalized) Then
                    candidates.Add(normalized)
                End If
            Next

            Return candidates
        End Function

        Private Shared Iterator Function ExpandTurnDiffPathCandidates(fullPathText As String) As IEnumerable(Of String)
            Dim rawText = If(fullPathText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(rawText) Then
                Return
            End If

            Yield rawText

            Dim renameArrowIndex = rawText.IndexOf(" -> ", StringComparison.Ordinal)
            If renameArrowIndex > 0 AndAlso renameArrowIndex + 4 < rawText.Length Then
                Yield rawText.Substring(0, renameArrowIndex).Trim()
                Yield rawText.Substring(renameArrowIndex + 4).Trim()
            End If

            Dim baseCandidates As New List(Of String)()
            baseCandidates.Add(rawText)
            If renameArrowIndex > 0 AndAlso renameArrowIndex + 4 < rawText.Length Then
                baseCandidates.Add(rawText.Substring(0, renameArrowIndex).Trim())
                baseCandidates.Add(rawText.Substring(renameArrowIndex + 4).Trim())
            End If

            For Each baseCandidate In baseCandidates
                Dim normalized = If(baseCandidate, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalized) Then
                    Continue For
                End If

                If normalized.Contains("\"c) Then
                    Yield normalized.Replace("\"c, "/"c)
                End If
                If normalized.Contains("/"c) Then
                    Yield normalized.Replace("/"c, "\"c)
                End If

                If normalized.Length > 2 AndAlso normalized(1) = "/"c AndAlso (normalized(0) = "a"c OrElse normalized(0) = "b"c) Then
                    Continue For
                End If

                Yield "a/" & normalized
                Yield "b/" & normalized
            Next
        End Function

        Private Shared Function LineContainsAnyTurnDiffPathCandidate(lineText As String,
                                                                     candidates As IEnumerable(Of String)) As Boolean
            If String.IsNullOrWhiteSpace(lineText) OrElse candidates Is Nothing Then
                Return False
            End If

            For Each candidate In candidates
                If String.IsNullOrWhiteSpace(candidate) Then
                    Continue For
                End If

                If lineText.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function TryReadTurnDiffCountsFromText(lineText As String,
                                                              ByRef counts As TurnDiffFileCounts) As Boolean
            counts = Nothing
            If String.IsNullOrWhiteSpace(lineText) Then
                Return False
            End If

            Dim addedMatch = TurnDiffAddedCountRegex.Match(lineText)
            Dim removedMatch = TurnDiffRemovedCountRegex.Match(lineText)
            Dim addedCount = ParseRegexLineCount(addedMatch)
            Dim removedCount = ParseRegexLineCount(removedMatch)
            If Not addedCount.HasValue AndAlso Not removedCount.HasValue Then
                Dim insertionMatch = TurnDiffInsertionsRegex.Match(lineText)
                Dim deletionMatch = TurnDiffDeletionsRegex.Match(lineText)
                addedCount = ParseRegexLineCount(insertionMatch)
                removedCount = ParseRegexLineCount(deletionMatch)
            End If

            If Not addedCount.HasValue AndAlso Not removedCount.HasValue Then
                TryReadTurnDiffCountsFromDiffStatBar(lineText, addedCount, removedCount)
            End If

            If Not addedCount.HasValue AndAlso Not removedCount.HasValue Then
                Return False
            End If

            counts = New TurnDiffFileCounts() With {
                .Added = addedCount,
                .Removed = removedCount
            }
            Return True
        End Function

        Private Shared Function ParseRegexLineCount(match As Match) As Integer?
            If match Is Nothing OrElse Not match.Success Then
                Return Nothing
            End If

            Dim numberText = If(match.Groups("n")?.Value, String.Empty)
            If String.IsNullOrWhiteSpace(numberText) Then
                Return Nothing
            End If

            Dim parsed As Integer
            If Not Integer.TryParse(numberText, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                Return Nothing
            End If

            Return SanitizeOptionalLineCount(parsed)
        End Function

        Private Shared Sub TryReadTurnDiffCountsFromDiffStatBar(lineText As String,
                                                                ByRef addedCount As Integer?,
                                                                ByRef removedCount As Integer?)
            If String.IsNullOrWhiteSpace(lineText) Then
                Return
            End If

            Dim pipeIndex = lineText.IndexOf("|"c)
            If pipeIndex < 0 OrElse pipeIndex + 1 >= lineText.Length Then
                Return
            End If

            Dim rhs = lineText.Substring(pipeIndex + 1)
            Dim plusCount = 0
            Dim minusCount = 0
            For Each ch In rhs
                If ch = "+"c Then
                    plusCount += 1
                ElseIf ch = "-"c Then
                    minusCount += 1
                End If
            Next

            If plusCount > 0 Then
                addedCount = plusCount
            End If
            If minusCount > 0 Then
                removedCount = minusCount
            End If
        End Sub

        Private Shared Function TryReadTurnDiffOverallCounts(turnDiffSummary As String,
                                                             ByRef counts As TurnDiffFileCounts) As Boolean
            counts = Nothing
            If String.IsNullOrWhiteSpace(turnDiffSummary) Then
                Return False
            End If

            Dim normalized = turnDiffSummary.Replace(ControlChars.CrLf, ControlChars.Lf).
                                             Replace(ControlChars.Cr, ControlChars.Lf)
            Dim lines = normalized.Split(ControlChars.Lf)

            ' Prefer explicit totals lines ("N insertions(+), M deletions(-)") over raw +n/-n matches.
            For Each rawLine In lines
                Dim lineText = If(rawLine, String.Empty)
                If String.IsNullOrWhiteSpace(lineText) Then
                    Continue For
                End If

                Dim insertionMatch = TurnDiffInsertionsRegex.Match(lineText)
                Dim deletionMatch = TurnDiffDeletionsRegex.Match(lineText)
                Dim addedCount = ParseRegexLineCount(insertionMatch)
                Dim removedCount = ParseRegexLineCount(deletionMatch)
                If addedCount.HasValue OrElse removedCount.HasValue Then
                    counts = New TurnDiffFileCounts() With {
                        .Added = addedCount,
                        .Removed = removedCount
                    }
                    Return True
                End If
            Next

            ' Fallback: if the summary is a single line with +n/-n totals, use that.
            If lines.Length = 1 Then
                Dim lineCounts As TurnDiffFileCounts = Nothing
                If TryReadTurnDiffCountsFromText(lines(0), lineCounts) Then
                    counts = lineCounts
                    Return True
                End If
            End If

            Return False
        End Function

        Private Shared Function ParseSignedLineBadgeCount(value As String) As Integer?
            Dim text = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return Nothing
            End If

            If text.StartsWith("+", StringComparison.Ordinal) OrElse
               text.StartsWith("-", StringComparison.Ordinal) Then
                text = text.Substring(1)
            End If

            Dim parsed As Integer
            If Not Integer.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                Return Nothing
            End If

            Return SanitizeOptionalLineCount(parsed)
        End Function

        Private Sub AppendTranscriptOrderingSnapshotDump(stage As String)
            If Not EnableTranscriptOrderingProtocolDebug Then
                Return
            End If

            Try
                Dim normalizedStage = If(stage, String.Empty).Trim()
                AppendTranscriptOrderingProtocolDebug(
                    "snapshot_dump_begin",
                    $"stage={FormatProtocolDebugToken(normalizedStage)} item_count={_items.Count.ToString(CultureInfo.InvariantCulture)} runtime_key_count={_runtimeEntriesByKey.Count.ToString(CultureInfo.InvariantCulture)}")

                Const maxLoggedRows As Integer = 400
                Dim logged = 0
                For i = 0 To _items.Count - 1
                    If logged >= maxLoggedRows Then
                        Exit For
                    End If

                    Dim entry = _items(i)
                    If entry Is Nothing Then
                        Continue For
                    End If

                    Dim runtimeKey = If(entry.RuntimeKey, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(runtimeKey) Then
                        Continue For
                    End If

                    AppendTranscriptOrderingProtocolDebug(
                        "snapshot_row",
                        $"stage={FormatProtocolDebugToken(normalizedStage)} idx={i.ToString(CultureInfo.InvariantCulture)} kind={FormatProtocolDebugToken(entry.Kind)} turnId={FormatProtocolDebugToken(entry.TurnId)} runtimeKey={FormatProtocolDebugToken(runtimeKey)} seq={FormatOrderingDebugNullableLong(entry.TurnItemStreamSequence)} ord={FormatOrderingDebugNullableInt(entry.TurnItemOrderIndex)} ts={FormatOrderingDebugTimestamp(entry.TurnItemSortTimestampUtc)}")
                    logged += 1
                Next

                If _items.Count > logged Then
                    AppendTranscriptOrderingProtocolDebug(
                        "snapshot_dump_truncated",
                        $"stage={FormatProtocolDebugToken(normalizedStage)} logged={logged.ToString(CultureInfo.InvariantCulture)} total={_items.Count.ToString(CultureInfo.InvariantCulture)}")
                End If
            Catch ex As Exception
                Dim safeMessage = If(ex.Message, ex.GetType().Name)
                AppendTranscriptOrderingProtocolDebug("snapshot_dump_error", $"message={FormatProtocolDebugToken(safeMessage)}")
            End Try
        End Sub

        Private Sub AppendTranscriptOrderingResolveDecisionDebug(reason As String,
                                                                 runtimeKey As String,
                                                                 descriptor As TranscriptEntryDescriptor,
                                                                 insertIndex As Integer,
                                                                 candidate As TranscriptEntryViewModel)
            If Not EnableTranscriptOrderingProtocolDebug Then
                Return
            End If

            Dim candidateSummary As String = "candidate=_"
            If candidate IsNot Nothing Then
                candidateSummary =
                    $"candidate_kind={FormatProtocolDebugToken(candidate.Kind)} candidate_turnId={FormatProtocolDebugToken(candidate.TurnId)} candidate_runtimeKey={FormatProtocolDebugToken(candidate.RuntimeKey)} candidate_seq={FormatOrderingDebugNullableLong(candidate.TurnItemStreamSequence)} candidate_ord={FormatOrderingDebugNullableInt(candidate.TurnItemOrderIndex)} candidate_ts={FormatOrderingDebugTimestamp(candidate.TurnItemSortTimestampUtc)}"
            End If

            AppendTranscriptOrderingProtocolDebug(
                "resolve_insert_index",
                $"reason={FormatProtocolDebugToken(reason)} insert_index={insertIndex.ToString(CultureInfo.InvariantCulture)} target_kind={FormatProtocolDebugToken(descriptor?.Kind)} target_turnId={FormatProtocolDebugToken(descriptor?.TurnId)} target_runtimeKey={FormatProtocolDebugToken(runtimeKey)} target_seq={FormatOrderingDebugNullableLong(descriptor?.TurnItemStreamSequence)} target_ord={FormatOrderingDebugNullableInt(descriptor?.TurnItemOrderIndex)} target_ts={FormatOrderingDebugTimestamp(descriptor?.TurnItemSortTimestampUtc)} {candidateSummary}")
        End Sub

        Private Function BuildOrderingUpsertDebugMessage(runtimeKey As String,
                                                         descriptor As TranscriptEntryDescriptor,
                                                         oldIndex As Integer,
                                                         newIndex As Integer,
                                                         action As String) As String
            Dim beforeNeighbor As String = "_"
            Dim afterNeighbor As String = "_"

            If newIndex > 0 AndAlso newIndex - 1 < _items.Count Then
                beforeNeighbor = DescribeOrderingEntryCompact(_items(newIndex - 1))
            End If

            If newIndex >= 0 AndAlso newIndex + 1 < _items.Count Then
                afterNeighbor = DescribeOrderingEntryCompact(_items(newIndex + 1))
            End If

            Return $"action={FormatProtocolDebugToken(action)} old_index={oldIndex.ToString(CultureInfo.InvariantCulture)} new_index={newIndex.ToString(CultureInfo.InvariantCulture)} target_kind={FormatProtocolDebugToken(descriptor?.Kind)} target_turnId={FormatProtocolDebugToken(descriptor?.TurnId)} runtimeKey={FormatProtocolDebugToken(runtimeKey)} seq={FormatOrderingDebugNullableLong(descriptor?.TurnItemStreamSequence)} ord={FormatOrderingDebugNullableInt(descriptor?.TurnItemOrderIndex)} ts={FormatOrderingDebugTimestamp(descriptor?.TurnItemSortTimestampUtc)} before={FormatProtocolDebugToken(beforeNeighbor)} after={FormatProtocolDebugToken(afterNeighbor)}"
        End Function

        Private Shared Function DescribeOrderingEntryCompact(entry As TranscriptEntryViewModel) As String
            If entry Is Nothing Then
                Return "_"
            End If

            Return $"kind={If(entry.Kind, String.Empty).Trim()}|turn={If(entry.TurnId, String.Empty).Trim()}|key={If(entry.RuntimeKey, String.Empty).Trim()}|seq={FormatOrderingDebugNullableLong(entry.TurnItemStreamSequence)}|ord={FormatOrderingDebugNullableInt(entry.TurnItemOrderIndex)}"
        End Function

        Private Sub AppendTranscriptOrderingProtocolDebug(stage As String, details As String)
            If Not EnableTranscriptOrderingProtocolDebug Then
                Return
            End If

            Dim safeStage = If(stage, String.Empty).Trim()
            Dim safeDetails = If(details, String.Empty).Trim()
            AppendProtocolChunk($"[{Date.Now:HH:mm:ss}] debug: transcript_ordering stage={safeStage} {safeDetails}{Environment.NewLine}")
        End Sub

        Private Shared Function FormatOrderingDebugNullableInt(value As Integer?) As String
            If Not value.HasValue Then
                Return "_"
            End If

            Return value.Value.ToString(CultureInfo.InvariantCulture)
        End Function

        Private Shared Function FormatOrderingDebugNullableLong(value As Long?) As String
            If Not value.HasValue Then
                Return "_"
            End If

            Return value.Value.ToString(CultureInfo.InvariantCulture)
        End Function

        Private Shared Function FormatOrderingDebugTimestamp(value As DateTimeOffset?) As String
            If Not value.HasValue Then
                Return "_"
            End If

            Return value.Value.UtcDateTime.ToString("O", CultureInfo.InvariantCulture)
        End Function

        Private Sub AppendFileChangeTurnDiffMergeProtocolDebug(stage As String,
                                                               turnId As String,
                                                               runtimeKey As String,
                                                               turnDiffSummary As String,
                                                               descriptor As TranscriptEntryDescriptor,
                                                               Optional note As String = "")
            If Not EnableFileChangeTurnDiffMergeProtocolDebug Then
                Return
            End If

            Try
                Dim normalizedStage = If(stage, String.Empty).Trim()
                Dim normalizedTurnId = NormalizeTurnId(turnId)
                Dim normalizedRuntimeKey = If(runtimeKey, String.Empty).Trim()
                Dim diffText = If(turnDiffSummary, String.Empty)
                Dim diffTrimmed = diffText.Trim()
                Dim hasUnifiedDiff = diffTrimmed.StartsWith("diff --git ", StringComparison.Ordinal)

                Dim itemCount = 0
                Dim pathItemCount = 0
                If descriptor?.FileChangeItems IsNot Nothing Then
                    itemCount = descriptor.FileChangeItems.Count
                    For Each item In descriptor.FileChangeItems
                        If item Is Nothing OrElse item.IsOverflow Then
                            Continue For
                        End If
                        pathItemCount += 1
                    Next
                End If

                Dim parsedMap As Dictionary(Of String, TurnDiffFileCounts) = Nothing
                If descriptor?.FileChangeItems IsNot Nothing AndAlso descriptor.FileChangeItems.Count > 0 AndAlso
                   Not String.IsNullOrWhiteSpace(diffTrimmed) Then
                    parsedMap = BuildTurnDiffCountsByPath(diffTrimmed, descriptor.FileChangeItems)
                Else
                    parsedMap = New Dictionary(Of String, TurnDiffFileCounts)(StringComparer.OrdinalIgnoreCase)
                End If

                Dim builder As New StringBuilder()
                builder.Append("[")
                builder.Append(Date.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture))
                builder.Append("] debug: filechange_turndiff_merge")
                builder.Append(" stage=").Append(FormatProtocolDebugToken(normalizedStage))
                builder.Append(" turnId=").Append(FormatProtocolDebugToken(normalizedTurnId))
                builder.Append(" runtimeKey=").Append(FormatProtocolDebugToken(normalizedRuntimeKey))
                builder.Append(" diff_len=").Append(diffTrimmed.Length.ToString(CultureInfo.InvariantCulture))
                builder.Append(" unified=").Append(If(hasUnifiedDiff, "1", "0"))
                builder.Append(" descriptor_present=").Append(If(descriptor Is Nothing, "0", "1"))
                builder.Append(" file_items=").Append(pathItemCount.ToString(CultureInfo.InvariantCulture))
                builder.Append(" parsed_keys=").Append(parsedMap.Count.ToString(CultureInfo.InvariantCulture))
                If Not String.IsNullOrWhiteSpace(note) Then
                    builder.Append(" note=").Append(FormatProtocolDebugToken(note))
                End If

                If parsedMap.Count > 0 Then
                    builder.Append(" keys=")
                    AppendProtocolDebugParsedKeySample(builder, parsedMap)
                End If

                If descriptor?.FileChangeItems IsNot Nothing AndAlso descriptor.FileChangeItems.Count > 0 Then
                    builder.Append(" files=")
                    AppendProtocolDebugFileBadgeSummary(builder, descriptor.FileChangeItems, parsedMap)
                End If

                builder.AppendLine()
                AppendProtocolChunk(builder.ToString())
            Catch ex As Exception
                Dim safeMessage = If(ex.Message, ex.GetType().Name)
                AppendProtocolChunk($"[{Date.Now:HH:mm:ss}] debug: filechange_turndiff_merge_error message={FormatProtocolDebugToken(safeMessage)}{Environment.NewLine}")
            End Try
        End Sub

        Private Shared Sub AppendProtocolDebugParsedKeySample(builder As StringBuilder,
                                                              parsedMap As IDictionary(Of String, TurnDiffFileCounts))
            If builder Is Nothing OrElse parsedMap Is Nothing OrElse parsedMap.Count = 0 Then
                Return
            End If

            Dim written = 0
            For Each pair In parsedMap
                If written > 0 Then
                    builder.Append(";")
                End If
                builder.Append(FormatProtocolDebugToken(pair.Key))
                builder.Append(":")
                builder.Append(FormatProtocolDebugCountPair(pair.Value))
                written += 1
                If written >= 6 Then
                    Exit For
                End If
            Next

            If parsedMap.Count > written Then
                builder.Append(";...")
            End If
        End Sub

        Private Shared Sub AppendProtocolDebugFileBadgeSummary(builder As StringBuilder,
                                                               items As IEnumerable(Of TranscriptFileChangeListItemViewModel),
                                                               parsedMap As IDictionary(Of String, TurnDiffFileCounts))
            If builder Is Nothing OrElse items Is Nothing Then
                Return
            End If

            Dim written = 0
            For Each item In items
                If item Is Nothing OrElse item.IsOverflow Then
                    Continue For
                End If

                If written > 0 Then
                    builder.Append(";")
                End If

                Dim fullPath = If(item.FullPathText, String.Empty).Trim()
                builder.Append(FormatProtocolDebugToken(fullPath))
                builder.Append("=")
                builder.Append(FormatProtocolDebugToken(item.AddedLinesText))
                builder.Append("/")
                builder.Append(FormatProtocolDebugToken(item.RemovedLinesText))

                Dim parsedCounts As TurnDiffFileCounts = Nothing
                If parsedMap IsNot Nothing AndAlso
                   TryResolveTurnDiffCountsForFilePath(parsedMap, fullPath, parsedCounts) AndAlso
                   parsedCounts IsNot Nothing Then
                    builder.Append("(parsed:")
                    builder.Append(FormatProtocolDebugCountPair(parsedCounts))
                    builder.Append(")")
                Else
                    builder.Append("(parsed:none)")
                End If

                written += 1
                If written >= 8 Then
                    Exit For
                End If
            Next

            If written = 0 Then
                builder.Append("none")
            End If
        End Sub

        Private Shared Function FormatProtocolDebugCountPair(counts As TurnDiffFileCounts) As String
            If counts Is Nothing Then
                Return "none"
            End If

            Dim plusText = If(counts.Added.HasValue AndAlso counts.Added.Value > 0,
                              "+" & counts.Added.Value.ToString(CultureInfo.InvariantCulture),
                              "_")
            Dim minusText = If(counts.Removed.HasValue AndAlso counts.Removed.Value > 0,
                               "-" & counts.Removed.Value.ToString(CultureInfo.InvariantCulture),
                               "_")
            Return plusText & "/" & minusText
        End Function

        Private Shared Function FormatProtocolDebugToken(value As String) As String
            Dim text = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return "_"
            End If

            text = text.Replace(ControlChars.Cr, " "c).
                        Replace(ControlChars.Lf, " "c).
                        Replace(ControlChars.Tab, " "c)
            Do While text.Contains("  ", StringComparison.Ordinal)
                text = text.Replace("  ", " ", StringComparison.Ordinal)
            Loop

            If text.Length > 120 Then
                text = text.Substring(0, 117) & "..."
            End If

            If text.IndexOfAny(New Char() {" "c, ";"c, "="c}) >= 0 Then
                Return """" & text.Replace("""", "'"c) & """"
            End If

            Return text
        End Function

        Private Shared Sub ApplyFileChangeInlineListFormatting(entry As TranscriptEntryViewModel)
            If entry Is Nothing Then
                Return
            End If

            Dim parsed = ParseFileChangeInlineList(entry.DetailsText)
            entry.SetFileChangeItems(parsed.Items)
            entry.DetailsText = parsed.RemainingDetails
        End Sub

        Private Shared Function ParseFileChangeInlineList(detailsText As String) As (Items As List(Of TranscriptFileChangeListItemViewModel), RemainingDetails As String)
            Dim items As New List(Of TranscriptFileChangeListItemViewModel)()
            Dim normalized = If(detailsText, String.Empty).Replace(ControlChars.CrLf, ControlChars.Lf).
                                                Replace(ControlChars.Cr, ControlChars.Lf)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return (items, String.Empty)
            End If

            Dim lines = normalized.Split(ControlChars.Lf)
            Dim lineIndex = 0
            Dim parsedAny = False

            While lineIndex < lines.Length
                Dim rawLine = If(lines(lineIndex), String.Empty)
                If String.IsNullOrWhiteSpace(rawLine) Then
                    Exit While
                End If

                Dim parsedItem = TryParseFileChangeInlineListLine(rawLine)
                If parsedItem Is Nothing Then
                    If parsedAny Then
                        Exit While
                    End If

                    Return (New List(Of TranscriptFileChangeListItemViewModel)(), normalized.Trim())
                End If

                items.Add(parsedItem)
                parsedAny = True
                lineIndex += 1
            End While

            If Not parsedAny Then
                Return (items, normalized.Trim())
            End If

            While lineIndex < lines.Length AndAlso String.IsNullOrWhiteSpace(lines(lineIndex))
                lineIndex += 1
            End While

            Dim remainingBuilder As New StringBuilder()
            While lineIndex < lines.Length
                If remainingBuilder.Length > 0 Then
                    remainingBuilder.AppendLine()
                End If
                remainingBuilder.Append(lines(lineIndex))
                lineIndex += 1
            End While

            Return (items, remainingBuilder.ToString().Trim())
        End Function

        Private Shared Function TryParseFileChangeInlineListLine(rawLine As String) As TranscriptFileChangeListItemViewModel
            Dim lineText = If(rawLine, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(lineText) Then
                Return Nothing
            End If

            If lineText.StartsWith("... +", StringComparison.Ordinal) AndAlso
               lineText.EndsWith(" more", StringComparison.Ordinal) Then
                Return TranscriptFileChangeListItemViewModel.CreateOverflowItem(lineText)
            End If

            Dim separatorIndex = lineText.IndexOf(": ", StringComparison.Ordinal)
            If separatorIndex > 0 Then
                Dim candidatePath = lineText.Substring(separatorIndex + 2).Trim()
                If LooksLikeFileChangePath(candidatePath) Then
                    Return TranscriptFileChangeListItemViewModel.CreatePathItem(candidatePath)
                End If
            End If

            If LooksLikeFileChangePath(lineText) Then
                Return TranscriptFileChangeListItemViewModel.CreatePathItem(lineText)
            End If

            Return Nothing
        End Function

        Private Shared Function LooksLikeFileChangePath(value As String) As Boolean
            Dim text = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return False
            End If

            If text.Contains(" -> ", StringComparison.Ordinal) OrElse
               text.Contains("/"c) OrElse
               text.Contains("\"c) Then
                Return True
            End If

            If text.StartsWith("."c) Then
                Return True
            End If

            Dim dotIndex = text.LastIndexOf("."c)
            If dotIndex > 0 AndAlso dotIndex < text.Length - 1 Then
                Return True
            End If

            Return text.IndexOf(" "c) < 0
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

        Private Function AppendDescriptor(descriptor As TranscriptEntryDescriptor,
                                          appendToRaw As Boolean,
                                          Optional rawRole As String = Nothing,
                                          Optional rawBody As String = Nothing) As TranscriptEntryViewModel
            If descriptor Is Nothing Then
                Return Nothing
            End If

            If appendToRaw Then
                Dim role = If(rawRole, descriptor.RoleText)
                Dim body = If(rawBody, descriptor.BodyText)
                If Not String.IsNullOrWhiteSpace(role) AndAlso Not String.IsNullOrWhiteSpace(body) Then
                    AppendRawRoleLine(role, body)
                End If
            End If

            If Not ShouldDisplayDescriptor(descriptor) Then
                Return Nothing
            End If

            Dim entry = CreateEntryFromDescriptor(descriptor)
            _items.Add(entry)
            TrimEntriesIfNeeded()
            Return entry
        End Function

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
                .RuntimeKey = If(descriptor.RuntimeKey, String.Empty),
                .ThreadId = If(descriptor.ThreadId, String.Empty),
                .TurnId = If(descriptor.TurnId, String.Empty),
                .TurnItemStreamSequence = descriptor.TurnItemStreamSequence,
                .TurnItemOrderIndex = descriptor.TurnItemOrderIndex,
                .TurnItemSortTimestampUtc = descriptor.TurnItemSortTimestampUtc,
                .TimestampText = If(descriptor.TimestampText, String.Empty),
                .RoleText = If(descriptor.RoleText, String.Empty),
                .BodyText = If(descriptor.BodyText, String.Empty),
                .StatusText = If(descriptor.StatusText, String.Empty),
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

            If StringComparer.OrdinalIgnoreCase.Equals(entry.Kind, "fileChange") Then
                If descriptor.FileChangeItems IsNot Nothing AndAlso descriptor.FileChangeItems.Count > 0 Then
                    entry.SetFileChangeItems(descriptor.FileChangeItems)
                Else
                    ApplyFileChangeInlineListFormatting(entry)
                End If
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
            If StringComparer.OrdinalIgnoreCase.Equals(kind, "reasoningCard") OrElse
               StringComparer.OrdinalIgnoreCase.Equals(kind, "fileChange") Then
                Return True
            End If

            Return _collapseCommandDetailsByDefault
        End Function

        Private Shared Function ShouldAllowDetailsCollapse(kind As String) As Boolean
            Select Case If(kind, String.Empty).Trim().ToLowerInvariant()
                Case "command", "reasoningcard", "filechange",
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

        Private Shared Function NormalizeStatusToken(value As String) As String
            Dim text = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If

            Dim compact = text.Replace("-", String.Empty, StringComparison.Ordinal).
                               Replace("_", String.Empty, StringComparison.Ordinal).
                               Replace(" ", String.Empty, StringComparison.Ordinal).
                               ToLowerInvariant()

            Select Case compact
                Case "inprogress", "running", "started"
                    Return "in_progress"
                Case "completed", "complete", "succeeded", "success", "ok", "done"
                    Return "completed"
                Case "failed", "failure", "error"
                    Return "failed"
                Case Else
                    Return text.ToLowerInvariant()
            End Select
        End Function

        Private Shared Function SanitizeReasoningCardDisplayText(value As String) As String
            Dim source = If(value, String.Empty)
            If String.IsNullOrWhiteSpace(source) Then
                Return String.Empty
            End If

            Dim normalized = source.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)
            For i = 0 To lines.Length - 1
                lines(i) = SanitizeReasoningCardDisplayLine(lines(i))
            Next

            Return String.Join(vbLf, lines).Trim()
        End Function

        Private Shared Function SanitizeReasoningCardDisplayLine(value As String) As String
            Dim working = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(working) Then
                Return String.Empty
            End If

            ' Summary parts can arrive as "**...**" wrappers per line.
            ' Remove markdown bold markers anywhere in the line, then trim bracket wrappers.
            working = working.Replace("**", String.Empty, StringComparison.Ordinal)

            Dim changed = True
            Do While changed AndAlso working.Length > 0
                changed = False

                If working.StartsWith("[", StringComparison.Ordinal) Then
                    working = working.Substring(1).TrimStart()
                    changed = True
                End If

                If working.EndsWith("]", StringComparison.Ordinal) AndAlso working.Length >= 1 Then
                    working = working.Substring(0, working.Length - 1).TrimEnd()
                    changed = True
                End If
            Loop

            Return working
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

        Private Shared Function FormatRuntimeItemTimestamp(itemState As TurnItemRuntimeState) As String
            If itemState Is Nothing Then
                Return FormatLiveTimestamp()
            End If

            Dim timestamp = If(itemState.StartedAt, itemState.CompletedAt)
            If Not timestamp.HasValue Then
                Return FormatLiveTimestamp()
            End If

            Return timestamp.Value.ToLocalTime().ToString("HH:mm")
        End Function

        Private Shared Function SanitizeOptionalLineCount(value As Integer?) As Integer?
            If Not value.HasValue OrElse value.Value <= 0 Then
                Return Nothing
            End If

            Return value.Value
        End Function

        Private Shared Function SanitizeOptionalLineCountLong(value As Long?) As Integer?
            If Not value.HasValue OrElse value.Value <= 0 Then
                Return Nothing
            End If

            If value.Value > Integer.MaxValue Then
                Return Integer.MaxValue
            End If

            Return CInt(value.Value)
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
                RemoveRuntimeKeyAssociations(keyToRemove)
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
