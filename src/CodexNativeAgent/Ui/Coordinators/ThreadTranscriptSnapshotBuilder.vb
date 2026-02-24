Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Text.Json.Nodes
Imports CodexNativeAgent.Ui.ViewModels
Imports CodexNativeAgent.Ui.ViewModels.Transcript

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class ThreadTranscriptSnapshot
        Public Property RawText As String = String.Empty
        Public ReadOnly Property DisplayEntries As New List(Of TranscriptEntryDescriptor)()
        Public ReadOnly Property DebugMessages As New List(Of String)()
        Public ReadOnly Property AssistantPhaseHintsByItemKey As New Dictionary(Of String, String)(StringComparer.Ordinal)
    End Class

    Public NotInheritable Class ThreadTranscriptSnapshotBuilder
        Private NotInheritable Class SnapshotTurnContext
            Public Property ThreadId As String = String.Empty
            Public Property TurnId As String = String.Empty
        End Class

        Private NotInheritable Class SnapshotDescriptorAccumulator
            Public ReadOnly Property Order As New List(Of String)()
            Public ReadOnly Property ByKey As New Dictionary(Of String, TranscriptEntryDescriptor)(StringComparer.Ordinal)
        End Class

        Private NotInheritable Class SnapshotRolloutAssistantPhaseEntry
            Public Property TurnId As String = String.Empty
            Public Property Phase As String = String.Empty
            Public Property TextKey As String = String.Empty
            Public Property LineNumber As Integer
            Public Property Consumed As Boolean
        End Class

        Private NotInheritable Class SnapshotRolloutAssistantPhaseIndex
            Public ReadOnly Property ByTurnId As New Dictionary(Of String, List(Of SnapshotRolloutAssistantPhaseEntry))(StringComparer.Ordinal)
            Public ReadOnly Property SearchStartByTurnId As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Public Property EntryCount As Integer
            Public Property SourcePath As String = String.Empty
        End Class

        Private NotInheritable Class SnapshotPhaseEnrichmentStats
            Public Property ServerPhaseCount As Integer
            Public Property LocalCacheMatchedCount As Integer
            Public Property RolloutMatchedCount As Integer
            Public Property UnmatchedCount As Integer
            Public Property MissingTextCount As Integer
        End Class

        Private Sub New()
        End Sub

        Public Shared Function BuildFromThread(threadObject As JsonObject,
                                               Optional assistantPhaseHintsByItemKey As IReadOnlyDictionary(Of String, String) = Nothing) As ThreadTranscriptSnapshot
            Dim snapshot As New ThreadTranscriptSnapshot()
            If threadObject Is Nothing Then
                Return snapshot
            End If

            Dim turns = ReadArray(threadObject, "turns")
            If turns Is Nothing OrElse turns.Count = 0 Then
                Return snapshot
            End If

            Dim runtimeStore As New TurnFlowRuntimeStore()
            Dim rolloutPhaseIndex = TryLoadRolloutAssistantPhaseIndex(threadObject, snapshot.DebugMessages)
            Dim phaseStats As New SnapshotPhaseEnrichmentStats()
            Dim replayOrder As New List(Of SnapshotTurnContext)()
            Dim replayedTurnKeys As New HashSet(Of String)(StringComparer.Ordinal)

            Dim fallbackThreadId = ReadStringFirst(threadObject, "id")
            For Each turnNode In turns
                Dim turnObject = TryCast(turnNode, JsonObject)
                If turnObject Is Nothing Then
                    Continue For
                End If

                Dim turnId = ReadStringFirst(turnObject, "id")
                If String.IsNullOrWhiteSpace(turnId) Then
                    Continue For
                End If

                Dim threadId = ReadStringFirst(turnObject, "threadId", "thread_id")
                If String.IsNullOrWhiteSpace(threadId) Then
                    threadId = fallbackThreadId
                End If

                If String.IsNullOrWhiteSpace(threadId) Then
                    Continue For
                End If

                Dim normalizedThreadId = threadId.Trim()
                Dim normalizedTurnId = turnId.Trim()
                Dim turnKey = BuildTurnKey(normalizedThreadId, normalizedTurnId)
                If replayedTurnKeys.Add(turnKey) Then
                    replayOrder.Add(New SnapshotTurnContext() With {
                        .ThreadId = normalizedThreadId,
                        .TurnId = normalizedTurnId
                    })
                End If

                ReplayNotification(runtimeStore,
                                   "turn/started",
                                   BuildTurnScopeParams(normalizedThreadId, normalizedTurnId))

                Dim diffSummary = ReadTurnSummary(turnObject, "diff")
                If Not String.IsNullOrWhiteSpace(diffSummary) Then
                    ReplayNotification(runtimeStore,
                                       "turn/diff/updated",
                                       BuildTurnDiffParams(normalizedThreadId,
                                                           normalizedTurnId,
                                                           diffSummary))
                End If

                Dim planSummary = ReadTurnSummary(turnObject, "plan")
                If Not String.IsNullOrWhiteSpace(planSummary) Then
                    ReplayNotification(runtimeStore,
                                       "turn/plan/updated",
                                       BuildTurnPlanParams(normalizedThreadId,
                                                           normalizedTurnId,
                                                           planSummary))
                End If

                Dim items = ReadArray(turnObject, "items")
                If items IsNot Nothing Then
                    For Each itemNode In items
                        Dim itemObject = TryCast(itemNode, JsonObject)
                        If itemObject Is Nothing Then
                            Continue For
                        End If

                        Dim itemId = ReadStringFirst(itemObject, "id")
                        If String.IsNullOrWhiteSpace(itemId) Then
                            Continue For
                        End If

                        Dim replayItemObject = TryCast(CloneNode(itemObject), JsonObject)
                        If replayItemObject Is Nothing Then
                            Continue For
                        End If

                        TryEnrichSnapshotAgentMessagePhase(replayItemObject,
                                                          normalizedThreadId,
                                                          normalizedTurnId,
                                                          assistantPhaseHintsByItemKey,
                                                          rolloutPhaseIndex,
                                                          snapshot.AssistantPhaseHintsByItemKey,
                                                          snapshot.DebugMessages,
                                                          phaseStats)

                        ReplayNotification(runtimeStore,
                                           "item/completed",
                                           BuildItemParams(normalizedThreadId,
                                                           normalizedTurnId,
                                                           replayItemObject))
                    Next
                End If

                Dim tokenUsage = ReadObject(turnObject, "tokenUsage")
                If tokenUsage IsNot Nothing Then
                    ReplayNotification(runtimeStore,
                                       "thread/tokenUsage/updated",
                                       BuildTokenUsageParams(normalizedThreadId,
                                                            normalizedTurnId,
                                                            tokenUsage))
                End If

                Dim status = ReadStringFirst(turnObject, "status")
                If String.IsNullOrWhiteSpace(status) Then
                    status = "completed"
                End If

                ReplayNotification(runtimeStore,
                                   "turn/completed",
                                   BuildTurnCompletedParams(normalizedThreadId,
                                                            normalizedTurnId,
                                                            status,
                                                            ReadTurnErrorMessage(turnObject)))
            Next

            AppendPhaseEnrichmentSummary(snapshot.DebugMessages, rolloutPhaseIndex, phaseStats)

            Dim accumulator As New SnapshotDescriptorAccumulator()
            For Each turnContext In replayOrder
                Dim turnKey = BuildTurnKey(turnContext.ThreadId, turnContext.TurnId)
                Dim turnState As TurnRuntimeState = Nothing
                If runtimeStore.TurnsById.ContainsKey(turnKey) Then
                    turnState = runtimeStore.TurnsById(turnKey)
                End If

                Dim itemOrder As List(Of String) = Nothing
                If runtimeStore.TurnItemOrder.ContainsKey(turnKey) Then
                    itemOrder = runtimeStore.TurnItemOrder(turnKey)
                End If

                Dim insertedTurnStartSection = False
                If itemOrder IsNot Nothing Then
                    For Each itemId In itemOrder
                        If String.IsNullOrWhiteSpace(itemId) Then
                            Continue For
                        End If

                        If Not runtimeStore.ItemsById.ContainsKey(itemId) Then
                            Continue For
                        End If

                        Dim descriptor = TranscriptPanelViewModel.BuildRuntimeItemDescriptorForSnapshot(runtimeStore.ItemsById(itemId))
                        If descriptor Is Nothing Then
                            Continue For
                        End If

                        Dim isUserDescriptor = StringComparer.OrdinalIgnoreCase.Equals(If(descriptor.Kind, String.Empty), "user")
                        If Not insertedTurnStartSection AndAlso Not isUserDescriptor Then
                            AppendTurnStartSection(accumulator, turnContext, turnState)
                            insertedTurnStartSection = True
                        End If

                        UpsertDescriptor(accumulator, $"item:{itemId}", descriptor)

                        If Not insertedTurnStartSection AndAlso isUserDescriptor Then
                            AppendTurnStartSection(accumulator, turnContext, turnState)
                            insertedTurnStartSection = True
                        End If
                    Next
                End If

                If Not insertedTurnStartSection Then
                    AppendTurnStartSection(accumulator, turnContext, turnState)
                End If

                Dim completionStatus = If(turnState Is Nothing, "completed", turnState.TurnStatus)
                UpsertDescriptor(accumulator,
                                 BuildTurnLifecycleRuntimeKey(turnContext.TurnId, "end"),
                                 TranscriptPanelViewModel.BuildTurnLifecycleDescriptorForSnapshot(turnContext.TurnId,
                                                                                                  completionStatus))

                If turnState IsNot Nothing AndAlso
                   StringComparer.OrdinalIgnoreCase.Equals(turnState.TurnStatus, "failed") AndAlso
                   Not String.IsNullOrWhiteSpace(turnState.LastErrorMessage) Then
                    UpsertDescriptor(accumulator,
                                     BuildTurnErrorRuntimeKey(turnContext.ThreadId, turnContext.TurnId),
                                     New TranscriptEntryDescriptor() With {
                        .Kind = "error",
                        .RoleText = "Error",
                        .BodyText = turnState.LastErrorMessage,
                        .IsError = True
                    })
                End If
            Next

            Dim rawBuilder As New StringBuilder()
            For Each runtimeKey In accumulator.Order
                If String.IsNullOrWhiteSpace(runtimeKey) Then
                    Continue For
                End If

                Dim descriptor As TranscriptEntryDescriptor = Nothing
                If Not accumulator.ByKey.TryGetValue(runtimeKey, descriptor) OrElse descriptor Is Nothing Then
                    Continue For
                End If

                snapshot.DisplayEntries.Add(descriptor)

                Dim body = If(descriptor.BodyText, String.Empty).Trim()
                If Not String.IsNullOrWhiteSpace(body) Then
                    AppendSnapshotLine(rawBuilder, BuildSnapshotRole(descriptor), body)
                End If

                Dim details = If(descriptor.DetailsText, String.Empty).Trim()
                If Not String.IsNullOrWhiteSpace(details) Then
                    AppendSnapshotLine(rawBuilder, BuildSnapshotRole(descriptor), details)
                End If
            Next

            snapshot.RawText = rawBuilder.ToString().TrimEnd()
            Return snapshot
        End Function

        Private Shared Sub AppendPhaseEnrichmentSummary(debugMessages As IList(Of String),
                                                        rolloutPhaseIndex As SnapshotRolloutAssistantPhaseIndex,
                                                        stats As SnapshotPhaseEnrichmentStats)
            If debugMessages Is Nothing OrElse stats Is Nothing Then
                Return
            End If

            Dim totalInspected = stats.ServerPhaseCount + stats.LocalCacheMatchedCount + stats.RolloutMatchedCount + stats.UnmatchedCount + stats.MissingTextCount
            If totalInspected <= 0 Then
                Return
            End If

            Dim rolloutEntryCount = If(rolloutPhaseIndex Is Nothing, 0, rolloutPhaseIndex.EntryCount)
            Dim rolloutSource = If(rolloutPhaseIndex Is Nothing, "none", "jsonl")
            debugMessages.Add(
                $"snapshot_phase_enrich_summary inspected={totalInspected} server_phase={stats.ServerPhaseCount} local_cache_matched={stats.LocalCacheMatchedCount} rollout_matched={stats.RolloutMatchedCount} unmatched={stats.UnmatchedCount} missing_text={stats.MissingTextCount} rollout_source={rolloutSource} rollout_entries={rolloutEntryCount}")
        End Sub

        Private Shared Sub TryEnrichSnapshotAgentMessagePhase(itemObject As JsonObject,
                                                              threadId As String,
                                                              turnId As String,
                                                              assistantPhaseHintsByItemKey As IReadOnlyDictionary(Of String, String),
                                                              rolloutPhaseIndex As SnapshotRolloutAssistantPhaseIndex,
                                                              learnedAssistantPhaseHintsByItemKey As IDictionary(Of String, String),
                                                              debugMessages As IList(Of String),
                                                              stats As SnapshotPhaseEnrichmentStats)
            If itemObject Is Nothing Then
                Return
            End If

            Dim itemType = NormalizeIdentifier(ReadStringFirst(itemObject, "type"))
            If Not StringComparer.OrdinalIgnoreCase.Equals(itemType, "agentmessage") Then
                Return
            End If

            Dim existingPhase = NormalizeIdentifier(ReadStringFirst(itemObject, "phase"))
            If Not String.IsNullOrWhiteSpace(existingPhase) Then
                RecordAssistantPhaseHint(learnedAssistantPhaseHintsByItemKey,
                                         threadId,
                                         turnId,
                                         ReadStringFirst(itemObject, "id"),
                                         existingPhase)
                If stats IsNot Nothing Then
                    stats.ServerPhaseCount += 1
                End If
                Return
            End If

            Dim itemId = ReadStringFirst(itemObject, "id")
            Dim itemKey = BuildAssistantPhaseItemKey(threadId, turnId, itemId)
            If Not String.IsNullOrWhiteSpace(itemKey) AndAlso assistantPhaseHintsByItemKey IsNot Nothing Then
                Dim cachedPhase As String = Nothing
                If assistantPhaseHintsByItemKey.TryGetValue(itemKey, cachedPhase) AndAlso
                   Not String.IsNullOrWhiteSpace(NormalizeIdentifier(cachedPhase)) Then
                    itemObject("phase") = NormalizeIdentifier(cachedPhase)
                    RecordAssistantPhaseHint(learnedAssistantPhaseHintsByItemKey,
                                             threadId,
                                             turnId,
                                             itemId,
                                             cachedPhase)
                    If stats IsNot Nothing Then
                        stats.LocalCacheMatchedCount += 1
                    End If
                    Return
                End If
            End If

            Dim itemText = ReadSnapshotAgentMessageText(itemObject)
            If String.IsNullOrWhiteSpace(itemText) Then
                If stats IsNot Nothing Then
                    stats.MissingTextCount += 1
                End If
                Return
            End If

            Dim matchedPhase As String = Nothing
            Dim matchedLineNumber = 0
            If TryMatchRolloutAssistantPhase(rolloutPhaseIndex, turnId, itemText, matchedPhase, matchedLineNumber) Then
                itemObject("phase") = matchedPhase
                RecordAssistantPhaseHint(learnedAssistantPhaseHintsByItemKey,
                                         threadId,
                                         turnId,
                                         itemId,
                                         matchedPhase)
                If stats IsNot Nothing Then
                    stats.RolloutMatchedCount += 1
                End If
                Return
            End If

            If stats IsNot Nothing Then
                stats.UnmatchedCount += 1
            End If

            If debugMessages IsNot Nothing Then
                Dim preview = BuildTextPreview(itemText, 96)
                Dim rolloutState = If(rolloutPhaseIndex Is Nothing, "unavailable", "loaded")
                debugMessages.Add(
                    $"snapshot_phase_enrich_unmatched threadId={NormalizeIdentifier(threadId)} turnId={NormalizeIdentifier(turnId)} itemId={NormalizeIdentifier(itemId)} rollout={rolloutState} local_cache={If(assistantPhaseHintsByItemKey Is Nothing, "unavailable", "loaded")} text_len={itemText.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)} preview={preview}")
            End If
        End Sub

        Private Shared Sub RecordAssistantPhaseHint(target As IDictionary(Of String, String),
                                                    threadId As String,
                                                    turnId As String,
                                                    itemId As String,
                                                    phase As String)
            If target Is Nothing Then
                Return
            End If

            Dim key = BuildAssistantPhaseItemKey(threadId, turnId, itemId)
            Dim normalizedPhase = NormalizeIdentifier(phase)
            If String.IsNullOrWhiteSpace(key) OrElse String.IsNullOrWhiteSpace(normalizedPhase) Then
                Return
            End If

            target(key) = normalizedPhase
        End Sub

        Private Shared Function BuildAssistantPhaseItemKey(threadId As String,
                                                           turnId As String,
                                                           itemId As String) As String
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            Dim normalizedTurnId = NormalizeIdentifier(turnId)
            Dim normalizedItemId = NormalizeIdentifier(itemId)

            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse
               String.IsNullOrWhiteSpace(normalizedTurnId) OrElse
               String.IsNullOrWhiteSpace(normalizedItemId) Then
                Return String.Empty
            End If

            Return $"{normalizedThreadId}:{normalizedTurnId}:{normalizedItemId}"
        End Function

        Private Shared Function TryLoadRolloutAssistantPhaseIndex(threadObject As JsonObject,
                                                                  debugMessages As IList(Of String)) As SnapshotRolloutAssistantPhaseIndex
            If threadObject Is Nothing Then
                Return Nothing
            End If

            Dim rawPath = ReadStringFirst(threadObject, "path")
            If String.IsNullOrWhiteSpace(rawPath) Then
                Return Nothing
            End If

            Dim resolvedPath = ResolveExistingRolloutPath(rawPath)
            If String.IsNullOrWhiteSpace(resolvedPath) Then
                Return Nothing
            End If

            Dim index As New SnapshotRolloutAssistantPhaseIndex() With {
                .SourcePath = resolvedPath
            }

            Try
                Dim currentTurnId = String.Empty
                Dim lineNumber = 0
                For Each line In File.ReadLines(resolvedPath)
                    lineNumber += 1
                    If String.IsNullOrWhiteSpace(line) Then
                        Continue For
                    End If

                    Dim lineObject As JsonObject = Nothing
                    Try
                        lineObject = TryCast(JsonNode.Parse(line), JsonObject)
                    Catch
                        Continue For
                    End Try

                    If lineObject Is Nothing Then
                        Continue For
                    End If

                    Dim recordType = NormalizeIdentifier(ReadStringFirst(lineObject, "type"))
                    Select Case recordType
                        Case "turn_context"
                            Dim payload = ReadObject(lineObject, "payload")
                            If payload IsNot Nothing Then
                                currentTurnId = NormalizeIdentifier(ReadStringFirst(payload, "turn_id", "turnId"))
                            End If

                        Case "response_item"
                            Dim payload = ReadObject(lineObject, "payload")
                            If payload Is Nothing Then
                                Continue For
                            End If

                            If Not StringComparer.OrdinalIgnoreCase.Equals(NormalizeIdentifier(ReadStringFirst(payload, "type")), "message") Then
                                Continue For
                            End If

                            If Not StringComparer.OrdinalIgnoreCase.Equals(NormalizeIdentifier(ReadStringFirst(payload, "role")), "assistant") Then
                                Continue For
                            End If

                            Dim phase = NormalizeIdentifier(ReadStringFirst(payload, "phase"))
                            If String.IsNullOrWhiteSpace(phase) Then
                                Continue For
                            End If

                            Dim turnId = NormalizeIdentifier(currentTurnId)
                            If String.IsNullOrWhiteSpace(turnId) Then
                                Continue For
                            End If

                            Dim messageText = ReadRolloutAssistantMessageText(payload)
                            Dim textKey = NormalizeTextForPhaseMatch(messageText)
                            If String.IsNullOrWhiteSpace(textKey) Then
                                Continue For
                            End If

                            Dim entry As New SnapshotRolloutAssistantPhaseEntry() With {
                                .TurnId = turnId,
                                .Phase = phase,
                                .TextKey = textKey,
                                .LineNumber = lineNumber
                            }

                            If Not index.ByTurnId.ContainsKey(turnId) Then
                                index.ByTurnId(turnId) = New List(Of SnapshotRolloutAssistantPhaseEntry)()
                            End If

                            index.ByTurnId(turnId).Add(entry)
                            index.EntryCount += 1
                    End Select
                Next
            Catch ex As Exception
                If debugMessages IsNot Nothing Then
                    debugMessages.Add(
                        $"snapshot_phase_enrich_rollout_read_failed path={BuildTextPreview(resolvedPath, 200)} error={BuildTextPreview(ex.Message, 160)}")
                End If
                Return Nothing
            End Try

            If index.EntryCount <= 0 Then
                Return Nothing
            End If

            Return index
        End Function

        Private Shared Function ResolveExistingRolloutPath(rawPath As String) As String
            Dim candidates As New List(Of String)()
            Dim normalizedRaw = If(rawPath, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedRaw) Then
                Return String.Empty
            End If

            candidates.Add(normalizedRaw)

            If normalizedRaw.StartsWith("\\?\UNC\", StringComparison.OrdinalIgnoreCase) Then
                candidates.Add("\\" & normalizedRaw.Substring("\\?\UNC\".Length))
            ElseIf normalizedRaw.StartsWith("\\?\", StringComparison.OrdinalIgnoreCase) Then
                candidates.Add(normalizedRaw.Substring("\\?\".Length))
            End If

            For Each candidate In candidates
                If String.IsNullOrWhiteSpace(candidate) Then
                    Continue For
                End If

                Try
                    If File.Exists(candidate) Then
                        Return candidate
                    End If
                Catch
                End Try
            Next

            Return String.Empty
        End Function

        Private Shared Function TryMatchRolloutAssistantPhase(rolloutPhaseIndex As SnapshotRolloutAssistantPhaseIndex,
                                                              turnId As String,
                                                              itemText As String,
                                                              ByRef matchedPhase As String,
                                                              ByRef matchedLineNumber As Integer) As Boolean
            matchedPhase = String.Empty
            matchedLineNumber = 0

            If rolloutPhaseIndex Is Nothing Then
                Return False
            End If

            Dim normalizedTurnId = NormalizeIdentifier(turnId)
            If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return False
            End If

            Dim entries As List(Of SnapshotRolloutAssistantPhaseEntry) = Nothing
            If Not rolloutPhaseIndex.ByTurnId.TryGetValue(normalizedTurnId, entries) OrElse entries Is Nothing OrElse entries.Count = 0 Then
                Return False
            End If

            Dim textKey = NormalizeTextForPhaseMatch(itemText)
            If String.IsNullOrWhiteSpace(textKey) Then
                Return False
            End If

            Dim searchStart = 0
            rolloutPhaseIndex.SearchStartByTurnId.TryGetValue(normalizedTurnId, searchStart)
            If searchStart < 0 Then
                searchStart = 0
            End If

            If TryMatchRolloutAssistantPhaseInRange(entries, textKey, searchStart, entries.Count - 1, matchedPhase, matchedLineNumber, searchStart) Then
                rolloutPhaseIndex.SearchStartByTurnId(normalizedTurnId) = searchStart
                Return True
            End If

            If searchStart > 0 AndAlso
               TryMatchRolloutAssistantPhaseInRange(entries, textKey, 0, searchStart - 1, matchedPhase, matchedLineNumber, searchStart) Then
                rolloutPhaseIndex.SearchStartByTurnId(normalizedTurnId) = searchStart
                Return True
            End If

            Return False
        End Function

        Private Shared Function TryMatchRolloutAssistantPhaseInRange(entries As List(Of SnapshotRolloutAssistantPhaseEntry),
                                                                     textKey As String,
                                                                     startIndex As Integer,
                                                                     endIndex As Integer,
                                                                     ByRef matchedPhase As String,
                                                                     ByRef matchedLineNumber As Integer,
                                                                     ByRef nextSearchStart As Integer) As Boolean
            If entries Is Nothing OrElse entries.Count = 0 Then
                Return False
            End If

            If startIndex < 0 Then
                startIndex = 0
            End If

            If endIndex >= entries.Count Then
                endIndex = entries.Count - 1
            End If

            If endIndex < startIndex Then
                Return False
            End If

            For i = startIndex To endIndex
                Dim entry = entries(i)
                If entry Is Nothing OrElse entry.Consumed Then
                    Continue For
                End If

                If Not StringComparer.Ordinal.Equals(entry.TextKey, textKey) Then
                    Continue For
                End If

                entry.Consumed = True
                matchedPhase = entry.Phase
                matchedLineNumber = entry.LineNumber
                nextSearchStart = i + 1
                Return True
            Next

            Return False
        End Function

        Private Shared Function ReadSnapshotAgentMessageText(itemObject As JsonObject) As String
            If itemObject Is Nothing Then
                Return String.Empty
            End If

            Dim directText = ReadStringFirst(itemObject, "text")
            If Not String.IsNullOrWhiteSpace(directText) Then
                Return directText
            End If

            Dim content = ReadArray(itemObject, "content")
            If content Is Nothing OrElse content.Count = 0 Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder()
            For Each partNode In content
                Dim partObject = TryCast(partNode, JsonObject)
                Dim partText = If(partObject Is Nothing,
                                  String.Empty,
                                  ReadStringFirst(partObject, "text", "value", "delta", "summary"))
                If String.IsNullOrWhiteSpace(partText) Then
                    Continue For
                End If

                If builder.Length > 0 Then
                    builder.AppendLine()
                End If

                builder.Append(partText)
            Next

            Return builder.ToString().Trim()
        End Function

        Private Shared Function ReadRolloutAssistantMessageText(payload As JsonObject) As String
            If payload Is Nothing Then
                Return String.Empty
            End If

            Dim directText = ReadStringFirst(payload, "text", "message")
            If Not String.IsNullOrWhiteSpace(directText) Then
                Return directText
            End If

            Dim content = ReadArray(payload, "content")
            If content Is Nothing OrElse content.Count = 0 Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder()
            For Each partNode In content
                Dim partObject = TryCast(partNode, JsonObject)
                If partObject Is Nothing Then
                    Continue For
                End If

                Dim partType = NormalizeIdentifier(ReadStringFirst(partObject, "type"))
                Dim partText As String
                Select Case partType
                    Case "output_text", "text", "input_text"
                        partText = ReadStringFirst(partObject, "text", "value")
                    Case Else
                        partText = ReadStringFirst(partObject, "text", "value")
                End Select

                If String.IsNullOrWhiteSpace(partText) Then
                    Continue For
                End If

                If builder.Length > 0 Then
                    builder.AppendLine()
                End If

                builder.Append(partText)
            Next

            Return builder.ToString().Trim()
        End Function

        Private Shared Function NormalizeTextForPhaseMatch(value As String) As String
            Dim normalized = If(value, String.Empty)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            normalized = normalized.Replace(ControlChars.CrLf, ControlChars.Lf).
                                    Replace(ControlChars.Cr, ControlChars.Lf).
                                    Trim()
            Return normalized
        End Function

        Private Shared Function BuildTextPreview(value As String, maxLength As Integer) As String
            Dim normalized = NormalizeTextForPhaseMatch(value)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return "(empty)"
            End If

            normalized = normalized.Replace(ControlChars.Lf, " "c)
            If maxLength > 0 AndAlso normalized.Length > maxLength Then
                Return normalized.Substring(0, maxLength) & "..."
            End If

            Return normalized
        End Function

        Private Shared Function BuildSnapshotRole(descriptor As TranscriptEntryDescriptor) As String
            If descriptor Is Nothing Then
                Return "item"
            End If

            Select Case If(descriptor.Kind, String.Empty).Trim().ToLowerInvariant()
                Case "user"
                    Return "user"
                Case "assistant", "assistantcommentary", "assistantfinal"
                    Return "assistant"
                Case "reasoning", "reasoningcard"
                    Return "reasoning"
                Case "plan"
                    Return "plan"
                Case "command"
                    Return "command"
                Case "filechange"
                    Return "fileChange"
                Case "toolmcp", "toolcollab", "toolwebsearch", "toolimageview"
                    Return "tool"
                Case "turnmarker", "turndiff", "turnplan", "systemmarker", "system"
                    Return "system"
                Case "error"
                    Return "error"
                Case Else
                    Return "item"
            End Select
        End Function

        Private Shared Sub AppendSnapshotLine(builder As StringBuilder, role As String, text As String)
            If builder Is Nothing OrElse String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            builder.Append(role)
            builder.Append(": ")
            builder.Append(text)
            builder.AppendLine()
            builder.AppendLine()
        End Sub

        Private Shared Function BuildTurnScopeParams(threadId As String, turnId As String) As JsonObject
            Dim paramsObject As New JsonObject()
            paramsObject("threadId") = threadId
            paramsObject("turnId") = turnId
            Return paramsObject
        End Function

        Private Shared Function BuildTurnDiffParams(threadId As String,
                                                    turnId As String,
                                                    diffSummary As String) As JsonObject
            Dim paramsObject = BuildTurnScopeParams(threadId, turnId)
            paramsObject("diff") = If(diffSummary, String.Empty)
            Return paramsObject
        End Function

        Private Shared Function BuildTurnPlanParams(threadId As String,
                                                    turnId As String,
                                                    planSummary As String) As JsonObject
            Dim paramsObject = BuildTurnScopeParams(threadId, turnId)
            paramsObject("summary") = If(planSummary, String.Empty)
            Return paramsObject
        End Function

        Private Shared Function BuildItemParams(threadId As String,
                                                turnId As String,
                                                itemObject As JsonObject) As JsonObject
            Dim paramsObject As New JsonObject()
            paramsObject("threadId") = threadId
            paramsObject("turnId") = turnId
            paramsObject("item") = CloneNode(itemObject)
            Return paramsObject
        End Function

        Private Shared Function BuildTokenUsageParams(threadId As String,
                                                      turnId As String,
                                                      tokenUsage As JsonObject) As JsonObject
            Dim paramsObject = BuildTurnScopeParams(threadId, turnId)
            paramsObject("tokenUsage") = CloneNode(tokenUsage)
            Return paramsObject
        End Function

        Private Shared Function BuildTurnCompletedParams(threadId As String,
                                                         turnId As String,
                                                         status As String,
                                                         errorMessage As String) As JsonObject
            Dim paramsObject = BuildTurnScopeParams(threadId, turnId)
            paramsObject("status") = If(status, String.Empty)
            If Not String.IsNullOrWhiteSpace(errorMessage) Then
                Dim errorObject As New JsonObject()
                errorObject("message") = errorMessage
                paramsObject("error") = errorObject
            End If

            Return paramsObject
        End Function

        Private Shared Function CloneNode(node As JsonNode) As JsonNode
            If node Is Nothing Then
                Return Nothing
            End If

            Return JsonNode.Parse(node.ToJsonString())
        End Function

        Private Shared Function ReadTurnSummary(turnObject As JsonObject, propertyName As String) As String
            If turnObject Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return String.Empty
            End If

            Dim rawText = ReadStringFirst(turnObject, propertyName)
            If Not String.IsNullOrWhiteSpace(rawText) Then
                Return rawText
            End If

            Dim summaryObject = ReadObject(turnObject, propertyName)
            If summaryObject Is Nothing Then
                Return String.Empty
            End If

            rawText = ReadStringFirst(summaryObject, "summary", "text")
            If Not String.IsNullOrWhiteSpace(rawText) Then
                Return rawText
            End If

            Return summaryObject.ToJsonString()
        End Function

        Private Shared Function ReadTurnErrorMessage(turnObject As JsonObject) As String
            If turnObject Is Nothing Then
                Return String.Empty
            End If

            Dim errorObject = ReadObject(turnObject, "error")
            If errorObject IsNot Nothing Then
                Dim message = ReadStringFirst(errorObject, "message")
                If Not String.IsNullOrWhiteSpace(message) Then
                    Return message
                End If
            End If

            Return ReadStringFirst(turnObject, "errorMessage")
        End Function

        Private Shared Sub ReplayNotification(runtimeStore As TurnFlowRuntimeStore,
                                              methodName As String,
                                              paramsObject As JsonObject)
            If runtimeStore Is Nothing OrElse String.IsNullOrWhiteSpace(methodName) Then
                Return
            End If

            Dim parsedEvent = TurnFlowEventParser.ParseNotification(methodName, paramsObject)
            runtimeStore.Reduce(parsedEvent)
        End Sub

        Private Shared Sub UpsertDescriptor(accumulator As SnapshotDescriptorAccumulator,
                                            runtimeKey As String,
                                            descriptor As TranscriptEntryDescriptor)
            If accumulator Is Nothing OrElse descriptor Is Nothing Then
                Return
            End If

            Dim normalizedKey = If(runtimeKey, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedKey) Then
                Return
            End If

            If Not accumulator.ByKey.ContainsKey(normalizedKey) Then
                accumulator.Order.Add(normalizedKey)
            End If

            accumulator.ByKey(normalizedKey) = descriptor
        End Sub

        Private Shared Sub AppendTurnStartSection(accumulator As SnapshotDescriptorAccumulator,
                                                  turnContext As SnapshotTurnContext,
                                                  turnState As TurnRuntimeState)
            If accumulator Is Nothing OrElse turnContext Is Nothing Then
                Return
            End If

            UpsertDescriptor(accumulator,
                             BuildTurnLifecycleRuntimeKey(turnContext.TurnId, "start"),
                             TranscriptPanelViewModel.BuildTurnLifecycleDescriptorForSnapshot(turnContext.TurnId,
                                                                                              "started"))

            If turnState Is Nothing Then
                Return
            End If

            If Not String.IsNullOrWhiteSpace(turnState.DiffSummary) Then
                UpsertDescriptor(accumulator,
                                 BuildTurnMetadataRuntimeKey(turnContext.ThreadId, turnContext.TurnId, "diff"),
                                 TranscriptPanelViewModel.BuildTurnMetadataDescriptorForSnapshot("diff",
                                                                                                 turnState.DiffSummary))
            End If

            If Not String.IsNullOrWhiteSpace(turnState.PlanSummary) Then
                UpsertDescriptor(accumulator,
                                 BuildTurnMetadataRuntimeKey(turnContext.ThreadId, turnContext.TurnId, "plan"),
                                 TranscriptPanelViewModel.BuildTurnMetadataDescriptorForSnapshot("plan",
                                                                                                 turnState.PlanSummary))
            End If
        End Sub

        Private Shared Function BuildTurnLifecycleRuntimeKey(turnId As String, slot As String) As String
            Return $"turn:lifecycle:{NormalizeIdentifier(slot)}:{NormalizeIdentifier(turnId)}"
        End Function

        Private Shared Function BuildTurnMetadataRuntimeKey(threadId As String,
                                                            turnId As String,
                                                            kind As String) As String
            Return $"turn:meta:{NormalizeIdentifier(kind)}:{NormalizeIdentifier(threadId)}:{NormalizeIdentifier(turnId)}"
        End Function

        Private Shared Function BuildTurnErrorRuntimeKey(threadId As String, turnId As String) As String
            Return $"turn:error:{NormalizeIdentifier(threadId)}:{NormalizeIdentifier(turnId)}"
        End Function

        Private Shared Function BuildTurnKey(threadId As String, turnId As String) As String
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse
               String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return String.Empty
            End If

            Return $"{normalizedThreadId}:{normalizedTurnId}"
        End Function

        Private Shared Function ReadArray(obj As JsonObject, key As String) As JsonArray
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonArray)
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
                If Not obj.TryGetPropertyValue(key, node) OrElse node Is Nothing Then
                    Continue For
                End If

                Dim value = TryCast(node, JsonValue)
                If value Is Nothing Then
                    Continue For
                End If

                Dim stringValue As String = Nothing
                If value.TryGetValue(Of String)(stringValue) AndAlso
                   Not String.IsNullOrWhiteSpace(stringValue) Then
                    Return stringValue.Trim()
                End If
            Next

            Return String.Empty
        End Function
    End Class
End Namespace
