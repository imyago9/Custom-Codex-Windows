Imports System.Collections.Generic
Imports CodexNativeAgent.Ui.ViewModels.Transcript

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class ThreadTranscriptProjectionSnapshot
        Public Property RawText As String = String.Empty
        Public ReadOnly Property DisplayEntries As New List(Of TranscriptEntryDescriptor)()
    End Class

    Public NotInheritable Class ThreadLiveSessionState
        Public Property ThreadId As String = String.Empty
        Public Property VisibleTurnId As String = String.Empty
        Public Property HasLoadedHistoricalSnapshot As Boolean
        Public Property SnapshotRawText As String = String.Empty
        Public ReadOnly Property SnapshotDisplayEntries As New List(Of TranscriptEntryDescriptor)()
        Public ReadOnly Property OverlayTurnIds As New HashSet(Of String)(StringComparer.Ordinal)
        Public Property LastBoundUtc As DateTimeOffset
        Public Property LastRuntimeEventUtc As DateTimeOffset
        Public Property IsTurnActive As Boolean
        Public Property ActiveTurnId As String = String.Empty
        Public Property PendingRebuild As Boolean
    End Class

    Public NotInheritable Class ThreadLiveSessionRegistry
        Private ReadOnly _statesByThreadId As New Dictionary(Of String, ThreadLiveSessionState)(StringComparer.Ordinal)

        Public ReadOnly Property StatesByThreadId As IReadOnlyDictionary(Of String, ThreadLiveSessionState)
            Get
                Return _statesByThreadId
            End Get
        End Property

        Public Sub Clear()
            _statesByThreadId.Clear()
        End Sub

        Public Function ContainsThread(threadId As String) As Boolean
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Return _statesByThreadId.ContainsKey(normalizedThreadId)
        End Function

        Public Function TryGet(threadId As String, ByRef state As ThreadLiveSessionState) As Boolean
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                state = Nothing
                Return False
            End If

            Return _statesByThreadId.TryGetValue(normalizedThreadId, state)
        End Function

        Public Function GetOrCreate(threadId As String) As ThreadLiveSessionState
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Throw New ArgumentException("Thread id is required.", NameOf(threadId))
            End If

            Dim state As ThreadLiveSessionState = Nothing
            If _statesByThreadId.TryGetValue(normalizedThreadId, state) AndAlso state IsNot Nothing Then
                Return state
            End If

            state = New ThreadLiveSessionState() With {
                .ThreadId = normalizedThreadId
            }
            _statesByThreadId(normalizedThreadId) = state
            Return state
        End Function

        Public Function UpsertSnapshot(threadId As String,
                                       snapshotRawText As String,
                                       snapshotDisplayEntries As IEnumerable(Of TranscriptEntryDescriptor),
                                       Optional visibleTurnId As String = Nothing) As ThreadLiveSessionState
            Dim state = GetOrCreate(threadId)

            state.HasLoadedHistoricalSnapshot = True
            state.SnapshotRawText = If(snapshotRawText, String.Empty)
            state.SnapshotDisplayEntries.Clear()
            For Each descriptor In CloneDescriptors(snapshotDisplayEntries)
                state.SnapshotDisplayEntries.Add(descriptor)
            Next

            ' Keep overlay turn tracking across snapshot refreshes so runtime-only bubbles
            ' (command/file changes/tool activity) remain replayable for turns observed live
            ' during this app session, even after the turn later completes.
            state.PendingRebuild = False

            Dim normalizedVisibleTurnId = NormalizeIdentifier(visibleTurnId)
            If Not String.IsNullOrWhiteSpace(normalizedVisibleTurnId) Then
                state.VisibleTurnId = normalizedVisibleTurnId
            End If

            Return state
        End Function

        Public Function GetProjectionSnapshot(threadId As String) As ThreadTranscriptProjectionSnapshot
            Dim state As ThreadLiveSessionState = Nothing
            If Not TryGet(threadId, state) OrElse state Is Nothing Then
                Return New ThreadTranscriptProjectionSnapshot()
            End If

            Dim snapshot As New ThreadTranscriptProjectionSnapshot() With {
                .RawText = If(state.SnapshotRawText, String.Empty)
            }

            For Each descriptor In CloneDescriptors(state.SnapshotDisplayEntries)
                snapshot.DisplayEntries.Add(descriptor)
            Next

            Return snapshot
        End Function

        Public Sub MarkBound(threadId As String, Optional visibleTurnId As String = Nothing)
            Dim state = GetOrCreate(threadId)
            state.LastBoundUtc = DateTimeOffset.UtcNow

            Dim normalizedVisibleTurnId = NormalizeIdentifier(visibleTurnId)
            If Not String.IsNullOrWhiteSpace(normalizedVisibleTurnId) Then
                state.VisibleTurnId = normalizedVisibleTurnId
            End If
        End Sub

        Public Sub MarkRuntimeActivity(threadId As String,
                                       Optional activeTurnId As String = Nothing,
                                       Optional isTurnActive As Boolean? = Nothing)
            Dim state = GetOrCreate(threadId)
            state.LastRuntimeEventUtc = DateTimeOffset.UtcNow

            Dim normalizedActiveTurnId = NormalizeIdentifier(activeTurnId)
            If Not String.IsNullOrWhiteSpace(normalizedActiveTurnId) Then
                state.ActiveTurnId = normalizedActiveTurnId
            End If

            If isTurnActive.HasValue Then
                state.IsTurnActive = isTurnActive.Value
            End If
        End Sub

        Public Sub MarkOverlayTurn(threadId As String, turnId As String)
            Dim normalizedTurnId = NormalizeIdentifier(turnId)
            If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                Return
            End If

            Dim state = GetOrCreate(threadId)
            state.OverlayTurnIds.Add(normalizedTurnId)
        End Sub

        Public Function GetOverlayTurnIds(threadId As String) As IReadOnlyCollection(Of String)
            Dim state As ThreadLiveSessionState = Nothing
            If Not TryGet(threadId, state) OrElse state Is Nothing Then
                Return New List(Of String)()
            End If

            Dim result As New List(Of String)(state.OverlayTurnIds.Count)
            For Each turnId In state.OverlayTurnIds
                Dim normalizedTurnId = NormalizeIdentifier(turnId)
                If String.IsNullOrWhiteSpace(normalizedTurnId) Then
                    Continue For
                End If

                result.Add(normalizedTurnId)
            Next

            Return result
        End Function

        Public Sub SetPendingRebuild(threadId As String, pendingRebuild As Boolean)
            Dim state = GetOrCreate(threadId)
            state.PendingRebuild = pendingRebuild
        End Sub

        Public Function Remove(threadId As String) As Boolean
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Return _statesByThreadId.Remove(normalizedThreadId)
        End Function

        Private Shared Function NormalizeIdentifier(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function

        Private Shared Function CloneDescriptors(entries As IEnumerable(Of TranscriptEntryDescriptor)) As List(Of TranscriptEntryDescriptor)
            Dim result As New List(Of TranscriptEntryDescriptor)()
            If entries Is Nothing Then
                Return result
            End If

            For Each entry In entries
                If entry Is Nothing Then
                    Continue For
                End If

                result.Add(CloneDescriptor(entry))
            Next

            Return result
        End Function

        Private Shared Function CloneDescriptor(source As TranscriptEntryDescriptor) As TranscriptEntryDescriptor
            Dim clone As New TranscriptEntryDescriptor() With {
                .Kind = If(source.Kind, String.Empty),
                .RuntimeKey = If(source.RuntimeKey, String.Empty),
                .ThreadId = If(source.ThreadId, String.Empty),
                .TurnId = If(source.TurnId, String.Empty),
                .TurnItemStreamSequence = source.TurnItemStreamSequence,
                .TurnItemOrderIndex = source.TurnItemOrderIndex,
                .TurnItemSortTimestampUtc = source.TurnItemSortTimestampUtc,
                .TimestampText = If(source.TimestampText, String.Empty),
                .RoleText = If(source.RoleText, String.Empty),
                .BodyText = If(source.BodyText, String.Empty),
                .StatusText = If(source.StatusText, String.Empty),
                .SecondaryText = If(source.SecondaryText, String.Empty),
                .DetailsText = If(source.DetailsText, String.Empty),
                .AddedLineCount = source.AddedLineCount,
                .RemovedLineCount = source.RemovedLineCount,
                .IsMuted = source.IsMuted,
                .IsMonospaceBody = source.IsMonospaceBody,
                .IsCommandLike = source.IsCommandLike,
                .IsReasoning = source.IsReasoning,
                .IsError = source.IsError,
                .IsStreaming = source.IsStreaming,
                .UseRawReasoningLayout = source.UseRawReasoningLayout
            }

            If source.FileChangeItems IsNot Nothing Then
                clone.FileChangeItems = New List(Of TranscriptFileChangeListItemViewModel)()
                For Each item In source.FileChangeItems
                    If item Is Nothing Then
                        Continue For
                    End If

                    clone.FileChangeItems.Add(
                        New TranscriptFileChangeListItemViewModel() With {
                            .FullPathText = If(item.FullPathText, String.Empty),
                            .DisplayPathPrefixText = If(item.DisplayPathPrefixText, String.Empty),
                            .DisplayPathFileNameText = If(item.DisplayPathFileNameText, String.Empty),
                            .OverflowText = If(item.OverflowText, String.Empty),
                            .IsOverflow = item.IsOverflow,
                            .AddedLinesText = If(item.AddedLinesText, String.Empty),
                            .RemovedLinesText = If(item.RemovedLinesText, String.Empty),
                            .FileIconSource = item.FileIconSource
                        })
                Next
            End If

            Return clone
        End Function
    End Class
End Namespace
