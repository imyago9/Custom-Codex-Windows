Imports System
Imports System.Collections.Generic
Imports System.Threading.Tasks
Imports System.Windows.Controls
Imports System.Windows.Threading
Imports CodexNativeAgent.Ui.Coordinators
Imports CodexNativeAgent.Ui.ViewModels.Transcript

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Const TranscriptChunkTopProbeIntervalMs As Integer = 120
        Private Const TranscriptChunkTopTriggerMinPixels As Double = 64.0R
        Private Const TranscriptChunkTopTriggerViewportRatio As Double = 0.18R
        Private Const TranscriptChunkPageMaxRows As Integer = 140
        Private Const TranscriptChunkPageMaxRenderWeight As Integer = 280

        Private _transcriptChunkTopProbeTimer As DispatcherTimer
        Private _transcriptChunkTopProbeTickActive As Boolean

        Private NotInheritable Class TranscriptChunkPrependAnchorSnapshot
            Public Property VerticalOffset As Double
            Public Property ExtentHeight As Double
            Public Property ViewportHeight As Double
        End Class

        Private NotInheritable Class TranscriptOlderChunkLoadSnapshot
            Public Property ThreadId As String = String.Empty
            Public Property GenerationId As Integer
            Public Property Projection As ThreadTranscriptProjectionSnapshot
            Public Property CompletedOverlayTurnsForFullReplay As HashSet(Of String)
            Public Property CurrentLoadedRangeStart As Integer
            Public Property CurrentLoadedRangeEnd As Integer
            Public Property Anchor As TranscriptChunkPrependAnchorSnapshot
        End Class

        Private Sub EnsureTranscriptChunkTopProbeTimerStarted()
            If _transcriptChunkTopProbeTimer IsNot Nothing Then
                If Not _transcriptChunkTopProbeTimer.IsEnabled Then
                    _transcriptChunkTopProbeTimer.Start()
                End If

                Return
            End If

            _transcriptChunkTopProbeTimer = New DispatcherTimer(DispatcherPriority.Background, Dispatcher) With {
                .Interval = TimeSpan.FromMilliseconds(TranscriptChunkTopProbeIntervalMs)
            }
            AddHandler _transcriptChunkTopProbeTimer.Tick, AddressOf OnTranscriptChunkTopProbeTimerTick
            _transcriptChunkTopProbeTimer.Start()
            DebugTranscriptScroll("chunk_top_probe", $"timer_started intervalMs={TranscriptChunkTopProbeIntervalMs}")
        End Sub

        Private Async Sub OnTranscriptChunkTopProbeTimerTick(sender As Object, e As EventArgs)
            If _transcriptChunkTopProbeTickActive Then
                Return
            End If

            _transcriptChunkTopProbeTickActive = True
            Try
                Dim scroller = ResolveTranscriptScrollViewer()
                If Not ShouldRequestOlderTranscriptChunkFromScrollProbe(scroller) Then
                    Return
                End If

                Await TryLoadOlderTranscriptChunkForVisibleThreadAsync(scroller).ConfigureAwait(True)
            Catch ex As Exception
                AppendProtocol("debug", $"transcript_chunk_probe_error message={ex.Message}")
            Finally
                _transcriptChunkTopProbeTickActive = False
            End Try
        End Sub

        Private Function ShouldRequestOlderTranscriptChunkFromScrollProbe(scroller As ScrollViewer) As Boolean
            If scroller Is Nothing Then
                Return False
            End If

            If _threadContentLoading Then
                Return False
            End If

            If _transcriptScrollThumbDragActive OrElse _transcriptScrollProgrammaticMoveInProgress Then
                Return False
            End If

            If _transcriptScrollFollowMode <> TranscriptScrollFollowMode.DetachedByUser Then
                Return False
            End If

            Dim visibleThreadId = GetVisibleThreadId()
            If String.IsNullOrWhiteSpace(visibleThreadId) Then
                Return False
            End If

            Dim activeSession = _threadTranscriptChunkSessionCoordinator.ActiveSession
            If activeSession Is Nothing OrElse
               Not StringComparer.Ordinal.Equals(If(activeSession.ThreadId, String.Empty).Trim(), visibleThreadId) Then
                Return False
            End If

            If activeSession.IsLoadingOlderChunk OrElse Not activeSession.HasMoreOlderChunks Then
                Return False
            End If

            If Not activeSession.LoadedRangeStart.HasValue OrElse activeSession.LoadedRangeStart.Value <= 0 Then
                Return False
            End If

            If scroller.ScrollableHeight <= 0.01R Then
                Return False
            End If

            Return IsTranscriptScrollNearTopThreshold(scroller)
        End Function

        Private Shared Function IsTranscriptScrollNearTopThreshold(scroller As ScrollViewer) As Boolean
            If scroller Is Nothing Then
                Return False
            End If

            Dim viewportThreshold = Math.Max(TranscriptChunkTopTriggerMinPixels,
                                             scroller.ViewportHeight * TranscriptChunkTopTriggerViewportRatio)
            Return scroller.VerticalOffset <= Math.Max(0R, viewportThreshold)
        End Function

        Private Async Function TryLoadOlderTranscriptChunkForVisibleThreadAsync(scroller As ScrollViewer) As Task
            Dim activeSession = _threadTranscriptChunkSessionCoordinator.ActiveSession
            If activeSession Is Nothing Then
                Return
            End If

            Dim threadId = If(activeSession.ThreadId, String.Empty).Trim()
            Dim generationId = activeSession.GenerationId
            If String.IsNullOrWhiteSpace(threadId) OrElse generationId <= 0 Then
                Return
            End If

            If Not _threadTranscriptChunkSessionCoordinator.TryBeginOlderChunkLoad(threadId, generationId, "top_probe_near_top") Then
                Return
            End If

            TraceTranscriptChunkSession("older_chunk_load_begin",
                                        $"thread={threadId}; generation={generationId}")

            Dim loadSnapshot As TranscriptOlderChunkLoadSnapshot = Nothing
            Try
                loadSnapshot = CaptureOlderChunkLoadSnapshot(threadId, generationId, scroller)
                If loadSnapshot Is Nothing Then
                    _threadTranscriptChunkSessionCoordinator.TryCancelOlderChunkLoad(threadId, generationId, "capture_snapshot_failed")
                    TraceTranscriptChunkSession("older_chunk_load_cancel", $"thread={threadId}; generation={generationId}; stage=capture")
                    Return
                End If

                Dim chunkPlan = Await Task.Run(
                    Function()
                        Return ThreadTranscriptChunkPlanner.PrependPreviousDisplayChunk(loadSnapshot.Projection.DisplayEntries,
                                                                                        loadSnapshot.CurrentLoadedRangeStart,
                                                                                        loadSnapshot.CurrentLoadedRangeEnd,
                                                                                        TranscriptChunkPageMaxRows,
                                                                                        TranscriptChunkPageMaxRenderWeight)
                    End Function).ConfigureAwait(True)

                If chunkPlan Is Nothing Then
                    _threadTranscriptChunkSessionCoordinator.TryCancelOlderChunkLoad(threadId, generationId, "chunk_plan_null")
                    Return
                End If

                If Not _threadTranscriptChunkSessionCoordinator.IsActiveSessionGeneration(threadId, generationId) OrElse
                   Not StringComparer.Ordinal.Equals(GetVisibleThreadId(), threadId) Then
                    _threadTranscriptChunkSessionCoordinator.TryCancelOlderChunkLoad(threadId, generationId, "stale_after_plan")
                    TraceTranscriptChunkSession("older_chunk_load_cancel", $"thread={threadId}; generation={generationId}; stage=stale_after_plan")
                    Return
                End If

                Dim didExpand = chunkPlan.SelectedEntryCount > 0 AndAlso
                                chunkPlan.LoadedRangeStart < loadSnapshot.CurrentLoadedRangeStart
                If Not didExpand Then
                    _threadTranscriptChunkSessionCoordinator.TryCompleteOlderChunkLoad(threadId,
                                                                                      generationId,
                                                                                      chunkPlan.HasMoreOlderEntries,
                                                                                      chunkPlan.LoadedRangeStart,
                                                                                      chunkPlan.LoadedRangeEnd,
                                                                                      "older_chunk_noop_complete")
                    TraceTranscriptChunkSession("older_chunk_load_noop",
                                                $"thread={threadId}; generation={generationId}; rangeStart={chunkPlan.LoadedRangeStart}; rangeEnd={chunkPlan.LoadedRangeEnd}; hasMoreOlder={chunkPlan.HasMoreOlderEntries}")
                    Return
                End If

                ApplyOlderTranscriptChunkRender(threadId, generationId, loadSnapshot, chunkPlan)
            Catch ex As Exception
                _threadTranscriptChunkSessionCoordinator.TryCancelOlderChunkLoad(threadId, generationId, "older_chunk_exception")
                AppendProtocol("debug", $"transcript_chunk_load_older_error thread={threadId} gen={generationId} message={ex.Message}")
            End Try
        End Function

        Private Function CaptureOlderChunkLoadSnapshot(threadId As String,
                                                       generationId As Integer,
                                                       scroller As ScrollViewer) As TranscriptOlderChunkLoadSnapshot
            If String.IsNullOrWhiteSpace(threadId) OrElse generationId <= 0 OrElse scroller Is Nothing Then
                Return Nothing
            End If

            If Not _threadTranscriptChunkSessionCoordinator.IsActiveSessionGeneration(threadId, generationId) Then
                Return Nothing
            End If

            Dim projection = _threadLiveSessionRegistry.GetProjectionSnapshot(threadId)
            ApplyOverlayRuntimeOrderMetadataToProjection(threadId, projection, _sessionNotificationCoordinator.RuntimeStore)
            Dim completedOverlayTurnsForFullReplay = RemoveProjectionSnapshotItemRowsForCompletedOverlayTurns(threadId,
                                                                                                              projection,
                                                                                                              _sessionNotificationCoordinator.RuntimeStore)
            RemoveProjectionOverlayReplaySupersededMarkers(threadId,
                                                           projection,
                                                           _sessionNotificationCoordinator.RuntimeStore)

            Dim activeSession = _threadTranscriptChunkSessionCoordinator.ActiveSession
            If activeSession Is Nothing Then
                Return Nothing
            End If

            Dim loadedRangeStart = If(activeSession.LoadedRangeStart.HasValue, activeSession.LoadedRangeStart.Value, 0)
            Dim loadedRangeEnd = If(activeSession.LoadedRangeEnd.HasValue,
                                    activeSession.LoadedRangeEnd.Value,
                                    Math.Max(0, projection.DisplayEntries.Count - 1))

            Dim anchor As New TranscriptChunkPrependAnchorSnapshot() With {
                .VerticalOffset = Math.Max(0R, scroller.VerticalOffset),
                .ExtentHeight = Math.Max(0R, scroller.ExtentHeight),
                .ViewportHeight = Math.Max(0R, scroller.ViewportHeight)
            }

            Return New TranscriptOlderChunkLoadSnapshot() With {
                .ThreadId = threadId,
                .GenerationId = generationId,
                .Projection = projection,
                .CompletedOverlayTurnsForFullReplay = completedOverlayTurnsForFullReplay,
                .CurrentLoadedRangeStart = Math.Max(0, loadedRangeStart),
                .CurrentLoadedRangeEnd = Math.Max(Math.Max(0, loadedRangeStart), loadedRangeEnd),
                .Anchor = anchor
            }
        End Function

        Private Sub ApplyOlderTranscriptChunkRender(threadId As String,
                                                    generationId As Integer,
                                                    loadSnapshot As TranscriptOlderChunkLoadSnapshot,
                                                    chunkPlan As ThreadTranscriptDisplayChunkPlan)
            If loadSnapshot Is Nothing OrElse chunkPlan Is Nothing Then
                Return
            End If

            If Not _threadTranscriptChunkSessionCoordinator.IsActiveSessionGeneration(threadId, generationId) OrElse
               Not StringComparer.Ordinal.Equals(GetVisibleThreadId(), threadId) Then
                _threadTranscriptChunkSessionCoordinator.TryCancelOlderChunkLoad(threadId, generationId, "stale_before_apply")
                TraceTranscriptChunkSession("older_chunk_load_cancel", $"thread={threadId}; generation={generationId}; stage=stale_before_apply")
                Return
            End If

            ClearPendingUserEchoTracking()
            _viewModel.TranscriptPanel.ClearTranscript()
            _viewModel.TranscriptPanel.SetTranscriptSnapshot(loadSnapshot.Projection.RawText)
            _viewModel.TranscriptPanel.SetTranscriptDisplaySnapshot(chunkPlan.DisplayEntries)

            ApplyLiveRuntimeOverlayForThread(threadId,
                                            _sessionNotificationCoordinator.RuntimeStore,
                                            loadSnapshot.CompletedOverlayTurnsForFullReplay)

            _threadLiveSessionRegistry.MarkBound(threadId, GetVisibleTurnId())
            _threadLiveSessionRegistry.SetPendingRebuild(threadId, False)
            RefreshThreadRuntimeIndicatorsIfNeeded()

            _threadTranscriptChunkSessionCoordinator.TryCompleteOlderChunkLoad(threadId,
                                                                              generationId,
                                                                              chunkPlan.HasMoreOlderEntries,
                                                                              chunkPlan.LoadedRangeStart,
                                                                              chunkPlan.LoadedRangeEnd,
                                                                              "older_chunk_apply_complete")

            QueueRestoreTranscriptOffsetAfterChunkPrepend(threadId,
                                                          generationId,
                                                          loadSnapshot.Anchor)

            TraceTranscriptChunkSession("older_chunk_load_complete",
                                        $"thread={threadId}; generation={generationId}; total={chunkPlan.TotalEntryCount}; selected={chunkPlan.SelectedEntryCount}; rangeStart={chunkPlan.LoadedRangeStart}; rangeEnd={chunkPlan.LoadedRangeEnd}; hasMoreOlder={chunkPlan.HasMoreOlderEntries}; weight={chunkPlan.TotalSelectedRenderWeight}")
        End Sub

        Private Sub QueueRestoreTranscriptOffsetAfterChunkPrepend(threadId As String,
                                                                  generationId As Integer,
                                                                  anchor As TranscriptChunkPrependAnchorSnapshot)
            If anchor Is Nothing Then
                Return
            End If

            Dispatcher.BeginInvoke(
                DispatcherPriority.Loaded,
                New Action(
                    Sub()
                        If Not _threadTranscriptChunkSessionCoordinator.IsActiveSessionGeneration(threadId, generationId) OrElse
                           Not StringComparer.Ordinal.Equals(GetVisibleThreadId(), threadId) Then
                            Return
                        End If

                        Dim scroller = ResolveTranscriptScrollViewer()
                        If scroller Is Nothing Then
                            Return
                        End If

                        Dim extentDelta = Math.Max(0R, scroller.ExtentHeight - anchor.ExtentHeight)
                        Dim desiredOffset = anchor.VerticalOffset + extentDelta
                        desiredOffset = Math.Max(0R, Math.Min(desiredOffset, scroller.ScrollableHeight))

                        DebugTranscriptScroll("chunk_prepend_anchor",
                                              $"restore_offset={FormatTranscriptScrollMetric(desiredOffset)};extentDelta={FormatTranscriptScrollMetric(extentDelta)}",
                                              scroller)

                        _transcriptScrollProgrammaticMoveInProgress = True
                        _suppressTranscriptScrollTracking = True
                        Try
                            scroller.ScrollToVerticalOffset(desiredOffset)
                            If _transcriptScrollFollowMode = TranscriptScrollFollowMode.DetachedByUser Then
                                UpdateTranscriptDetachedAnchorOffset(scroller)
                            End If
                        Finally
                            _suppressTranscriptScrollTracking = False
                            _transcriptScrollProgrammaticMoveInProgress = False
                        End Try
                    End Sub))
        End Sub
    End Class
End Namespace
