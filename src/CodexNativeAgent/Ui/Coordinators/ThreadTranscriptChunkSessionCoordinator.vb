Imports System.Collections.Generic

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class ThreadTranscriptChunkSession
        Public Property ThreadId As String = String.Empty
        Public Property GenerationId As Integer
        Public Property LoadedRangeStart As Integer?
        Public Property LoadedRangeEnd As Integer?
        Public Property IsLoadingOlderChunk As Boolean
        Public Property HasMoreOlderChunks As Boolean
        Public Property PendingPrependRequest As Boolean

        Public Property OlderChunkLoadsRequested As Integer
        Public Property OlderChunkLoadsCompleted As Integer
        Public Property OlderChunkLoadsCanceled As Integer

        Public Property CreatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property LastUpdatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        Public Property LastLifecycleReason As String = String.Empty
    End Class

    Public NotInheritable Class ThreadTranscriptChunkSelectionLoadRegistration
        Public Property ThreadId As String = String.Empty
        Public Property UiLoadVersion As Integer
        Public Property SelectionGenerationId As Integer
        Public Property RegisteredUtc As DateTimeOffset = DateTimeOffset.UtcNow
    End Class

    Public NotInheritable Class ThreadTranscriptChunkSessionCoordinator
        Private ReadOnly _selectionLoadsByUiVersion As New Dictionary(Of Integer, ThreadTranscriptChunkSelectionLoadRegistration)()
        Private _activeSession As ThreadTranscriptChunkSession
        Private _nextSessionGenerationId As Integer = 1
        Private _selectionLoadGenerationId As Integer

        Public ReadOnly Property ActiveSession As ThreadTranscriptChunkSession
            Get
                Return _activeSession
            End Get
        End Property

        Public Function ActivateVisibleThread(threadId As String,
                                              Optional reason As String = Nothing) As ThreadTranscriptChunkSession
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                ResetActiveSession(If(reason, "activate_visible_thread_empty"))
                Return Nothing
            End If

            If _activeSession IsNot Nothing AndAlso
               StringComparer.Ordinal.Equals(_activeSession.ThreadId, normalizedThreadId) Then
                TouchSession(_activeSession, If(reason, "activate_visible_thread_existing"))
                Return _activeSession
            End If

            ResetActiveSession(If(reason, "activate_visible_thread_switch"))

            _activeSession = New ThreadTranscriptChunkSession() With {
                .ThreadId = normalizedThreadId,
                .GenerationId = NextSessionGenerationId(),
                .HasMoreOlderChunks = False
            }
            TouchSession(_activeSession, If(reason, "activate_visible_thread_new"))
            Return _activeSession
        End Function

        Public Function ResetActiveSession(Optional reason As String = Nothing) As ThreadTranscriptChunkSession
            Dim previous = _activeSession
            If previous Is Nothing Then
                Return Nothing
            End If

            TouchSession(previous, If(reason, "reset_active_session"))
            If previous.IsLoadingOlderChunk Then
                previous.IsLoadingOlderChunk = False
                previous.OlderChunkLoadsCanceled += 1
            End If
            previous.PendingPrependRequest = False

            _activeSession = Nothing
            Return previous
        End Function

        Public Function BumpActiveSessionGeneration(Optional reason As String = Nothing) As ThreadTranscriptChunkSession
            If _activeSession Is Nothing Then
                Return Nothing
            End If

            If _activeSession.IsLoadingOlderChunk Then
                _activeSession.IsLoadingOlderChunk = False
                _activeSession.OlderChunkLoadsCanceled += 1
            End If

            _activeSession.PendingPrependRequest = False
            _activeSession.GenerationId = NextSessionGenerationId()
            TouchSession(_activeSession, If(reason, "bump_generation"))
            Return _activeSession
        End Function

        Public Function RegisterThreadSelectionLoad(uiLoadVersion As Integer,
                                                    threadId As String,
                                                    Optional reason As String = Nothing) As ThreadTranscriptChunkSelectionLoadRegistration
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            _selectionLoadGenerationId += 1
            _selectionLoadsByUiVersion.Clear()

            BumpActiveSessionGeneration(If(reason, "register_thread_selection_load"))

            Dim registration As New ThreadTranscriptChunkSelectionLoadRegistration() With {
                .ThreadId = normalizedThreadId,
                .UiLoadVersion = uiLoadVersion,
                .SelectionGenerationId = _selectionLoadGenerationId
            }

            If uiLoadVersion > 0 Then
                _selectionLoadsByUiVersion(uiLoadVersion) = registration
            End If

            Return registration
        End Function

        Public Function IsCurrentThreadSelectionLoad(uiLoadVersion As Integer,
                                                     threadId As String) As Boolean
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If uiLoadVersion <= 0 OrElse String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Dim registration As ThreadTranscriptChunkSelectionLoadRegistration = Nothing
            If Not _selectionLoadsByUiVersion.TryGetValue(uiLoadVersion, registration) OrElse registration Is Nothing Then
                Return False
            End If

            Return registration.SelectionGenerationId = _selectionLoadGenerationId AndAlso
                   StringComparer.Ordinal.Equals(registration.ThreadId, normalizedThreadId)
        End Function

        Public Sub CompleteThreadSelectionLoad(uiLoadVersion As Integer)
            If uiLoadVersion <= 0 Then
                Return
            End If

            _selectionLoadsByUiVersion.Remove(uiLoadVersion)
        End Sub

        Public Sub CancelPendingThreadSelectionLoads(Optional reason As String = Nothing)
            _selectionLoadGenerationId += 1
            _selectionLoadsByUiVersion.Clear()
            BumpActiveSessionGeneration(If(reason, "cancel_pending_thread_selection_loads"))
        End Sub

        Public Function IsActiveSessionGeneration(threadId As String, generationId As Integer) As Boolean
            Dim normalizedThreadId = NormalizeIdentifier(threadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) OrElse generationId <= 0 Then
                Return False
            End If

            If _activeSession Is Nothing Then
                Return False
            End If

            Return StringComparer.Ordinal.Equals(_activeSession.ThreadId, normalizedThreadId) AndAlso
                   _activeSession.GenerationId = generationId
        End Function

        Public Function TryBeginOlderChunkLoad(threadId As String,
                                               generationId As Integer,
                                               Optional reason As String = Nothing) As Boolean
            If Not IsActiveSessionGeneration(threadId, generationId) Then
                Return False
            End If

            If _activeSession.IsLoadingOlderChunk Then
                Return False
            End If

            _activeSession.IsLoadingOlderChunk = True
            _activeSession.OlderChunkLoadsRequested += 1
            TouchSession(_activeSession, If(reason, "begin_older_chunk_load"))
            Return True
        End Function

        Public Function TryCompleteOlderChunkLoad(threadId As String,
                                                  generationId As Integer,
                                                  hasMoreOlderChunks As Boolean,
                                                  Optional loadedRangeStart As Integer? = Nothing,
                                                  Optional loadedRangeEnd As Integer? = Nothing,
                                                  Optional reason As String = Nothing) As Boolean
            If Not IsActiveSessionGeneration(threadId, generationId) Then
                Return False
            End If

            _activeSession.IsLoadingOlderChunk = False
            _activeSession.HasMoreOlderChunks = hasMoreOlderChunks
            _activeSession.LoadedRangeStart = loadedRangeStart
            _activeSession.LoadedRangeEnd = loadedRangeEnd
            _activeSession.PendingPrependRequest = False
            _activeSession.OlderChunkLoadsCompleted += 1
            TouchSession(_activeSession, If(reason, "complete_older_chunk_load"))
            Return True
        End Function

        Public Function TryCancelOlderChunkLoad(threadId As String,
                                                generationId As Integer,
                                                Optional reason As String = Nothing) As Boolean
            If Not IsActiveSessionGeneration(threadId, generationId) Then
                Return False
            End If

            If Not _activeSession.IsLoadingOlderChunk Then
                Return False
            End If

            _activeSession.IsLoadingOlderChunk = False
            _activeSession.PendingPrependRequest = False
            _activeSession.OlderChunkLoadsCanceled += 1
            TouchSession(_activeSession, If(reason, "cancel_older_chunk_load"))
            Return True
        End Function

        Public Function SetPendingPrependRequest(threadId As String,
                                                 generationId As Integer,
                                                 pending As Boolean,
                                                 Optional reason As String = Nothing) As Boolean
            If Not IsActiveSessionGeneration(threadId, generationId) Then
                Return False
            End If

            _activeSession.PendingPrependRequest = pending
            TouchSession(_activeSession, If(reason, "set_pending_prepend_request"))
            Return True
        End Function

        Private Function NextSessionGenerationId() As Integer
            Dim generationId = _nextSessionGenerationId
            _nextSessionGenerationId += 1
            If generationId <= 0 Then
                generationId = 1
                _nextSessionGenerationId = 2
            End If

            Return generationId
        End Function

        Private Shared Sub TouchSession(session As ThreadTranscriptChunkSession, reason As String)
            If session Is Nothing Then
                Return
            End If

            session.LastUpdatedUtc = DateTimeOffset.UtcNow
            session.LastLifecycleReason = If(reason, String.Empty)
        End Sub

        Private Shared Function NormalizeIdentifier(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function
    End Class
End Namespace
