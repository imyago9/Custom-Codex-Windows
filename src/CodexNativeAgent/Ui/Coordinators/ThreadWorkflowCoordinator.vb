Imports System.Threading
Imports System.Threading.Tasks
Imports System.Text.Json.Nodes
Imports CodexNativeAgent.Services
Imports CodexNativeAgent.Ui.ViewModels.Threads

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class ThreadWorkflowCoordinator
        Public Async Function RunStartThreadAsync(resolveTargetCwd As Func(Of String),
                                                  buildThreadRequestOptions As Func(Of Boolean, ThreadRequestOptions),
                                                  startThreadAsync As Func(Of ThreadRequestOptions, CancellationToken, Task(Of JsonObject)),
                                                  applyCurrentThreadFromObject As Action(Of JsonObject),
                                                  clearTranscript As Action,
                                                  renderThreadObject As Action(Of JsonObject),
                                                  appendSystemMessage As Action(Of String),
                                                  showStatus As Action(Of String, Boolean, Boolean),
                                                  refreshThreadsAsync As Func(Of Task)) As Task
            If resolveTargetCwd Is Nothing Then Throw New ArgumentNullException(NameOf(resolveTargetCwd))
            If buildThreadRequestOptions Is Nothing Then Throw New ArgumentNullException(NameOf(buildThreadRequestOptions))
            If startThreadAsync Is Nothing Then Throw New ArgumentNullException(NameOf(startThreadAsync))
            If applyCurrentThreadFromObject Is Nothing Then Throw New ArgumentNullException(NameOf(applyCurrentThreadFromObject))
            If clearTranscript Is Nothing Then Throw New ArgumentNullException(NameOf(clearTranscript))
            If renderThreadObject Is Nothing Then Throw New ArgumentNullException(NameOf(renderThreadObject))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If showStatus Is Nothing Then Throw New ArgumentNullException(NameOf(showStatus))
            If refreshThreadsAsync Is Nothing Then Throw New ArgumentNullException(NameOf(refreshThreadsAsync))

            Dim targetCwd = resolveTargetCwd()
            Dim options = buildThreadRequestOptions(True)
            If Not String.IsNullOrWhiteSpace(targetCwd) Then
                options.Cwd = targetCwd
            End If

            Dim threadObject = Await startThreadAsync(options, CancellationToken.None).ConfigureAwait(True)

            applyCurrentThreadFromObject(threadObject)
            clearTranscript()
            renderThreadObject(threadObject)
            appendSystemMessage("Started thread.")
            showStatus("Started thread.", False, True)
            Await refreshThreadsAsync().ConfigureAwait(True)
        End Function

        Public Async Function RunRefreshThreadsAsync(canBegin As Func(Of Boolean),
                                                     beginUi As Action,
                                                     showArchived As Boolean,
                                                     filterByWorkingDir As Boolean,
                                                     resolveWorkingDir As Func(Of String),
                                                     listThreadsAsync As Func(Of Boolean, String, CancellationToken, Task(Of IReadOnlyList(Of ThreadSummary))),
                                                     applySummaries As Action(Of IReadOnlyList(Of ThreadSummary)),
                                                     completeUi As Action,
                                                     failUi As Action(Of Exception),
                                                     finalizeUi As Action) As Task
            If canBegin Is Nothing Then Throw New ArgumentNullException(NameOf(canBegin))
            If beginUi Is Nothing Then Throw New ArgumentNullException(NameOf(beginUi))
            If resolveWorkingDir Is Nothing Then Throw New ArgumentNullException(NameOf(resolveWorkingDir))
            If listThreadsAsync Is Nothing Then Throw New ArgumentNullException(NameOf(listThreadsAsync))
            If applySummaries Is Nothing Then Throw New ArgumentNullException(NameOf(applySummaries))
            If completeUi Is Nothing Then Throw New ArgumentNullException(NameOf(completeUi))
            If failUi Is Nothing Then Throw New ArgumentNullException(NameOf(failUi))
            If finalizeUi Is Nothing Then Throw New ArgumentNullException(NameOf(finalizeUi))

            If Not canBegin() Then
                Return
            End If

            beginUi()

            Try
                Dim cwd = If(filterByWorkingDir, resolveWorkingDir(), String.Empty)
                Dim summaries = Await listThreadsAsync(showArchived, cwd, CancellationToken.None).ConfigureAwait(True)
                applySummaries(summaries)
                completeUi()
            Catch ex As Exception
                failUi(ex)
                Throw
            Finally
                finalizeUi()
            End Try
        End Function

        Public Async Function RunAutoLoadThreadSelectionAsync(selected As ThreadListEntry,
                                                              forceReload As Boolean,
                                                              tryPrepareThreadId As Func(Of ThreadListEntry, Boolean, String),
                                                              beginRequest As Func(Of String, Object),
                                                              getRequestThreadId As Func(Of Object, String),
                                                              getRequestCancellationToken As Func(Of Object, CancellationToken),
                                                              loadPayloadAsync As Func(Of String, CancellationToken, Task(Of Object)),
                                                              applyPayloadUiAsync As Func(Of Object, Object, Task),
                                                              handleFailureUi As Action(Of Object, Exception),
                                                              finalizeRequestUi As Action(Of Object),
                                                              disposeRequest As Action(Of Object)) As Task
            If tryPrepareThreadId Is Nothing Then Throw New ArgumentNullException(NameOf(tryPrepareThreadId))
            If beginRequest Is Nothing Then Throw New ArgumentNullException(NameOf(beginRequest))
            If getRequestThreadId Is Nothing Then Throw New ArgumentNullException(NameOf(getRequestThreadId))
            If getRequestCancellationToken Is Nothing Then Throw New ArgumentNullException(NameOf(getRequestCancellationToken))
            If loadPayloadAsync Is Nothing Then Throw New ArgumentNullException(NameOf(loadPayloadAsync))
            If applyPayloadUiAsync Is Nothing Then Throw New ArgumentNullException(NameOf(applyPayloadUiAsync))
            If handleFailureUi Is Nothing Then Throw New ArgumentNullException(NameOf(handleFailureUi))
            If finalizeRequestUi Is Nothing Then Throw New ArgumentNullException(NameOf(finalizeRequestUi))
            If disposeRequest Is Nothing Then Throw New ArgumentNullException(NameOf(disposeRequest))

            Dim selectedThreadId = tryPrepareThreadId(selected, forceReload)
            If String.IsNullOrWhiteSpace(selectedThreadId) Then
                Return
            End If

            Dim request = beginRequest(selectedThreadId)
            Try
                Dim requestThreadId = getRequestThreadId(request)
                Dim requestToken = getRequestCancellationToken(request)
                Dim payload = Await loadPayloadAsync(requestThreadId, requestToken).ConfigureAwait(False)
                Await applyPayloadUiAsync(request, payload).ConfigureAwait(False)
            Catch ex As OperationCanceledException
            Catch ex As Exception
                handleFailureUi(request, ex)
            Finally
                finalizeRequestUi(request)
                disposeRequest(request)
            End Try
        End Function

        Public Async Function RunThreadContextActionAsync(Of TTarget As Class)(resolveTarget As Func(Of TTarget),
                                                                               missingTargetMessage As String,
                                                                               executeAsync As Func(Of TTarget, Task)) As Task
            If resolveTarget Is Nothing Then Throw New ArgumentNullException(NameOf(resolveTarget))
            If executeAsync Is Nothing Then Throw New ArgumentNullException(NameOf(executeAsync))

            Dim target = resolveTarget()
            If target Is Nothing Then
                Throw New InvalidOperationException(If(missingTargetMessage, "Missing target."))
            End If

            Await executeAsync(target).ConfigureAwait(True)
        End Function

        Public Async Function RunForkThreadAsync(selected As ThreadListEntry,
                                                 buildThreadRequestOptions As Func(Of Boolean, ThreadRequestOptions),
                                                 forkThreadAsync As Func(Of String, ThreadRequestOptions, CancellationToken, Task(Of JsonObject)),
                                                 applyCurrentThreadFromObject As Action(Of JsonObject),
                                                 renderThreadObject As Action(Of JsonObject),
                                                 appendSystemMessage As Action(Of String),
                                                 showStatus As Action(Of String, Boolean, Boolean),
                                                 refreshThreadsAsync As Func(Of Task)) As Task
            If selected Is Nothing Then Throw New ArgumentNullException(NameOf(selected))
            If buildThreadRequestOptions Is Nothing Then Throw New ArgumentNullException(NameOf(buildThreadRequestOptions))
            If forkThreadAsync Is Nothing Then Throw New ArgumentNullException(NameOf(forkThreadAsync))
            If applyCurrentThreadFromObject Is Nothing Then Throw New ArgumentNullException(NameOf(applyCurrentThreadFromObject))
            If renderThreadObject Is Nothing Then Throw New ArgumentNullException(NameOf(renderThreadObject))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If showStatus Is Nothing Then Throw New ArgumentNullException(NameOf(showStatus))
            If refreshThreadsAsync Is Nothing Then Throw New ArgumentNullException(NameOf(refreshThreadsAsync))

            Dim options = buildThreadRequestOptions(False)
            Dim threadObject = Await forkThreadAsync(selected.Id, options, CancellationToken.None).ConfigureAwait(True)

            applyCurrentThreadFromObject(threadObject)
            renderThreadObject(threadObject)
            appendSystemMessage("Forked thread.")
            showStatus("Forked thread.", False, True)
            Await refreshThreadsAsync().ConfigureAwait(True)
        End Function

        Public Async Function RunArchiveThreadAsync(selected As ThreadListEntry,
                                                    archiveThreadAsync As Func(Of String, CancellationToken, Task),
                                                    appendSystemMessage As Action(Of String),
                                                    showStatus As Action(Of String, Boolean, Boolean),
                                                    refreshThreadsAsync As Func(Of Task)) As Task
            If selected Is Nothing Then Throw New ArgumentNullException(NameOf(selected))
            If archiveThreadAsync Is Nothing Then Throw New ArgumentNullException(NameOf(archiveThreadAsync))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If showStatus Is Nothing Then Throw New ArgumentNullException(NameOf(showStatus))
            If refreshThreadsAsync Is Nothing Then Throw New ArgumentNullException(NameOf(refreshThreadsAsync))

            Await archiveThreadAsync(selected.Id, CancellationToken.None).ConfigureAwait(True)
            appendSystemMessage($"Archived thread {selected.Id}.")
            showStatus($"Archived thread {selected.Id}.", False, False)
            Await refreshThreadsAsync().ConfigureAwait(True)
        End Function

        Public Async Function RunUnarchiveThreadAsync(selected As ThreadListEntry,
                                                      unarchiveThreadAsync As Func(Of String, CancellationToken, Task),
                                                      appendSystemMessage As Action(Of String),
                                                      showStatus As Action(Of String, Boolean, Boolean),
                                                      refreshThreadsAsync As Func(Of Task)) As Task
            If selected Is Nothing Then Throw New ArgumentNullException(NameOf(selected))
            If unarchiveThreadAsync Is Nothing Then Throw New ArgumentNullException(NameOf(unarchiveThreadAsync))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If showStatus Is Nothing Then Throw New ArgumentNullException(NameOf(showStatus))
            If refreshThreadsAsync Is Nothing Then Throw New ArgumentNullException(NameOf(refreshThreadsAsync))

            Await unarchiveThreadAsync(selected.Id, CancellationToken.None).ConfigureAwait(True)
            appendSystemMessage($"Unarchived thread {selected.Id}.")
            showStatus($"Unarchived thread {selected.Id}.", False, False)
            Await refreshThreadsAsync().ConfigureAwait(True)
        End Function
    End Class
End Namespace
