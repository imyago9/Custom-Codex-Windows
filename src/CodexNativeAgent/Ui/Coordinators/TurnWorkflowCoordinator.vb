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

        Private ReadOnly _viewModel As MainWindowViewModel
        Private ReadOnly _turnService As ITurnService
        Private ReadOnly _approvalService As IApprovalService
        Private ReadOnly _runUiActionAsync As Func(Of Func(Of Task), Task)
        Private ReadOnly _approvalQueue As New Queue(Of PendingApprovalInfo)()
        Private _activeApproval As PendingApprovalInfo

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
                Return _activeApproval IsNot Nothing
            End Get
        End Property

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
            _approvalQueue.Clear()
            _activeApproval = Nothing
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
                                                       refreshControlStates As Action,
                                                       appendSystemMessage As Action(Of String),
                                                       showStatus As Action(Of String, Boolean, Boolean)) As Task
            If request Is Nothing Then Throw New ArgumentNullException(NameOf(request))
            If handleToolRequestUserInputAsync Is Nothing Then Throw New ArgumentNullException(NameOf(handleToolRequestUserInputAsync))
            If handleUnsupportedToolCallAsync Is Nothing Then Throw New ArgumentNullException(NameOf(handleUnsupportedToolCallAsync))
            If handleChatgptTokenRefreshAsync Is Nothing Then Throw New ArgumentNullException(NameOf(handleChatgptTokenRefreshAsync))
            If sendUnsupportedServerRequestErrorAsync Is Nothing Then Throw New ArgumentNullException(NameOf(sendUnsupportedServerRequestErrorAsync))
            If refreshControlStates Is Nothing Then Throw New ArgumentNullException(NameOf(refreshControlStates))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If showStatus Is Nothing Then Throw New ArgumentNullException(NameOf(showStatus))

            Dim approvalInfo As PendingApprovalInfo = Nothing
            If _approvalService.TryCreateApproval(request, approvalInfo) Then
                QueueApproval(approvalInfo, refreshControlStates, appendSystemMessage, showStatus)
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
                                                   sendResultAsync As Func(Of JsonNode, JsonObject, Task),
                                                   refreshControlStates As Action,
                                                   appendSystemMessage As Action(Of String),
                                                   showStatus As Action(Of String, Boolean, Boolean)) As Task
            If sendResultAsync Is Nothing Then Throw New ArgumentNullException(NameOf(sendResultAsync))
            If refreshControlStates Is Nothing Then Throw New ArgumentNullException(NameOf(refreshControlStates))
            If appendSystemMessage Is Nothing Then Throw New ArgumentNullException(NameOf(appendSystemMessage))
            If showStatus Is Nothing Then Throw New ArgumentNullException(NameOf(showStatus))

            If _activeApproval Is Nothing Then
                _viewModel.ApprovalPanel.RecordError("No active approval.")
                Return
            End If

            Dim decisionLabel As String = Nothing
            Dim decisionPayload = _approvalService.ResolveDecisionPayload(_activeApproval, action, decisionLabel)
            If decisionPayload Is Nothing OrElse String.IsNullOrWhiteSpace(decisionLabel) Then
                _viewModel.ApprovalPanel.RecordError("No decision mapping is available for this approval type.")
                Throw New InvalidOperationException("No decision mapping is available for this approval type.")
            End If

            Try
                Dim resultNode As New JsonObject()
                resultNode("decision") = decisionPayload

                Await sendResultAsync(_activeApproval.RequestId, resultNode).ConfigureAwait(True)
                appendSystemMessage($"Approval sent: {decisionLabel}")
                showStatus($"Approval sent: {decisionLabel}", False, True)
                _viewModel.ApprovalPanel.OnApprovalResolved(action, decisionLabel, _approvalQueue.Count)

                RaiseEvent ApprovalResolved(Me,
                                            New ApprovalResolvedEventArgs() With {
                                                .RequestId = CloneJson(_activeApproval.RequestId),
                                                .MethodName = _activeApproval.MethodName,
                                                .Decision = decisionLabel,
                                                .ThreadId = _activeApproval.ThreadId,
                                                .TurnId = _activeApproval.TurnId,
                                                .ItemId = _activeApproval.ItemId
                                            })

                _activeApproval = Nothing
                ShowNextApprovalIfNeeded(refreshControlStates)
            Catch ex As Exception
                _viewModel.ApprovalPanel.RecordError(ex.Message)
                Throw
            End Try
        End Function

        Private Sub QueueApproval(approvalInfo As PendingApprovalInfo,
                                  refreshControlStates As Action,
                                  appendSystemMessage As Action(Of String),
                                  showStatus As Action(Of String, Boolean, Boolean))
            If approvalInfo Is Nothing Then
                Return
            End If

            _approvalQueue.Enqueue(approvalInfo)
            _viewModel.ApprovalPanel.OnApprovalQueued(approvalInfo.MethodName, _approvalQueue.Count)
            RaiseEvent ApprovalQueued(Me,
                                      New ApprovalQueuedEventArgs() With {
                                          .Approval = approvalInfo
                                      })
            ShowNextApprovalIfNeeded(refreshControlStates)
            appendSystemMessage($"Approval queued: {approvalInfo.MethodName}")
            showStatus($"Approval queued: {approvalInfo.MethodName}", False, True)
        End Sub

        Private Sub ShowNextApprovalIfNeeded(refreshControlStates As Action)
            If refreshControlStates Is Nothing Then
                Throw New ArgumentNullException(NameOf(refreshControlStates))
            End If

            If _activeApproval IsNot Nothing Then
                Return
            End If

            If _approvalQueue.Count = 0 Then
                _viewModel.ApprovalPanel.OnApprovalQueueEmpty()
                refreshControlStates()
                Return
            End If

            _activeApproval = _approvalQueue.Dequeue()
            _viewModel.ApprovalPanel.OnApprovalActivated(_activeApproval.MethodName,
                                                         _activeApproval.Summary,
                                                         _activeApproval.SupportsExecpolicyAmendment,
                                                         _approvalQueue.Count)
            RaiseEvent ApprovalActivated(Me,
                                         New ApprovalQueuedEventArgs() With {
                                             .Approval = _activeApproval
                                         })
            refreshControlStates()
        End Sub
    End Class
End Namespace
