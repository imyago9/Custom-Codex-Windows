Imports System.Collections.Generic
Imports System.Text
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media.Animation
Imports CodexNativeAgent.AppServer
Imports CodexNativeAgent.Services

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Async Function RefreshModelsAsync() As Task
            Dim previousModelId = SelectedModelId()
            Dim models = Await _threadService.ListModelsAsync(CancellationToken.None)

            WorkspacePaneHost.CmbModel.Items.Clear()
            For Each model In models
                WorkspacePaneHost.CmbModel.Items.Add(New ModelListEntry() With {
                    .Id = model.Id,
                    .DisplayName = model.DisplayName,
                    .IsDefault = model.IsDefault
                })
            Next

            Dim selectedIndex = -1
            For i = 0 To WorkspacePaneHost.CmbModel.Items.Count - 1
                Dim item = TryCast(WorkspacePaneHost.CmbModel.Items(i), ModelListEntry)
                If item Is Nothing Then
                    Continue For
                End If

                If StringComparer.Ordinal.Equals(item.Id, previousModelId) Then
                    selectedIndex = i
                    Exit For
                End If

                If selectedIndex = -1 AndAlso item.IsDefault Then
                    selectedIndex = i
                End If
            Next

            If selectedIndex >= 0 Then
                WorkspacePaneHost.CmbModel.SelectedIndex = selectedIndex
            End If

            _modelsLoadedAtLeastOnce = True
            ShowStatus($"Loaded {WorkspacePaneHost.CmbModel.Items.Count} model(s).")
            AppendSystemMessage($"Loaded {WorkspacePaneHost.CmbModel.Items.Count} models.")
        End Function

        Private Function SelectedModelId() As String
            Dim selected = TryCast(WorkspacePaneHost.CmbModel.SelectedItem, ModelListEntry)
            If selected Is Nothing Then
                Return String.Empty
            End If

            Return selected.Id
        End Function

        Private Shared Function SelectedComboValue(comboBox As ComboBox) As String
            If comboBox Is Nothing OrElse comboBox.SelectedItem Is Nothing Then
                Return String.Empty
            End If

            Dim comboItem = TryCast(comboBox.SelectedItem, ComboBoxItem)
            If comboItem IsNot Nothing Then
                Return If(comboItem.Content, String.Empty).ToString().Trim()
            End If

            Return comboBox.SelectedItem.ToString().Trim()
        End Function

        Private Shared Function ParseIntegerCommandParameter(parameter As Object) As Integer
            If parameter Is Nothing Then
                Return -1
            End If

            Dim parsed As Integer
            If Integer.TryParse(parameter.ToString(), parsed) Then
                Return parsed
            End If

            Return -1
        End Function

        Private Shared Function DensityValueFromIndex(selectedIndex As Integer) As String
            If selectedIndex = 1 Then
                Return AppAppearanceManager.CompactDensity
            End If

            Return AppAppearanceManager.ComfortableDensity
        End Function

        Private Shared Function ThreadSortLabel(sortIndex As Integer) As String
            Select Case sortIndex
                Case 1
                    Return "Oldest first"
                Case 2
                    Return "Preview A-Z"
                Case 3
                    Return "Preview Z-A"
                Case Else
                    Return "Newest first"
            End Select
        End Function

        Private Sub ToggleTurnComposerPickersCollapsed()
            _turnComposerPickersCollapsed = Not _turnComposerPickersCollapsed
            ApplyTurnComposerPickersCollapsedState(animated:=True, persist:=True)
        End Sub

        Private Sub ApplyTurnComposerPickersCollapsedState(animated As Boolean, persist As Boolean)
            If WorkspacePaneHost Is Nothing Then
                Return
            End If

            Dim container = WorkspacePaneHost.TurnComposerPickersContainer
            Dim toggleButton = WorkspacePaneHost.BtnTurnComposerPickersToggle
            Dim reasoningCombo = WorkspacePaneHost.CmbReasoningEffort
            Dim modelCombo = WorkspacePaneHost.CmbModel

            If container Is Nothing OrElse toggleButton Is Nothing Then
                Return
            End If

            If Not _turnComposerPickersCollapsed Then
                CaptureTurnComposerPickersExpandedWidth()
            End If

            Dim expandedWidth = ResolveTurnComposerPickersExpandedWidth()
            Dim targetWidth = If(_turnComposerPickersCollapsed, 0.0R, expandedWidth)
            Dim targetOpacity = If(_turnComposerPickersCollapsed, 0.0R, 1.0R)
            Dim currentWidth = EffectiveAnimatedWidth(container, expandedWidth)
            Dim currentOpacity = If(Double.IsNaN(container.Opacity), If(currentWidth <= 0.5R, 0.0R, 1.0R), container.Opacity)

            If _turnComposerPickersCollapsed AndAlso
               ((modelCombo IsNot Nothing AndAlso modelCombo.IsKeyboardFocusWithin) OrElse
                (reasoningCombo IsNot Nothing AndAlso reasoningCombo.IsKeyboardFocusWithin)) Then
                toggleButton.Focus()
            End If

            UpdateTurnComposerPickersToggleVisual()

            container.BeginAnimation(FrameworkElement.WidthProperty, Nothing)
            container.BeginAnimation(UIElement.OpacityProperty, Nothing)
            container.Visibility = Visibility.Visible
            container.IsHitTestVisible = Not _turnComposerPickersCollapsed

            If Not animated Then
                container.Width = targetWidth
                container.Opacity = targetOpacity
                If persist Then
                    SaveSettings()
                End If

                Return
            End If

            ' Set the final base values first so they persist after the animation clock completes.
            container.Width = targetWidth
            container.Opacity = targetOpacity

            Dim duration = TimeSpan.FromMilliseconds(170)
            Dim easing As New CubicEase() With {.EasingMode = EasingMode.EaseInOut}

            Dim widthAnimation As New DoubleAnimation() With {
                .From = currentWidth,
                .To = targetWidth,
                .Duration = duration,
                .EasingFunction = easing,
                .FillBehavior = FillBehavior.Stop
            }
            Dim opacityAnimation As New DoubleAnimation() With {
                .From = currentOpacity,
                .To = targetOpacity,
                .Duration = TimeSpan.FromMilliseconds(140),
                .EasingFunction = easing,
                .FillBehavior = FillBehavior.Stop
            }

            container.BeginAnimation(FrameworkElement.WidthProperty, widthAnimation)
            container.BeginAnimation(UIElement.OpacityProperty, opacityAnimation)

            If persist Then
                SaveSettings()
            End If
        End Sub

        Private Sub UpdateTurnComposerPickersToggleVisual()
            If WorkspacePaneHost Is Nothing Then
                Return
            End If

            If WorkspacePaneHost.TxtTurnComposerPickersToggleGlyph IsNot Nothing Then
                WorkspacePaneHost.TxtTurnComposerPickersToggleGlyph.Text = If(_turnComposerPickersCollapsed, ChrW(&HE76C), ChrW(&HE76B))
            End If

            If WorkspacePaneHost.BtnTurnComposerPickersToggle IsNot Nothing Then
                WorkspacePaneHost.BtnTurnComposerPickersToggle.ToolTip =
                    If(_turnComposerPickersCollapsed, "Expand model and reasoning", "Collapse model and reasoning")
            End If
        End Sub

        Private Sub CaptureTurnComposerPickersExpandedWidth()
            Dim width = ComputeTurnComposerPickersExpandedWidth()
            If width > 1.0R Then
                _turnComposerPickersExpandedWidth = Math.Max(width, 1.0R)
            End If
        End Sub

        Private Function ResolveTurnComposerPickersExpandedWidth() As Double
            Dim measured = ComputeTurnComposerPickersExpandedWidth()
            If measured > 1.0R Then
                _turnComposerPickersExpandedWidth = Math.Max(measured, 1.0R)
            End If

            If _turnComposerPickersExpandedWidth <= 1.0R Then
                _turnComposerPickersExpandedWidth = 434.0R
            End If

            Return _turnComposerPickersExpandedWidth
        End Function

        Private Function ComputeTurnComposerPickersExpandedWidth() As Double
            If WorkspacePaneHost Is Nothing Then
                Return 0.0R
            End If

            Dim container = WorkspacePaneHost.TurnComposerPickersContainer
            If container IsNot Nothing AndAlso container.ActualWidth > 1.0R AndAlso container.Width > 0.0R Then
                Return container.ActualWidth
            End If

            Dim totalWidth = 0.0R
            totalWidth += MeasureFrameworkElementWidth(WorkspacePaneHost.CmbModel)
            totalWidth += MeasureFrameworkElementWidth(WorkspacePaneHost.CmbReasoningEffort)
            Return totalWidth
        End Function

        Private Shared Function MeasureFrameworkElementWidth(element As FrameworkElement) As Double
            If element Is Nothing Then
                Return 0.0R
            End If

            Dim width = element.ActualWidth
            If width <= 0.0R OrElse Double.IsNaN(width) Then
                width = element.Width
            End If

            If width <= 0.0R OrElse Double.IsNaN(width) Then
                width = element.MinWidth
            End If

            If width < 0.0R OrElse Double.IsNaN(width) Then
                width = 0.0R
            End If

            width += element.Margin.Left + element.Margin.Right
            Return Math.Max(width, 0.0R)
        End Function

        Private Shared Function EffectiveAnimatedWidth(container As FrameworkElement, fallbackExpandedWidth As Double) As Double
            If container Is Nothing Then
                Return Math.Max(fallbackExpandedWidth, 1.0R)
            End If

            Dim width = container.ActualWidth
            If width <= 0.0R OrElse Double.IsNaN(width) Then
                width = container.Width
            End If

            If width <= 0.0R OrElse Double.IsNaN(width) Then
                width = fallbackExpandedWidth
            End If

            Return Math.Max(width, 0.0R)
        End Function

        Private Async Function StartTurnAsync() As Task
            Dim submittedInputText As String = String.Empty
            Dim shouldRefreshThreadsAfterTurnStart As Boolean = False

            If _pendingNewThreadFirstPromptSelection AndAlso String.IsNullOrWhiteSpace(_currentThreadId) Then
                Await EnsurePendingDraftThreadCreatedAsync()
            End If

            Await _turnWorkflowCoordinator.RunStartTurnAsync(
                _currentThreadId,
                _viewModel.TurnComposer.InputText,
                _viewModel.TurnComposer.SelectedModelId,
                _viewModel.TurnComposer.SelectedReasoningEffort,
                _viewModel.TurnComposer.SelectedApprovalPolicy,
                AddressOf EnsureThreadSelected,
                Sub(inputText)
                    submittedInputText = If(inputText, String.Empty)
                    AppendTranscript("user", inputText)
                    TrackPendingUserEcho(inputText)
                End Sub,
                Sub(returnedTurnId)
                    If Not String.IsNullOrWhiteSpace(returnedTurnId) Then
                        _currentTurnId = returnedTurnId
                    End If

                    MarkThreadLastActive(_currentThreadId)
                    shouldRefreshThreadsAfterTurnStart = SyncThreadListAfterUserPrompt(_currentThreadId, submittedInputText)
                    FinalizePendingNewThreadFirstPromptSelection()
                    UpdateThreadTurnLabels()
                    RefreshControlStates()
                    WorkspacePaneHost.TxtTurnInput.Clear()
                    ShowStatus($"Turn started: {_currentTurnId}")
                End Sub)

            If shouldRefreshThreadsAfterTurnStart Then
                Await RefreshThreadsCoreAsync(silent:=True)
                FinalizePendingNewThreadFirstPromptSelection()
            End If
        End Function

        Private Async Function SteerTurnAsync() As Task
            Await _turnWorkflowCoordinator.RunSteerTurnAsync(
                _currentThreadId,
                _currentTurnId,
                _viewModel.TurnComposer.InputText,
                AddressOf EnsureThreadSelected,
                Sub(steerText)
                    AppendTranscript("user (steer)", steerText)
                    TrackPendingUserEcho(steerText)
                End Sub,
                Sub(returnedTurnId)
                    If Not String.IsNullOrWhiteSpace(returnedTurnId) Then
                        _currentTurnId = returnedTurnId
                    End If

                    MarkThreadLastActive(_currentThreadId)
                    UpdateThreadTurnLabels()
                    RefreshControlStates()
                    WorkspacePaneHost.TxtTurnInput.Clear()
                    ShowStatus($"Turn steered: {_currentTurnId}")
                End Sub)
        End Function

        Private Async Function InterruptTurnAsync() As Task
            Await _turnWorkflowCoordinator.RunInterruptTurnAsync(
                _currentThreadId,
                _currentTurnId,
                AddressOf EnsureThreadSelected,
                Sub(turnId)
                    AppendSystemMessage($"Interrupt requested for turn {turnId}.")
                    ShowStatus($"Interrupt requested for turn {turnId}.", displayToast:=True)
                End Sub)
        End Function

        Private Sub AppendTranscript(role As String, text As String)
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            _viewModel.TranscriptPanel.AppendRoleMessage(role, text)
            ScrollTranscriptToBottom()
        End Sub

        Private Sub AppendSystemMessage(message As String)
            If String.IsNullOrWhiteSpace(message) Then
                Return
            End If

            _viewModel.TranscriptPanel.AppendSystemMessage(message)
            ScrollTranscriptToBottom()
            ShowStatus(message)
        End Sub

        Private Sub AppendProtocol(direction As String, payload As String)
            Dim safePayload = If(payload, String.Empty)
            _viewModel.TranscriptPanel.AppendProtocolChunk($"[{Now:HH:mm:ss}] {direction}: {safePayload}{Environment.NewLine}")
            ScrollTextBoxToBottom(WorkspacePaneHost.TxtProtocol)
        End Sub

        Private Sub RenderItem(itemObject As JsonObject)
            Dim itemType = GetPropertyString(itemObject, "type")

            Select Case itemType
                Case "userMessage"
                    Dim content = GetPropertyArray(itemObject, "content")
                    Dim userText = FlattenUserInput(content)
                    If Not ShouldSuppressUserEcho(userText) Then
                        AppendTranscript("user", userText)
                    End If

                Case "agentMessage"
                    Dim agentText = GetPropertyString(itemObject, "text")
                    If IsCommentaryAgentMessage(itemObject) Then
                        _viewModel.TranscriptPanel.AppendAssistantCommentaryReasoningStep(agentText)
                        ScrollTranscriptToBottom()
                    Else
                        AppendTranscript("assistant", agentText)
                    End If

                Case "plan"
                    _viewModel.TranscriptPanel.AppendPlan(GetPropertyString(itemObject, "text"))
                    ScrollTranscriptToBottom()

                Case "reasoning"
                    Dim reasoningText = ExtractReasoningText(itemObject)
                    If Not String.IsNullOrWhiteSpace(reasoningText) Then
                        _viewModel.TranscriptPanel.AppendReasoning(reasoningText)
                        ScrollTranscriptToBottom()
                    End If

                Case "commandExecution"
                    Dim command = GetPropertyString(itemObject, "command")
                    Dim status = GetPropertyString(itemObject, "status")
                    Dim output = GetPropertyString(itemObject, "aggregatedOutput")
                    _viewModel.TranscriptPanel.AppendCommandExecution(command, status, output)
                    ScrollTranscriptToBottom()

                Case "fileChange"
                    Dim status = GetPropertyString(itemObject, "status")
                    Dim changes = GetPropertyArray(itemObject, "changes")
                    Dim count = If(changes Is Nothing, 0, changes.Count)
                    Dim lineStats = BuildFileChangeLineStats(changes)
                    _viewModel.TranscriptPanel.AppendFileChangeSummary(status,
                                                                       count,
                                                                       BuildFileChangeDetails(changes),
                                                                       lineStats.AddedLineCount,
                                                                       lineStats.RemovedLineCount)
                    ScrollTranscriptToBottom()

                Case Else
                    Dim itemId = GetPropertyString(itemObject, "id")
                    _viewModel.TranscriptPanel.AppendUnknownItem(itemType, itemId)
                    ScrollTranscriptToBottom()
            End Select
        End Sub

        Private Sub UpdateThreadTurnLabels()
            Dim showDraftNewThreadTitle = _pendingNewThreadFirstPromptSelection OrElse String.IsNullOrWhiteSpace(_currentThreadId)
            _viewModel.CurrentThreadText = If(showDraftNewThreadTitle,
                                              "New thread",
                                              _currentThreadId)

            _viewModel.CurrentTurnText = If(String.IsNullOrWhiteSpace(_currentTurnId),
                                            "Turn: 0",
                                            $"Turn: {_currentTurnId}")
            SyncSessionStateViewModel()
        End Sub

        Private Shared Function FlattenUserInput(content As JsonArray) As String
            If content Is Nothing OrElse content.Count = 0 Then
                Return String.Empty
            End If

            Dim parts As New List(Of String)()

            For Each entryNode In content
                Dim entryObject = AsObject(entryNode)
                If entryObject Is Nothing Then
                    Continue For
                End If

                Dim kind = GetPropertyString(entryObject, "type")
                Select Case kind
                    Case "text"
                        Dim value = GetPropertyString(entryObject, "text")
                        If Not String.IsNullOrWhiteSpace(value) Then
                            parts.Add(value)
                        End If
                    Case "image"
                        parts.Add($"[image] {GetPropertyString(entryObject, "url")}")
                    Case "localImage"
                        parts.Add($"[localImage] {GetPropertyString(entryObject, "path")}")
                    Case "mention"
                        parts.Add($"[mention] {GetPropertyString(entryObject, "name")}")
                    Case "skill"
                        parts.Add($"[skill] {GetPropertyString(entryObject, "name")}")
                End Select
            Next

            Return String.Join(Environment.NewLine, parts)
        End Function

        Private Sub TrackPendingUserEcho(text As String)
            Dim normalized = NormalizeUserEchoText(text)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return
            End If

            _pendingLocalUserEchoes.Enqueue(New PendingUserEcho With {
                .Text = normalized,
                .AddedUtc = DateTimeOffset.UtcNow
            })

            PrunePendingUserEchoes(DateTimeOffset.UtcNow)
        End Sub

        Private Function ShouldSuppressUserEcho(text As String) As Boolean
            Dim normalized = NormalizeUserEchoText(text)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            PrunePendingUserEchoes(DateTimeOffset.UtcNow)

            If _pendingLocalUserEchoes.Count = 0 Then
                Return False
            End If

            If StringComparer.Ordinal.Equals(_pendingLocalUserEchoes.Peek().Text, normalized) Then
                _pendingLocalUserEchoes.Dequeue()
                Return True
            End If

            Dim size = _pendingLocalUserEchoes.Count
            Dim matched = False
            For i = 0 To size - 1
                Dim candidate = _pendingLocalUserEchoes.Dequeue()
                If Not matched AndAlso StringComparer.Ordinal.Equals(candidate.Text, normalized) Then
                    matched = True
                    Continue For
                End If

                _pendingLocalUserEchoes.Enqueue(candidate)
            Next

            Return matched
        End Function

        Private Sub PrunePendingUserEchoes(nowUtc As DateTimeOffset)
            While _pendingLocalUserEchoes.Count > 16
                _pendingLocalUserEchoes.Dequeue()
            End While

            While _pendingLocalUserEchoes.Count > 0
                Dim pending = _pendingLocalUserEchoes.Peek()
                If nowUtc - pending.AddedUtc <= PendingUserEchoMaxAge Then
                    Exit While
                End If

                _pendingLocalUserEchoes.Dequeue()
            End While
        End Sub

        Private Shared Function NormalizeUserEchoText(text As String) As String
            If text Is Nothing Then
                Return String.Empty
            End If

            Return text.Trim()
        End Function

        Private Shared Function BuildFileChangeDetails(changes As JsonArray) As String
            If changes Is Nothing OrElse changes.Count = 0 Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder()
            Dim shown = 0
            For Each changeNode In changes
                If shown >= 12 Then
                    Exit For
                End If

                Dim changeObject = AsObject(changeNode)
                If changeObject Is Nothing Then
                    Continue For
                End If

                Dim path = GetPropertyString(changeObject, "path")
                If String.IsNullOrWhiteSpace(path) Then
                    path = GetPropertyString(changeObject, "file")
                End If
                If String.IsNullOrWhiteSpace(path) Then
                    Continue For
                End If

                Dim status = GetPropertyString(changeObject, "status")
                If String.IsNullOrWhiteSpace(status) Then
                    builder.AppendLine(path)
                Else
                    builder.AppendLine($"{status}: {path}")
                End If

                shown += 1
            Next

            If builder.Length = 0 Then
                Return String.Empty
            End If

            If changes.Count > shown Then
                builder.Append("... +")
                builder.Append(changes.Count - shown)
                builder.Append(" more")
            End If

            Return builder.ToString().TrimEnd()
        End Function

        Private Structure FileChangeLineStats
            Public Property AddedLineCount As Integer?
            Public Property RemovedLineCount As Integer?
        End Structure

        Private Shared Function BuildFileChangeLineStats(changes As JsonArray) As FileChangeLineStats
            Dim totalAdded As Integer = 0
            Dim totalRemoved As Integer = 0
            Dim hasAnyStats = False

            If changes Is Nothing OrElse changes.Count = 0 Then
                Return New FileChangeLineStats()
            End If

            For Each changeNode In changes
                Dim changeObject = AsObject(changeNode)
                If changeObject Is Nothing Then
                    Continue For
                End If

                Dim addedValue As Integer
                Dim removedValue As Integer
                Dim hasAdded = TryGetFileChangeLineCount(changeObject, addedValue, isAdded:=True)
                Dim hasRemoved = TryGetFileChangeLineCount(changeObject, removedValue, isAdded:=False)

                If Not hasAdded AndAlso Not hasRemoved Then
                    Dim diffStats = CountFileChangeDiffLines(changeObject)
                    If diffStats.AddedLineCount.HasValue Then
                        addedValue = diffStats.AddedLineCount.Value
                        hasAdded = True
                    End If
                    If diffStats.RemovedLineCount.HasValue Then
                        removedValue = diffStats.RemovedLineCount.Value
                        hasRemoved = True
                    End If
                End If

                If hasAdded AndAlso addedValue > 0 Then
                    totalAdded += addedValue
                    hasAnyStats = True
                End If

                If hasRemoved AndAlso removedValue > 0 Then
                    totalRemoved += removedValue
                    hasAnyStats = True
                End If
            Next

            If Not hasAnyStats Then
                Return New FileChangeLineStats()
            End If

            Return New FileChangeLineStats() With {
                .AddedLineCount = If(totalAdded > 0, CType(totalAdded, Integer?), Nothing),
                .RemovedLineCount = If(totalRemoved > 0, CType(totalRemoved, Integer?), Nothing)
            }
        End Function

        Private Shared Function TryGetFileChangeLineCount(changeObject As JsonObject,
                                                          ByRef value As Integer,
                                                          isAdded As Boolean) As Boolean
            value = 0
            If changeObject Is Nothing Then
                Return False
            End If

            Dim keys = If(isAdded,
                          New String() {"addedLines", "linesAdded", "additions", "added_count", "added"},
                          New String() {"removedLines", "linesRemoved", "deletions", "removed_count", "removed"})

            For Each key In keys
                If TryGetNonNegativeIntegerProperty(changeObject, key, value) Then
                    Return True
                End If
            Next

            Dim nestedKeys = {"stats", "summary", "lineStats", "counts"}
            For Each nestedKey In nestedKeys
                Dim nestedObject = GetPropertyObject(changeObject, nestedKey)
                If nestedObject Is Nothing Then
                    Continue For
                End If

                For Each key In keys
                    If TryGetNonNegativeIntegerProperty(nestedObject, key, value) Then
                        Return True
                    End If
                Next
            Next

            Return False
        End Function

        Private Shared Function TryGetNonNegativeIntegerProperty(obj As JsonObject,
                                                                 propertyName As String,
                                                                 ByRef value As Integer) As Boolean
            value = 0
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return False
            End If

            Dim raw = GetPropertyString(obj, propertyName)
            If String.IsNullOrWhiteSpace(raw) Then
                Return False
            End If

            Dim parsed As Integer
            If Integer.TryParse(raw.Trim(), parsed) Then
                value = Math.Max(0, parsed)
                Return True
            End If

            Return False
        End Function

        Private Shared Function CountFileChangeDiffLines(changeObject As JsonObject) As FileChangeLineStats
            If changeObject Is Nothing Then
                Return New FileChangeLineStats()
            End If

            Dim diffCandidates = {
                GetPropertyString(changeObject, "patch"),
                GetPropertyString(changeObject, "diff"),
                GetPropertyString(changeObject, "unifiedDiff"),
                GetPropertyString(changeObject, "aggregatedOutput"),
                GetPropertyString(changeObject, "output")
            }

            For Each candidate In diffCandidates
                Dim stats = CountUnifiedDiffLines(candidate)
                If stats.AddedLineCount.HasValue OrElse stats.RemovedLineCount.HasValue Then
                    Return stats
                End If
            Next

            Return New FileChangeLineStats()
        End Function

        Private Shared Function CountUnifiedDiffLines(diffText As String) As FileChangeLineStats
            If String.IsNullOrWhiteSpace(diffText) Then
                Return New FileChangeLineStats()
            End If

            Dim added = 0
            Dim removed = 0
            Dim normalized = diffText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)

            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty)
                If line.Length = 0 Then
                    Continue For
                End If

                If line.StartsWith("+++", StringComparison.Ordinal) OrElse
                   line.StartsWith("---", StringComparison.Ordinal) OrElse
                   line.StartsWith("@@", StringComparison.Ordinal) OrElse
                   line.StartsWith("diff ", StringComparison.OrdinalIgnoreCase) OrElse
                   line.StartsWith("index ", StringComparison.OrdinalIgnoreCase) OrElse
                   line.StartsWith("\ No newline", StringComparison.Ordinal) Then
                    Continue For
                End If

                If line(0) = "+"c Then
                    added += 1
                ElseIf line(0) = "-"c Then
                    removed += 1
                End If
            Next

            If added <= 0 AndAlso removed <= 0 Then
                Return New FileChangeLineStats()
            End If

            Return New FileChangeLineStats() With {
                .AddedLineCount = If(added > 0, CType(added, Integer?), Nothing),
                .RemovedLineCount = If(removed > 0, CType(removed, Integer?), Nothing)
            }
        End Function

        Private Shared Function ExtractReasoningText(itemObject As JsonObject) As String
            If itemObject Is Nothing Then
                Return String.Empty
            End If

            Dim text = GetPropertyString(itemObject, "text")
            If Not String.IsNullOrWhiteSpace(text) Then
                Return text
            End If

            text = GetPropertyString(itemObject, "summary")
            If Not String.IsNullOrWhiteSpace(text) Then
                Return text
            End If

            Dim content = GetPropertyArray(itemObject, "content")
            If content IsNot Nothing Then
                Dim builder As New StringBuilder()
                For Each entryNode In content
                    Dim entryObject = AsObject(entryNode)
                    If entryObject Is Nothing Then
                        Continue For
                    End If

                    Dim part = GetPropertyString(entryObject, "text")
                    If String.IsNullOrWhiteSpace(part) Then
                        Continue For
                    End If

                    If builder.Length > 0 Then
                        builder.AppendLine()
                    End If
                    builder.Append(part.Trim())
                Next

                If builder.Length > 0 Then
                    Return builder.ToString()
                End If
            End If

            Return String.Empty
        End Function

        Private Async Function HandleServerRequestAsync(request As RpcServerRequest) As Task
            Await _turnWorkflowCoordinator.HandleServerRequestAsync(
                request,
                AddressOf HandleToolRequestUserInputAsync,
                AddressOf HandleUnsupportedToolCallAsync,
                AddressOf HandleChatgptTokenRefreshAsync,
                Async Function(serverRequest, code, message)
                    Await CurrentClient().SendErrorAsync(serverRequest.Id, code, message)
                End Function,
                AddressOf RefreshControlStates,
                AddressOf AppendSystemMessage,
                Sub(message, isError, displayToast)
                    ShowStatus(message, isError:=isError, displayToast:=displayToast)
                End Sub)
        End Function

        Private Async Function ResolveApprovalAsync(action As String) As Task
            Await _turnWorkflowCoordinator.ResolveApprovalAsync(
                action,
                Async Function(requestId, resultNode)
                    Await CurrentClient().SendResultAsync(requestId, resultNode)
                End Function,
                AddressOf RefreshControlStates,
                AddressOf AppendSystemMessage,
                Sub(message, isError, displayToast)
                    ShowStatus(message, isError:=isError, displayToast:=displayToast)
                End Sub)
        End Function

        Private Async Function HandleToolRequestUserInputAsync(request As RpcServerRequest) As Task
            Dim paramsObject = AsObject(request.ParamsNode)
            Dim questions = GetPropertyArray(paramsObject, "questions")
            If questions Is Nothing OrElse questions.Count = 0 Then
                Await CurrentClient().SendErrorAsync(request.Id, -32602, "No questions were provided in item/tool/requestUserInput.")
                Return
            End If

            Dim answersRoot As New JsonObject()

            For Each questionNode In questions
                Dim questionObject = AsObject(questionNode)
                If questionObject Is Nothing Then
                    Continue For
                End If

                Dim questionId = GetPropertyString(questionObject, "id")
                If String.IsNullOrWhiteSpace(questionId) Then
                    Continue For
                End If

                Dim header = GetPropertyString(questionObject, "header", "Input required")
                Dim prompt = GetPropertyString(questionObject, "question", questionId)
                Dim isSecret = GetPropertyBoolean(questionObject, "isSecret", False)

                Dim options As New List(Of String)()
                Dim optionsArray = GetPropertyArray(questionObject, "options")
                If optionsArray IsNot Nothing Then
                    For Each optionNode In optionsArray
                        Dim optionObject = AsObject(optionNode)
                        If optionObject Is Nothing Then
                            Continue For
                        End If

                        Dim label = GetPropertyString(optionObject, "label")
                        Dim description = GetPropertyString(optionObject, "description")
                        If String.IsNullOrWhiteSpace(label) Then
                            Continue For
                        End If

                        If String.IsNullOrWhiteSpace(description) Then
                            options.Add(label)
                        Else
                            options.Add($"{label} - {description}")
                        End If
                    Next
                End If

                Dim answer As String = Nothing
                Dim promptDialog As New QuestionPromptWindow(header, prompt, options, isSecret) With {
                    .Owner = Me
                }
                Dim dialogResult = promptDialog.ShowDialog()
                If Not dialogResult.HasValue OrElse Not dialogResult.Value Then
                    Await CurrentClient().SendErrorAsync(request.Id, -32800, "User canceled request_user_input.")
                    Return
                End If

                answer = promptDialog.Answer

                If String.IsNullOrWhiteSpace(answer) Then
                    Await CurrentClient().SendErrorAsync(request.Id, -32602, $"No answer provided for question '{questionId}'.")
                    Return
                End If

                Dim answerObject As New JsonObject()
                Dim answerList As New JsonArray()
                answerList.Add(answer)
                answerObject("answers") = answerList
                answersRoot(questionId) = answerObject
            Next

            Dim response As New JsonObject()
            response("answers") = answersRoot

            Await CurrentClient().SendResultAsync(request.Id, response)
            AppendSystemMessage("Submitted request_user_input answers.")
            ShowStatus("Submitted request_user_input answers.", displayToast:=True)
        End Function

        Private Async Function HandleUnsupportedToolCallAsync(request As RpcServerRequest) As Task
            Dim response As New JsonObject()
            response("success") = False

            Dim contentItems As New JsonArray()
            Dim textItem As New JsonObject()
            textItem("type") = "inputText"
            textItem("text") = "This host does not expose dynamic tool callbacks yet."
            contentItems.Add(textItem)

            response("contentItems") = contentItems
            Await CurrentClient().SendResultAsync(request.Id, response)
            AppendSystemMessage("Declined dynamic tool call (unsupported in this host).")
            ShowStatus("Declined dynamic tool call.", isError:=True)
        End Function

    End Class
End Namespace
