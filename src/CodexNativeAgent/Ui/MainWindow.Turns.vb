Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports CodexNativeAgent.AppServer
Imports CodexNativeAgent.Services
Imports CodexNativeAgent.Ui.Coordinators

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
            Dim turnComposerVm = If(_viewModel?.TurnComposer, Nothing)

            If container Is Nothing OrElse toggleButton Is Nothing Then
                Return
            End If

            If Not _turnComposerPickersCollapsed Then
                CaptureTurnComposerPickersExpandedWidth()
            End If

            Dim targetExpanded = Not _turnComposerPickersCollapsed
            Dim expandedWidth = ResolveTurnComposerPickersExpandedWidth()
            Dim targetWidth = If(_turnComposerPickersCollapsed, 0.0R, expandedWidth)
            Dim targetOpacity = If(_turnComposerPickersCollapsed, 0.0R, 1.0R)
            Dim currentWidth = EffectiveAnimatedWidth(container, expandedWidth)
            Dim currentOpacity = If(Double.IsNaN(container.Opacity), If(currentWidth <= 0.5R, 0.0R, 1.0R), container.Opacity)
            Dim pickerTranslate = EnsureTranslateTransform(container)
            If turnComposerVm IsNot Nothing Then
                turnComposerVm.ArePickersExpanded = targetExpanded
            End If

            If _turnComposerPickersCollapsed AndAlso
               ((modelCombo IsNot Nothing AndAlso modelCombo.IsKeyboardFocusWithin) OrElse
                (reasoningCombo IsNot Nothing AndAlso reasoningCombo.IsKeyboardFocusWithin)) Then
                toggleButton.Focus()
            End If

            UpdateTurnComposerPickersToggleVisual()

            container.BeginAnimation(FrameworkElement.WidthProperty, Nothing)
            container.BeginAnimation(UIElement.OpacityProperty, Nothing)
            pickerTranslate.BeginAnimation(TranslateTransform.XProperty, Nothing)
            container.Visibility = Visibility.Visible
            container.IsHitTestVisible = targetExpanded

            If Not animated Then
                container.Width = targetWidth
                container.Opacity = targetOpacity
                pickerTranslate.X = If(targetExpanded, 0.0R, -14.0R)
                pickerTranslate.Y = 0.0R
                If persist Then
                    SaveSettings()
                End If

                Return
            End If

            ' Set the final base values first so they persist after the animation clock completes.
            container.Width = targetWidth
            container.Opacity = targetOpacity
            pickerTranslate.X = If(targetExpanded, 0.0R, -14.0R)
            pickerTranslate.Y = 0.0R

            Dim duration = TimeSpan.FromMilliseconds(210)
            Dim easing As New CubicEase() With {.EasingMode = EasingMode.EaseInOut}
            Dim fromSlideX = If(targetExpanded, -14.0R, 0.0R)
            Dim toSlideX = If(targetExpanded, 0.0R, -14.0R)

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
            Dim slideXAnimation As New DoubleAnimation() With {
                .From = fromSlideX,
                .To = toSlideX,
                .Duration = duration,
                .EasingFunction = easing,
                .FillBehavior = FillBehavior.Stop
            }

            container.BeginAnimation(FrameworkElement.WidthProperty, widthAnimation)
            container.BeginAnimation(UIElement.OpacityProperty, opacityAnimation)
            pickerTranslate.BeginAnimation(TranslateTransform.XProperty, slideXAnimation)

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

        Private Shared Function EnsureTranslateTransform(element As UIElement) As TranslateTransform
            If element Is Nothing Then
                Return New TranslateTransform()
            End If

            Dim existing = TryCast(element.RenderTransform, TranslateTransform)
            If existing IsNot Nothing Then
                Return existing
            End If

            existing = New TranslateTransform()
            element.RenderTransform = existing
            Return existing
        End Function

        Private Async Function StartTurnAsync() As Task
            Dim submittedInputText As String = String.Empty
            Dim shouldRefreshThreadsAfterTurnStart As Boolean = False
            Dim threadIdToRefreshAfterTurnStart As String = String.Empty

            Dim startedFromDraftNewThread = Await EnsureThreadReadyForTurnSubmissionAsync()
            Dim visibleThreadIdAtDispatch = GetVisibleThreadId()
            SyncTurnComposerStateForCurrentSelection()
            Dim composerState = ResolveTurnComposerStateForThread(visibleThreadIdAtDispatch, allowDraftWhenNoThread:=True)
            If String.IsNullOrWhiteSpace(composerState.InputText) Then
                Throw New InvalidOperationException("Enter turn input before sending.")
            End If

            Await _turnWorkflowCoordinator.RunStartTurnAsync(
                visibleThreadIdAtDispatch,
                composerState.InputText,
                composerState.ModelId,
                composerState.ReasoningEffort,
                composerState.ApprovalPolicy,
                AddressOf EnsureThreadSelected,
                Sub(inputText)
                    submittedInputText = If(inputText, String.Empty)
                    AppendTranscript("user", inputText)
                    TrackPendingUserEcho(inputText)
                End Sub,
                Sub(returnedTurnId)
                    If Not String.IsNullOrWhiteSpace(returnedTurnId) Then
                        SetVisibleTurnId(returnedTurnId)
                    End If

                    SyncCurrentTurnFromRuntimeStore(keepExistingWhenRuntimeIsIdle:=True)
                    Dim visibleThreadIdAfterStart = GetVisibleThreadId()
                    threadIdToRefreshAfterTurnStart = visibleThreadIdAfterStart
                    MarkThreadLastActive(visibleThreadIdAfterStart)
                    Dim threadMissingFromList = SyncThreadListAfterUserPrompt(visibleThreadIdAfterStart, submittedInputText)
                    shouldRefreshThreadsAfterTurnStart = startedFromDraftNewThread OrElse threadMissingFromList
                    FinalizePendingNewThreadFirstPromptSelection()
                    UpdateThreadTurnLabels()
                    RefreshControlStates()
                    WorkspacePaneHost.TxtTurnInput.Clear()
                    ShowStatus($"Turn started: {GetVisibleTurnId()}")
                End Sub)

            If shouldRefreshThreadsAfterTurnStart Then
                Await RefreshThreadsCoreAsync(silent:=True)
                If startedFromDraftNewThread Then
                    Await RetrySilentThreadRefreshUntilListedAsync(threadIdToRefreshAfterTurnStart)
                    If Not String.IsNullOrWhiteSpace(threadIdToRefreshAfterTurnStart) AndAlso
                       Not HasThreadEntry(threadIdToRefreshAfterTurnStart) Then
                        SyncThreadListAfterUserPrompt(threadIdToRefreshAfterTurnStart, submittedInputText)
                    End If
                End If
                FinalizePendingNewThreadFirstPromptSelection()
            End If
        End Function

        Private Async Function EnsureThreadReadyForTurnSubmissionAsync() As Task(Of Boolean)
            Dim startedFromDraftNewThread = _pendingNewThreadFirstPromptSelection

            If String.IsNullOrWhiteSpace(GetVisibleThreadId()) AndAlso
               Not startedFromDraftNewThread AndAlso
               IsPendingNewThreadTranscriptTabActive() Then
                SetPendingNewThreadFirstPromptSelectionActive(True, clearThreadSelection:=False)
                startedFromDraftNewThread = True
            End If

            If String.IsNullOrWhiteSpace(GetVisibleThreadId()) AndAlso Not _pendingNewThreadFirstPromptSelection Then
                Await StartThreadAsync()
            End If

            startedFromDraftNewThread = startedFromDraftNewThread OrElse _pendingNewThreadFirstPromptSelection

            If _pendingNewThreadFirstPromptSelection AndAlso String.IsNullOrWhiteSpace(GetVisibleThreadId()) Then
                Await EnsurePendingDraftThreadCreatedAsync()
                startedFromDraftNewThread = True
            End If

            Return startedFromDraftNewThread
        End Function

        Private Async Function SteerTurnAsync() As Task
            Dim visibleThreadId = GetVisibleThreadId()
            Dim visibleTurnId = GetVisibleTurnId()
            SyncTurnComposerStateForCurrentSelection()
            Dim composerState = ResolveTurnComposerStateForThread(visibleThreadId, allowDraftWhenNoThread:=True)
            Await _turnWorkflowCoordinator.RunSteerTurnAsync(
                visibleThreadId,
                visibleTurnId,
                composerState.InputText,
                AddressOf EnsureThreadSelected,
                Sub(steerText)
                    AppendTranscript("user (steer)", steerText)
                    TrackPendingUserEcho(steerText)
                End Sub,
                Sub(returnedTurnId)
                    If Not String.IsNullOrWhiteSpace(returnedTurnId) Then
                        SetVisibleTurnId(returnedTurnId)
                    End If

                    SyncCurrentTurnFromRuntimeStore(keepExistingWhenRuntimeIsIdle:=True)
                    MarkThreadLastActive(GetVisibleThreadId())
                    UpdateThreadTurnLabels()
                    RefreshControlStates()
                    WorkspacePaneHost.TxtTurnInput.Clear()
                    ShowStatus($"Turn steered: {GetVisibleTurnId()}")
                End Sub)
        End Function

        Private Async Function InterruptTurnAsync() As Task
            Dim visibleThreadId = GetVisibleThreadId()
            Dim visibleTurnId = GetVisibleTurnId()
            Await _turnWorkflowCoordinator.RunInterruptTurnAsync(
                visibleThreadId,
                visibleTurnId,
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
            ScrollTranscriptToBottom(reason:=TranscriptScrollRequestReason.UserMessage)
        End Sub

        Private Sub AppendSystemMessage(message As String)
            If String.IsNullOrWhiteSpace(message) Then
                Return
            End If

            _viewModel.TranscriptPanel.AppendSystemMessage(message)
            ScrollTranscriptToBottom(reason:=TranscriptScrollRequestReason.SystemMessage)
            ShowStatus(message)
        End Sub

        Private Sub AppendProtocol(direction As String, payload As String)
            Dim safePayload = If(payload, String.Empty)
            _viewModel.TranscriptPanel.AppendProtocolChunk($"[{Now:HH:mm:ss}] {direction}: {safePayload}{Environment.NewLine}")
            ScrollTextBoxToBottom(WorkspacePaneHost.TxtProtocol)
        End Sub

        Private Sub RenderItem(itemObject As JsonObject)
            If itemObject Is Nothing Then
                Return
            End If

            Dim itemId = GetPropertyString(itemObject, "id")
            Dim itemType = GetPropertyString(itemObject, "type")
            If String.IsNullOrWhiteSpace(itemId) OrElse String.IsNullOrWhiteSpace(itemType) Then
                Return
            End If

            Dim runtimeItem As New TurnItemRuntimeState() With {
                .ThreadId = GetVisibleThreadId(),
                .TurnId = GetVisibleTurnId(),
                .ItemId = itemId,
                .ItemType = itemType,
                .Status = "completed",
                .IsCompleted = True,
                .RawItemPayload = itemObject
            }

            RenderItem(runtimeItem)
        End Sub

        Private Sub RenderItem(itemState As TurnItemRuntimeState)
            If itemState Is Nothing Then
                Return
            End If

            If StringComparer.OrdinalIgnoreCase.Equals(itemState.ItemType, "userMessage") Then
                Dim suppressedUserItemKey = BuildSuppressedUserEchoItemKey(itemState)
                If Not String.IsNullOrWhiteSpace(suppressedUserItemKey) AndAlso
                   _suppressedServerUserEchoItemKeys.Contains(suppressedUserItemKey) Then
                    Return
                End If

                If ShouldSuppressUserEcho(itemState.GenericText) Then
                    If Not String.IsNullOrWhiteSpace(suppressedUserItemKey) Then
                        _suppressedServerUserEchoItemKeys.Add(suppressedUserItemKey)
                    End If
                    Return
                End If
            End If

            CacheAssistantPhaseHintFromRuntimeItem(itemState)
            _viewModel.TranscriptPanel.UpsertRuntimeItem(itemState)
        End Sub

        Private Sub AppendTurnLifecycleMarker(threadId As String, turnId As String, status As String)
            Dim lifecycleTimestampUtc As DateTimeOffset? = Nothing
            Dim runtimeStore = _sessionNotificationCoordinator?.RuntimeStore
            If runtimeStore IsNot Nothing Then
                Dim turnState = runtimeStore.GetTurnState(threadId, turnId)
                If turnState IsNot Nothing Then
                    Dim normalizedStatus = If(status, String.Empty).Trim()
                    If StringComparer.OrdinalIgnoreCase.Equals(normalizedStatus, "started") Then
                        lifecycleTimestampUtc = turnState.StartedAt
                    Else
                        lifecycleTimestampUtc = turnState.CompletedAt
                    End If
                End If
            End If

            _viewModel.TranscriptPanel.UpsertTurnLifecycleMarker(threadId, turnId, status, lifecycleTimestampUtc)
        End Sub

        Private Sub UpsertTurnMetadata(threadId As String, turnId As String, kind As String, summaryText As String)
            _viewModel.TranscriptPanel.UpsertTurnMetadataMarker(threadId, turnId, kind, summaryText)
        End Sub

        Private Sub UpdateTokenUsageWidget(threadId As String, turnId As String, tokenUsage As JsonObject)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            Dim normalizedTurnId = If(turnId, String.Empty).Trim()
            TraceContextUsageDebug("update_token_usage_widget",
                                   $"thread={normalizedThreadId}; turn={normalizedTurnId}; hasPayload={(tokenUsage IsNot Nothing).ToString()}; keys={DescribeJsonObjectKeys(tokenUsage)}")
            If Not String.IsNullOrWhiteSpace(normalizedThreadId) AndAlso tokenUsage IsNot Nothing Then
                CacheThreadContextUsageSnapshot(normalizedThreadId, normalizedTurnId, tokenUsage)
                TraceContextUsageDebug("context_cache_store",
                                       $"thread={normalizedThreadId}; turn={normalizedTurnId}; keys={DescribeJsonObjectKeys(tokenUsage)}")
            End If
            _viewModel.TranscriptPanel.SetTokenUsageSummary(threadId, turnId, tokenUsage)
            SyncTurnComposerContextUsageIndicator(turnId, tokenUsage)
        End Sub

        Private Sub SyncTurnComposerContextUsageIndicator(turnId As String, tokenUsage As JsonObject)
            If _viewModel Is Nothing OrElse _viewModel.TurnComposer Is Nothing Then
                Return
            End If

            Dim usedTokens As Long? = Nothing
            Dim maxTokens As Long? = Nothing
            Dim contextResolution As String = String.Empty
            Dim contextPercent = ResolveContextUsagePercent(tokenUsage,
                                                           usedTokens,
                                                           maxTokens,
                                                           contextResolution)
            If Not contextPercent.HasValue Then
                TraceContextUsageDebug("context_indicator_unresolved",
                                       $"turn={If(turnId, String.Empty).Trim()}; used={FormatNullableLong(usedTokens)}; max={FormatNullableLong(maxTokens)}; resolution={If(contextResolution, String.Empty)}; keys={DescribeJsonObjectKeys(tokenUsage)}")
                _viewModel.TurnComposer.SetContextUsageIndicator(Nothing, String.Empty)
                Return
            End If

            Dim normalizedPercent = ClampPercent(contextPercent.Value)
            Dim tooltipText = BuildContextUsageTooltipText(turnId,
                                                           normalizedPercent,
                                                           usedTokens,
                                                           maxTokens)
            TraceContextUsageDebug("context_indicator_applied",
                                   $"turn={If(turnId, String.Empty).Trim()}; percent={normalizedPercent.ToString("0.###", CultureInfo.InvariantCulture)}; used={FormatNullableLong(usedTokens)}; max={FormatNullableLong(maxTokens)}; resolution={If(contextResolution, String.Empty)}")
            _viewModel.TurnComposer.SetContextUsageIndicator(normalizedPercent, tooltipText)
        End Sub

        Private Shared Function BuildContextUsageTooltipText(turnId As String,
                                                             contextPercent As Double,
                                                             usedTokens As Long?,
                                                             maxTokens As Long?) As String
            Dim lines As New List(Of String)()
            lines.Add($"Thread context used: {contextPercent.ToString("0.#", CultureInfo.InvariantCulture)}%")

            If usedTokens.HasValue AndAlso maxTokens.HasValue AndAlso maxTokens.Value > 0 Then
                lines.Add($"Usage: {usedTokens.Value.ToString(CultureInfo.InvariantCulture)} / {maxTokens.Value.ToString(CultureInfo.InvariantCulture)} tokens")
            ElseIf maxTokens.HasValue AndAlso maxTokens.Value > 0 Then
                lines.Add($"Usage: -- / {maxTokens.Value.ToString(CultureInfo.InvariantCulture)} tokens")
            ElseIf usedTokens.HasValue Then
                lines.Add($"Usage: {usedTokens.Value.ToString(CultureInfo.InvariantCulture)} tokens")
            Else
                lines.Add("Usage: unavailable")
            End If

            Return String.Join(Environment.NewLine, lines)
        End Function

        Private Shared Function ResolveContextUsagePercent(tokenUsage As JsonObject,
                                                           ByRef usedTokens As Long?,
                                                           ByRef maxTokens As Long?,
                                                           ByRef resolution As String) As Double?
            usedTokens = Nothing
            maxTokens = Nothing
            resolution = String.Empty

            If tokenUsage Is Nothing Then
                resolution = "token_usage_payload_missing"
                Return Nothing
            End If

            Dim modelContextWindow As Long
            If TryReadJsonLong(tokenUsage,
                               modelContextWindow,
                               "modelContextWindow",
                               "model_context_window",
                               "contextWindowTokens",
                               "context_window_tokens",
                               "maxContextTokens",
                               "max_context_tokens") Then
                If modelContextWindow > 0 Then
                    maxTokens = modelContextWindow
                    resolution = $"window_from_payload={modelContextWindow.ToString(CultureInfo.InvariantCulture)}"
                End If
            End If

            Dim usageSummaryCandidates = {
                New KeyValuePair(Of String, JsonObject)("last", ReadNestedObject(tokenUsage, "last")),
                New KeyValuePair(Of String, JsonObject)("lastUsage", ReadNestedObject(tokenUsage, "lastUsage")),
                New KeyValuePair(Of String, JsonObject)("last_usage", ReadNestedObject(tokenUsage, "last_usage")),
                New KeyValuePair(Of String, JsonObject)("lastTokenUsage", ReadNestedObject(tokenUsage, "lastTokenUsage")),
                New KeyValuePair(Of String, JsonObject)("last_token_usage", ReadNestedObject(tokenUsage, "last_token_usage")),
                New KeyValuePair(Of String, JsonObject)("total", ReadNestedObject(tokenUsage, "total")),
                New KeyValuePair(Of String, JsonObject)("totalUsage", ReadNestedObject(tokenUsage, "totalUsage")),
                New KeyValuePair(Of String, JsonObject)("total_usage", ReadNestedObject(tokenUsage, "total_usage")),
                New KeyValuePair(Of String, JsonObject)("totalTokenUsage", ReadNestedObject(tokenUsage, "totalTokenUsage")),
                New KeyValuePair(Of String, JsonObject)("total_token_usage", ReadNestedObject(tokenUsage, "total_token_usage"))
            }

            Dim directPercent As Double
            If TryReadJsonDouble(tokenUsage,
                                 directPercent,
                                 "contextPercent",
                                 "context_percent",
                                 "contextUsagePercent",
                                 "context_usage_percent",
                                 "contextUsagePercentage",
                                 "context_usage_percentage",
                                 "contextUsedPercent",
                                 "context_used_percent",
                                 "contextWindowPercent",
                                 "context_window_percent",
                                 "contextWindowUsedPercent",
                                 "context_window_used_percent") Then
                resolution = "direct_context_percent_field"
                Return ClampPercent(directPercent)
            End If

            If TryReadJsonDouble(tokenUsage,
                                 directPercent,
                                 "contextRemainingPercent",
                                 "context_remaining_percent",
                                 "contextRemainingPercentage",
                                 "context_remaining_percentage") Then
                resolution = "direct_context_remaining_percent_field"
                Return ClampPercent(100.0R - directPercent)
            End If

            Dim usedForContext As Long
            Dim usedResolved As Boolean = False
            Dim resolvedUsageSummaryLabel As String = String.Empty
            For Each usageSummaryCandidate In usageSummaryCandidates
                Dim usageSummary = usageSummaryCandidate.Value
                If usageSummary Is Nothing Then
                    Continue For
                End If

                If ResolveUsedTokensForContext(usageSummary, usedForContext) Then
                    usedTokens = usedForContext
                    usedResolved = True
                    resolvedUsageSummaryLabel = usageSummaryCandidate.Key
                    Exit For
                End If
            Next

            If Not usedResolved Then
                If ResolveUsedTokensForContext(tokenUsage, usedForContext) Then
                    usedTokens = usedForContext
                    usedResolved = True
                    resolvedUsageSummaryLabel = "payload_root"
                End If
            End If

            If usedResolved AndAlso maxTokens.HasValue AndAlso maxTokens.Value > 0 Then
                usedTokens = usedForContext
                Dim ratio = CDbl(usedForContext) / CDbl(maxTokens.Value)
                resolution = $"ratio_from_{resolvedUsageSummaryLabel}; used={usedForContext.ToString(CultureInfo.InvariantCulture)}; max={maxTokens.Value.ToString(CultureInfo.InvariantCulture)}"
                Return ClampPercent(ratio * 100.0R)
            End If

            If usedResolved AndAlso Not maxTokens.HasValue Then
                usedTokens = usedForContext
                resolution = $"used_tokens_resolved_but_window_missing source={resolvedUsageSummaryLabel}"
            ElseIf Not usedResolved AndAlso maxTokens.HasValue Then
                resolution = "window_resolved_but_used_tokens_missing"
            Else
                resolution = "context_usage_fields_not_found"
            End If

            Return Nothing
        End Function

        Private Shared Function ResolveUsedTokensForContext(source As JsonObject, ByRef usedTokens As Long) As Boolean
            usedTokens = 0
            If source Is Nothing Then
                Return False
            End If

            Dim parsedContext As Long
            If TryReadJsonLong(source,
                               parsedContext,
                               "contextUsedTokens",
                               "context_used_tokens",
                               "usedContextTokens",
                               "used_context_tokens",
                               "usedTokens",
                               "used_tokens",
                               "contextTokens",
                                "context_tokens") Then
                usedTokens = Math.Max(0L, parsedContext)
                Return True
            End If

            Dim parsedInput As Long
            Dim parsedOutput As Long
            Dim hasInput = TryReadJsonLong(source, parsedInput, "inputTokens", "input_tokens")
            Dim hasOutput = TryReadJsonLong(source, parsedOutput, "outputTokens", "output_tokens")

            If hasInput OrElse hasOutput Then
                usedTokens = Math.Max(0L, parsedInput) + Math.Max(0L, parsedOutput)
                Return True
            End If

            Dim parsedTotal As Long
            Dim hasTotal = TryReadJsonLong(source,
                                           parsedTotal,
                                           "totalTokens",
                                           "total_tokens",
                                           "total")

            Dim parsedReasoning As Long
            Dim hasReasoning = TryReadJsonLong(source,
                                               parsedReasoning,
                                               "reasoningTokens",
                                               "reasoning_tokens",
                                               "reasoningOutputTokens",
                                               "reasoning_output_tokens")

            If hasTotal Then
                If hasReasoning Then
                    usedTokens = Math.Max(0L, parsedTotal - parsedReasoning)
                Else
                    usedTokens = Math.Max(0L, parsedTotal)
                End If
                Return True
            End If

            Return False
        End Function

        Private Sub TraceContextUsageDebug(eventName As String, details As String)
            AppendProtocol("debug",
                           $"context_usage_debug event={If(eventName, String.Empty)} {If(details, String.Empty)}")
        End Sub

        Private Shared Function DescribeJsonObjectKeys(obj As JsonObject, Optional maxKeys As Integer = 24) As String
            If obj Is Nothing Then
                Return "<null>"
            End If

            Dim keys As New List(Of String)()
            For Each kvp In obj
                If keys.Count >= maxKeys Then
                    Exit For
                End If

                keys.Add(If(kvp.Key, String.Empty))
            Next

            If obj.Count > keys.Count Then
                keys.Add($"...+{(obj.Count - keys.Count).ToString(CultureInfo.InvariantCulture)}")
            End If

            If keys.Count = 0 Then
                Return "<empty>"
            End If

            Return String.Join(",", keys)
        End Function

        Private Shared Function FormatNullableLong(value As Long?) As String
            If value.HasValue Then
                Return value.Value.ToString(CultureInfo.InvariantCulture)
            End If

            Return "<none>"
        End Function

        Private Shared Function ReadNestedObject(source As JsonObject, propertyName As String) As JsonObject
            If source Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not source.TryGetPropertyValue(propertyName, node) OrElse node Is Nothing Then
                Return Nothing
            End If

            Return TryCast(node, JsonObject)
        End Function

        Private Sub AppendRuntimeDiagnosticEvent(message As String)
            _viewModel.TranscriptPanel.AppendRuntimeDiagnosticEvent(message)
        End Sub

        Private Sub UpdateThreadTurnLabels()
            Dim visibleThreadId = GetVisibleThreadId()
            Dim visibleTurnId = GetVisibleTurnId()
            Dim showDraftNewThreadTitle = _pendingNewThreadFirstPromptSelection OrElse String.IsNullOrWhiteSpace(visibleThreadId)
            _viewModel.CurrentThreadText = If(showDraftNewThreadTitle,
                                              "New thread",
                                              ResolveThreadTitleForUi(visibleThreadId, 32))

            _viewModel.CurrentTurnText = If(String.IsNullOrWhiteSpace(visibleTurnId),
                                            "Turn: 0",
                                            $"Turn: {visibleTurnId}")
            RefreshActiveTranscriptTabCaption()
            SyncSessionStateViewModel()
        End Sub

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

        Private Sub ClearPendingUserEchoTracking()
            _pendingLocalUserEchoes.Clear()
            _suppressedServerUserEchoItemKeys.Clear()
        End Sub

        Private Shared Function BuildSuppressedUserEchoItemKey(itemState As TurnItemRuntimeState) As String
            If itemState Is Nothing Then
                Return String.Empty
            End If

            Dim scoped = If(itemState.ScopedItemKey, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(scoped) Then
                Return scoped
            End If

            Dim itemId = If(itemState.ItemId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(itemId) Then
                Return String.Empty
            End If

            Dim threadId = If(itemState.ThreadId, String.Empty).Trim()
            Dim turnId = If(itemState.TurnId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(threadId) OrElse String.IsNullOrWhiteSpace(turnId) Then
                Return itemId
            End If

            Return $"{threadId}:{turnId}:{itemId}"
        End Function

        Private Shared Function NormalizeUserEchoText(text As String) As String
            If text Is Nothing Then
                Return String.Empty
            End If

            Dim normalized = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim lines = normalized.Split({vbLf}, StringSplitOptions.None)
            Dim keptLines As New List(Of String)()

            For Each rawLine In lines
                Dim trimmedLine = If(rawLine, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(trimmedLine) Then
                    If keptLines.Count > 0 AndAlso keptLines(keptLines.Count - 1).Length > 0 Then
                        keptLines.Add(String.Empty)
                    End If
                    Continue For
                End If

                If IsSyntheticUserEchoMetadataLine(trimmedLine) Then
                    Continue For
                End If

                keptLines.Add(CollapseWhitespaceForUserEcho(trimmedLine))
            Next

            While keptLines.Count > 0 AndAlso keptLines(0).Length = 0
                keptLines.RemoveAt(0)
            End While

            While keptLines.Count > 0 AndAlso keptLines(keptLines.Count - 1).Length = 0
                keptLines.RemoveAt(keptLines.Count - 1)
            End While

            Return String.Join(vbLf, keptLines).Trim()
        End Function

        Private Shared Function IsSyntheticUserEchoMetadataLine(line As String) As Boolean
            Dim normalized = If(line, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            Dim lower = normalized.ToLowerInvariant()
            Return lower.StartsWith("[mention] ", StringComparison.Ordinal) OrElse
                   lower.StartsWith("[skill] ", StringComparison.Ordinal) OrElse
                   lower.StartsWith("[image] ", StringComparison.Ordinal) OrElse
                   lower.StartsWith("[localimage] ", StringComparison.Ordinal)
        End Function

        Private Shared Function CollapseWhitespaceForUserEcho(value As String) As String
            Dim source = If(value, String.Empty)
            If String.IsNullOrWhiteSpace(source) Then
                Return String.Empty
            End If

            Dim builder As New System.Text.StringBuilder(source.Length)
            Dim previousWasWhitespace = False

            For Each ch In source
                If Char.IsWhiteSpace(ch) Then
                    If Not previousWasWhitespace Then
                        builder.Append(" "c)
                        previousWasWhitespace = True
                    End If
                Else
                    builder.Append(ch)
                    previousWasWhitespace = False
                End If
            Next

            Return builder.ToString().Trim()
        End Function

        Private Async Function HandleServerRequestAsync(request As RpcServerRequest) As Task
            Dim dispatch = _sessionNotificationCoordinator.DispatchServerRequest(request)
            ApplyServerRequestDispatchResult(dispatch)
            SyncCurrentTurnFromRuntimeStore(keepExistingWhenRuntimeIsIdle:=True)

            Await _turnWorkflowCoordinator.HandleServerRequestAsync(
                request,
                AddressOf HandleToolRequestUserInputAsync,
                AddressOf HandleUnsupportedToolCallAsync,
                AddressOf HandleChatgptTokenRefreshAsync,
                Async Function(serverRequest, code, message)
                    Await CurrentClient().SendErrorAsync(serverRequest.Id, code, message)
                End Function,
                AddressOf GetVisibleThreadId,
                AddressOf RefreshControlStates,
                AddressOf RefreshThreadRuntimeIndicatorsIfNeeded,
                AddressOf AppendSystemMessage,
                Sub(message, isError, displayToast)
                    ShowStatus(message, isError:=isError, displayToast:=displayToast)
                End Sub)
        End Function

        Private Async Function ResolveApprovalAsync(action As String) As Task
            Await _turnWorkflowCoordinator.ResolveApprovalAsync(
                action,
                GetVisibleThreadId(),
                Async Function(requestId, resultNode)
                    Await CurrentClient().SendResultAsync(requestId, resultNode)
                End Function,
                AddressOf RefreshControlStates,
                AddressOf RefreshThreadRuntimeIndicatorsIfNeeded,
                AddressOf AppendSystemMessage,
                Sub(message, isError, displayToast)
                    ShowStatus(message, isError:=isError, displayToast:=displayToast)
                End Sub)
        End Function

        Private Async Function HandleToolRequestUserInputAsync(request As RpcServerRequest) As Task
            Dim paramsObject = AsObject(request.ParamsNode)
            Dim questions = GetPropertyArray(paramsObject, "questions")
            If questions Is Nothing OrElse questions.Count = 0 Then
                Await CurrentClient().SendErrorAsync(request.Id, -32602, $"No questions were provided in {ToolRequestUserInputMethod}.")
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
