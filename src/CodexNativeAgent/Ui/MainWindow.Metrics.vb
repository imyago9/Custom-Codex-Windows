Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Const MetricsCalendarVisibleWeeks As Integer = 6
        Private Const MetricsProjectRowLimit As Integer = 80
        Private Const MetricsUnknownProjectLabel As String = "(Unknown Project)"
        Private Const MetricsDefaultLastDaysFilter As Integer = 30

        Private Enum MetricsFilterMode
            AllActivity = 0
            SingleDay = 1
            LastNDays = 2
        End Enum

        Private NotInheritable Class MetricsSnapshotFilterOptions
            Public Property CalendarAnchorMonth As Date = New Date(Date.Today.Year, Date.Today.Month, 1)
            Public Property FilterMode As MetricsFilterMode = MetricsFilterMode.AllActivity
            Public Property SelectedDay As Date?
            Public Property LastNDays As Integer = MetricsDefaultLastDaysFilter
        End Class

        Private _metricsCalendarAnchorMonth As Date = New Date(Date.Today.Year, Date.Today.Month, 1)
        Private _metricsFilterMode As MetricsFilterMode = MetricsFilterMode.AllActivity
        Private _metricsFilterSelectedDay As Date = Date.Today
        Private _metricsFilterLastNDays As Integer = MetricsDefaultLastDaysFilter
        Private _metricsFilterUiHandlersBound As Boolean

        Private NotInheritable Class MetricsPanelSnapshot
            Public Property CodexHome As String = String.Empty
            Public Property ActivityStartUtc As DateTimeOffset?
            Public Property ActivityEndUtc As DateTimeOffset?
            Public Property SessionCount As Integer
            Public Property PromptCount As Integer
            Public Property ProjectCount As Integer
            Public Property InputTokens As Long
            Public Property OutputTokens As Long
            Public Property ToolCallCount As Integer
            Public Property HistoryRecordCount As Integer
            Public Property SessionFileCount As Integer
            Public Property ArchivedSessionFileCount As Integer
            Public Property ParseErrorCount As Integer
            Public Property CalendarMonthText As String = String.Empty
            Public Property FilterSummaryText As String = String.Empty
            Public Property CalendarDays As New List(Of MetricsCalendarDayCell)()
            Public Property ProjectRows As New List(Of MetricsProjectRow)()
            Public Property ModelRows As New List(Of MetricsModelRow)()
            Public Property ErrorMessage As String = String.Empty
        End Class

        Private NotInheritable Class MetricsCalendarDayCell
            Public Property DayNumberText As String = String.Empty
            Public Property TooltipTitle As String = String.Empty
            Public Property TooltipBody As String = String.Empty
            Public Property IntensityLevel As Integer
            Public Property IsOutsideCurrentMonth As Boolean
            Public Property IsToday As Boolean
            Public Property ActivityDotOpacity As Double
        End Class

        Private NotInheritable Class MetricsProjectRow
            Public Property ProjectText As String = String.Empty
            Public Property SessionCountText As String = "0"
            Public Property PromptCountText As String = "0"
            Public Property InputTokensText As String = "0"
            Public Property OutputTokensText As String = "0"
            Public Property ToolCallCountText As String = "0"
            Public Property LastActiveText As String = "n/a"
        End Class

        Private NotInheritable Class MetricsModelRow
            Public Property ModelText As String = String.Empty
            Public Property TurnCountText As String = "0"
            Public Property SessionCountText As String = "0"
            Public Property InputTokensText As String = "0"
            Public Property OutputTokensText As String = "0"
            Public Property ToolCallCountText As String = "0"
            Public Property TurnShareText As String = "0%"
        End Class

        Private NotInheritable Class MetricsPromptRecord
            Public Property SessionId As String = String.Empty
            Public Property TimestampUtc As DateTimeOffset?
            Public Property Text As String = String.Empty
        End Class

        Private NotInheritable Class MetricsSessionAccumulator
            Public Property SessionId As String = String.Empty
            Public Property ProjectPath As String = String.Empty
            Public Property PromptCount As Integer
            Public Property InputTokens As Long
            Public Property OutputTokens As Long
            Public Property ToolCallCount As Integer
            Public Property TurnContextCount As Integer
            Public Property FirstActivityUtc As DateTimeOffset?
            Public Property LastActivityUtc As DateTimeOffset?
            Public Property LastSeenTotalInputTokens As Long?
            Public Property LastSeenTotalOutputTokens As Long?
            Public Property HasSeenTotalTokenUsage As Boolean
            Public Property CurrentModel As String = String.Empty
        End Class

        Private NotInheritable Class MetricsDailyAccumulator
            Public ReadOnly SessionIds As New HashSet(Of String)(StringComparer.Ordinal)
            Public Property PromptCount As Integer
            Public Property InputTokens As Long
            Public Property OutputTokens As Long
            Public Property ToolCallCount As Integer
        End Class

        Private NotInheritable Class MetricsProjectAccumulator
            Public Property ProjectPath As String = String.Empty
            Public Property SessionCount As Integer
            Public Property PromptCount As Integer
            Public Property InputTokens As Long
            Public Property OutputTokens As Long
            Public Property ToolCallCount As Integer
            Public Property LastActiveUtc As DateTimeOffset?
        End Class

        Private NotInheritable Class MetricsModelAccumulator
            Public Property ModelId As String = String.Empty
            Public Property TurnCount As Integer
            Public ReadOnly SessionIds As New HashSet(Of String)(StringComparer.Ordinal)
            Public Property InputTokens As Long
            Public Property OutputTokens As Long
            Public Property ToolCallCount As Integer
        End Class

        Private NotInheritable Class MetricsParseAccumulator
            Public ReadOnly SessionsById As New Dictionary(Of String, MetricsSessionAccumulator)(StringComparer.Ordinal)
            Public ReadOnly DailyByDate As New Dictionary(Of Date, MetricsDailyAccumulator)()
            Public ReadOnly CalendarDailyByDate As New Dictionary(Of Date, MetricsDailyAccumulator)()
            Public ReadOnly ModelStatsById As New Dictionary(Of String, MetricsModelAccumulator)(StringComparer.OrdinalIgnoreCase)
            Public ReadOnly HistoryPrompts As New List(Of MetricsPromptRecord)()
            Public ReadOnly FallbackPrompts As New List(Of MetricsPromptRecord)()
            Public Property ActivityStartUtc As DateTimeOffset?
            Public Property ActivityEndUtc As DateTimeOffset?
            Public Property HistoryRecordCount As Integer
            Public Property SessionFileCount As Integer
            Public Property ArchivedSessionFileCount As Integer
            Public Property ParseErrorCount As Integer
        End Class

        Private Structure MetricsTokenDelta
            Public Property InputTokens As Long
            Public Property OutputTokens As Long
        End Structure

        Private Sub InitializeMetricsPanelUi()
            If MetricsPaneHost Is Nothing Then
                Return
            End If

            If Not _metricsFilterUiHandlersBound Then
                If MetricsPaneHost.BtnMetricsMonthPrev IsNot Nothing Then
                    AddHandler MetricsPaneHost.BtnMetricsMonthPrev.Click, Sub(sender, e) ShiftMetricsCalendarMonth(-1)
                End If
                If MetricsPaneHost.BtnMetricsMonthNext IsNot Nothing Then
                    AddHandler MetricsPaneHost.BtnMetricsMonthNext.Click, Sub(sender, e) ShiftMetricsCalendarMonth(1)
                End If
                If MetricsPaneHost.BtnMetricsMonthToday IsNot Nothing Then
                    AddHandler MetricsPaneHost.BtnMetricsMonthToday.Click, Sub(sender, e) JumpMetricsCalendarToToday()
                End If
                If MetricsPaneHost.BtnMetricsApplyFilter IsNot Nothing Then
                    AddHandler MetricsPaneHost.BtnMetricsApplyFilter.Click,
                        Sub(sender, e)
                            UpdateMetricsFiltersFromUi()
                            FireAndForget(RefreshMetricsPanelAsync())
                        End Sub
                End If
                If MetricsPaneHost.CmbMetricsFilterMode IsNot Nothing Then
                    AddHandler MetricsPaneHost.CmbMetricsFilterMode.SelectionChanged,
                        Sub(sender, e)
                            UpdateMetricsFiltersFromUi()
                            RefreshMetricsFilterInputState()
                        End Sub
                End If
                If MetricsPaneHost.DpMetricsFilterDay IsNot Nothing Then
                    AddHandler MetricsPaneHost.DpMetricsFilterDay.SelectedDateChanged,
                        Sub(sender, e)
                            UpdateMetricsFiltersFromUi()
                        End Sub
                End If
                If MetricsPaneHost.TxtMetricsFilterDays IsNot Nothing Then
                    AddHandler MetricsPaneHost.TxtMetricsFilterDays.LostFocus,
                        Sub(sender, e)
                            UpdateMetricsFiltersFromUi()
                        End Sub
                End If
                _metricsFilterUiHandlersBound = True
            End If

            If MetricsPaneHost.MetricsInspectorPanel IsNot Nothing Then
                MetricsPaneHost.MetricsInspectorPanel.Visibility = Visibility.Collapsed
            End If
            MetricsPaneHost.IsHitTestVisible = False

            If MetricsPaneHost.LstMetricsCalendarDays IsNot Nothing Then
                MetricsPaneHost.LstMetricsCalendarDays.ItemsSource = Nothing
            End If
            If MetricsPaneHost.LstMetricsProjects IsNot Nothing Then
                MetricsPaneHost.LstMetricsProjects.ItemsSource = Nothing
            End If
            If MetricsPaneHost.LstMetricsModels IsNot Nothing Then
                MetricsPaneHost.LstMetricsModels.ItemsSource = Nothing
            End If

            If MetricsPaneHost.LblMetricsPanelState IsNot Nothing Then
                MetricsPaneHost.LblMetricsPanelState.Text = "Open Metrics to inspect Codex usage."
            End If
            If MetricsPaneHost.LblMetricsCalendarMonth IsNot Nothing Then
                MetricsPaneHost.LblMetricsCalendarMonth.Text = _metricsCalendarAnchorMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture)
            End If
            If MetricsPaneHost.LblMetricsActivityRange IsNot Nothing Then
                MetricsPaneHost.LblMetricsActivityRange.Text = "No activity loaded."
            End If
            If MetricsPaneHost.LblMetricsDataSources IsNot Nothing Then
                MetricsPaneHost.LblMetricsDataSources.Text = String.Empty
            End If
            If MetricsPaneHost.LblMetricsFilterHint IsNot Nothing Then
                MetricsPaneHost.LblMetricsFilterHint.Text = $"Filter: {DescribeMetricsFilterSummary(BuildCurrentMetricsFilterOptions())}"
            End If

            If MetricsPaneHost.CmbMetricsFilterMode IsNot Nothing Then
                Dim selectedModeIndex = Math.Max(0, Math.Min(MetricsPaneHost.CmbMetricsFilterMode.Items.Count - 1, CInt(_metricsFilterMode)))
                MetricsPaneHost.CmbMetricsFilterMode.SelectedIndex = selectedModeIndex
            End If
            If MetricsPaneHost.DpMetricsFilterDay IsNot Nothing Then
                MetricsPaneHost.DpMetricsFilterDay.SelectedDate = _metricsFilterSelectedDay
            End If
            If MetricsPaneHost.TxtMetricsFilterDays IsNot Nothing Then
                MetricsPaneHost.TxtMetricsFilterDays.Text = _metricsFilterLastNDays.ToString(CultureInfo.InvariantCulture)
            End If
            RefreshMetricsFilterInputState()

            If MetricsPaneHost.LblMetricsSummarySessions IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummarySessions.Text = "0"
            End If
            If MetricsPaneHost.LblMetricsSummaryPrompts IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummaryPrompts.Text = "0"
            End If
            If MetricsPaneHost.LblMetricsSummaryProjects IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummaryProjects.Text = "0"
            End If
            If MetricsPaneHost.LblMetricsSummaryTokens IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummaryTokens.Text = "0 / 0"
            End If
            If MetricsPaneHost.LblMetricsSummaryToolCalls IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummaryToolCalls.Text = "0"
            End If

            SetMetricsPanelLoadingState(False, "Ready.")
        End Sub

        Private Sub ToggleMetricsPanel()
            If MetricsPaneHost Is Nothing OrElse MetricsPaneHost.MetricsInspectorPanel Is Nothing Then
                Return
            End If

            If IsMetricsPanelVisible() Then
                CloseMetricsPanel()
                Return
            End If

            HideGitPanelForPanelSwitch()
            ShowGitPanelDock()
            MetricsPaneHost.IsHitTestVisible = True
            MetricsPaneHost.MetricsInspectorPanel.Visibility = Visibility.Visible
            UpdateSidebarSelectionState(showSettings:=(_viewModel.SidebarSettingsViewVisibility = Visibility.Visible))
            FireAndForget(RefreshMetricsPanelAsync())
        End Sub

        Private Sub CloseMetricsPanel()
            If MetricsPaneHost Is Nothing OrElse MetricsPaneHost.MetricsInspectorPanel Is Nothing Then
                Return
            End If

            Dim actualWidth = MetricsPaneHost.MetricsInspectorPanel.ActualWidth
            If Not Double.IsNaN(actualWidth) AndAlso Not Double.IsInfinity(actualWidth) AndAlso actualWidth >= 280 Then
                _gitPanelDockWidth = actualWidth
            End If

            MetricsPaneHost.MetricsInspectorPanel.Visibility = Visibility.Collapsed
            MetricsPaneHost.IsHitTestVisible = False
            Interlocked.Increment(_metricsPanelLoadVersion)

            If Not IsGitPanelVisible() Then
                If RightGitPaneSplitter IsNot Nothing Then
                    RightGitPaneSplitter.Visibility = Visibility.Collapsed
                End If
                If RightGitPaneSplitterColumn IsNot Nothing Then
                    RightGitPaneSplitterColumn.Width = New GridLength(0, GridUnitType.Pixel)
                End If
                If RightGitPaneColumn IsNot Nothing Then
                    RightGitPaneColumn.MinWidth = 0
                    RightGitPaneColumn.Width = New GridLength(0, GridUnitType.Pixel)
                End If
                If RightGitPaneShell IsNot Nothing Then
                    RightGitPaneShell.Visibility = Visibility.Collapsed
                End If
            End If

            UpdateMainPaneResizeBounds()
            UpdateSidebarSelectionState(showSettings:=(_viewModel.SidebarSettingsViewVisibility = Visibility.Visible))
        End Sub

        Private Function IsMetricsPanelVisible() As Boolean
            Return MetricsPaneHost IsNot Nothing AndAlso
                   MetricsPaneHost.MetricsInspectorPanel IsNot Nothing AndAlso
                   MetricsPaneHost.MetricsInspectorPanel.Visibility = Visibility.Visible
        End Function

        Private Function IsGitPanelVisible() As Boolean
            Return GitPaneHost IsNot Nothing AndAlso
                   GitPaneHost.GitInspectorPanel IsNot Nothing AndAlso
                   GitPaneHost.GitInspectorPanel.Visibility = Visibility.Visible
        End Function

        Private Sub HideGitPanelForPanelSwitch()
            If Not IsGitPanelVisible() Then
                Return
            End If

            Dim actualWidth = GitPaneHost.GitInspectorPanel.ActualWidth
            If Not Double.IsNaN(actualWidth) AndAlso Not Double.IsInfinity(actualWidth) AndAlso actualWidth >= 280 Then
                _gitPanelDockWidth = actualWidth
            End If

            GitPaneHost.GitInspectorPanel.Visibility = Visibility.Collapsed
            Interlocked.Increment(_gitPanelLoadVersion)
            Interlocked.Increment(_gitPanelDiffPreviewLoadVersion)
            Interlocked.Increment(_gitPanelCommitPreviewLoadVersion)
            Interlocked.Increment(_gitPanelBranchPreviewLoadVersion)
        End Sub

        Private Sub HideMetricsPanelForPanelSwitch()
            If Not IsMetricsPanelVisible() Then
                Return
            End If

            Dim actualWidth = MetricsPaneHost.MetricsInspectorPanel.ActualWidth
            If Not Double.IsNaN(actualWidth) AndAlso Not Double.IsInfinity(actualWidth) AndAlso actualWidth >= 280 Then
                _gitPanelDockWidth = actualWidth
            End If

            MetricsPaneHost.MetricsInspectorPanel.Visibility = Visibility.Collapsed
            MetricsPaneHost.IsHitTestVisible = False
            Interlocked.Increment(_metricsPanelLoadVersion)
        End Sub

        Private Async Function RefreshMetricsPanelAsync() As Task
            If Not IsMetricsPanelVisible() Then
                Return
            End If

            UpdateMetricsFiltersFromUi()
            Dim filterOptions = BuildCurrentMetricsFilterOptions()
            Dim codexHome = ResolveMetricsCodexHome()
            Dim loadVersion = Interlocked.Increment(_metricsPanelLoadVersion)
            SetMetricsPanelLoadingState(True, $"Loading Codex usage metrics from {codexHome} ({DescribeMetricsFilterSummary(filterOptions)})...")

            Dim snapshot = Await Task.Run(Function() BuildMetricsPanelSnapshot(codexHome, filterOptions)).ConfigureAwait(True)

            If loadVersion <> _metricsPanelLoadVersion Then
                Return
            End If

            If Not IsMetricsPanelVisible() Then
                Return
            End If

            ApplyMetricsPanelSnapshot(snapshot)
        End Function

        Private Sub ShiftMetricsCalendarMonth(monthDelta As Integer)
            If monthDelta = 0 Then
                Return
            End If

            _metricsCalendarAnchorMonth = New Date(_metricsCalendarAnchorMonth.Year, _metricsCalendarAnchorMonth.Month, 1).AddMonths(monthDelta)
            If MetricsPaneHost IsNot Nothing AndAlso MetricsPaneHost.LblMetricsCalendarMonth IsNot Nothing Then
                MetricsPaneHost.LblMetricsCalendarMonth.Text = _metricsCalendarAnchorMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture)
            End If

            FireAndForget(RefreshMetricsPanelAsync())
        End Sub

        Private Sub JumpMetricsCalendarToToday()
            _metricsCalendarAnchorMonth = New Date(Date.Today.Year, Date.Today.Month, 1)
            If MetricsPaneHost IsNot Nothing AndAlso MetricsPaneHost.LblMetricsCalendarMonth IsNot Nothing Then
                MetricsPaneHost.LblMetricsCalendarMonth.Text = _metricsCalendarAnchorMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture)
            End If

            FireAndForget(RefreshMetricsPanelAsync())
        End Sub

        Private Sub UpdateMetricsFiltersFromUi()
            Dim selectedMode = _metricsFilterMode
            If MetricsPaneHost IsNot Nothing AndAlso MetricsPaneHost.CmbMetricsFilterMode IsNot Nothing Then
                selectedMode = ResolveMetricsFilterMode(MetricsPaneHost.CmbMetricsFilterMode.SelectedIndex)
            End If
            _metricsFilterMode = selectedMode

            If MetricsPaneHost IsNot Nothing AndAlso MetricsPaneHost.DpMetricsFilterDay IsNot Nothing AndAlso
               MetricsPaneHost.DpMetricsFilterDay.SelectedDate.HasValue Then
                _metricsFilterSelectedDay = MetricsPaneHost.DpMetricsFilterDay.SelectedDate.Value.Date
            End If

            If MetricsPaneHost IsNot Nothing AndAlso MetricsPaneHost.TxtMetricsFilterDays IsNot Nothing Then
                Dim parsedLastDays As Integer
                If Integer.TryParse(If(MetricsPaneHost.TxtMetricsFilterDays.Text, String.Empty).Trim(),
                                    NumberStyles.Integer,
                                    CultureInfo.InvariantCulture,
                                    parsedLastDays) Then
                    parsedLastDays = Math.Max(1, Math.Min(3650, parsedLastDays))
                    _metricsFilterLastNDays = parsedLastDays
                End If

                MetricsPaneHost.TxtMetricsFilterDays.Text = _metricsFilterLastNDays.ToString(CultureInfo.InvariantCulture)
            End If

            If _metricsFilterMode = MetricsFilterMode.SingleDay Then
                _metricsCalendarAnchorMonth = New Date(_metricsFilterSelectedDay.Year, _metricsFilterSelectedDay.Month, 1)
            End If

            RefreshMetricsFilterInputState()
        End Sub

        Private Sub RefreshMetricsFilterInputState()
            If MetricsPaneHost Is Nothing Then
                Return
            End If

            If MetricsPaneHost.DpMetricsFilterDay IsNot Nothing Then
                MetricsPaneHost.DpMetricsFilterDay.IsEnabled = (_metricsFilterMode = MetricsFilterMode.SingleDay)
            End If
            If MetricsPaneHost.TxtMetricsFilterDays IsNot Nothing Then
                MetricsPaneHost.TxtMetricsFilterDays.IsEnabled = (_metricsFilterMode = MetricsFilterMode.LastNDays)
            End If
            If MetricsPaneHost.LblMetricsFilterHint IsNot Nothing Then
                MetricsPaneHost.LblMetricsFilterHint.Text = $"Filter: {DescribeMetricsFilterSummary(BuildCurrentMetricsFilterOptions())}"
            End If
        End Sub

        Private Function BuildCurrentMetricsFilterOptions() As MetricsSnapshotFilterOptions
            Dim options As New MetricsSnapshotFilterOptions() With {
                .CalendarAnchorMonth = New Date(_metricsCalendarAnchorMonth.Year, _metricsCalendarAnchorMonth.Month, 1),
                .FilterMode = _metricsFilterMode,
                .LastNDays = Math.Max(1, _metricsFilterLastNDays)
            }

            If _metricsFilterMode = MetricsFilterMode.SingleDay Then
                options.SelectedDay = _metricsFilterSelectedDay.Date
            End If

            Return options
        End Function

        Private Shared Function ResolveMetricsFilterMode(selectedIndex As Integer) As MetricsFilterMode
            Select Case selectedIndex
                Case 1
                    Return MetricsFilterMode.SingleDay
                Case 2
                    Return MetricsFilterMode.LastNDays
                Case Else
                    Return MetricsFilterMode.AllActivity
            End Select
        End Function

        Private Shared Function DescribeMetricsFilterSummary(options As MetricsSnapshotFilterOptions) As String
            If options Is Nothing Then
                Return "All activity"
            End If

            Select Case options.FilterMode
                Case MetricsFilterMode.SingleDay
                    If options.SelectedDay.HasValue Then
                        Return $"Single day ({options.SelectedDay.Value:yyyy-MM-dd})"
                    End If

                    Return "Single day"

                Case MetricsFilterMode.LastNDays
                    Return $"Last {Math.Max(1, options.LastNDays).ToString(CultureInfo.InvariantCulture)} days"

                Case Else
                    Return "All activity"
            End Select
        End Function

        Private Function ResolveMetricsCodexHome() As String
            Dim configuredHome = If(_viewModel?.SettingsPanel?.WindowsCodexHome, String.Empty).Trim()
            If Not String.IsNullOrWhiteSpace(configuredHome) Then
                Return configuredHome
            End If

            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex")
        End Function

        Private Sub SetMetricsPanelLoadingState(isLoading As Boolean, statusText As String)
            If MetricsPaneHost Is Nothing Then
                Return
            End If

            If MetricsPaneHost.MetricsPanelLoadingOverlay IsNot Nothing Then
                MetricsPaneHost.MetricsPanelLoadingOverlay.Visibility = If(isLoading, Visibility.Visible, Visibility.Collapsed)
            End If

            If MetricsPaneHost.LblMetricsPanelLoading IsNot Nothing Then
                MetricsPaneHost.LblMetricsPanelLoading.Text = If(statusText, String.Empty)
            End If
        End Sub

        Private Sub ApplyMetricsPanelSnapshot(snapshot As MetricsPanelSnapshot)
            If MetricsPaneHost Is Nothing Then
                Return
            End If

            If snapshot Is Nothing Then
                SetMetricsPanelLoadingState(False, "No metrics loaded.")
                Return
            End If

            Dim panelStateText As String
            If Not String.IsNullOrWhiteSpace(snapshot.ErrorMessage) Then
                panelStateText = snapshot.ErrorMessage
            ElseIf snapshot.SessionCount = 0 AndAlso
                   snapshot.PromptCount = 0 AndAlso
                   snapshot.InputTokens = 0 AndAlso
                   snapshot.OutputTokens = 0 AndAlso
                   snapshot.ToolCallCount = 0 Then
                panelStateText = "No Codex activity records were found in this workspace."
            Else
                Dim avgPromptsPerSession As Double = 0
                If snapshot.SessionCount > 0 Then
                    avgPromptsPerSession = CDbl(snapshot.PromptCount) / CDbl(snapshot.SessionCount)
                End If

                panelStateText = $"Loaded {FormatMetricCount(snapshot.SessionCount)} sessions across {FormatMetricCount(snapshot.ProjectCount)} projects ({avgPromptsPerSession.ToString("0.0", CultureInfo.InvariantCulture)} prompts/session)."
                If snapshot.ParseErrorCount > 0 Then
                    panelStateText &= $" Skipped {FormatMetricCount(snapshot.ParseErrorCount)} malformed records."
                End If
            End If

            If MetricsPaneHost.LblMetricsPanelState IsNot Nothing Then
                MetricsPaneHost.LblMetricsPanelState.Text = panelStateText
            End If
            If MetricsPaneHost.LblMetricsSummarySessions IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummarySessions.Text = FormatMetricCount(snapshot.SessionCount)
            End If
            If MetricsPaneHost.LblMetricsSummaryPrompts IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummaryPrompts.Text = FormatMetricCount(snapshot.PromptCount)
            End If
            If MetricsPaneHost.LblMetricsSummaryProjects IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummaryProjects.Text = FormatMetricCount(snapshot.ProjectCount)
            End If
            If MetricsPaneHost.LblMetricsSummaryTokens IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummaryTokens.Text =
                    $"{FormatMetricCount(snapshot.InputTokens)} / {FormatMetricCount(snapshot.OutputTokens)}"
            End If
            If MetricsPaneHost.LblMetricsSummaryToolCalls IsNot Nothing Then
                MetricsPaneHost.LblMetricsSummaryToolCalls.Text = FormatMetricCount(snapshot.ToolCallCount)
            End If

            If MetricsPaneHost.LblMetricsActivityRange IsNot Nothing Then
                If snapshot.ActivityStartUtc.HasValue AndAlso snapshot.ActivityEndUtc.HasValue Then
                    MetricsPaneHost.LblMetricsActivityRange.Text =
                        $"{snapshot.ActivityStartUtc.Value.LocalDateTime:yyyy-MM-dd HH:mm} to {snapshot.ActivityEndUtc.Value.LocalDateTime:yyyy-MM-dd HH:mm} (local time)"
                Else
                    MetricsPaneHost.LblMetricsActivityRange.Text = "No activity timestamps found."
                End If
            End If

            If MetricsPaneHost.LblMetricsDataSources IsNot Nothing Then
                Dim sourceParts As New List(Of String) From {
                    $"Codex home: {snapshot.CodexHome}",
                    $"history entries: {FormatMetricCount(snapshot.HistoryRecordCount)}",
                    $"session files: {FormatMetricCount(snapshot.SessionFileCount)}",
                    $"archived files: {FormatMetricCount(snapshot.ArchivedSessionFileCount)}"
                }
                If snapshot.ParseErrorCount > 0 Then
                    sourceParts.Add($"parse issues: {FormatMetricCount(snapshot.ParseErrorCount)}")
                End If

                MetricsPaneHost.LblMetricsDataSources.Text = String.Join(" | ", sourceParts)
            End If

            If MetricsPaneHost.LblMetricsCalendarMonth IsNot Nothing Then
                MetricsPaneHost.LblMetricsCalendarMonth.Text = snapshot.CalendarMonthText
            End If
            If MetricsPaneHost.LblMetricsFilterHint IsNot Nothing Then
                MetricsPaneHost.LblMetricsFilterHint.Text = $"Filter: {snapshot.FilterSummaryText}"
            End If
            If MetricsPaneHost.LstMetricsCalendarDays IsNot Nothing Then
                MetricsPaneHost.LstMetricsCalendarDays.ItemsSource = snapshot.CalendarDays
            End If
            If MetricsPaneHost.LstMetricsProjects IsNot Nothing Then
                MetricsPaneHost.LstMetricsProjects.ItemsSource = snapshot.ProjectRows
            End If
            If MetricsPaneHost.LstMetricsModels IsNot Nothing Then
                MetricsPaneHost.LstMetricsModels.ItemsSource = snapshot.ModelRows
            End If

            SetMetricsPanelLoadingState(False, "Metrics loaded.")
        End Sub

        Private Shared Function BuildMetricsPanelSnapshot(codexHome As String,
                                                          filterOptions As MetricsSnapshotFilterOptions) As MetricsPanelSnapshot
            Dim resolvedFilterOptions = NormalizeMetricsFilterOptions(filterOptions)
            Dim snapshot As New MetricsPanelSnapshot() With {
                .CodexHome = If(codexHome, String.Empty).Trim()
            }
            snapshot.FilterSummaryText = DescribeMetricsFilterSummary(resolvedFilterOptions)

            If String.IsNullOrWhiteSpace(snapshot.CodexHome) Then
                snapshot.ErrorMessage = "Codex home is not configured. Set it in Settings."
                Return snapshot
            End If

            If Not Directory.Exists(snapshot.CodexHome) Then
                snapshot.ErrorMessage = $"Codex home does not exist: {snapshot.CodexHome}"
                Return snapshot
            End If

            Dim accumulator As New MetricsParseAccumulator()

            ParseMetricsHistoryFile(Path.Combine(snapshot.CodexHome, "history.jsonl"), accumulator, resolvedFilterOptions)
            ParseMetricsSessionDirectory(Path.Combine(snapshot.CodexHome, "sessions"), isArchived:=False, accumulator, resolvedFilterOptions)
            ParseMetricsSessionDirectory(Path.Combine(snapshot.CodexHome, "archived_sessions"), isArchived:=True, accumulator, resolvedFilterOptions)

            ApplyPromptRecordsToAccumulator(accumulator, resolvedFilterOptions)

            snapshot.ActivityStartUtc = accumulator.ActivityStartUtc
            snapshot.ActivityEndUtc = accumulator.ActivityEndUtc
            snapshot.HistoryRecordCount = accumulator.HistoryRecordCount
            snapshot.SessionFileCount = accumulator.SessionFileCount
            snapshot.ArchivedSessionFileCount = accumulator.ArchivedSessionFileCount
            snapshot.ParseErrorCount = accumulator.ParseErrorCount

            Dim allSessions = accumulator.SessionsById.Values.
                Where(Function(item) HasFilteredSessionActivity(item)).
                ToList()
            snapshot.SessionCount = allSessions.Count
            snapshot.PromptCount = allSessions.Sum(Function(item) item.PromptCount)
            snapshot.InputTokens = allSessions.Sum(Function(item) item.InputTokens)
            snapshot.OutputTokens = allSessions.Sum(Function(item) item.OutputTokens)
            snapshot.ToolCallCount = allSessions.Sum(Function(item) item.ToolCallCount)

            Dim projectsByPath As New Dictionary(Of String, MetricsProjectAccumulator)(StringComparer.OrdinalIgnoreCase)
            For Each session In allSessions
                Dim projectPath = NormalizeProjectPath(session.ProjectPath)
                If String.IsNullOrWhiteSpace(projectPath) Then
                    projectPath = MetricsUnknownProjectLabel
                End If

                Dim project As MetricsProjectAccumulator = Nothing
                If Not projectsByPath.TryGetValue(projectPath, project) Then
                    project = New MetricsProjectAccumulator() With {
                        .ProjectPath = projectPath
                    }
                    projectsByPath(projectPath) = project
                End If

                project.SessionCount += 1
                project.PromptCount += session.PromptCount
                project.InputTokens += session.InputTokens
                project.OutputTokens += session.OutputTokens
                project.ToolCallCount += session.ToolCallCount

                If session.LastActivityUtc.HasValue Then
                    If Not project.LastActiveUtc.HasValue OrElse session.LastActivityUtc.Value > project.LastActiveUtc.Value Then
                        project.LastActiveUtc = session.LastActivityUtc
                    End If
                End If
            Next
            snapshot.ProjectCount = projectsByPath.Count

            BuildMetricsCalendarDays(snapshot, accumulator.CalendarDailyByDate, resolvedFilterOptions.CalendarAnchorMonth)

            For Each projectEntry In projectsByPath.Values.
                OrderByDescending(Function(item) item.PromptCount).
                ThenByDescending(Function(item) If(item.LastActiveUtc.HasValue, item.LastActiveUtc.Value, DateTimeOffset.MinValue)).
                ThenBy(Function(item) item.ProjectPath, StringComparer.OrdinalIgnoreCase).
                Take(MetricsProjectRowLimit)
                snapshot.ProjectRows.Add(New MetricsProjectRow() With {
                    .ProjectText = BuildProjectDisplayText(projectEntry.ProjectPath),
                    .SessionCountText = FormatMetricCount(projectEntry.SessionCount),
                    .PromptCountText = FormatMetricCount(projectEntry.PromptCount),
                    .InputTokensText = FormatMetricCount(projectEntry.InputTokens),
                    .OutputTokensText = FormatMetricCount(projectEntry.OutputTokens),
                    .ToolCallCountText = FormatMetricCount(projectEntry.ToolCallCount),
                    .LastActiveText = If(projectEntry.LastActiveUtc.HasValue,
                                         projectEntry.LastActiveUtc.Value.LocalDateTime.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
                                         "n/a")
                })
            Next

            Dim totalModelTurns = accumulator.ModelStatsById.Values.Sum(Function(item) item.TurnCount)
            For Each modelEntry In accumulator.ModelStatsById.Values.
                OrderByDescending(Function(item) item.TurnCount).
                ThenBy(Function(item) item.ModelId, StringComparer.OrdinalIgnoreCase)
                Dim turnShare As Double = 0
                If totalModelTurns > 0 Then
                    turnShare = CDbl(modelEntry.TurnCount) / CDbl(totalModelTurns)
                End If

                snapshot.ModelRows.Add(New MetricsModelRow() With {
                    .ModelText = modelEntry.ModelId,
                    .TurnCountText = FormatMetricCount(modelEntry.TurnCount),
                    .SessionCountText = FormatMetricCount(modelEntry.SessionIds.Count),
                    .InputTokensText = FormatMetricCount(modelEntry.InputTokens),
                    .OutputTokensText = FormatMetricCount(modelEntry.OutputTokens),
                    .ToolCallCountText = FormatMetricCount(modelEntry.ToolCallCount),
                    .TurnShareText = turnShare.ToString("P1", CultureInfo.InvariantCulture)
                })
            Next

            Return snapshot
        End Function

        Private Shared Function NormalizeMetricsFilterOptions(filterOptions As MetricsSnapshotFilterOptions) As MetricsSnapshotFilterOptions
            Dim normalized As New MetricsSnapshotFilterOptions()

            If filterOptions Is Nothing Then
                Return normalized
            End If

            Dim anchor = filterOptions.CalendarAnchorMonth
            If anchor.Year < Date.MinValue.Year OrElse anchor.Year > Date.MaxValue.Year Then
                anchor = Date.Today
            End If
            normalized.CalendarAnchorMonth = New Date(anchor.Year, anchor.Month, 1)

            normalized.FilterMode = filterOptions.FilterMode
            normalized.LastNDays = Math.Max(1, filterOptions.LastNDays)
            If filterOptions.SelectedDay.HasValue Then
                normalized.SelectedDay = filterOptions.SelectedDay.Value.Date
            End If

            If normalized.FilterMode = MetricsFilterMode.SingleDay AndAlso Not normalized.SelectedDay.HasValue Then
                normalized.SelectedDay = Date.Today
            End If

            Return normalized
        End Function

        Private Shared Function ShouldIncludeMetricsRecord(hasTimestamp As Boolean,
                                                           timestampUtc As DateTimeOffset,
                                                           filterOptions As MetricsSnapshotFilterOptions) As Boolean
            Dim activeFilter = If(filterOptions, New MetricsSnapshotFilterOptions())
            Select Case activeFilter.FilterMode
                Case MetricsFilterMode.AllActivity
                    Return True

                Case MetricsFilterMode.SingleDay
                    If Not hasTimestamp OrElse Not activeFilter.SelectedDay.HasValue Then
                        Return False
                    End If

                    Return timestampUtc.LocalDateTime.Date = activeFilter.SelectedDay.Value

                Case MetricsFilterMode.LastNDays
                    If Not hasTimestamp Then
                        Return False
                    End If

                    Dim localDate = timestampUtc.LocalDateTime.Date
                    Dim rangeEnd = Date.Today
                    Dim rangeStart = rangeEnd.AddDays(-Math.Max(1, activeFilter.LastNDays) + 1)
                    Return localDate >= rangeStart AndAlso localDate <= rangeEnd

                Case Else
                    Return True
            End Select
        End Function

        Private Shared Function HasFilteredSessionActivity(session As MetricsSessionAccumulator) As Boolean
            If session Is Nothing Then
                Return False
            End If

            Return session.PromptCount > 0 OrElse
                   session.InputTokens > 0 OrElse
                   session.OutputTokens > 0 OrElse
                   session.ToolCallCount > 0 OrElse
                   session.TurnContextCount > 0
        End Function

        Private Shared Sub BuildMetricsCalendarDays(snapshot As MetricsPanelSnapshot,
                                                    dailyByDate As Dictionary(Of Date, MetricsDailyAccumulator),
                                                    anchorMonth As Date)
            If snapshot Is Nothing Then
                Return
            End If

            Dim monthStart As New Date(anchorMonth.Year, anchorMonth.Month, 1)
            snapshot.CalendarMonthText = monthStart.ToString("MMMM yyyy", CultureInfo.InvariantCulture)
            snapshot.CalendarDays.Clear()

            Dim firstCellDate = monthStart.AddDays(-CInt(monthStart.DayOfWeek))
            Dim maxScore As Double = 0

            Dim daysInMonth = Date.DaysInMonth(monthStart.Year, monthStart.Month)
            For offset = 0 To daysInMonth - 1
                Dim day = monthStart.AddDays(offset)
                Dim dayMetrics As MetricsDailyAccumulator = Nothing
                If dailyByDate IsNot Nothing AndAlso dailyByDate.TryGetValue(day, dayMetrics) Then
                    maxScore = Math.Max(maxScore, ComputeDailyActivityScore(dayMetrics))
                End If
            Next

            If maxScore <= 0 Then
                maxScore = 1
            End If

            Dim totalCells = MetricsCalendarVisibleWeeks * 7
            For cellIndex = 0 To totalCells - 1
                Dim cellDate = firstCellDate.AddDays(cellIndex)
                Dim dayMetrics As MetricsDailyAccumulator = Nothing
                Dim hasMetrics = dailyByDate IsNot Nothing AndAlso dailyByDate.TryGetValue(cellDate, dayMetrics)
                Dim activityScore = If(hasMetrics, ComputeDailyActivityScore(dayMetrics), 0)
                Dim intensityLevel = ResolveCalendarIntensityLevel(activityScore, maxScore)
                Dim isOutsideCurrentMonth = cellDate.Month <> monthStart.Month OrElse cellDate.Year <> monthStart.Year
                Dim hasActivity = hasMetrics AndAlso activityScore > 0

                snapshot.CalendarDays.Add(New MetricsCalendarDayCell() With {
                    .DayNumberText = cellDate.Day.ToString(CultureInfo.InvariantCulture),
                    .TooltipTitle = cellDate.ToString("dddd, MMM d, yyyy", CultureInfo.InvariantCulture),
                    .TooltipBody = BuildCalendarTooltipBody(dayMetrics, isOutsideCurrentMonth),
                    .IntensityLevel = intensityLevel,
                    .IsOutsideCurrentMonth = isOutsideCurrentMonth,
                    .IsToday = cellDate = Date.Today,
                    .ActivityDotOpacity = If(hasActivity AndAlso Not isOutsideCurrentMonth, 1, 0)
                })
            Next
        End Sub

        Private Shared Function ComputeDailyActivityScore(dayMetrics As MetricsDailyAccumulator) As Double
            If dayMetrics Is Nothing Then
                Return 0
            End If

            Dim tokenWeight = CDbl(dayMetrics.InputTokens + dayMetrics.OutputTokens) / 2000.0R
            Dim interactionWeight = CDbl(dayMetrics.PromptCount * 3) + CDbl(dayMetrics.ToolCallCount * 2)
            Dim sessionWeight = CDbl(dayMetrics.SessionIds.Count)
            Return Math.Max(0, interactionWeight + sessionWeight + tokenWeight)
        End Function

        Private Shared Function ResolveCalendarIntensityLevel(score As Double, maxScore As Double) As Integer
            If score <= 0 OrElse maxScore <= 0 Then
                Return 0
            End If

            Dim ratio = score / maxScore
            If ratio < 0.25R Then
                Return 1
            End If
            If ratio < 0.5R Then
                Return 2
            End If
            If ratio < 0.75R Then
                Return 3
            End If

            Return 4
        End Function

        Private Shared Function BuildCalendarTooltipBody(dayMetrics As MetricsDailyAccumulator,
                                                         isOutsideCurrentMonth As Boolean) As String
            If dayMetrics Is Nothing OrElse
               (dayMetrics.PromptCount <= 0 AndAlso
                dayMetrics.ToolCallCount <= 0 AndAlso
                dayMetrics.InputTokens <= 0 AndAlso
                dayMetrics.OutputTokens <= 0 AndAlso
                dayMetrics.SessionIds.Count <= 0) Then
                Return If(isOutsideCurrentMonth,
                          "Outside displayed month" & Environment.NewLine & "No recorded activity.",
                          "No recorded activity.")
            End If

            Dim lines As New List(Of String) From {
                $"Sessions: {FormatMetricCount(dayMetrics.SessionIds.Count)}",
                $"Prompts: {FormatMetricCount(dayMetrics.PromptCount)}",
                $"Input tokens: {FormatMetricCount(dayMetrics.InputTokens)}",
                $"Output tokens: {FormatMetricCount(dayMetrics.OutputTokens)}",
                $"Tool calls: {FormatMetricCount(dayMetrics.ToolCallCount)}"
            }

            If isOutsideCurrentMonth Then
                lines.Insert(0, "Outside displayed month")
            End If

            Return String.Join(Environment.NewLine, lines)
        End Function

        Private Shared Sub ParseMetricsHistoryFile(historyPath As String,
                                                   accumulator As MetricsParseAccumulator,
                                                   filterOptions As MetricsSnapshotFilterOptions)
            If accumulator Is Nothing OrElse String.IsNullOrWhiteSpace(historyPath) OrElse Not File.Exists(historyPath) Then
                Return
            End If

            Dim lines As IEnumerable(Of String)
            Try
                lines = File.ReadLines(historyPath)
            Catch
                accumulator.ParseErrorCount += 1
                Return
            End Try

            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(line) Then
                    Continue For
                End If

                Try
                    Using document = JsonDocument.Parse(line)
                        Dim root = document.RootElement
                        Dim sessionId = GetJsonString(root, "session_id").Trim()
                        If String.IsNullOrWhiteSpace(sessionId) Then
                            Continue For
                        End If

                        Dim timestampValue As JsonElement
                        Dim timestampUtc As DateTimeOffset?
                        If root.TryGetProperty("ts", timestampValue) Then
                            Dim parsedTimestamp As DateTimeOffset
                            If TryParseTimestampValue(timestampValue, parsedTimestamp) Then
                                timestampUtc = parsedTimestamp
                            End If
                        End If

                        accumulator.HistoryPrompts.Add(New MetricsPromptRecord() With {
                            .SessionId = sessionId,
                            .TimestampUtc = timestampUtc,
                            .Text = GetJsonString(root, "text")
                        })
                        accumulator.HistoryRecordCount += 1
                    End Using
                Catch
                    accumulator.ParseErrorCount += 1
                End Try
            Next
        End Sub

        Private Shared Sub ParseMetricsSessionDirectory(rootDirectory As String,
                                                        isArchived As Boolean,
                                                        accumulator As MetricsParseAccumulator,
                                                        filterOptions As MetricsSnapshotFilterOptions)
            If accumulator Is Nothing OrElse String.IsNullOrWhiteSpace(rootDirectory) OrElse Not Directory.Exists(rootDirectory) Then
                Return
            End If

            Dim files As IEnumerable(Of String)
            Try
                files = Directory.EnumerateFiles(rootDirectory, "*.jsonl", SearchOption.AllDirectories)
            Catch
                accumulator.ParseErrorCount += 1
                Return
            End Try

            For Each filePath In files
                If isArchived Then
                    accumulator.ArchivedSessionFileCount += 1
                Else
                    accumulator.SessionFileCount += 1
                End If

                ParseMetricsSessionFile(filePath, accumulator, filterOptions)
            Next
        End Sub

        Private Shared Sub ParseMetricsSessionFile(filePath As String,
                                                   accumulator As MetricsParseAccumulator,
                                                   filterOptions As MetricsSnapshotFilterOptions)
            If accumulator Is Nothing OrElse String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                Return
            End If

            Dim fileSessionId = Path.GetFileNameWithoutExtension(filePath)
            If String.IsNullOrWhiteSpace(fileSessionId) Then
                fileSessionId = filePath
            End If

            Dim session = GetOrCreateSessionAccumulator(accumulator, fileSessionId)

            Dim lines As IEnumerable(Of String)
            Try
                lines = File.ReadLines(filePath)
            Catch
                accumulator.ParseErrorCount += 1
                Return
            End Try

            For Each rawLine In lines
                Dim line = If(rawLine, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(line) Then
                    Continue For
                End If

                Try
                    Using document = JsonDocument.Parse(line)
                        Dim root = document.RootElement
                        Dim timestampUtc As DateTimeOffset = DateTimeOffset.MinValue
                        Dim hasTimestamp = TryGetJsonTimestamp(root, "timestamp", timestampUtc)
                        Dim includeInFilteredMetrics = ShouldIncludeMetricsRecord(hasTimestamp, timestampUtc, filterOptions)

                        Dim recordType = GetJsonString(root, "type")
                        Dim payload As JsonElement
                        Dim hasPayload = root.TryGetProperty("payload", payload) AndAlso payload.ValueKind = JsonValueKind.Object

                        If hasPayload Then
                            Select Case recordType
                                Case "session_meta"
                                    Dim metaSessionId = GetJsonString(payload, "id")
                                    If Not String.IsNullOrWhiteSpace(metaSessionId) Then
                                        session = RebindSessionAccumulator(accumulator, session, metaSessionId)
                                    End If

                                    SetSessionProjectPath(session, GetJsonString(payload, "cwd"))

                                    If Not hasTimestamp Then
                                        hasTimestamp = TryGetJsonTimestamp(payload, "timestamp", timestampUtc)
                                    End If

                                Case "turn_context"
                                    Dim model = GetJsonString(payload, "model").Trim()
                                    If includeInFilteredMetrics Then
                                        session.CurrentModel = model
                                        session.TurnContextCount += 1
                                        RegisterModelTurnContext(accumulator, model, session.SessionId)
                                    End If
                                    SetSessionProjectPath(session, GetJsonString(payload, "cwd"))

                                Case "event_msg"
                                    Dim eventKind = GetJsonString(payload, "type")
                                    If StringComparer.Ordinal.Equals(eventKind, "token_count") Then
                                        Dim delta = ExtractTokenDelta(payload, session)
                                        If hasTimestamp Then
                                            AddCalendarTokenUsage(accumulator, timestampUtc, session.SessionId, delta.InputTokens, delta.OutputTokens)
                                        End If

                                        If includeInFilteredMetrics Then
                                            session.InputTokens += delta.InputTokens
                                            session.OutputTokens += delta.OutputTokens
                                            RegisterModelTokenUsage(accumulator, session.CurrentModel, session.SessionId, delta.InputTokens, delta.OutputTokens)
                                            AddDailyTokenUsage(accumulator, timestampUtc, session.SessionId, delta.InputTokens, delta.OutputTokens)
                                        End If
                                    End If

                                Case "response_item"
                                    Dim payloadType = GetJsonString(payload, "type")
                                    If StringComparer.Ordinal.Equals(payloadType, "function_call") Then
                                        If hasTimestamp Then
                                            AddCalendarToolCall(accumulator, timestampUtc, session.SessionId)
                                        End If

                                        If includeInFilteredMetrics Then
                                            session.ToolCallCount += 1
                                            RegisterModelToolCall(accumulator, session.CurrentModel, session.SessionId, 1)
                                            AddDailyToolCall(accumulator, timestampUtc, session.SessionId)
                                        End If
                                    ElseIf StringComparer.Ordinal.Equals(payloadType, "message") AndAlso
                                           StringComparer.Ordinal.Equals(GetJsonString(payload, "role"), "user") Then
                                        Dim promptText = ExtractPromptTextFromResponsePayload(payload)
                                        If Not String.IsNullOrWhiteSpace(promptText) Then
                                            Dim fallbackRecord As New MetricsPromptRecord() With {
                                                .SessionId = session.SessionId,
                                                .Text = promptText
                                            }
                                            If hasTimestamp Then
                                                fallbackRecord.TimestampUtc = timestampUtc
                                            End If

                                            accumulator.FallbackPrompts.Add(fallbackRecord)
                                        End If
                                    End If
                            End Select
                        End If

                        If hasTimestamp Then
                            RegisterCalendarSessionActivity(accumulator, session.SessionId, timestampUtc)
                            If includeInFilteredMetrics Then
                                RegisterSessionActivity(accumulator, session, timestampUtc)
                            End If
                        End If
                    End Using
                Catch
                    accumulator.ParseErrorCount += 1
                End Try
            Next
        End Sub

        Private Shared Sub ApplyPromptRecordsToAccumulator(accumulator As MetricsParseAccumulator,
                                                           filterOptions As MetricsSnapshotFilterOptions)
            Dim selectedPromptRecords As IEnumerable(Of MetricsPromptRecord) =
                If(accumulator.HistoryPrompts.Count > 0, accumulator.HistoryPrompts, accumulator.FallbackPrompts)

            For Each promptRecord In selectedPromptRecords
                If promptRecord Is Nothing Then
                    Continue For
                End If

                Dim sessionId = If(promptRecord.SessionId, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(sessionId) Then
                    Continue For
                End If

                Dim session = GetOrCreateSessionAccumulator(accumulator, sessionId)
                Dim hasTimestamp = promptRecord.TimestampUtc.HasValue
                Dim timestampUtc = If(hasTimestamp, promptRecord.TimestampUtc.Value, DateTimeOffset.MinValue)
                Dim includeInFilteredMetrics = ShouldIncludeMetricsRecord(hasTimestamp, timestampUtc, filterOptions)

                If hasTimestamp Then
                    AddCalendarPrompt(accumulator, timestampUtc, sessionId)
                    RegisterCalendarSessionActivity(accumulator, sessionId, timestampUtc)
                End If

                If includeInFilteredMetrics Then
                    session.PromptCount += 1
                    If hasTimestamp Then
                        RegisterSessionActivity(accumulator, session, timestampUtc)
                        AddDailyPrompt(accumulator, timestampUtc, sessionId)
                    End If
                End If
            Next
        End Sub

        Private Shared Function ExtractTokenDelta(payload As JsonElement, session As MetricsSessionAccumulator) As MetricsTokenDelta
            If session Is Nothing Then
                Return New MetricsTokenDelta()
            End If

            Dim infoElement As JsonElement
            If Not payload.TryGetProperty("info", infoElement) OrElse infoElement.ValueKind <> JsonValueKind.Object Then
                Return New MetricsTokenDelta()
            End If

            Dim lastUsageElement As JsonElement
            If infoElement.TryGetProperty("last_token_usage", lastUsageElement) AndAlso
               lastUsageElement.ValueKind = JsonValueKind.Object Then
                Return New MetricsTokenDelta() With {
                    .InputTokens = Math.Max(0, GetJsonInt64(lastUsageElement, "input_tokens")),
                    .OutputTokens = Math.Max(0, GetJsonInt64(lastUsageElement, "output_tokens"))
                }
            End If

            Dim totalUsageElement As JsonElement
            If infoElement.TryGetProperty("total_token_usage", totalUsageElement) AndAlso
               totalUsageElement.ValueKind = JsonValueKind.Object Then
                Dim totalInputTokens = Math.Max(0, GetJsonInt64(totalUsageElement, "input_tokens"))
                Dim totalOutputTokens = Math.Max(0, GetJsonInt64(totalUsageElement, "output_tokens"))

                Dim deltaInput As Long
                Dim deltaOutput As Long

                If session.HasSeenTotalTokenUsage AndAlso session.LastSeenTotalInputTokens.HasValue Then
                    deltaInput = Math.Max(0, totalInputTokens - session.LastSeenTotalInputTokens.Value)
                Else
                    deltaInput = totalInputTokens
                End If

                If session.HasSeenTotalTokenUsage AndAlso session.LastSeenTotalOutputTokens.HasValue Then
                    deltaOutput = Math.Max(0, totalOutputTokens - session.LastSeenTotalOutputTokens.Value)
                Else
                    deltaOutput = totalOutputTokens
                End If

                session.HasSeenTotalTokenUsage = True
                session.LastSeenTotalInputTokens = totalInputTokens
                session.LastSeenTotalOutputTokens = totalOutputTokens

                Return New MetricsTokenDelta() With {
                    .InputTokens = deltaInput,
                    .OutputTokens = deltaOutput
                }
            End If

            Return New MetricsTokenDelta()
        End Function

        Private Shared Function ExtractPromptTextFromResponsePayload(payload As JsonElement) As String
            Dim contentElement As JsonElement
            If payload.TryGetProperty("content", contentElement) Then
                If contentElement.ValueKind = JsonValueKind.String Then
                    Return If(contentElement.GetString(), String.Empty)
                End If

                If contentElement.ValueKind = JsonValueKind.Array Then
                    Dim builder As New StringBuilder()
                    For Each item In contentElement.EnumerateArray()
                        If item.ValueKind <> JsonValueKind.Object Then
                            Continue For
                        End If

                        Dim text = GetJsonString(item, "text")
                        If String.IsNullOrWhiteSpace(text) Then
                            Continue For
                        End If

                        If builder.Length > 0 Then
                            builder.AppendLine()
                        End If
                        builder.Append(text)
                    Next

                    If builder.Length > 0 Then
                        Return builder.ToString()
                    End If
                End If
            End If

            Return GetJsonString(payload, "message")
        End Function

        Private Shared Sub RegisterModelTurnContext(accumulator As MetricsParseAccumulator,
                                                    modelId As String,
                                                    sessionId As String)
            Dim modelStats = GetOrCreateModelAccumulator(accumulator, modelId)
            If modelStats Is Nothing Then
                Return
            End If

            modelStats.TurnCount += 1
            RegisterModelSession(modelStats, sessionId)
        End Sub

        Private Shared Sub RegisterModelTokenUsage(accumulator As MetricsParseAccumulator,
                                                   modelId As String,
                                                   sessionId As String,
                                                   inputTokens As Long,
                                                   outputTokens As Long)
            If inputTokens = 0 AndAlso outputTokens = 0 Then
                Return
            End If

            Dim modelStats = GetOrCreateModelAccumulator(accumulator, modelId)
            If modelStats Is Nothing Then
                Return
            End If

            modelStats.InputTokens += Math.Max(0, inputTokens)
            modelStats.OutputTokens += Math.Max(0, outputTokens)
            RegisterModelSession(modelStats, sessionId)
        End Sub

        Private Shared Sub RegisterModelToolCall(accumulator As MetricsParseAccumulator,
                                                 modelId As String,
                                                 sessionId As String,
                                                 callCount As Integer)
            If callCount <= 0 Then
                Return
            End If

            Dim modelStats = GetOrCreateModelAccumulator(accumulator, modelId)
            If modelStats Is Nothing Then
                Return
            End If

            modelStats.ToolCallCount += callCount
            RegisterModelSession(modelStats, sessionId)
        End Sub

        Private Shared Sub RegisterModelSession(modelStats As MetricsModelAccumulator, sessionId As String)
            If modelStats Is Nothing Then
                Return
            End If

            Dim normalizedSessionId = If(sessionId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedSessionId) Then
                Return
            End If

            modelStats.SessionIds.Add(normalizedSessionId)
        End Sub

        Private Shared Function GetOrCreateModelAccumulator(accumulator As MetricsParseAccumulator,
                                                            modelId As String) As MetricsModelAccumulator
            If accumulator Is Nothing Then
                Return Nothing
            End If

            Dim normalizedModelId = If(modelId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedModelId) Then
                normalizedModelId = "(unknown-model)"
            End If

            Dim modelStats As MetricsModelAccumulator = Nothing
            If Not accumulator.ModelStatsById.TryGetValue(normalizedModelId, modelStats) Then
                modelStats = New MetricsModelAccumulator() With {
                    .ModelId = normalizedModelId
                }
                accumulator.ModelStatsById(normalizedModelId) = modelStats
            End If

            Return modelStats
        End Function

        Private Shared Sub AddDailyPrompt(accumulator As MetricsParseAccumulator,
                                          timestampUtc As DateTimeOffset,
                                          sessionId As String)
            Dim dayMetrics = GetOrCreateDailyAccumulator(accumulator, timestampUtc)
            dayMetrics.PromptCount += 1
            RegisterDailySession(dayMetrics, sessionId)
        End Sub

        Private Shared Sub AddCalendarPrompt(accumulator As MetricsParseAccumulator,
                                             timestampUtc As DateTimeOffset,
                                             sessionId As String)
            Dim dayMetrics = GetOrCreateCalendarDailyAccumulator(accumulator, timestampUtc)
            dayMetrics.PromptCount += 1
            RegisterDailySession(dayMetrics, sessionId)
        End Sub

        Private Shared Sub AddDailyTokenUsage(accumulator As MetricsParseAccumulator,
                                              timestampUtc As DateTimeOffset,
                                              sessionId As String,
                                              inputTokens As Long,
                                              outputTokens As Long)
            If inputTokens = 0 AndAlso outputTokens = 0 Then
                Return
            End If

            Dim dayMetrics = GetOrCreateDailyAccumulator(accumulator, timestampUtc)
            dayMetrics.InputTokens += Math.Max(0, inputTokens)
            dayMetrics.OutputTokens += Math.Max(0, outputTokens)
            RegisterDailySession(dayMetrics, sessionId)
        End Sub

        Private Shared Sub AddCalendarTokenUsage(accumulator As MetricsParseAccumulator,
                                                 timestampUtc As DateTimeOffset,
                                                 sessionId As String,
                                                 inputTokens As Long,
                                                 outputTokens As Long)
            If inputTokens = 0 AndAlso outputTokens = 0 Then
                Return
            End If

            Dim dayMetrics = GetOrCreateCalendarDailyAccumulator(accumulator, timestampUtc)
            dayMetrics.InputTokens += Math.Max(0, inputTokens)
            dayMetrics.OutputTokens += Math.Max(0, outputTokens)
            RegisterDailySession(dayMetrics, sessionId)
        End Sub

        Private Shared Sub AddDailyToolCall(accumulator As MetricsParseAccumulator,
                                            timestampUtc As DateTimeOffset,
                                            sessionId As String)
            Dim dayMetrics = GetOrCreateDailyAccumulator(accumulator, timestampUtc)
            dayMetrics.ToolCallCount += 1
            RegisterDailySession(dayMetrics, sessionId)
        End Sub

        Private Shared Sub AddCalendarToolCall(accumulator As MetricsParseAccumulator,
                                               timestampUtc As DateTimeOffset,
                                               sessionId As String)
            Dim dayMetrics = GetOrCreateCalendarDailyAccumulator(accumulator, timestampUtc)
            dayMetrics.ToolCallCount += 1
            RegisterDailySession(dayMetrics, sessionId)
        End Sub

        Private Shared Function GetOrCreateDailyAccumulator(accumulator As MetricsParseAccumulator,
                                                            timestampUtc As DateTimeOffset) As MetricsDailyAccumulator
            Return GetOrCreateDailyAccumulatorFromDictionary(accumulator?.DailyByDate, timestampUtc)
        End Function

        Private Shared Function GetOrCreateCalendarDailyAccumulator(accumulator As MetricsParseAccumulator,
                                                                    timestampUtc As DateTimeOffset) As MetricsDailyAccumulator
            Return GetOrCreateDailyAccumulatorFromDictionary(accumulator?.CalendarDailyByDate, timestampUtc)
        End Function

        Private Shared Function GetOrCreateDailyAccumulatorFromDictionary(dayDictionary As Dictionary(Of Date, MetricsDailyAccumulator),
                                                                          timestampUtc As DateTimeOffset) As MetricsDailyAccumulator
            If dayDictionary Is Nothing Then
                Return New MetricsDailyAccumulator()
            End If

            Dim dayKey = timestampUtc.LocalDateTime.Date
            Dim dayMetrics As MetricsDailyAccumulator = Nothing
            If Not dayDictionary.TryGetValue(dayKey, dayMetrics) Then
                dayMetrics = New MetricsDailyAccumulator()
                dayDictionary(dayKey) = dayMetrics
            End If

            Return dayMetrics
        End Function

        Private Shared Sub RegisterDailySession(dayMetrics As MetricsDailyAccumulator, sessionId As String)
            If dayMetrics Is Nothing Then
                Return
            End If

            Dim normalizedSessionId = If(sessionId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedSessionId) Then
                Return
            End If

            dayMetrics.SessionIds.Add(normalizedSessionId)
        End Sub

        Private Shared Sub RegisterSessionActivity(accumulator As MetricsParseAccumulator,
                                                   session As MetricsSessionAccumulator,
                                                   timestampUtc As DateTimeOffset)
            If session Is Nothing Then
                Return
            End If

            If Not session.FirstActivityUtc.HasValue OrElse timestampUtc < session.FirstActivityUtc.Value Then
                session.FirstActivityUtc = timestampUtc
            End If
            If Not session.LastActivityUtc.HasValue OrElse timestampUtc > session.LastActivityUtc.Value Then
                session.LastActivityUtc = timestampUtc
            End If

            If Not accumulator.ActivityStartUtc.HasValue OrElse timestampUtc < accumulator.ActivityStartUtc.Value Then
                accumulator.ActivityStartUtc = timestampUtc
            End If
            If Not accumulator.ActivityEndUtc.HasValue OrElse timestampUtc > accumulator.ActivityEndUtc.Value Then
                accumulator.ActivityEndUtc = timestampUtc
            End If

            Dim dayMetrics = GetOrCreateDailyAccumulator(accumulator, timestampUtc)
            RegisterDailySession(dayMetrics, session.SessionId)
        End Sub

        Private Shared Sub RegisterCalendarSessionActivity(accumulator As MetricsParseAccumulator,
                                                           sessionId As String,
                                                           timestampUtc As DateTimeOffset)
            Dim dayMetrics = GetOrCreateCalendarDailyAccumulator(accumulator, timestampUtc)
            RegisterDailySession(dayMetrics, sessionId)
        End Sub

        Private Shared Function RebindSessionAccumulator(accumulator As MetricsParseAccumulator,
                                                         currentSession As MetricsSessionAccumulator,
                                                         desiredSessionId As String) As MetricsSessionAccumulator
            Dim normalizedDesiredId = If(desiredSessionId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedDesiredId) Then
                Return currentSession
            End If

            Dim targetSession = GetOrCreateSessionAccumulator(accumulator, normalizedDesiredId)
            If currentSession Is Nothing OrElse
               StringComparer.Ordinal.Equals(currentSession.SessionId, targetSession.SessionId) Then
                Return targetSession
            End If

            MergeSessionAccumulators(targetSession, currentSession)
            accumulator.SessionsById.Remove(currentSession.SessionId)
            Return targetSession
        End Function

        Private Shared Sub MergeSessionAccumulators(target As MetricsSessionAccumulator,
                                                    source As MetricsSessionAccumulator)
            If target Is Nothing OrElse source Is Nothing Then
                Return
            End If

            target.PromptCount += source.PromptCount
            target.InputTokens += source.InputTokens
            target.OutputTokens += source.OutputTokens
            target.ToolCallCount += source.ToolCallCount
            target.TurnContextCount += source.TurnContextCount

            If String.IsNullOrWhiteSpace(target.ProjectPath) Then
                target.ProjectPath = source.ProjectPath
            End If

            If source.FirstActivityUtc.HasValue Then
                If Not target.FirstActivityUtc.HasValue OrElse source.FirstActivityUtc.Value < target.FirstActivityUtc.Value Then
                    target.FirstActivityUtc = source.FirstActivityUtc
                End If
            End If

            If source.LastActivityUtc.HasValue Then
                If Not target.LastActivityUtc.HasValue OrElse source.LastActivityUtc.Value > target.LastActivityUtc.Value Then
                    target.LastActivityUtc = source.LastActivityUtc
                End If
            End If
        End Sub

        Private Shared Function GetOrCreateSessionAccumulator(accumulator As MetricsParseAccumulator,
                                                              sessionId As String) As MetricsSessionAccumulator
            Dim normalizedSessionId = If(sessionId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedSessionId) Then
                normalizedSessionId = "(unknown-session)"
            End If

            Dim session As MetricsSessionAccumulator = Nothing
            If Not accumulator.SessionsById.TryGetValue(normalizedSessionId, session) Then
                session = New MetricsSessionAccumulator() With {
                    .SessionId = normalizedSessionId
                }
                accumulator.SessionsById(normalizedSessionId) = session
            End If

            Return session
        End Function

        Private Shared Sub SetSessionProjectPath(session As MetricsSessionAccumulator, candidatePath As String)
            If session Is Nothing Then
                Return
            End If

            Dim normalizedCandidatePath = NormalizeProjectPath(candidatePath)
            If String.IsNullOrWhiteSpace(normalizedCandidatePath) Then
                Return
            End If

            If String.IsNullOrWhiteSpace(session.ProjectPath) Then
                session.ProjectPath = normalizedCandidatePath
                Return
            End If

            If session.ProjectPath.StartsWith(normalizedCandidatePath, StringComparison.OrdinalIgnoreCase) AndAlso
               normalizedCandidatePath.Length < session.ProjectPath.Length Then
                session.ProjectPath = normalizedCandidatePath
            End If
        End Sub

        Private Shared Function BuildProjectDisplayText(projectPath As String) As String
            Dim normalizedProjectPath = If(projectPath, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedProjectPath) Then
                Return MetricsUnknownProjectLabel
            End If

            If StringComparer.Ordinal.Equals(normalizedProjectPath, MetricsUnknownProjectLabel) Then
                Return MetricsUnknownProjectLabel
            End If

            Dim leaf = normalizedProjectPath
            Try
                Dim maybeLeaf = Path.GetFileName(normalizedProjectPath)
                If Not String.IsNullOrWhiteSpace(maybeLeaf) Then
                    leaf = maybeLeaf
                End If
            Catch
                leaf = normalizedProjectPath
            End Try

            If StringComparer.OrdinalIgnoreCase.Equals(leaf, normalizedProjectPath) Then
                Return leaf
            End If

            Return $"{leaf} ({normalizedProjectPath})"
        End Function

        Private Shared Function TryGetJsonTimestamp(source As JsonElement,
                                                    propertyName As String,
                                                    ByRef timestampUtc As DateTimeOffset) As Boolean
            If source.ValueKind <> JsonValueKind.Object OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return False
            End If

            Dim timestampElement As JsonElement
            If Not source.TryGetProperty(propertyName, timestampElement) Then
                Return False
            End If

            Return TryParseTimestampValue(timestampElement, timestampUtc)
        End Function

        Private Shared Function TryParseTimestampValue(timestampElement As JsonElement,
                                                       ByRef timestampUtc As DateTimeOffset) As Boolean
            Select Case timestampElement.ValueKind
                Case JsonValueKind.String
                    Dim raw = If(timestampElement.GetString(), String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(raw) Then
                        Return False
                    End If

                    Return DateTimeOffset.TryParse(raw,
                                                   CultureInfo.InvariantCulture,
                                                   DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal,
                                                   timestampUtc)

                Case JsonValueKind.Number
                    Dim unixValue As Long
                    If timestampElement.TryGetInt64(unixValue) Then
                        Return TryParseUnixEpoch(unixValue, timestampUtc)
                    End If

                    Dim unixDouble As Double
                    If timestampElement.TryGetDouble(unixDouble) Then
                        If Double.IsNaN(unixDouble) OrElse Double.IsInfinity(unixDouble) Then
                            Return False
                        End If

                        Return TryParseUnixEpoch(CLng(Math.Truncate(unixDouble)), timestampUtc)
                    End If
            End Select

            Return False
        End Function

        Private Shared Function TryParseUnixEpoch(unixValue As Long,
                                                  ByRef timestampUtc As DateTimeOffset) As Boolean
            Try
                If Math.Abs(unixValue) >= 100000000000L Then
                    timestampUtc = DateTimeOffset.FromUnixTimeMilliseconds(unixValue)
                Else
                    timestampUtc = DateTimeOffset.FromUnixTimeSeconds(unixValue)
                End If

                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function GetJsonString(source As JsonElement, propertyName As String) As String
            If source.ValueKind <> JsonValueKind.Object OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return String.Empty
            End If

            Dim value As JsonElement
            If Not source.TryGetProperty(propertyName, value) Then
                Return String.Empty
            End If

            Select Case value.ValueKind
                Case JsonValueKind.String
                    Return If(value.GetString(), String.Empty)
                Case JsonValueKind.Number, JsonValueKind.True, JsonValueKind.False
                    Return value.ToString()
                Case Else
                    Return String.Empty
            End Select
        End Function

        Private Shared Function GetJsonInt64(source As JsonElement, propertyName As String) As Long
            If source.ValueKind <> JsonValueKind.Object OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return 0
            End If

            Dim value As JsonElement
            If Not source.TryGetProperty(propertyName, value) Then
                Return 0
            End If

            Select Case value.ValueKind
                Case JsonValueKind.Number
                    Dim parsedInteger As Long
                    If value.TryGetInt64(parsedInteger) Then
                        Return parsedInteger
                    End If

                    Dim parsedDouble As Double
                    If value.TryGetDouble(parsedDouble) AndAlso
                       Not Double.IsNaN(parsedDouble) AndAlso
                       Not Double.IsInfinity(parsedDouble) Then
                        Return CLng(Math.Truncate(parsedDouble))
                    End If

                Case JsonValueKind.String
                    Dim raw = If(value.GetString(), String.Empty).Trim()
                    Dim parsedFromString As Long
                    If Long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, parsedFromString) Then
                        Return parsedFromString
                    End If
            End Select

            Return 0
        End Function

        Private Shared Function FormatMetricCount(value As Integer) As String
            Return value.ToString("N0", CultureInfo.InvariantCulture)
        End Function

        Private Shared Function FormatMetricCount(value As Long) As String
            Return value.ToString("N0", CultureInfo.InvariantCulture)
        End Function
    End Class
End Namespace
