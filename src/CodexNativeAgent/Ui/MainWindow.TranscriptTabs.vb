Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Threading
Imports CodexNativeAgent.Ui.ViewModels.Threads

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Enum TranscriptTabKind
            Thread = 0
            PendingNewThread = 1
        End Enum

        Private NotInheritable Class TranscriptTabState
            Public Property TabId As String = String.Empty
            Public Property Kind As TranscriptTabKind = TranscriptTabKind.Thread
            Public Property ThreadId As String = String.Empty
            Public Property CreatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
            Public Property LastActivatedUtc As DateTimeOffset = DateTimeOffset.UtcNow
            Public Property LastLifecycleReason As String = String.Empty
            Public Property IsClosable As Boolean = True
            Public Property AutoRemoveIfEmptyOnExistingSelection As Boolean
        End Class

        Private NotInheritable Class TranscriptTabSurfaceHandle
            Public Property ThreadId As String = String.Empty
            Public Property SurfaceListBox As ListBox
            Public Property TabButton As Button
            Public Property TabChipBorder As Border
            Public Property TabCloseButton As Button
            Public Property IsPrimarySurface As Boolean
        End Class

        Private NotInheritable Class TranscriptTabSurfaceRetireWorkItem
            Public Property ThreadId As String = String.Empty
            Public Property SurfaceListBox As ListBox
            Public Property IsPrimarySurface As Boolean
            Public Property QueuedUtc As DateTimeOffset = DateTimeOffset.UtcNow
        End Class

        Private ReadOnly _transcriptTabSurfacesByThreadId As New Dictionary(Of String, TranscriptTabSurfaceHandle)(StringComparer.Ordinal)
        Private ReadOnly _transcriptTabStatesByTabId As New Dictionary(Of String, TranscriptTabState)(StringComparer.Ordinal)
        Private ReadOnly _transcriptInteractionHandlersAttached As New List(Of ListBox)()
        Private ReadOnly _transcriptTabSurfaceRetireQueue As New Queue(Of TranscriptTabSurfaceRetireWorkItem)()
        Private ReadOnly _dormantTranscriptTabSurfaces As New Stack(Of ListBox)()
        Private Const PendingNewThreadTranscriptTabId As String = "__pending_new_thread__"
        Private Const TranscriptTabStripHorizontalInset As Double = 14.0R
        Private Const TranscriptTabButtonCompactMinWidth As Double = 56.0R
        Private Const TranscriptTabButtonExpandedMaxWidth As Double = 320.0R
        Private Const TranscriptTabButtonDefaultMaxChars As Integer = 32
        Private Const TranscriptTabButtonApproxCharWidth As Double = 7.0R
        Private Const TranscriptTabButtonPaddingHorizontalWidth As Double = 20.0R
        Private Const TranscriptTabChipCloseAreaReservedWidth As Double = 30.0R
        Private Const TranscriptTabChipSpacingReservedWidth As Double = 4.0R
        Private Const TranscriptTabCaptionMinChars As Integer = 6
        Private Const TranscriptTabCaptionMaxChars As Integer = 48
        Private _transcriptSurfaceHostPanel As Panel
        Private _transcriptTabStripBorder As Border
        Private _transcriptTabStripScrollViewer As ScrollViewer
        Private _transcriptTabStripPanel As StackPanel
        Private _activeTranscriptSurfaceListBox As ListBox
        Private _activeTranscriptSurfaceThreadId As String = String.Empty
        Private _transcriptTabSurfaceRetireDrainScheduled As Boolean
        Private _primaryTranscriptSurfaceResetScheduled As Boolean
        Private _primaryTranscriptSurfaceResetDeferredNeeded As Boolean
        Private _deferredBlankCloseFinalizeVersion As Integer
        Private _transcriptTabInteractionEnabled As Boolean = True
        Private Const TranscriptTabDormantSurfacePoolMax As Integer = 4

        Private Function CurrentTranscriptListControl() As ListBox
            If _activeTranscriptSurfaceListBox IsNot Nothing Then
                Return _activeTranscriptSurfaceListBox
            End If

            If WorkspacePaneHost Is Nothing Then
                Return Nothing
            End If

            Return WorkspacePaneHost.LstTranscript
        End Function

        Private Shared Function NormalizeTranscriptTabId(value As String) As String
            Return If(value, String.Empty).Trim()
        End Function

        Private Shared Function TranscriptTabKindForTabId(tabId As String) As TranscriptTabKind
            If StringComparer.Ordinal.Equals(NormalizeTranscriptTabId(tabId), PendingNewThreadTranscriptTabId) Then
                Return TranscriptTabKind.PendingNewThread
            End If

            Return TranscriptTabKind.Thread
        End Function

        Private Function GetOrCreateTranscriptTabState(tabId As String,
                                                       Optional lifecycleReason As String = Nothing) As TranscriptTabState
            Dim normalizedTabId = NormalizeTranscriptTabId(tabId)
            If String.IsNullOrWhiteSpace(normalizedTabId) Then
                Return Nothing
            End If

            Dim state As TranscriptTabState = Nothing
            If Not _transcriptTabStatesByTabId.TryGetValue(normalizedTabId, state) OrElse state Is Nothing Then
                state = New TranscriptTabState() With {
                    .TabId = normalizedTabId,
                    .Kind = TranscriptTabKindForTabId(normalizedTabId),
                    .ThreadId = If(StringComparer.Ordinal.Equals(normalizedTabId, PendingNewThreadTranscriptTabId), String.Empty, normalizedTabId),
                    .AutoRemoveIfEmptyOnExistingSelection = StringComparer.Ordinal.Equals(normalizedTabId, PendingNewThreadTranscriptTabId)
                }
                _transcriptTabStatesByTabId(normalizedTabId) = state
                TraceTranscriptTabStateSnapshot("tab_state_created",
                                               $"tab={normalizedTabId}; kind={state.Kind}")
            End If

            state.Kind = TranscriptTabKindForTabId(normalizedTabId)
            state.ThreadId = If(state.Kind = TranscriptTabKind.PendingNewThread, String.Empty, normalizedTabId)
            state.AutoRemoveIfEmptyOnExistingSelection = (state.Kind = TranscriptTabKind.PendingNewThread)
            If Not String.IsNullOrWhiteSpace(lifecycleReason) Then
                state.LastLifecycleReason = lifecycleReason
            End If

            Return state
        End Function

        Private Function TryGetTranscriptTabState(tabId As String,
                                                  ByRef state As TranscriptTabState) As Boolean
            state = Nothing
            Dim normalizedTabId = NormalizeTranscriptTabId(tabId)
            If String.IsNullOrWhiteSpace(normalizedTabId) Then
                Return False
            End If

            Return _transcriptTabStatesByTabId.TryGetValue(normalizedTabId, state) AndAlso state IsNot Nothing
        End Function

        Private Sub RemoveTranscriptTabState(tabId As String,
                                             Optional reason As String = Nothing)
            Dim normalizedTabId = NormalizeTranscriptTabId(tabId)
            If String.IsNullOrWhiteSpace(normalizedTabId) Then
                Return
            End If

            If _transcriptTabStatesByTabId.Remove(normalizedTabId) Then
                RecomputeTranscriptTabStatePolicies($"remove:{If(reason, String.Empty)}")
                TraceTranscriptTabStateSnapshot("tab_state_removed",
                                               $"tab={normalizedTabId}; reason={If(reason, String.Empty)}")
            End If
        End Sub

        Private Sub TouchTranscriptTabStateActivation(tabId As String,
                                                      Optional reason As String = Nothing)
            Dim state = GetOrCreateTranscriptTabState(tabId, reason)
            If state Is Nothing Then
                Return
            End If

            state.LastActivatedUtc = DateTimeOffset.UtcNow
            If Not String.IsNullOrWhiteSpace(reason) Then
                state.LastLifecycleReason = reason
            End If
            RecomputeTranscriptTabStatePolicies($"activate:{If(reason, String.Empty)}")
        End Sub

        Private Sub RecomputeTranscriptTabStatePolicies(Optional reason As String = Nothing)
            Dim totalTabs = _transcriptTabStatesByTabId.Count
            For Each kvp In _transcriptTabStatesByTabId
                Dim state = kvp.Value
                If state Is Nothing Then
                    Continue For
                End If

                Dim isOnlyTab = (totalTabs <= 1)
                If state.Kind = TranscriptTabKind.PendingNewThread Then
                    state.IsClosable = Not isOnlyTab
                    state.AutoRemoveIfEmptyOnExistingSelection = True
                Else
                    state.IsClosable = True
                    state.AutoRemoveIfEmptyOnExistingSelection = False
                End If

                If Not String.IsNullOrWhiteSpace(reason) Then
                    state.LastLifecycleReason = reason
                End If
            Next

            TraceTranscriptTabStateSnapshot("policies_recomputed",
                                           $"reason={If(reason, String.Empty)}")
        End Sub

        Private Function IsTranscriptTabClosable(tabId As String) As Boolean
            Dim state As TranscriptTabState = Nothing
            If TryGetTranscriptTabState(tabId, state) AndAlso state IsNot Nothing Then
                Return state.IsClosable
            End If

            Return True
        End Function

        Private Sub TraceTranscriptTabStateSnapshot(eventName As String,
                                                    Optional details As String = Nothing)
            Dim totalTabs = _transcriptTabStatesByTabId.Count
            Dim pendingExists = _transcriptTabStatesByTabId.ContainsKey(PendingNewThreadTranscriptTabId)
            Dim pendingClosable = IsTranscriptTabClosable(PendingNewThreadTranscriptTabId)
            AppendProtocol("debug",
                           $"transcript_tab_state event={If(eventName, String.Empty)} totalTabs={totalTabs} pendingExists={pendingExists} pendingClosable={pendingClosable} activeTab={If(_activeTranscriptSurfaceThreadId, String.Empty)} details={If(details, String.Empty)}")
        End Sub

        Private Sub AttachTranscriptInteractionHandlers(listBox As ListBox)
            If listBox Is Nothing Then
                Return
            End If

            For Each existing In _transcriptInteractionHandlersAttached
                If ReferenceEquals(existing, listBox) Then
                    Return
                End If
            Next

            listBox.AddHandler(ScrollViewer.ScrollChangedEvent,
                               New ScrollChangedEventHandler(AddressOf OnTranscriptScrollChanged))
            AddHandler listBox.PreviewMouseWheel,
                New MouseWheelEventHandler(AddressOf OnTranscriptPreviewMouseWheel)
            AddHandler listBox.PreviewMouseLeftButtonDown,
                New MouseButtonEventHandler(AddressOf OnTranscriptPreviewMouseLeftButtonDown)
            AddHandler listBox.PreviewKeyDown,
                New KeyEventHandler(AddressOf OnTranscriptPreviewKeyDown)
            listBox.AddHandler(Thumb.DragStartedEvent,
                               New DragStartedEventHandler(AddressOf OnTranscriptScrollThumbDragStarted),
                               True)
            listBox.AddHandler(Thumb.DragCompletedEvent,
                               New DragCompletedEventHandler(AddressOf OnTranscriptScrollThumbDragCompleted),
                               True)

            _transcriptInteractionHandlersAttached.Add(listBox)
        End Sub

        Private Sub EnsureTranscriptTabsUiInitialized()
            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return
            End If

            If _transcriptSurfaceHostPanel Is Nothing Then
                _transcriptSurfaceHostPanel = TryCast(WorkspacePaneHost.LstTranscript.Parent, Panel)
            End If

            If _activeTranscriptSurfaceListBox Is Nothing Then
                _activeTranscriptSurfaceListBox = WorkspacePaneHost.LstTranscript
            End If

            WorkspacePaneHost.LstTranscript.DataContext = _viewModel
            AttachTranscriptInteractionHandlers(WorkspacePaneHost.LstTranscript)

            If _transcriptTabStripPanel Is Nothing Then
                _transcriptTabStripPanel = New StackPanel() With {
                    .Orientation = Orientation.Horizontal,
                    .FlowDirection = FlowDirection.LeftToRight
                }

                _transcriptTabStripScrollViewer = New ScrollViewer() With {
                    .HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
                    .VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    .Content = _transcriptTabStripPanel,
                    .CanContentScroll = False,
                    .FlowDirection = FlowDirection.LeftToRight
                }

                _transcriptTabStripBorder = New Border() With {
                    .BorderThickness = New Thickness(0, 0, 0, 2),
                    .Opacity = 1.0R,
                    .Padding = New Thickness(0, 2, 0, 1),
                    .Background = Brushes.Transparent,
                    .HorizontalAlignment = HorizontalAlignment.Stretch,
                    .Visibility = Visibility.Collapsed,
                    .Child = _transcriptTabStripScrollViewer
                }
                _transcriptTabStripBorder.SetResourceReference(Border.BorderBrushProperty, "BorderStrongBrush")
                AddHandler _transcriptTabStripBorder.SizeChanged, AddressOf OnTranscriptTabStripSizeChanged
                AddHandler _transcriptTabStripScrollViewer.SizeChanged, AddressOf OnTranscriptTabStripSizeChanged

                Dim headerDockPanel = FindVisualAncestor(Of DockPanel)(WorkspacePaneHost.LblCurrentThread)
                If headerDockPanel IsNot Nothing Then
                    _transcriptTabStripBorder.Margin = New Thickness(8, 2, 8, 2)
                    _transcriptTabStripBorder.VerticalAlignment = VerticalAlignment.Center
                    headerDockPanel.Children.Add(_transcriptTabStripBorder)
                Else
                    Dim titleStack = TryCast(WorkspacePaneHost.LblCurrentThread.Parent, StackPanel)
                    If titleStack IsNot Nothing Then
                        _transcriptTabStripBorder.Margin = New Thickness(TranscriptTabStripHorizontalInset, 0, TranscriptTabStripHorizontalInset, 2)
                        titleStack.Children.Insert(0, _transcriptTabStripBorder)
                    End If
                End If
            End If

            ApplyTranscriptTabStripInteractionState()
            ApplyTranscriptTabResponsiveLayout()
        End Sub

        Private Sub SetTranscriptTabInteractionEnabled(isEnabled As Boolean)
            _transcriptTabInteractionEnabled = isEnabled
            ApplyTranscriptTabStripInteractionState()
            UpdateTranscriptTabButtonVisuals()
        End Sub

        Private Sub ApplyTranscriptTabStripInteractionState()
            If _transcriptTabStripBorder Is Nothing Then
                Return
            End If

            _transcriptTabStripBorder.IsHitTestVisible = _transcriptTabInteractionEnabled
            _transcriptTabStripBorder.Opacity = If(_transcriptTabInteractionEnabled, 1.0R, 0.74R)
        End Sub

        Private Sub AppendTranscriptTabChipAtStripEnd(tabChipBorder As Border)
            If _transcriptTabStripPanel Is Nothing OrElse tabChipBorder Is Nothing Then
                Return
            End If

            If _transcriptTabStripPanel.Children.Contains(tabChipBorder) Then
                _transcriptTabStripPanel.Children.Remove(tabChipBorder)
            End If

            _transcriptTabStripPanel.Children.Add(tabChipBorder)
            ApplyTranscriptTabResponsiveLayout()
        End Sub

        Private Sub OnTranscriptTabStripSizeChanged(sender As Object, e As SizeChangedEventArgs)
            If e Is Nothing Then
                Return
            End If

            If Math.Abs(e.NewSize.Width - e.PreviousSize.Width) < 0.5R Then
                Return
            End If

            ApplyTranscriptTabResponsiveLayout()
        End Sub

        Private Function ResolveTranscriptTabStripContentWidth() As Double
            If _transcriptTabStripBorder IsNot Nothing Then
                Dim borderContentWidth = _transcriptTabStripBorder.ActualWidth -
                                         _transcriptTabStripBorder.Padding.Left -
                                         _transcriptTabStripBorder.Padding.Right
                If borderContentWidth > 0 Then
                    Return borderContentWidth
                End If
            End If

            If _transcriptTabStripScrollViewer IsNot Nothing AndAlso _transcriptTabStripScrollViewer.ActualWidth > 0 Then
                Return _transcriptTabStripScrollViewer.ActualWidth
            End If

            Return 0
        End Function

        Private Sub ApplyTranscriptTabResponsiveLayout()
            Dim tabCount = _transcriptTabSurfacesByThreadId.Count
            If tabCount <= 0 Then
                Return
            End If

            Dim stripContentWidth = ResolveTranscriptTabStripContentWidth()
            If stripContentWidth <= 0 Then
                stripContentWidth = 640.0R
            End If

            If _transcriptTabStripPanel IsNot Nothing Then
                _transcriptTabStripPanel.MaxWidth = Math.Max(0, stripContentWidth)
            End If

            Dim perTabChipBudget = Math.Max(44.0R,
                                            (stripContentWidth / tabCount) - TranscriptTabChipSpacingReservedWidth)
            Dim buttonMaxWidth = Math.Max(TranscriptTabButtonCompactMinWidth,
                                          Math.Min(TranscriptTabButtonExpandedMaxWidth,
                                                   perTabChipBudget - TranscriptTabChipCloseAreaReservedWidth))
            Dim buttonMinWidth = Math.Min(96.0R, buttonMaxWidth)
            Dim captionBudgetWidth = Math.Max(8.0R, buttonMaxWidth - TranscriptTabButtonPaddingHorizontalWidth)
            Dim maxChars = CInt(Math.Floor(captionBudgetWidth / TranscriptTabButtonApproxCharWidth))
            maxChars = Math.Max(TranscriptTabCaptionMinChars, Math.Min(TranscriptTabCaptionMaxChars, maxChars))

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim handle = kvp.Value
                If handle Is Nothing OrElse handle.TabButton Is Nothing Then
                    Continue For
                End If

                handle.TabButton.MinWidth = buttonMinWidth
                handle.TabButton.MaxWidth = buttonMaxWidth
                UpdateTranscriptTabButtonCaption(handle, maxChars)
            Next
        End Sub

        Private Function EnsureTranscriptTabSurfaceHandle(threadId As String,
                                                         Optional preferSecondarySurface As Boolean = False,
                                                         Optional avoidDormantSurfaceReuse As Boolean = False) As TranscriptTabSurfaceHandle
            EnsureTranscriptTabsUiInitialized()

            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return Nothing
            End If

            Dim existing As TranscriptTabSurfaceHandle = Nothing
            If _transcriptTabSurfacesByThreadId.TryGetValue(normalizedThreadId, existing) AndAlso existing IsNot Nothing Then
                GetOrCreateTranscriptTabState(normalizedThreadId, "surface_handle_reuse")
                RecomputeTranscriptTabStatePolicies("surface_handle_reuse")
                Return existing
            End If

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return Nothing
            End If

            Dim primarySurfaceDeferredDirty = _primaryTranscriptSurfaceResetDeferredNeeded
            Dim usePrimarySurface = (_transcriptTabSurfacesByThreadId.Count = 0) AndAlso
                                    Not preferSecondarySurface AndAlso
                                    Not primarySurfaceDeferredDirty
            If preferSecondarySurface AndAlso _transcriptSurfaceHostPanel Is Nothing Then
                usePrimarySurface = True
            End If
            If Not usePrimarySurface AndAlso _transcriptSurfaceHostPanel Is Nothing Then
                Return Nothing
            End If

            Dim reusedDormantSurface = False
            Dim listBox As ListBox = Nothing
            If usePrimarySurface Then
                listBox = WorkspacePaneHost.LstTranscript
            Else
                If Not avoidDormantSurfaceReuse Then
                    listBox = TakeDormantTranscriptTabSurface()
                    reusedDormantSurface = listBox IsNot Nothing
                End If
                If listBox Is Nothing Then
                    listBox = CreateRetainedTranscriptSurfaceClone()
                End If
            End If
            If listBox Is Nothing Then
                Return Nothing
            End If

            If Not usePrimarySurface AndAlso Not reusedDormantSurface AndAlso _transcriptSurfaceHostPanel IsNot Nothing Then
                _transcriptSurfaceHostPanel.Children.Add(listBox)
            End If

            If usePrimarySurface Then
                _primaryTranscriptSurfaceResetDeferredNeeded = False
            End If

            listBox.DataContext = _viewModel
            listBox.Visibility = Visibility.Collapsed

            Dim tabButton As Button = Nothing
            Dim tabCloseButton As Button = Nothing
            Dim tabChipBorder = CreateTranscriptTabChip(normalizedThreadId,
                                                        tabButton,
                                                        tabCloseButton)

            Dim handle As New TranscriptTabSurfaceHandle() With {
                .ThreadId = normalizedThreadId,
                .SurfaceListBox = listBox,
                .TabButton = tabButton,
                .TabChipBorder = tabChipBorder,
                .TabCloseButton = tabCloseButton,
                .IsPrimarySurface = usePrimarySurface
            }

            _transcriptTabSurfacesByThreadId(normalizedThreadId) = handle
            GetOrCreateTranscriptTabState(normalizedThreadId, "surface_handle_created")
            RecomputeTranscriptTabStatePolicies("surface_handle_created")
            If _transcriptTabStripPanel IsNot Nothing AndAlso tabChipBorder IsNot Nothing Then
                AppendTranscriptTabChipAtStripEnd(tabChipBorder)
            End If

            AttachTranscriptInteractionHandlers(listBox)
            RefreshTranscriptTabStripVisibility()
            UpdateTranscriptTabButtonCaptions()
            UpdateTranscriptTabButtonVisual(handle,
                                            StringComparer.Ordinal.Equals(normalizedThreadId, _activeTranscriptSurfaceThreadId))

            If reusedDormantSurface Then
                AppendProtocol("debug",
                               $"transcript_tab_perf event=surface_reuse_dormant thread={normalizedThreadId} dormantRemaining={_dormantTranscriptTabSurfaces.Count}")
            End If

            Return handle
        End Function

        Private Function CreateRetainedTranscriptSurfaceClone() As ListBox
            Dim source = If(WorkspacePaneHost, Nothing)?.LstTranscript
            If source Is Nothing Then
                Return Nothing
            End If

            Dim clone As New ListBox() With {
                .Background = source.Background,
                .BorderBrush = source.BorderBrush,
                .BorderThickness = source.BorderThickness,
                .FontFamily = source.FontFamily,
                .FontSize = source.FontSize,
                .HorizontalContentAlignment = source.HorizontalContentAlignment,
                .HorizontalAlignment = source.HorizontalAlignment,
                .VerticalAlignment = source.VerticalAlignment,
                .Margin = source.Margin,
                .Padding = source.Padding,
                .SelectionMode = source.SelectionMode,
                .ItemTemplate = source.ItemTemplate,
                .ItemContainerStyle = source.ItemContainerStyle,
                .ItemsPanel = source.ItemsPanel,
                .Visibility = Visibility.Collapsed
            }

            clone.SetValue(ScrollViewer.CanContentScrollProperty, source.GetValue(ScrollViewer.CanContentScrollProperty))
            clone.SetValue(VirtualizingPanel.IsVirtualizingProperty, source.GetValue(VirtualizingPanel.IsVirtualizingProperty))
            clone.SetValue(VirtualizingPanel.VirtualizationModeProperty, source.GetValue(VirtualizingPanel.VirtualizationModeProperty))
            clone.SetValue(VirtualizingPanel.ScrollUnitProperty, source.GetValue(VirtualizingPanel.ScrollUnitProperty))
            clone.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, source.GetValue(ScrollViewer.VerticalScrollBarVisibilityProperty))
            clone.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, source.GetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty))

            If _transcriptSurfaceHostPanel IsNot Nothing Then
                Grid.SetRow(clone, Grid.GetRow(source))
                Grid.SetColumn(clone, Grid.GetColumn(source))
                Grid.SetRowSpan(clone, Grid.GetRowSpan(source))
                Grid.SetColumnSpan(clone, Grid.GetColumnSpan(source))
                Panel.SetZIndex(clone, Panel.GetZIndex(source))
            End If

            For Each key As Object In source.Resources.Keys
                If clone.Resources.Contains(key) Then
                    Continue For
                End If

                clone.Resources.Add(key, source.Resources(key))
            Next

            Return clone
        End Function

        Private Shared Function CreateContentOnlyButtonTemplate(horizontalAlignment As HorizontalAlignment,
                                                                verticalAlignment As VerticalAlignment) As ControlTemplate
            Dim template As New ControlTemplate(GetType(Button))
            Dim chrome As New FrameworkElementFactory(GetType(Border))
            chrome.SetValue(Border.BackgroundProperty,
                            New TemplateBindingExtension(Control.BackgroundProperty))
            chrome.SetValue(Border.BorderBrushProperty,
                            New TemplateBindingExtension(Control.BorderBrushProperty))
            chrome.SetValue(Border.BorderThicknessProperty,
                            New TemplateBindingExtension(Control.BorderThicknessProperty))

            Dim presenterFactory As New FrameworkElementFactory(GetType(ContentPresenter))
            presenterFactory.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, horizontalAlignment)
            presenterFactory.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, verticalAlignment)
            presenterFactory.SetValue(System.Windows.Controls.ContentPresenter.RecognizesAccessKeyProperty, True)
            presenterFactory.SetValue(System.Windows.Controls.ContentPresenter.MarginProperty,
                                      New TemplateBindingExtension(Control.PaddingProperty))

            chrome.AppendChild(presenterFactory)
            template.VisualTree = chrome
            Return template
        End Function

        Private Shared Function CreateChromeButtonTemplate(cornerRadius As Double) As ControlTemplate
            Dim template As New ControlTemplate(GetType(Button))
            Dim chrome As New FrameworkElementFactory(GetType(Border))
            chrome.SetValue(Border.CornerRadiusProperty, New CornerRadius(cornerRadius))
            chrome.SetValue(Border.BackgroundProperty,
                            New TemplateBindingExtension(Control.BackgroundProperty))
            chrome.SetValue(Border.BorderBrushProperty,
                            New TemplateBindingExtension(Control.BorderBrushProperty))
            chrome.SetValue(Border.BorderThicknessProperty,
                            New TemplateBindingExtension(Control.BorderThicknessProperty))

            Dim presenterFactory As New FrameworkElementFactory(GetType(ContentPresenter))
            presenterFactory.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center)
            presenterFactory.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center)
            presenterFactory.SetValue(System.Windows.Controls.ContentPresenter.RecognizesAccessKeyProperty, True)
            presenterFactory.SetValue(System.Windows.Controls.ContentPresenter.MarginProperty,
                                      New TemplateBindingExtension(Control.PaddingProperty))

            chrome.AppendChild(presenterFactory)
            template.VisualTree = chrome
            Return template
        End Function

        Private Function CreateTranscriptTabChip(threadId As String,
                                                 ByRef tabButton As Button,
                                                 ByRef closeButton As Button) As Border
            tabButton = New Button() With {
                .Tag = threadId,
                .Padding = New Thickness(12, 7, 8, 7),
                .MinHeight = 32,
                .MinWidth = 96,
                .MaxWidth = 320,
                .FontSize = 12.6R,
                .Cursor = Cursors.Hand,
                .FocusVisualStyle = Nothing,
                .Template = CreateContentOnlyButtonTemplate(HorizontalAlignment.Left, VerticalAlignment.Center),
                .Background = Brushes.Transparent,
                .BorderBrush = Brushes.Transparent,
                .BorderThickness = New Thickness(0),
                .HorizontalContentAlignment = HorizontalAlignment.Left
            }

            AddHandler tabButton.Click, AddressOf OnTranscriptTabButtonClick

            closeButton = New Button() With {
                .Tag = threadId,
                .Width = 20,
                .Height = 20,
                .MinWidth = 20,
                .MinHeight = 20,
                .Margin = New Thickness(0, 0, 6, 0),
                .Padding = New Thickness(0),
                .Cursor = Cursors.Hand,
                .FocusVisualStyle = Nothing,
                .Template = CreateChromeButtonTemplate(6),
                .Background = Brushes.Transparent,
                .BorderBrush = Brushes.Transparent,
                .BorderThickness = New Thickness(0),
                .Opacity = 0.0R,
                .IsHitTestVisible = False,
                .ToolTip = "Close tab",
                .Content = New TextBlock() With {
                    .Text = "x",
                    .FontSize = 11.2R,
                    .FontWeight = FontWeights.SemiBold,
                    .HorizontalAlignment = HorizontalAlignment.Center,
                    .VerticalAlignment = VerticalAlignment.Center
                }
            }

            AddHandler closeButton.Click, AddressOf OnTranscriptTabCloseButtonClick
            AddHandler closeButton.MouseEnter, AddressOf OnTranscriptTabCloseButtonMouseStateChanged
            AddHandler closeButton.MouseLeave, AddressOf OnTranscriptTabCloseButtonMouseStateChanged

            Dim contentGrid As New Grid()
            contentGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(1, GridUnitType.Star)})
            contentGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = GridLength.Auto})
            Grid.SetColumn(tabButton, 0)
            Grid.SetColumn(closeButton, 1)
            contentGrid.Children.Add(tabButton)
            contentGrid.Children.Add(closeButton)

            Dim chipBorder As New Border() With {
                .Tag = threadId,
                .CornerRadius = New CornerRadius(6, 6, 0, 0),
                .BorderThickness = New Thickness(0),
                .Margin = New Thickness(0, 2, 4, 0),
                .Child = contentGrid
            }

            AddHandler chipBorder.MouseEnter, AddressOf OnTranscriptTabChipMouseEnter
            AddHandler chipBorder.MouseLeave, AddressOf OnTranscriptTabChipMouseLeave

            Return chipBorder
        End Function

        Private Async Sub OnTranscriptTabButtonClick(sender As Object, e As RoutedEventArgs)
            Dim button = TryCast(sender, Button)
            If button Is Nothing Then
                Return
            End If

            If Not _transcriptTabInteractionEnabled Then
                Return
            End If

            Await ActivateTranscriptTabAsync(TryCast(button.Tag, String)).ConfigureAwait(True)
        End Sub

        Private Async Function ActivateTranscriptTabAsync(threadId As String) As Task
            If Not _transcriptTabInteractionEnabled Then
                Return
            End If

            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            If StringComparer.Ordinal.Equals(normalizedThreadId, PendingNewThreadTranscriptTabId) Then
                Await StartThreadAsync().ConfigureAwait(True)
                Return
            End If

            Dim entry = FindVisibleThreadListEntryById(normalizedThreadId)
            If entry Is Nothing Then
                ShowStatus("This tab's thread is hidden by the current filter/search.", isError:=True, displayToast:=True)
                Return
            End If

            SelectThreadEntry(entry, suppressAutoLoad:=False)
        End Function

        Private Sub OnTranscriptTabChipMouseEnter(sender As Object, e As MouseEventArgs)
            Dim chipBorder = TryCast(sender, Border)
            If chipBorder Is Nothing Then
                Return
            End If

            Dim handle = FindTranscriptTabSurfaceHandleByThreadId(TryCast(chipBorder.Tag, String))
            If handle Is Nothing Then
                Return
            End If

            UpdateTranscriptTabButtonVisual(handle,
                                            StringComparer.Ordinal.Equals(handle.ThreadId, _activeTranscriptSurfaceThreadId))
        End Sub

        Private Sub OnTranscriptTabChipMouseLeave(sender As Object, e As MouseEventArgs)
            Dim chipBorder = TryCast(sender, Border)
            If chipBorder Is Nothing Then
                Return
            End If

            Dim handle = FindTranscriptTabSurfaceHandleByThreadId(TryCast(chipBorder.Tag, String))
            If handle Is Nothing Then
                Return
            End If

            UpdateTranscriptTabButtonVisual(handle,
                                            StringComparer.Ordinal.Equals(handle.ThreadId, _activeTranscriptSurfaceThreadId))
        End Sub

        Private Sub OnTranscriptTabCloseButtonMouseStateChanged(sender As Object, e As MouseEventArgs)
            Dim closeButton = TryCast(sender, Button)
            If closeButton Is Nothing Then
                Return
            End If

            Dim handle = FindTranscriptTabSurfaceHandleByThreadId(TryCast(closeButton.Tag, String))
            If handle Is Nothing Then
                Return
            End If

            UpdateTranscriptTabButtonVisual(handle,
                                            StringComparer.Ordinal.Equals(handle.ThreadId, _activeTranscriptSurfaceThreadId))
        End Sub

        Private Sub OnTranscriptTabCloseButtonClick(sender As Object, e As RoutedEventArgs)
            e.Handled = True
            Dim closePerf = Stopwatch.StartNew()

            Dim button = TryCast(sender, Button)
            If button Is Nothing Then
                Return
            End If

            If Not _transcriptTabInteractionEnabled Then
                Return
            End If

            Dim normalizedThreadId = If(TryCast(button.Tag, String), String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            If Not IsTranscriptTabClosable(normalizedThreadId) Then
                AppendProtocol("debug",
                               $"transcript_tab_perf event=close_blocked thread={normalizedThreadId} reason=tab_not_closable openTabs={_transcriptTabSurfacesByThreadId.Count}")
                Return
            End If

            AppendProtocol("debug",
                           $"transcript_tab_perf event=close_begin thread={normalizedThreadId} active={StringComparer.Ordinal.Equals(normalizedThreadId, _activeTranscriptSurfaceThreadId)} openTabs={_transcriptTabSurfacesByThreadId.Count}")

            Dim isActiveTab = StringComparer.Ordinal.Equals(normalizedThreadId, _activeTranscriptSurfaceThreadId)
            Dim fallbackEntry As ThreadListEntry = Nothing
            If isActiveTab Then
                fallbackEntry = FindFirstVisibleTranscriptTabEntryExcluding(normalizedThreadId)
            End If

            Dim removePerf = Stopwatch.StartNew()
            RemoveRetainedTranscriptTabSurface(normalizedThreadId)
            _inactiveTranscriptDocumentsByThreadId.Remove(normalizedThreadId)
            Dim removeMs = removePerf.ElapsedMilliseconds

            If Not isActiveTab Then
                AppendProtocol("debug",
                               $"transcript_tab_perf event=close_complete thread={normalizedThreadId} mode=inactive_tab elapsedMs={closePerf.ElapsedMilliseconds} removeMs={removeMs} remainingTabs={_transcriptTabSurfacesByThreadId.Count}")
                Return
            End If

            If fallbackEntry IsNot Nothing Then
                SelectThreadEntry(fallbackEntry, suppressAutoLoad:=False)
                AppendProtocol("debug",
                               $"transcript_tab_perf event=close_complete thread={normalizedThreadId} mode=switch_to_fallback elapsedMs={closePerf.ElapsedMilliseconds} removeMs={removeMs} fallback={fallbackEntry.Id} remainingTabs={_transcriptTabSurfacesByThreadId.Count}")
                Return
            End If

            Dim cancelLoadPerf = Stopwatch.StartNew()
            CancelActiveThreadSelectionLoad()
            Dim cancelLoadMs = cancelLoadPerf.ElapsedMilliseconds

            Dim resetLoadUiPerf = Stopwatch.StartNew()
            ResetThreadSelectionLoadUiState(hideTranscriptLoader:=True)
            Dim resetLoadUiMs = resetLoadUiPerf.ElapsedMilliseconds

            Dim blankSurfacePerf = Stopwatch.StartNew()
            ActivateFreshTranscriptDocument("tab_close_last_tab", activateBlankSurface:=False)
            Dim transitionedToPendingDraft = EnsurePendingNewThreadTranscriptTabActivated(preferSecondarySurface:=True,
                                                                                          avoidDormantSurfaceReuse:=True)
            If Not transitionedToPendingDraft Then
                ActivateBlankTranscriptSurfacePlaceholder()
            Else
                MarkPrimaryTranscriptSurfaceResetDeferredIfDetached()
            End If
            Dim blankSurfaceMs = blankSurfacePerf.ElapsedMilliseconds

            Dim deferQueueMs As Long = 0
            Dim deferredFinalizeVersion As Integer = 0
            Dim finalizeMode = "queued_blank_placeholder"
            If transitionedToPendingDraft Then
                Dim immediateFinalizePerf = Stopwatch.StartNew()
                ClearPendingUserEchoTracking()
                ClearVisibleSelection()
                SetPendingNewThreadFirstPromptSelectionActive(True, clearThreadSelection:=True)
                RefreshControlStates()
                ShowStatus("New thread ready. Send your first instruction.")
                deferQueueMs = immediateFinalizePerf.ElapsedMilliseconds
                finalizeMode = "immediate_pending_draft"
                AppendProtocol("debug",
                               $"transcript_tab_perf event=close_start_blank_finalize_immediate thread={normalizedThreadId} totalMs={deferQueueMs} pendingDraft=True")
            Else
                Dim deferredFinalizePerf = Stopwatch.StartNew()
                deferredFinalizeVersion = QueueDeferredBlankCloseFinalize(normalizedThreadId)
                deferQueueMs = deferredFinalizePerf.ElapsedMilliseconds
            End If

            _inactiveTranscriptDocumentsByThreadId.Remove(normalizedThreadId)
            AppendProtocol("debug",
                           $"transcript_tab_perf event=close_complete thread={normalizedThreadId} mode=start_blank_deferred elapsedMs={closePerf.ElapsedMilliseconds} removeMs={removeMs} cancelLoadMs={cancelLoadMs} resetLoadUiMs={resetLoadUiMs} blankSurfaceMs={blankSurfaceMs} nextSurfaceMode={If(transitionedToPendingDraft, "pending_draft", "blank_placeholder")} finalizeMode={finalizeMode} queueFinalizeMs={deferQueueMs} finalizeVersion={deferredFinalizeVersion} remainingTabs={_transcriptTabSurfacesByThreadId.Count}")
        End Sub

        Private Sub ActivateBlankTranscriptSurfacePlaceholder()
            EnsureTranscriptTabsUiInitialized()

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return
            End If

            ' Avoid re-binding the primary transcript list during close; that teardown is expensive.
            WorkspacePaneHost.LstTranscript.Visibility = Visibility.Collapsed
            WorkspacePaneHost.LstTranscript.IsHitTestVisible = False

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim candidate = kvp.Value
                If candidate Is Nothing OrElse candidate.SurfaceListBox Is Nothing Then
                    Continue For
                End If

                candidate.SurfaceListBox.Visibility = Visibility.Collapsed
                candidate.SurfaceListBox.IsHitTestVisible = False
            Next

            _activeTranscriptSurfaceListBox = WorkspacePaneHost.LstTranscript
            _activeTranscriptSurfaceThreadId = String.Empty
            _transcriptScrollViewer = Nothing

            UpdateTranscriptTabButtonVisuals()
            RefreshTranscriptTabStripVisibility()
            UpdateWorkspaceEmptyStateVisibility()
        End Sub

        Private Function QueueDeferredBlankCloseFinalize(closedThreadId As String) As Integer
            _deferredBlankCloseFinalizeVersion += 1
            Dim version = _deferredBlankCloseFinalizeVersion
            Dim normalizedClosedThreadId = If(closedThreadId, String.Empty).Trim()
            Dim queuedUtc = DateTimeOffset.UtcNow

            AppendProtocol("debug",
                           $"transcript_tab_perf event=close_start_blank_finalize_queued thread={normalizedClosedThreadId} version={version}")

            Dispatcher.BeginInvoke(
                DispatcherPriority.ApplicationIdle,
                New Action(
                    Sub()
                        RunDeferredBlankCloseFinalize(version, normalizedClosedThreadId, queuedUtc)
                    End Sub))

            Return version
        End Function

        Private Sub RunDeferredBlankCloseFinalize(version As Integer,
                                                  closedThreadId As String,
                                                  queuedUtc As DateTimeOffset)
            If version <> _deferredBlankCloseFinalizeVersion Then
                AppendProtocol("debug",
                               $"transcript_tab_perf event=close_start_blank_finalize_skip reason=stale version={version} currentVersion={_deferredBlankCloseFinalizeVersion}")
                Return
            End If

            If Not String.IsNullOrWhiteSpace(_activeTranscriptSurfaceThreadId) Then
                AppendProtocol("debug",
                               $"transcript_tab_perf event=close_start_blank_finalize_skip reason=active_surface_changed version={version} activeSurfaceThread={_activeTranscriptSurfaceThreadId}")
                Return
            End If

            Dim visibleThreadIdAtFinalize = GetVisibleThreadId()
            If Not String.IsNullOrWhiteSpace(visibleThreadIdAtFinalize) AndAlso
               Not StringComparer.Ordinal.Equals(visibleThreadIdAtFinalize, closedThreadId) Then
                AppendProtocol("debug",
                               $"transcript_tab_perf event=close_start_blank_finalize_skip reason=visible_thread_changed version={version} visibleThread={visibleThreadIdAtFinalize} closedThread={closedThreadId}")
                Return
            End If

            Dim finalizePerf = Stopwatch.StartNew()

            Dim clearPendingUserEchoPerf = Stopwatch.StartNew()
            ClearPendingUserEchoTracking()
            Dim clearPendingUserEchoMs = clearPendingUserEchoPerf.ElapsedMilliseconds

            Dim clearSelectionPerf = Stopwatch.StartNew()
            ClearVisibleSelection()
            Dim clearSelectionMs = clearSelectionPerf.ElapsedMilliseconds

            Dim pendingFirstPromptPerf = Stopwatch.StartNew()
            SetPendingNewThreadFirstPromptSelectionActive(True, clearThreadSelection:=True)
            Dim pendingFirstPromptMs = pendingFirstPromptPerf.ElapsedMilliseconds

            Dim updateLabelsPerf = Stopwatch.StartNew()
            UpdateThreadTurnLabels()
            Dim updateLabelsMs = updateLabelsPerf.ElapsedMilliseconds

            Dim refreshControlsPerf = Stopwatch.StartNew()
            RefreshControlStates()
            Dim refreshControlsMs = refreshControlsPerf.ElapsedMilliseconds

            Dim statusPerf = Stopwatch.StartNew()
            ShowStatus("New thread ready. Send your first instruction.")
            Dim statusMs = statusPerf.ElapsedMilliseconds

            AppendProtocol("debug",
                           $"transcript_tab_perf event=close_start_blank_finalize_complete thread={closedThreadId} version={version} queuedMs={CLng((DateTimeOffset.UtcNow - queuedUtc).TotalMilliseconds)} totalMs={finalizePerf.ElapsedMilliseconds} clearPendingUserEchoMs={clearPendingUserEchoMs} clearSelectionMs={clearSelectionMs} pendingFirstPromptMs={pendingFirstPromptMs} updateLabelsMs={updateLabelsMs} refreshControlsMs={refreshControlsMs} statusMs={statusMs}")
        End Sub

        Private Function FindFirstVisibleTranscriptTabEntryExcluding(threadId As String) As ThreadListEntry
            Dim excludedThreadId = If(threadId, String.Empty).Trim()

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim candidateHandle = kvp.Value
                If candidateHandle Is Nothing OrElse String.IsNullOrWhiteSpace(candidateHandle.ThreadId) Then
                    Continue For
                End If

                If StringComparer.Ordinal.Equals(candidateHandle.ThreadId, excludedThreadId) Then
                    Continue For
                End If

                Dim entry = FindVisibleThreadListEntryById(candidateHandle.ThreadId)
                If entry IsNot Nothing Then
                    Return entry
                End If
            Next

            Return Nothing
        End Function

        Private Function FindTranscriptTabSurfaceHandleByThreadId(threadId As String) As TranscriptTabSurfaceHandle
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return Nothing
            End If

            Dim handle As TranscriptTabSurfaceHandle = Nothing
            If _transcriptTabSurfacesByThreadId.TryGetValue(normalizedThreadId, handle) Then
                Return handle
            End If

            Return Nothing
        End Function

        Private Sub EnsureTranscriptTabSurfaceActivatedForThread(threadId As String)
            EnsureTranscriptTabsUiInitialized()

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return
            End If

            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                ActivateBlankTranscriptSurface()
                Return
            End If

            PromotePendingNewThreadTranscriptTabHandleIfActive(normalizedThreadId)

            Dim handle = EnsureTranscriptTabSurfaceHandle(normalizedThreadId)
            If handle Is Nothing OrElse handle.SurfaceListBox Is Nothing Then
                Return
            End If

            handle.SurfaceListBox.DataContext = _viewModel
            handle.SurfaceListBox.ItemsSource = _viewModel.TranscriptPanel.Items

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim candidate = kvp.Value
                If candidate Is Nothing OrElse candidate.SurfaceListBox Is Nothing Then
                    Continue For
                End If

                Dim isVisible = StringComparer.Ordinal.Equals(candidate.ThreadId, normalizedThreadId)
                candidate.SurfaceListBox.Visibility = If(isVisible, Visibility.Visible, Visibility.Collapsed)
                candidate.SurfaceListBox.IsHitTestVisible = isVisible
            Next

            If Not HandleOwnsPrimaryTranscriptSurface(normalizedThreadId) Then
                WorkspacePaneHost.LstTranscript.Visibility = Visibility.Collapsed
                WorkspacePaneHost.LstTranscript.IsHitTestVisible = False
            End If

            _activeTranscriptSurfaceListBox = handle.SurfaceListBox
            _activeTranscriptSurfaceThreadId = normalizedThreadId
            _transcriptScrollViewer = Nothing
            TouchTranscriptTabStateActivation(normalizedThreadId, "activate_thread_surface")

            UpdateTranscriptTabButtonCaptions()
            UpdateTranscriptTabButtonVisuals()
            RefreshTranscriptTabStripVisibility()
            UpdateWorkspaceEmptyStateVisibility()
        End Sub

        Private Function EnsurePendingNewThreadTranscriptTabActivated(Optional preferSecondarySurface As Boolean = False,
                                                                     Optional avoidDormantSurfaceReuse As Boolean = False) As Boolean
            EnsureTranscriptTabsUiInitialized()

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return False
            End If

            Dim handle = EnsureTranscriptTabSurfaceHandle(PendingNewThreadTranscriptTabId,
                                                          preferSecondarySurface,
                                                          avoidDormantSurfaceReuse)
            If handle Is Nothing OrElse handle.SurfaceListBox Is Nothing Then
                Return False
            End If

            handle.SurfaceListBox.DataContext = _viewModel
            handle.SurfaceListBox.ItemsSource = _viewModel.TranscriptPanel.Items

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim candidate = kvp.Value
                If candidate Is Nothing OrElse candidate.SurfaceListBox Is Nothing Then
                    Continue For
                End If

                Dim isVisible = StringComparer.Ordinal.Equals(candidate.ThreadId, PendingNewThreadTranscriptTabId)
                candidate.SurfaceListBox.Visibility = If(isVisible, Visibility.Visible, Visibility.Collapsed)
                candidate.SurfaceListBox.IsHitTestVisible = isVisible
            Next

            If Not HandleOwnsPrimaryTranscriptSurface(PendingNewThreadTranscriptTabId) Then
                WorkspacePaneHost.LstTranscript.Visibility = Visibility.Collapsed
                WorkspacePaneHost.LstTranscript.IsHitTestVisible = False
            Else
                WorkspacePaneHost.LstTranscript.Visibility = Visibility.Visible
                WorkspacePaneHost.LstTranscript.IsHitTestVisible = True
            End If

            _activeTranscriptSurfaceListBox = handle.SurfaceListBox
            _activeTranscriptSurfaceThreadId = PendingNewThreadTranscriptTabId
            _transcriptScrollViewer = Nothing
            TouchTranscriptTabStateActivation(PendingNewThreadTranscriptTabId, "activate_pending_new_thread")

            UpdateTranscriptTabButtonCaptions()
            UpdateTranscriptTabButtonVisuals()
            RefreshTranscriptTabStripVisibility()
            UpdateWorkspaceEmptyStateVisibility()

            AppendProtocol("debug",
                           $"transcript_tab_perf event=pending_new_thread_tab_activated activeSurfaceThread={_activeTranscriptSurfaceThreadId} openTabs={_transcriptTabSurfacesByThreadId.Count}")
            Return True
        End Function

        Private Function AnyTranscriptTabOwnsPrimarySurface() As Boolean
            Dim primaryListBox = If(WorkspacePaneHost, Nothing)?.LstTranscript
            If primaryListBox Is Nothing Then
                Return False
            End If

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim handle = kvp.Value
                If handle Is Nothing OrElse handle.SurfaceListBox Is Nothing Then
                    Continue For
                End If

                If handle.IsPrimarySurface AndAlso ReferenceEquals(handle.SurfaceListBox, primaryListBox) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Sub MarkPrimaryTranscriptSurfaceResetDeferredIfDetached()
            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return
            End If

            If AnyTranscriptTabOwnsPrimarySurface() Then
                Return
            End If

            _primaryTranscriptSurfaceResetDeferredNeeded = True
            AppendProtocol("debug",
                           $"transcript_tab_perf event=primary_surface_reset_deferred activeTab={If(_activeTranscriptSurfaceThreadId, String.Empty)} dormantCount={_dormantTranscriptTabSurfaces.Count}")
        End Sub

        Private Sub ResetPrimaryTranscriptSurfaceIfDetached()
            _primaryTranscriptSurfaceResetScheduled = False

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return
            End If

            If AnyTranscriptTabOwnsPrimarySurface() Then
                Return
            End If

            If ReferenceEquals(_activeTranscriptSurfaceListBox, WorkspacePaneHost.LstTranscript) Then
                Return
            End If

            Dim resetPerf = Stopwatch.StartNew()
            Dim primaryListBox = WorkspacePaneHost.LstTranscript
            primaryListBox.DataContext = _viewModel
            If _viewModel IsNot Nothing AndAlso _viewModel.TranscriptPanel IsNot Nothing Then
                primaryListBox.ItemsSource = _viewModel.TranscriptPanel.Items
            Else
                primaryListBox.ItemsSource = Nothing
            End If

            _primaryTranscriptSurfaceResetDeferredNeeded = False

            AppendProtocol("debug",
                           $"transcript_tab_perf event=primary_surface_reset_if_detached elapsedMs={resetPerf.ElapsedMilliseconds} activeTab={If(_activeTranscriptSurfaceThreadId, String.Empty)}")
        End Sub

        Private Function IsPendingNewThreadTranscriptTabActive() As Boolean
            Return StringComparer.Ordinal.Equals(_activeTranscriptSurfaceThreadId, PendingNewThreadTranscriptTabId)
        End Function

        Private Sub TryRemoveEmptyPendingNewThreadDraftTabOnExistingSelection(selectedThreadId As String)
            Dim normalizedSelectedThreadId = If(selectedThreadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedSelectedThreadId) Then
                Return
            End If

            Dim pendingState As TranscriptTabState = Nothing
            If Not TryGetTranscriptTabState(PendingNewThreadTranscriptTabId, pendingState) OrElse
               pendingState Is Nothing OrElse
               Not pendingState.AutoRemoveIfEmptyOnExistingSelection Then
                Return
            End If

            Dim pendingTabIsActive = IsPendingNewThreadTranscriptTabActive()
            If Not pendingTabIsActive AndAlso Not _pendingNewThreadFirstPromptSelection Then
                Return
            End If

            If Not String.IsNullOrWhiteSpace(GetVisibleThreadId()) Then
                Return
            End If

            If Not HasRetainedTranscriptTabSurface(PendingNewThreadTranscriptTabId) Then
                Return
            End If

            Dim transcriptItemsCount = 0
            If _viewModel IsNot Nothing AndAlso _viewModel.TranscriptPanel IsNot Nothing Then
                transcriptItemsCount = _viewModel.TranscriptPanel.Items.Count
            End If

            Dim composerHasDraftText = _viewModel IsNot Nothing AndAlso
                                       _viewModel.TurnComposer IsNot Nothing AndAlso
                                       Not String.IsNullOrWhiteSpace(_viewModel.TurnComposer.InputText)
            If composerHasDraftText Then
                Return
            End If

            ' Keep the tab strip stable while switching from pending-new-thread to an existing thread:
            ' promote/retag the pending chip first so tab count does not transiently hit zero.
            If pendingTabIsActive AndAlso Not HasRetainedTranscriptTabSurface(normalizedSelectedThreadId) Then
                PromotePendingNewThreadTranscriptTabHandleIfActive(normalizedSelectedThreadId)
                If HasRetainedTranscriptTabSurface(normalizedSelectedThreadId) Then
                    SetPendingNewThreadFirstPromptSelectionActive(False, clearThreadSelection:=False)
                    TraceTranscriptTabStateSnapshot("pending_promoted_on_existing_selection",
                                                   $"selected={normalizedSelectedThreadId}; transcriptItems={transcriptItemsCount}")
                    AppendProtocol("debug",
                                   $"transcript_tab_perf event=pending_new_thread_tab_promoted_on_select thread={normalizedSelectedThreadId} transcriptItems={transcriptItemsCount}")
                    Return
                End If
            End If

            RemoveRetainedTranscriptTabSurface(PendingNewThreadTranscriptTabId)
            SetPendingNewThreadFirstPromptSelectionActive(False, clearThreadSelection:=False)
            TraceTranscriptTabStateSnapshot("pending_removed_on_existing_selection",
                                           $"selected={normalizedSelectedThreadId}; transcriptItems={transcriptItemsCount}")
            AppendProtocol("debug",
                           $"transcript_tab_perf event=pending_new_thread_tab_removed_on_select thread={normalizedSelectedThreadId} transcriptItems={transcriptItemsCount}")
        End Sub

        Private Sub PromotePendingNewThreadTranscriptTabHandleIfActive(targetThreadId As String)
            Dim normalizedTargetThreadId = If(targetThreadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedTargetThreadId) Then
                Return
            End If

            If Not _pendingNewThreadFirstPromptSelection Then
                Return
            End If

            If Not StringComparer.Ordinal.Equals(_activeTranscriptSurfaceThreadId, PendingNewThreadTranscriptTabId) Then
                Return
            End If

            If _transcriptTabSurfacesByThreadId.ContainsKey(normalizedTargetThreadId) Then
                Return
            End If

            Dim draftHandle As TranscriptTabSurfaceHandle = Nothing
            If Not _transcriptTabSurfacesByThreadId.TryGetValue(PendingNewThreadTranscriptTabId, draftHandle) OrElse
               draftHandle Is Nothing Then
                Return
            End If

            _transcriptTabSurfacesByThreadId.Remove(PendingNewThreadTranscriptTabId)
            RemoveTranscriptTabState(PendingNewThreadTranscriptTabId, "promote_pending_to_thread")
            draftHandle.ThreadId = normalizedTargetThreadId
            If draftHandle.TabButton IsNot Nothing Then
                draftHandle.TabButton.Tag = normalizedTargetThreadId
            End If

            If draftHandle.TabCloseButton IsNot Nothing Then
                draftHandle.TabCloseButton.Tag = normalizedTargetThreadId
            End If

            If draftHandle.TabChipBorder IsNot Nothing Then
                draftHandle.TabChipBorder.Tag = normalizedTargetThreadId
            End If

            _transcriptTabSurfacesByThreadId(normalizedTargetThreadId) = draftHandle
            GetOrCreateTranscriptTabState(normalizedTargetThreadId, "promote_pending_to_thread")
            RecomputeTranscriptTabStatePolicies("promote_pending_to_thread")
            _activeTranscriptSurfaceThreadId = normalizedTargetThreadId
            AppendTranscriptTabChipAtStripEnd(draftHandle.TabChipBorder)

            UpdateTranscriptTabButtonCaptions()
            UpdateTranscriptTabButtonVisual(draftHandle, True)
            AppendProtocol("debug",
                           $"transcript_tab_perf event=pending_new_thread_tab_promoted thread={normalizedTargetThreadId}")
            TraceTranscriptTabStateSnapshot("pending_promoted",
                                           $"tab={normalizedTargetThreadId}")
        End Sub

        Private Sub ActivateBlankTranscriptSurface()
            EnsureTranscriptTabsUiInitialized()

            If WorkspacePaneHost Is Nothing OrElse WorkspacePaneHost.LstTranscript Is Nothing Then
                Return
            End If

            WorkspacePaneHost.LstTranscript.DataContext = _viewModel
            WorkspacePaneHost.LstTranscript.ItemsSource = _viewModel.TranscriptPanel.Items
            WorkspacePaneHost.LstTranscript.Visibility = Visibility.Visible
            WorkspacePaneHost.LstTranscript.IsHitTestVisible = True

            For Each kvp In _transcriptTabSurfacesByThreadId
                Dim candidate = kvp.Value
                If candidate Is Nothing OrElse candidate.SurfaceListBox Is Nothing Then
                    Continue For
                End If

                candidate.SurfaceListBox.Visibility = Visibility.Collapsed
                candidate.SurfaceListBox.IsHitTestVisible = False
            Next

            _activeTranscriptSurfaceListBox = WorkspacePaneHost.LstTranscript
            _activeTranscriptSurfaceThreadId = String.Empty
            _transcriptScrollViewer = Nothing
            TraceTranscriptTabStateSnapshot("activate_blank_surface")

            UpdateTranscriptTabButtonVisuals()
            RefreshTranscriptTabStripVisibility()
            UpdateWorkspaceEmptyStateVisibility()
        End Sub

        Private Function HandleOwnsPrimaryTranscriptSurface(threadId As String) As Boolean
            Dim handle As TranscriptTabSurfaceHandle = Nothing
            If Not _transcriptTabSurfacesByThreadId.TryGetValue(If(threadId, String.Empty), handle) OrElse handle Is Nothing Then
                Return False
            End If

            Return handle.IsPrimarySurface AndAlso ReferenceEquals(handle.SurfaceListBox, If(WorkspacePaneHost, Nothing)?.LstTranscript)
        End Function

        Private Function HasRetainedTranscriptTabSurface(threadId As String) As Boolean
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return False
            End If

            Dim handle As TranscriptTabSurfaceHandle = Nothing
            Return _transcriptTabSurfacesByThreadId.TryGetValue(normalizedThreadId, handle) AndAlso
                   handle IsNot Nothing AndAlso
                   handle.SurfaceListBox IsNot Nothing
        End Function

        Private Sub QueueTranscriptTabSurfaceRetireForIdle(handle As TranscriptTabSurfaceHandle)
            If handle Is Nothing OrElse handle.SurfaceListBox Is Nothing Then
                Return
            End If

            Dim listBox = handle.SurfaceListBox
            listBox.Visibility = Visibility.Collapsed
            listBox.IsHitTestVisible = False

            If TryStoreDormantTranscriptTabSurface(handle) Then
                Return
            End If

            _transcriptTabSurfaceRetireQueue.Enqueue(New TranscriptTabSurfaceRetireWorkItem() With {
                .ThreadId = If(handle.ThreadId, String.Empty).Trim(),
                .SurfaceListBox = listBox,
                .IsPrimarySurface = handle.IsPrimarySurface
            })

            AppendProtocol("debug",
                           $"transcript_tab_perf event=surface_retire_queued thread={If(handle.ThreadId, String.Empty)} queueCount={_transcriptTabSurfaceRetireQueue.Count}")

            ScheduleTranscriptTabSurfaceRetireDrain()
        End Sub

        Private Function TryStoreDormantTranscriptTabSurface(handle As TranscriptTabSurfaceHandle) As Boolean
            If handle Is Nothing OrElse handle.SurfaceListBox Is Nothing OrElse handle.IsPrimarySurface Then
                Return False
            End If

            If _dormantTranscriptTabSurfaces.Count >= TranscriptTabDormantSurfacePoolMax Then
                Return False
            End If

            _dormantTranscriptTabSurfaces.Push(handle.SurfaceListBox)
            AppendProtocol("debug",
                           $"transcript_tab_perf event=surface_retire_dormant thread={If(handle.ThreadId, String.Empty)} dormantCount={_dormantTranscriptTabSurfaces.Count}")
            Return True
        End Function

        Private Function TakeDormantTranscriptTabSurface() As ListBox
            Do While _dormantTranscriptTabSurfaces.Count > 0
                Dim listBox = _dormantTranscriptTabSurfaces.Pop()
                If listBox IsNot Nothing Then
                    Return listBox
                End If
            Loop

            Return Nothing
        End Function

        Private Sub ScheduleTranscriptTabSurfaceRetireDrain()
            If _transcriptTabSurfaceRetireDrainScheduled Then
                Return
            End If

            _transcriptTabSurfaceRetireDrainScheduled = True
            Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle,
                                   New Action(AddressOf DrainTranscriptTabSurfaceRetireQueue))
        End Sub

        Private Sub DrainTranscriptTabSurfaceRetireQueue()
            _transcriptTabSurfaceRetireDrainScheduled = False

            If _transcriptTabSurfaceRetireQueue.Count = 0 Then
                Return
            End If

            Dim retirePerf = Stopwatch.StartNew()
            Dim processed = 0
            Const maxItemsPerIdlePass As Integer = 1
            Const maxIdleBudgetMs As Long = 12

            Do While _transcriptTabSurfaceRetireQueue.Count > 0 AndAlso
                     processed < maxItemsPerIdlePass AndAlso
                     retirePerf.ElapsedMilliseconds <= maxIdleBudgetMs
                Dim workItem = _transcriptTabSurfaceRetireQueue.Dequeue()
                If workItem Is Nothing OrElse workItem.SurfaceListBox Is Nothing Then
                    Continue Do
                End If

                Dim disposePerf = Stopwatch.StartNew()
                DisposeRetiredTranscriptSurfaceWorkItem(workItem)
                Dim disposeMs = disposePerf.ElapsedMilliseconds

                processed += 1

                AppendProtocol("debug",
                               $"transcript_tab_perf event=surface_retire_disposed thread={If(workItem.ThreadId, String.Empty)} queuedMs={CLng((DateTimeOffset.UtcNow - workItem.QueuedUtc).TotalMilliseconds)} disposeMs={disposeMs} remaining={_transcriptTabSurfaceRetireQueue.Count}")
            Loop

            If _transcriptTabSurfaceRetireQueue.Count > 0 Then
                ScheduleTranscriptTabSurfaceRetireDrain()
            End If
        End Sub

        Private Sub DisposeRetiredTranscriptSurfaceWorkItem(workItem As TranscriptTabSurfaceRetireWorkItem)
            If workItem Is Nothing OrElse workItem.SurfaceListBox Is Nothing Then
                Return
            End If

            Dim listBox = workItem.SurfaceListBox

            If Not workItem.IsPrimarySurface AndAlso _transcriptSurfaceHostPanel IsNot Nothing Then
                _transcriptSurfaceHostPanel.Children.Remove(listBox)
            End If

            listBox.ItemsSource = Nothing
            listBox.DataContext = Nothing

            For i = _transcriptInteractionHandlersAttached.Count - 1 To 0 Step -1
                If ReferenceEquals(_transcriptInteractionHandlersAttached(i), listBox) Then
                    _transcriptInteractionHandlersAttached.RemoveAt(i)
                End If
            Next

            For i = _transcriptChunkScrollChangedHandlerAttachedLists.Count - 1 To 0 Step -1
                If ReferenceEquals(_transcriptChunkScrollChangedHandlerAttachedLists(i), listBox) Then
                    _transcriptChunkScrollChangedHandlerAttachedLists.RemoveAt(i)
                End If
            Next
        End Sub

        Private Sub FlushQueuedTranscriptTabSurfaceRetiresImmediately()
            If _transcriptTabSurfaceRetireQueue.Count = 0 Then
                _transcriptTabSurfaceRetireDrainScheduled = False
            Else
                Do While _transcriptTabSurfaceRetireQueue.Count > 0
                    DisposeRetiredTranscriptSurfaceWorkItem(_transcriptTabSurfaceRetireQueue.Dequeue())
                Loop
            End If

            Do While _dormantTranscriptTabSurfaces.Count > 0
                Dim dormantListBox = _dormantTranscriptTabSurfaces.Pop()
                If dormantListBox Is Nothing Then
                    Continue Do
                End If

                DisposeRetiredTranscriptSurfaceWorkItem(New TranscriptTabSurfaceRetireWorkItem() With {
                    .ThreadId = String.Empty,
                    .SurfaceListBox = dormantListBox,
                    .IsPrimarySurface = False
                })
            Loop

            _transcriptTabSurfaceRetireDrainScheduled = False
        End Sub

        Private Sub RemoveRetainedTranscriptTabSurface(threadId As String)
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return
            End If

            Dim handle As TranscriptTabSurfaceHandle = Nothing
            If Not _transcriptTabSurfacesByThreadId.TryGetValue(normalizedThreadId, handle) OrElse handle Is Nothing Then
                Return
            End If

            _transcriptTabSurfacesByThreadId.Remove(normalizedThreadId)
            RemoveTranscriptTabState(normalizedThreadId, "remove_surface")

            If handle.TabChipBorder IsNot Nothing AndAlso _transcriptTabStripPanel IsNot Nothing Then
                _transcriptTabStripPanel.Children.Remove(handle.TabChipBorder)
            End If

            If handle.SurfaceListBox IsNot Nothing Then
                If handle.IsPrimarySurface AndAlso WorkspacePaneHost IsNot Nothing AndAlso ReferenceEquals(handle.SurfaceListBox, WorkspacePaneHost.LstTranscript) Then
                    handle.SurfaceListBox.Visibility = Visibility.Collapsed
                    handle.SurfaceListBox.IsHitTestVisible = False
                Else
                    QueueTranscriptTabSurfaceRetireForIdle(handle)
                End If
            End If

            If StringComparer.Ordinal.Equals(_activeTranscriptSurfaceThreadId, normalizedThreadId) Then
                _activeTranscriptSurfaceThreadId = String.Empty
                _activeTranscriptSurfaceListBox = If(WorkspacePaneHost, Nothing)?.LstTranscript
            End If

            RefreshTranscriptTabStripVisibility()
            UpdateTranscriptTabButtonVisuals()
            TraceTranscriptTabStateSnapshot("surface_removed",
                                           $"tab={normalizedThreadId}")
        End Sub

        Private Sub ClearRetainedTranscriptTabSurfaces()
            EnsureTranscriptTabsUiInitialized()

            Dim surfaceHandles As New List(Of TranscriptTabSurfaceHandle)(_transcriptTabSurfacesByThreadId.Values)
            _transcriptTabSurfacesByThreadId.Clear()
            _transcriptTabStatesByTabId.Clear()
            FlushQueuedTranscriptTabSurfaceRetiresImmediately()
            _primaryTranscriptSurfaceResetDeferredNeeded = False
            _primaryTranscriptSurfaceResetScheduled = False

            For Each handle In surfaceHandles
                If handle Is Nothing Then
                    Continue For
                End If

                If handle.TabChipBorder IsNot Nothing AndAlso _transcriptTabStripPanel IsNot Nothing Then
                    _transcriptTabStripPanel.Children.Remove(handle.TabChipBorder)
                End If

                If handle.SurfaceListBox Is Nothing Then
                    Continue For
                End If

                If WorkspacePaneHost IsNot Nothing AndAlso ReferenceEquals(handle.SurfaceListBox, WorkspacePaneHost.LstTranscript) Then
                    handle.SurfaceListBox.Visibility = Visibility.Visible
                    handle.SurfaceListBox.IsHitTestVisible = True
                    handle.SurfaceListBox.DataContext = _viewModel
                    handle.SurfaceListBox.ItemsSource = _viewModel.TranscriptPanel.Items
                ElseIf _transcriptSurfaceHostPanel IsNot Nothing Then
                    _transcriptSurfaceHostPanel.Children.Remove(handle.SurfaceListBox)
                    handle.SurfaceListBox.ItemsSource = Nothing
                    handle.SurfaceListBox.DataContext = Nothing
                End If
            Next

            Dim primaryListBox = If(WorkspacePaneHost, Nothing)?.LstTranscript
            For i = _transcriptInteractionHandlersAttached.Count - 1 To 0 Step -1
                If primaryListBox Is Nothing OrElse
                   Not ReferenceEquals(_transcriptInteractionHandlersAttached(i), primaryListBox) Then
                    _transcriptInteractionHandlersAttached.RemoveAt(i)
                End If
            Next

            For i = _transcriptChunkScrollChangedHandlerAttachedLists.Count - 1 To 0 Step -1
                If primaryListBox Is Nothing OrElse
                   Not ReferenceEquals(_transcriptChunkScrollChangedHandlerAttachedLists(i), primaryListBox) Then
                    _transcriptChunkScrollChangedHandlerAttachedLists.RemoveAt(i)
                End If
            Next

            If primaryListBox IsNot Nothing Then
                AttachTranscriptInteractionHandlers(primaryListBox)
                _activeTranscriptSurfaceListBox = WorkspacePaneHost.LstTranscript
                primaryListBox.IsHitTestVisible = True
            Else
                _activeTranscriptSurfaceListBox = Nothing
            End If

            _activeTranscriptSurfaceThreadId = String.Empty
            _transcriptScrollViewer = Nothing
            RefreshTranscriptTabStripVisibility()
            TraceTranscriptTabStateSnapshot("clear_all_surfaces")
        End Sub

        Private Sub RefreshTranscriptTabStripVisibility()
            If _transcriptTabStripBorder Is Nothing Then
                Return
            End If

            If _transcriptTabStatesByTabId.Count <> _transcriptTabSurfacesByThreadId.Count Then
                AppendProtocol("debug",
                               $"transcript_tab_state event=surface_state_mismatch surfaceCount={_transcriptTabSurfacesByThreadId.Count} stateCount={_transcriptTabStatesByTabId.Count} activeTab={If(_activeTranscriptSurfaceThreadId, String.Empty)}")
            End If

            _transcriptTabStripBorder.Visibility = If(_transcriptTabSurfacesByThreadId.Count > 0,
                                                      Visibility.Visible,
                                                      Visibility.Collapsed)
            ApplyTranscriptTabStripInteractionState()
            ApplyTranscriptTabResponsiveLayout()
        End Sub

        Private Sub UpdateTranscriptTabButtonCaptions()
            ApplyTranscriptTabResponsiveLayout()
        End Sub

        Private Sub UpdateTranscriptTabButtonCaption(handle As TranscriptTabSurfaceHandle,
                                                     Optional maxChars As Integer = TranscriptTabButtonDefaultMaxChars)
            If handle Is Nothing OrElse handle.TabButton Is Nothing Then
                Return
            End If

            Dim threadId = If(handle.ThreadId, String.Empty).Trim()
            Dim text = ResolveTranscriptTabCaption(threadId, maxChars)
            handle.TabButton.Content = text
            handle.TabButton.ToolTip = threadId
        End Sub

        Private Function ResolveTranscriptTabCaption(threadId As String,
                                                     Optional maxChars As Integer = TranscriptTabButtonDefaultMaxChars) As String
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()
            Dim normalizedMaxChars = Math.Max(1, maxChars)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                Return CompactTranscriptTabCaption("New thread", normalizedMaxChars)
            End If

            If StringComparer.Ordinal.Equals(normalizedThreadId, PendingNewThreadTranscriptTabId) Then
                Return CompactTranscriptTabCaption("New thread", normalizedMaxChars)
            End If

            Return ResolveThreadTitleForUi(normalizedThreadId, normalizedMaxChars)
        End Function
        Private Shared Function CompactTranscriptTabCaption(value As String,
                                                            Optional maxLen As Integer = TranscriptTabButtonDefaultMaxChars) As String
            Dim text = If(value, String.Empty).Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return "Thread"
            End If

            Dim normalizedMaxLen = Math.Max(1, maxLen)
            If text.Length <= normalizedMaxLen Then
                Return text
            End If

            Return text.Substring(0, normalizedMaxLen).TrimEnd() & "..."
        End Function

        Private Sub UpdateTranscriptTabButtonVisuals()
            For Each kvp In _transcriptTabSurfacesByThreadId
                UpdateTranscriptTabButtonVisual(kvp.Value,
                                               StringComparer.Ordinal.Equals(kvp.Key, _activeTranscriptSurfaceThreadId))
            Next
        End Sub

        Private Sub UpdateTranscriptTabButtonVisual(handle As TranscriptTabSurfaceHandle, isActive As Boolean)
            If handle Is Nothing OrElse handle.TabButton Is Nothing Then
                Return
            End If

            Dim button = handle.TabButton
            Dim chipBorder = handle.TabChipBorder
            Dim canInteract = _transcriptTabInteractionEnabled
            Dim isHover = canInteract AndAlso chipBorder IsNot Nothing AndAlso chipBorder.IsMouseOver
            Dim activeBackground = TryCast(TryFindResource("SurfaceBrush"), Brush)
            Dim activeHoverBackground = TryCast(TryFindResource("SurfaceRaisedBrush"), Brush)
            Dim inactiveBackground = TryCast(TryFindResource("SurfaceContrastBrush"), Brush)
            Dim inactiveHoverBackground = TryCast(TryFindResource("TabHoverBackgroundBrush"), Brush)
            Dim activeForeground = TryCast(TryFindResource("TextPrimaryBrush"), Brush)
            Dim inactiveForeground = TryCast(TryFindResource("TextTertiaryBrush"), Brush)

            If chipBorder IsNot Nothing Then
                chipBorder.CornerRadius = New CornerRadius(6, 6, 0, 0)
                If isActive Then
                    chipBorder.Background = If(If(isHover, activeHoverBackground, activeBackground), Brushes.Transparent)
                    chipBorder.BorderBrush = Brushes.Transparent
                    chipBorder.BorderThickness = New Thickness(0)
                    chipBorder.Margin = New Thickness(0, 2, 4, 0)
                    chipBorder.Opacity = 1.0R
                    Panel.SetZIndex(chipBorder, 20)
                Else
                    chipBorder.Background = If(If(isHover, inactiveHoverBackground, inactiveBackground), Brushes.Transparent)
                    chipBorder.BorderBrush = Brushes.Transparent
                    chipBorder.BorderThickness = New Thickness(0)
                    chipBorder.Margin = New Thickness(0, 2, 4, 0)
                    chipBorder.Opacity = If(isHover, 0.92R, 0.72R)
                    Panel.SetZIndex(chipBorder, If(isHover, 12, 8))
                End If

                chipBorder.IsHitTestVisible = canInteract
                If Not canInteract Then
                    chipBorder.Opacity = 0.68R
                End If
            End If

            button.Background = Brushes.Transparent
            button.BorderBrush = Brushes.Transparent
            button.BorderThickness = New Thickness(0)
            button.Foreground = If(If(isActive OrElse isHover, activeForeground, inactiveForeground), Brushes.Black)
            button.FontWeight = If(isActive OrElse isHover, FontWeights.SemiBold, FontWeights.Normal)
            button.Opacity = If(canInteract, 1.0R, 0.82R)
            button.IsEnabled = canInteract
            button.IsHitTestVisible = canInteract
            button.Cursor = If(canInteract, Cursors.Hand, Cursors.Arrow)

            UpdateTranscriptTabCloseButtonVisual(handle,
                                                 isActive,
                                                 If(If(isActive OrElse isHover, activeForeground, inactiveForeground), Brushes.Black),
                                                 canInteract)
        End Sub

        Private Sub UpdateTranscriptTabCloseButtonVisual(handle As TranscriptTabSurfaceHandle,
                                                         isActive As Boolean,
                                                         foreground As Brush,
                                                         canInteract As Boolean)
            If handle Is Nothing OrElse handle.TabCloseButton Is Nothing Then
                Return
            End If

            Dim closeButton = handle.TabCloseButton
            Dim chipBorder = handle.TabChipBorder
            Dim isChipHover = canInteract AndAlso chipBorder IsNot Nothing AndAlso chipBorder.IsMouseOver
            Dim showClose = isActive OrElse isChipHover
            Dim isClosable = IsTranscriptTabClosable(handle.ThreadId)
            showClose = showClose AndAlso isClosable AndAlso canInteract

            closeButton.Opacity = If(showClose, 1.0R, 0.0R)
            closeButton.IsHitTestVisible = showClose
            closeButton.IsEnabled = canInteract AndAlso isClosable
            closeButton.Cursor = If(canInteract, Cursors.Hand, Cursors.Arrow)
            closeButton.Foreground = If(foreground, Brushes.Black)
            closeButton.ToolTip = If(Not canInteract,
                                     "Tabs are unavailable while signed out.",
                                     If(isClosable, "Close tab", "Keep at least one New thread tab open"))
            closeButton.BorderThickness = New Thickness(0)
            closeButton.Background = If(closeButton.IsMouseOver AndAlso showClose,
                                        If(TryCast(TryFindResource("ButtonHoverBrush"), Brush), Brushes.Transparent),
                                        Brushes.Transparent)
            closeButton.BorderBrush = Brushes.Transparent
            closeButton.Opacity = If(showClose, If(closeButton.IsMouseOver, 1.0R, 0.72R), 0.0R)

            Dim closeGlyph = TryCast(closeButton.Content, TextBlock)
            If closeGlyph IsNot Nothing Then
                closeGlyph.Foreground = If(foreground, Brushes.Black)
                closeGlyph.FontWeight = FontWeights.SemiBold
                closeGlyph.Opacity = If(showClose, If(closeButton.IsMouseOver, 1.0R, 0.95R), 0.0R)
            End If
        End Sub

        Private Sub RefreshActiveTranscriptTabCaption()
            If String.IsNullOrWhiteSpace(_activeTranscriptSurfaceThreadId) Then
                Return
            End If

            Dim handle As TranscriptTabSurfaceHandle = Nothing
            If _transcriptTabSurfacesByThreadId.TryGetValue(_activeTranscriptSurfaceThreadId, handle) Then
                UpdateTranscriptTabButtonCaption(handle)
                UpdateTranscriptTabButtonVisual(handle, True)
            End If
        End Sub
    End Class
End Namespace
