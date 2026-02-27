Imports System
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Threading

Namespace CodexNativeAgent.Ui
    Public NotInheritable Class ScrollViewerActivityBehavior
        Inherits DependencyObject

        Private NotInheritable Class ScrollActivityState
            Public Sub New(owner As ScrollViewer)
                HideTimer = New DispatcherTimer() With {
                    .Interval = TimeSpan.FromMilliseconds(900)
                }
                AddHandler HideTimer.Tick,
                    Sub(sender, e)
                        HandleHideTimerTick(owner)
                    End Sub
            End Sub

            Public ReadOnly Property HideTimer As DispatcherTimer
        End Class

        Private Shared ReadOnly _states As New Dictionary(Of ScrollViewer, ScrollActivityState)()

        Private Sub New()
        End Sub

        Public Shared ReadOnly EnableAutoHideProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("EnableAutoHide",
                                                GetType(Boolean),
                                                GetType(ScrollViewerActivityBehavior),
                                                New PropertyMetadata(False, AddressOf OnEnableAutoHideChanged))

        Public Shared Function GetEnableAutoHide(element As DependencyObject) As Boolean
            If element Is Nothing Then
                Return False
            End If

            Return CBool(element.GetValue(EnableAutoHideProperty))
        End Function

        Public Shared Sub SetEnableAutoHide(element As DependencyObject, value As Boolean)
            If element Is Nothing Then
                Return
            End If

            element.SetValue(EnableAutoHideProperty, value)
        End Sub

        Public Shared ReadOnly IsScrollActivityVisibleProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("IsScrollActivityVisible",
                                                GetType(Boolean),
                                                GetType(ScrollViewerActivityBehavior),
                                                New PropertyMetadata(False))

        Public Shared Function GetIsScrollActivityVisible(element As DependencyObject) As Boolean
            If element Is Nothing Then
                Return False
            End If

            Return CBool(element.GetValue(IsScrollActivityVisibleProperty))
        End Function

        Public Shared Sub SetIsScrollActivityVisible(element As DependencyObject, value As Boolean)
            If element Is Nothing Then
                Return
            End If

            element.SetValue(IsScrollActivityVisibleProperty, value)
        End Sub

        Private Shared Sub OnEnableAutoHideChanged(dependencyObject As DependencyObject,
                                                   e As DependencyPropertyChangedEventArgs)
            Dim viewer = TryCast(dependencyObject, ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            Dim enabled = False
            If e.NewValue IsNot Nothing Then
                enabled = CBool(e.NewValue)
            End If

            If enabled Then
                Attach(viewer)
            Else
                Detach(viewer)
            End If
        End Sub

        Private Shared Sub Attach(viewer As ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            Dim alreadyAttached = _states.ContainsKey(viewer)
            Dim state = GetOrCreateState(viewer)
            state.HideTimer.Stop()
            SetIsScrollActivityVisible(viewer, viewer.IsMouseOver)

            If alreadyAttached Then
                Return
            End If

            AddHandler viewer.ScrollChanged, AddressOf OnViewerScrollChanged
            AddHandler viewer.MouseEnter, AddressOf OnViewerMouseEnter
            AddHandler viewer.MouseLeave, AddressOf OnViewerMouseLeave
            AddHandler viewer.Unloaded, AddressOf OnViewerUnloaded
        End Sub

        Private Shared Sub Detach(viewer As ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            RemoveHandler viewer.ScrollChanged, AddressOf OnViewerScrollChanged
            RemoveHandler viewer.MouseEnter, AddressOf OnViewerMouseEnter
            RemoveHandler viewer.MouseLeave, AddressOf OnViewerMouseLeave
            RemoveHandler viewer.Unloaded, AddressOf OnViewerUnloaded

            Dim state As ScrollActivityState = Nothing
            If _states.TryGetValue(viewer, state) Then
                state.HideTimer.Stop()
                _states.Remove(viewer)
            End If

            SetIsScrollActivityVisible(viewer, False)
        End Sub

        Private Shared Function GetOrCreateState(viewer As ScrollViewer) As ScrollActivityState
            Dim state As ScrollActivityState = Nothing
            If _states.TryGetValue(viewer, state) Then
                Return state
            End If

            state = New ScrollActivityState(viewer)
            _states(viewer) = state
            Return state
        End Function

        Private Shared Sub OnViewerUnloaded(sender As Object, e As RoutedEventArgs)
            Dim viewer = TryCast(sender, ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            Detach(viewer)
        End Sub

        Private Shared Sub OnViewerMouseEnter(sender As Object, e As MouseEventArgs)
            Dim viewer = TryCast(sender, ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            Dim state = GetOrCreateState(viewer)
            state.HideTimer.Stop()
            SetIsScrollActivityVisible(viewer, True)
        End Sub

        Private Shared Sub OnViewerMouseLeave(sender As Object, e As MouseEventArgs)
            Dim viewer = TryCast(sender, ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            ScheduleHide(viewer)
        End Sub

        Private Shared Sub OnViewerScrollChanged(sender As Object, e As ScrollChangedEventArgs)
            Dim viewer = TryCast(sender, ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            If Math.Abs(e.HorizontalChange) > 0.0R OrElse
               Math.Abs(e.VerticalChange) > 0.0R OrElse
               Math.Abs(e.ExtentHeightChange) > 0.0R OrElse
               Math.Abs(e.ExtentWidthChange) > 0.0R Then
                ShowForActivity(viewer)
            End If
        End Sub

        Private Shared Sub ShowForActivity(viewer As ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            Dim state = GetOrCreateState(viewer)
            SetIsScrollActivityVisible(viewer, True)
            state.HideTimer.Stop()
            If Not viewer.IsMouseOver Then
                state.HideTimer.Start()
            End If
        End Sub

        Private Shared Sub ScheduleHide(viewer As ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            Dim state = GetOrCreateState(viewer)
            state.HideTimer.Stop()
            If viewer.IsMouseOver Then
                Return
            End If

            state.HideTimer.Start()
        End Sub

        Private Shared Sub HandleHideTimerTick(viewer As ScrollViewer)
            If viewer Is Nothing Then
                Return
            End If

            Dim state As ScrollActivityState = Nothing
            If Not _states.TryGetValue(viewer, state) Then
                Return
            End If

            If viewer.IsMouseOver Then
                state.HideTimer.Stop()
                Return
            End If

            state.HideTimer.Stop()
            SetIsScrollActivityVisible(viewer, False)
        End Sub
    End Class
End Namespace
