Imports System
Imports System.Windows
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Animation

Namespace CodexNativeAgent.Ui
    Public NotInheritable Class ScrollBarInitialFadeBehavior
        Inherits DependencyObject

        Private Sub New()
        End Sub

        Public Shared ReadOnly EnableInitialFadeProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("EnableInitialFade",
                                                GetType(Boolean),
                                                GetType(ScrollBarInitialFadeBehavior),
                                                New PropertyMetadata(False, AddressOf OnEnableInitialFadeChanged))

        Private Shared ReadOnly HasPlayedInitialFadeProperty As DependencyProperty =
            DependencyProperty.RegisterAttached("HasPlayedInitialFade",
                                                GetType(Boolean),
                                                GetType(ScrollBarInitialFadeBehavior),
                                                New PropertyMetadata(False))

        Public Shared Function GetEnableInitialFade(element As DependencyObject) As Boolean
            If element Is Nothing Then
                Return False
            End If

            Return CBool(element.GetValue(EnableInitialFadeProperty))
        End Function

        Public Shared Sub SetEnableInitialFade(element As DependencyObject, value As Boolean)
            If element Is Nothing Then
                Return
            End If

            element.SetValue(EnableInitialFadeProperty, value)
        End Sub

        Private Shared Sub OnEnableInitialFadeChanged(dependencyObject As DependencyObject,
                                                      e As DependencyPropertyChangedEventArgs)
            Dim scrollBar = TryCast(dependencyObject, ScrollBar)
            If scrollBar Is Nothing Then
                Return
            End If

            Dim enabled = False
            If e.NewValue IsNot Nothing Then
                enabled = CBool(e.NewValue)
            End If

            RemoveHandler scrollBar.Loaded, AddressOf OnScrollBarLoaded
            If enabled Then
                AddHandler scrollBar.Loaded, AddressOf OnScrollBarLoaded
            End If
        End Sub

        Private Shared Sub OnScrollBarLoaded(sender As Object, e As RoutedEventArgs)
            Dim scrollBar = TryCast(sender, ScrollBar)
            If scrollBar Is Nothing Then
                Return
            End If

            Dim hasPlayed = CBool(scrollBar.GetValue(HasPlayedInitialFadeProperty))
            If hasPlayed Then
                Return
            End If

            scrollBar.SetValue(HasPlayedInitialFadeProperty, True)

            Dim currentOpacity = scrollBar.Opacity
            If Double.IsNaN(currentOpacity) OrElse Double.IsInfinity(currentOpacity) OrElse currentOpacity <= 0.001R Then
                Return
            End If

            Dim animation As New DoubleAnimation() With {
                .From = 0.0R,
                .To = currentOpacity,
                .Duration = TimeSpan.FromMilliseconds(180),
                .EasingFunction = New QuadraticEase() With {.EasingMode = EasingMode.EaseOut},
                .FillBehavior = FillBehavior.Stop
            }

            scrollBar.BeginAnimation(UIElement.OpacityProperty, animation, HandoffBehavior.SnapshotAndReplace)
        End Sub
    End Class
End Namespace
