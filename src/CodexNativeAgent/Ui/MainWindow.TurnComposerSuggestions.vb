Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Threading

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Structure TurnComposerSuggestionContext
            Public Property Query As String
            Public Property TokenStart As Integer
            Public Property TokenEnd As Integer
        End Structure

        Private _turnComposerSuggestionTokenStart As Integer = -1
        Private _turnComposerSuggestionTokenEnd As Integer = -1
        Private _turnComposerSuggestionCatalogRefreshPending As Boolean
        Private _suppressTurnComposerSuggestionRefresh As Boolean

        Private Sub OnTurnInputTextChangedForSuggestions(sender As Object, e As TextChangedEventArgs)
            RefreshTurnComposerTokenSuggestionsPopup(triggerCatalogWarmup:=True)
        End Sub

        Private Sub OnTurnInputSelectionChangedForSuggestions(sender As Object, e As RoutedEventArgs)
            RefreshTurnComposerTokenSuggestionsPopup(triggerCatalogWarmup:=False)
        End Sub

        Private Sub OnTurnInputLostKeyboardFocusForSuggestions(sender As Object, e As KeyboardFocusChangedEventArgs)
            Dispatcher.BeginInvoke(
                Sub()
                    Dim inputBox = If(WorkspacePaneHost?.TxtTurnInput, Nothing)
                    Dim suggestionList = If(WorkspacePaneHost?.LstTurnComposerTokenSuggestions, Nothing)
                    If inputBox Is Nothing OrElse suggestionList Is Nothing Then
                        HideTurnComposerTokenSuggestionsPopup()
                        Return
                    End If

                    If inputBox.IsKeyboardFocusWithin OrElse suggestionList.IsKeyboardFocusWithin Then
                        Return
                    End If

                    HideTurnComposerTokenSuggestionsPopup()
                End Sub,
                DispatcherPriority.Background)
        End Sub

        Private Sub OnTurnInputPreviewKeyDownForSuggestions(sender As Object, e As KeyEventArgs)
            If e Is Nothing Then
                Return
            End If

            Dim popup = If(WorkspacePaneHost?.TurnComposerTokenSuggestionsPopup, Nothing)
            If popup Is Nothing OrElse Not popup.IsOpen Then
                Return
            End If

            If e.Key = Key.Enter AndAlso Keyboard.Modifiers = ModifierKeys.Control Then
                Return
            End If

            Select Case e.Key
                Case Key.Down
                    If Keyboard.Modifiers = ModifierKeys.None Then
                        MoveTurnComposerSuggestionSelection(1)
                        e.Handled = True
                    End If

                Case Key.Up
                    If Keyboard.Modifiers = ModifierKeys.None Then
                        MoveTurnComposerSuggestionSelection(-1)
                        e.Handled = True
                    End If

                Case Key.Enter
                    If Keyboard.Modifiers = ModifierKeys.None AndAlso ApplySelectedTurnComposerTokenSuggestion() Then
                        e.Handled = True
                    End If

                Case Key.Tab
                    If Keyboard.Modifiers = ModifierKeys.None AndAlso ApplySelectedTurnComposerTokenSuggestion() Then
                        e.Handled = True
                    End If

                Case Key.Escape
                    HideTurnComposerTokenSuggestionsPopup()
                    e.Handled = True
            End Select
        End Sub

        Private Sub OnTurnComposerTokenSuggestionsPreviewMouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
            Dim suggestionList = TryCast(sender, ListBox)
            If suggestionList Is Nothing OrElse e Is Nothing Then
                Return
            End If

            Dim source = TryCast(e.OriginalSource, DependencyObject)
            Dim item = FindVisualAncestor(Of ListBoxItem)(source)
            If item Is Nothing Then
                Return
            End If

            suggestionList.SelectedItem = item.DataContext
            If ApplySelectedTurnComposerTokenSuggestion() Then
                e.Handled = True
            End If
        End Sub

        Private Sub OnTurnComposerTokenSuggestionsPreviewKeyDown(sender As Object, e As KeyEventArgs)
            If e Is Nothing Then
                Return
            End If

            Select Case e.Key
                Case Key.Enter, Key.Tab
                    If Keyboard.Modifiers = ModifierKeys.None AndAlso ApplySelectedTurnComposerTokenSuggestion() Then
                        e.Handled = True
                    End If
                Case Key.Escape
                    HideTurnComposerTokenSuggestionsPopup()
                    Dim inputBox = If(WorkspacePaneHost?.TxtTurnInput, Nothing)
                    If inputBox IsNot Nothing Then
                        inputBox.Focus()
                    End If
                    e.Handled = True
            End Select
        End Sub

        Private Sub RefreshTurnComposerTokenSuggestionsPopup(Optional triggerCatalogWarmup As Boolean = False)
            Dim inputBox = If(WorkspacePaneHost?.TxtTurnInput, Nothing)
            Dim popup = If(WorkspacePaneHost?.TurnComposerTokenSuggestionsPopup, Nothing)
            Dim suggestionList = If(WorkspacePaneHost?.LstTurnComposerTokenSuggestions, Nothing)

            If inputBox Is Nothing OrElse popup Is Nothing OrElse suggestionList Is Nothing Then
                Return
            End If

            If _suppressTurnComposerSuggestionRefresh OrElse
               Not inputBox.IsEnabled OrElse
               Not inputBox.IsKeyboardFocusWithin Then
                HideTurnComposerTokenSuggestionsPopup()
                Return
            End If

            Dim context As New TurnComposerSuggestionContext()
            If Not TryResolveTurnComposerSuggestionContext(inputBox.Text, inputBox.SelectionStart, context) Then
                HideTurnComposerTokenSuggestionsPopup()
                Return
            End If

            Dim suggestions = BuildTurnComposerTokenSuggestions(context.Query, maxItems:=12)
            If suggestions Is Nothing OrElse suggestions.Count = 0 Then
                HideTurnComposerTokenSuggestionsPopup()
                If triggerCatalogWarmup Then
                    TriggerTurnComposerSuggestionCatalogWarmup()
                End If
                Return
            End If

            _turnComposerSuggestionTokenStart = context.TokenStart
            _turnComposerSuggestionTokenEnd = context.TokenEnd

            _suppressTurnComposerSuggestionRefresh = True
            Try
                suggestionList.ItemsSource = suggestions
                If suggestionList.SelectedIndex < 0 OrElse suggestionList.SelectedIndex >= suggestions.Count Then
                    suggestionList.SelectedIndex = 0
                End If
            Finally
                _suppressTurnComposerSuggestionRefresh = False
            End Try

            If Not popup.IsOpen Then
                popup.IsOpen = True
            End If
        End Sub

        Private Sub HideTurnComposerTokenSuggestionsPopup()
            Dim popup = If(WorkspacePaneHost?.TurnComposerTokenSuggestionsPopup, Nothing)
            Dim suggestionList = If(WorkspacePaneHost?.LstTurnComposerTokenSuggestions, Nothing)

            If popup IsNot Nothing Then
                popup.IsOpen = False
            End If

            If suggestionList IsNot Nothing Then
                suggestionList.ItemsSource = Nothing
                suggestionList.SelectedIndex = -1
            End If

            ClearTurnComposerTokenSuggestionTracking()
        End Sub

        Private Sub ClearTurnComposerTokenSuggestionTracking()
            _turnComposerSuggestionTokenStart = -1
            _turnComposerSuggestionTokenEnd = -1
        End Sub

        Private Sub MoveTurnComposerSuggestionSelection(delta As Integer)
            Dim suggestionList = If(WorkspacePaneHost?.LstTurnComposerTokenSuggestions, Nothing)
            If suggestionList Is Nothing Then
                Return
            End If

            Dim count = suggestionList.Items.Count
            If count <= 0 Then
                Return
            End If

            Dim currentIndex = suggestionList.SelectedIndex
            If currentIndex < 0 OrElse currentIndex >= count Then
                currentIndex = 0
            End If

            Dim nextIndex = (currentIndex + delta) Mod count
            If nextIndex < 0 Then
                nextIndex += count
            End If

            suggestionList.SelectedIndex = nextIndex
            suggestionList.ScrollIntoView(suggestionList.SelectedItem)
        End Sub

        Private Function ApplySelectedTurnComposerTokenSuggestion() As Boolean
            Dim inputBox = If(WorkspacePaneHost?.TxtTurnInput, Nothing)
            Dim suggestionList = If(WorkspacePaneHost?.LstTurnComposerTokenSuggestions, Nothing)
            If inputBox Is Nothing OrElse suggestionList Is Nothing Then
                Return False
            End If

            Dim selected = TryCast(suggestionList.SelectedItem, TurnComposerTokenSuggestionEntry)
            If selected Is Nothing Then
                If suggestionList.Items.Count > 0 Then
                    selected = TryCast(suggestionList.Items(0), TurnComposerTokenSuggestionEntry)
                End If
            End If

            If selected Is Nothing Then
                Return False
            End If

            Dim currentText = If(inputBox.Text, String.Empty)
            Dim tokenStart = _turnComposerSuggestionTokenStart
            Dim tokenEnd = _turnComposerSuggestionTokenEnd

            If tokenStart < 0 OrElse tokenEnd < tokenStart OrElse tokenEnd > currentText.Length Then
                Dim context As New TurnComposerSuggestionContext()
                If Not TryResolveTurnComposerSuggestionContext(currentText, inputBox.SelectionStart, context) Then
                    Return False
                End If

                tokenStart = context.TokenStart
                tokenEnd = context.TokenEnd
            End If

            Dim insertToken = NormalizeLookupToken(selected.InsertToken)
            If String.IsNullOrWhiteSpace(insertToken) Then
                Return False
            End If

            Dim replacement = "$" & insertToken
            Dim suffix = If(tokenEnd < currentText.Length, currentText.Substring(tokenEnd), String.Empty)
            Dim spacer = If(tokenEnd >= currentText.Length, " ", String.Empty)
            Dim nextText = currentText.Substring(0, tokenStart) & replacement & spacer & suffix
            Dim nextCaret = tokenStart + replacement.Length + spacer.Length

            _suppressTurnComposerSuggestionRefresh = True
            Try
                inputBox.Text = nextText
                inputBox.SelectionStart = Math.Min(nextCaret, inputBox.Text.Length)
                inputBox.SelectionLength = 0
            Finally
                _suppressTurnComposerSuggestionRefresh = False
            End Try

            HideTurnComposerTokenSuggestionsPopup()
            inputBox.Focus()
            Return True
        End Function

        Private Shared Function TryResolveTurnComposerSuggestionContext(inputText As String,
                                                                        caretIndex As Integer,
                                                                        ByRef context As TurnComposerSuggestionContext) As Boolean
            context = New TurnComposerSuggestionContext() With {
                .Query = String.Empty,
                .TokenStart = -1,
                .TokenEnd = -1
            }

            Dim text = If(inputText, String.Empty)
            If text.Length = 0 Then
                Return False
            End If

            Dim clampedCaret = Math.Max(0, Math.Min(caretIndex, text.Length))
            Dim scanIndex = clampedCaret - 1

            While scanIndex >= 0 AndAlso IsTokenChar(text(scanIndex))
                scanIndex -= 1
            End While

            If scanIndex < 0 OrElse text(scanIndex) <> "$"c Then
                Return False
            End If

            If scanIndex > 0 AndAlso IsTokenChar(text(scanIndex - 1)) Then
                Return False
            End If

            Dim tokenStart = scanIndex
            Dim tokenBodyStart = tokenStart + 1
            Dim tokenEnd = tokenBodyStart

            While tokenEnd < text.Length AndAlso IsTokenChar(text(tokenEnd))
                tokenEnd += 1
            End While

            If clampedCaret < tokenBodyStart OrElse clampedCaret > tokenEnd Then
                Return False
            End If

            context = New TurnComposerSuggestionContext() With {
                .Query = text.Substring(tokenBodyStart, clampedCaret - tokenBodyStart),
                .TokenStart = tokenStart,
                .TokenEnd = tokenEnd
            }
            Return True
        End Function

        Private Sub TriggerTurnComposerSuggestionCatalogWarmup()
            If _turnComposerSuggestionCatalogRefreshPending Then
                Return
            End If

            _turnComposerSuggestionCatalogRefreshPending = True
            FireAndForget(RefreshTurnComposerSuggestionCatalogAndPopupAsync())
        End Sub

        Private Async Function RefreshTurnComposerSuggestionCatalogAndPopupAsync() As Task
            Try
                Await EnsureSkillAndAppCatalogFreshForTurnAsync(GetVisibleThreadId(), forceRefresh:=False).ConfigureAwait(True)
            Finally
                _turnComposerSuggestionCatalogRefreshPending = False
            End Try

            RunOnUi(Sub() RefreshTurnComposerTokenSuggestionsPopup(triggerCatalogWarmup:=False))
        End Function
    End Class
End Namespace
