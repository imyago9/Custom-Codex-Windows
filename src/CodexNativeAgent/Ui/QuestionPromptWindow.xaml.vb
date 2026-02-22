Imports System.Collections.Generic
Imports System.Windows

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class QuestionPromptWindow
        Inherits Window

        Private Const CustomOptionLabel As String = "<Custom...>"
        Private ReadOnly _isSecret As Boolean

        Public Sub New(header As String, prompt As String, options As IEnumerable(Of String), isSecret As Boolean)
            InitializeComponent()

            _isSecret = isSecret
            Title = If(String.IsNullOrWhiteSpace(header), "Input Required", header)
            LblPrompt.Text = If(String.IsNullOrWhiteSpace(prompt), "Provide input:", prompt)

            Dim optionList As New List(Of String)()
            If options IsNot Nothing Then
                For Each optionValue In options
                    If String.IsNullOrWhiteSpace(optionValue) Then
                        Continue For
                    End If

                    optionList.Add(optionValue)
                Next
            End If

            If optionList.Count > 0 Then
                CmbOptions.Visibility = Visibility.Visible
                For Each optionValue In optionList
                    CmbOptions.Items.Add(optionValue)
                Next

                CmbOptions.Items.Add(CustomOptionLabel)
                CmbOptions.SelectedIndex = 0
            End If

            UpdateAnswerInputVisibility()

            AddHandler CmbOptions.SelectionChanged, AddressOf CmbOptionsOnSelectionChanged
            AddHandler BtnOk.Click, AddressOf BtnOkOnClick
        End Sub

        Public ReadOnly Property Answer As String
            Get
                If CmbOptions.Visibility = Visibility.Visible AndAlso CmbOptions.SelectedItem IsNot Nothing Then
                    Dim selected = CmbOptions.SelectedItem.ToString()
                    If Not StringComparer.Ordinal.Equals(selected, CustomOptionLabel) Then
                        Return selected
                    End If
                End If

                If _isSecret AndAlso PwdAnswer.Visibility = Visibility.Visible Then
                    Return PwdAnswer.Password.Trim()
                End If

                Return TxtAnswer.Text.Trim()
            End Get
        End Property

        Private Sub CmbOptionsOnSelectionChanged(sender As Object, e As Controls.SelectionChangedEventArgs)
            UpdateAnswerInputVisibility()
        End Sub

        Private Sub UpdateAnswerInputVisibility()
            Dim shouldShowCustomInput = CmbOptions.Visibility <> Visibility.Visible OrElse
                                        (CmbOptions.SelectedItem IsNot Nothing AndAlso
                                         StringComparer.Ordinal.Equals(CmbOptions.SelectedItem.ToString(), CustomOptionLabel))

            If _isSecret Then
                PwdAnswer.Visibility = If(shouldShowCustomInput, Visibility.Visible, Visibility.Collapsed)
                TxtAnswer.Visibility = Visibility.Collapsed
                If shouldShowCustomInput Then
                    PwdAnswer.Focus()
                End If
            Else
                TxtAnswer.Visibility = If(shouldShowCustomInput, Visibility.Visible, Visibility.Collapsed)
                PwdAnswer.Visibility = Visibility.Collapsed
                If shouldShowCustomInput Then
                    TxtAnswer.Focus()
                End If
            End If
        End Sub

        Private Sub BtnOkOnClick(sender As Object, e As RoutedEventArgs)
            If String.IsNullOrWhiteSpace(Answer) Then
                MessageBox.Show(Me,
                                "Please provide an answer before continuing.",
                                "Missing input",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning)
                Return
            End If

            DialogResult = True
            Close()
        End Sub
    End Class
End Namespace
