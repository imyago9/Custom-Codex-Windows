Imports System
Imports System.Collections.Generic
Imports System.Windows

Namespace CodexNativeAgent.Ui
    Friend NotInheritable Class AppAppearanceManager
        Public Const LightTheme As String = "light"
        Public Const DarkTheme As String = "dark"
        Public Const ComfortableDensity As String = "comfortable"
        Public Const CompactDensity As String = "compact"

        Private Const ControlsPath As String = "Ui/Styles/Controls.Shared.xaml"
        Private Const LightThemePath As String = "Ui/Styles/Theme.Light.xaml"
        Private Const DarkThemePath As String = "Ui/Styles/Theme.Dark.xaml"
        Private Const ComfortableDensityPath As String = "Ui/Styles/Density.Comfortable.xaml"
        Private Const CompactDensityPath As String = "Ui/Styles/Density.Compact.xaml"

        Private Sub New()
        End Sub

        Public Shared Sub Initialize(app As Application)
            If app Is Nothing Then
                Return
            End If

            Dim merged = app.Resources.MergedDictionaries
            EnsureDictionary(merged, "Controls.Shared.xaml", ControlsPath)
            ReplaceDictionary(merged, "Density.", ComfortableDensityPath)
            ReplaceDictionary(merged, "Theme.", LightThemePath)
        End Sub

        Public Shared Function NormalizeTheme(value As String) As String
            If StringComparer.OrdinalIgnoreCase.Equals(value, DarkTheme) Then
                Return DarkTheme
            End If

            Return LightTheme
        End Function

        Public Shared Function NormalizeDensity(value As String) As String
            If StringComparer.OrdinalIgnoreCase.Equals(value, CompactDensity) Then
                Return CompactDensity
            End If

            Return ComfortableDensity
        End Function

        Public Shared Sub ApplyTheme(theme As String)
            Dim app = Application.Current
            If app Is Nothing Then
                Return
            End If

            EnsureDictionary(app.Resources.MergedDictionaries, "Controls.Shared.xaml", ControlsPath)
            Dim normalized = NormalizeTheme(theme)
            Dim path = If(StringComparer.OrdinalIgnoreCase.Equals(normalized, DarkTheme),
                          DarkThemePath,
                          LightThemePath)
            ReplaceDictionary(app.Resources.MergedDictionaries, "Theme.", path)
        End Sub

        Public Shared Sub ApplyDensity(density As String)
            Dim app = Application.Current
            If app Is Nothing Then
                Return
            End If

            EnsureDictionary(app.Resources.MergedDictionaries, "Controls.Shared.xaml", ControlsPath)
            Dim normalized = NormalizeDensity(density)
            Dim path = If(StringComparer.OrdinalIgnoreCase.Equals(normalized, CompactDensity),
                          CompactDensityPath,
                          ComfortableDensityPath)
            ReplaceDictionary(app.Resources.MergedDictionaries, "Density.", path)
        End Sub

        Public Shared Function ToggleTheme(currentTheme As String) As String
            If StringComparer.OrdinalIgnoreCase.Equals(NormalizeTheme(currentTheme), DarkTheme) Then
                Return LightTheme
            End If

            Return DarkTheme
        End Function

        Public Shared Function ThemeButtonLabel(currentTheme As String) As String
            If StringComparer.OrdinalIgnoreCase.Equals(NormalizeTheme(currentTheme), DarkTheme) Then
                Return "Switch to Light Mode"
            End If

            Return "Switch to Dark Mode"
        End Function

        Public Shared Function DisplayTheme(currentTheme As String) As String
            If StringComparer.OrdinalIgnoreCase.Equals(NormalizeTheme(currentTheme), DarkTheme) Then
                Return "Dark"
            End If

            Return "Light"
        End Function

        Private Shared Sub EnsureDictionary(merged As IList(Of ResourceDictionary),
                                            sourceToken As String,
                                            path As String)
            If merged Is Nothing Then
                Return
            End If

            For Each dictionary In merged
                If DictionarySourceContains(dictionary, sourceToken) Then
                    Return
                End If
            Next

            merged.Add(New ResourceDictionary With {.Source = New Uri(path, UriKind.Relative)})
        End Sub

        Private Shared Sub ReplaceDictionary(merged As IList(Of ResourceDictionary),
                                             sourceToken As String,
                                             path As String)
            If merged Is Nothing Then
                Return
            End If

            For index = merged.Count - 1 To 0 Step -1
                If DictionarySourceContains(merged(index), sourceToken) Then
                    merged.RemoveAt(index)
                End If
            Next

            merged.Add(New ResourceDictionary With {.Source = New Uri(path, UriKind.Relative)})
        End Sub

        Private Shared Function DictionarySourceContains(dictionary As ResourceDictionary,
                                                         sourceToken As String) As Boolean
            If dictionary Is Nothing OrElse dictionary.Source Is Nothing Then
                Return False
            End If

            Return dictionary.Source.OriginalString.IndexOf(sourceToken, StringComparison.OrdinalIgnoreCase) >= 0
        End Function
    End Class
End Namespace
