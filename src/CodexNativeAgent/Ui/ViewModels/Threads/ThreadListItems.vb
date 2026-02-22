Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media

Namespace CodexNativeAgent.Ui.ViewModels.Threads
    Public NotInheritable Class ThreadListEntry
        Public Property Id As String = String.Empty
        Public Property Preview As String = String.Empty
        Public Property LastActiveAt As String = String.Empty
        Public Property LastActiveSortTimestamp As Long
        Public Property Cwd As String = String.Empty
        Public Property IsArchived As Boolean

        Public ReadOnly Property ListLeftText As String
            Get
                Return NormalizePreviewSnippet(Preview)
            End Get
        End Property

        Public ReadOnly Property ListRightText As String
            Get
                Return FormatCompactAge(LastActiveSortTimestamp)
            End Get
        End Property

        Public ReadOnly Property ListLeftMargin As Thickness
            Get
                Return New Thickness(18, 0, 0, 0)
            End Get
        End Property

        Public ReadOnly Property ListLeftFontWeight As FontWeight
            Get
                Return FontWeights.Normal
            End Get
        End Property

        Public Overrides Function ToString() As String
            Dim snippet = ListLeftText
            Dim age = ListRightText
            If String.IsNullOrWhiteSpace(age) Then
                Return $"    {snippet}"
            End If

            Return $"    {snippet} | {age}"
        End Function

        Private Shared Function NormalizePreviewSnippet(value As String) As String
            Dim text = If(String.IsNullOrWhiteSpace(value), "(untitled)", value)
            text = text.Replace(ControlChars.Cr, " "c).
                        Replace(ControlChars.Lf, " "c).
                        Replace(ControlChars.Tab, " "c).
                        Trim()

            Do While text.Contains("  ", StringComparison.Ordinal)
                text = text.Replace("  ", " ", StringComparison.Ordinal)
            Loop

            Const maxLength As Integer = 72
            If text.Length > maxLength Then
                Return text.Substring(0, maxLength - 3) & "..."
            End If

            Return text
        End Function

        Private Shared Function FormatCompactAge(unixMilliseconds As Long) As String
            If unixMilliseconds <= 0 OrElse unixMilliseconds = Long.MinValue Then
                Return String.Empty
            End If

            Try
                Dim age = DateTimeOffset.UtcNow - DateTimeOffset.FromUnixTimeMilliseconds(unixMilliseconds)
                If age < TimeSpan.Zero Then
                    age = TimeSpan.Zero
                End If

                If age.TotalMinutes < 1 Then
                    Return "now"
                End If

                If age.TotalHours < 1 Then
                    Return $"{Math.Max(1, CInt(Math.Floor(age.TotalMinutes)))}m"
                End If

                If age.TotalDays < 1 Then
                    Return $"{Math.Max(1, CInt(Math.Floor(age.TotalHours)))}h"
                End If

                If age.TotalDays < 7 Then
                    Return $"{Math.Max(1, CInt(Math.Floor(age.TotalDays)))}d"
                End If

                If age.TotalDays < 30 Then
                    Return $"{Math.Max(1, CInt(Math.Floor(age.TotalDays / 7)))}w"
                End If

                If age.TotalDays < 365 Then
                    Return $"{Math.Max(1, CInt(Math.Floor(age.TotalDays / 30)))}mo"
                End If

                Return $"{Math.Max(1, CInt(Math.Floor(age.TotalDays / 365)))}y"
            Catch
                Return String.Empty
            End Try
        End Function
    End Class

    Public NotInheritable Class ThreadGroupHeaderEntry
        Public Property GroupKey As String = String.Empty
        Public Property ProjectPath As String = String.Empty
        Public Property FolderName As String = String.Empty
        Public Property Count As Integer
        Public Property IsExpanded As Boolean

        Public ReadOnly Property ListLeftText As String
            Get
                Dim folderIcon = Char.ConvertFromUtf32(If(IsExpanded, &H1F4C2, &H1F4C1))
                Return $"{folderIcon} {FolderName}"
            End Get
        End Property

        Public ReadOnly Property ListRightText As String
            Get
                Return Count.ToString()
            End Get
        End Property

        Public ReadOnly Property ListLeftMargin As Thickness
            Get
                Return New Thickness(0)
            End Get
        End Property

        Public ReadOnly Property ListLeftFontWeight As FontWeight
            Get
                Return FontWeights.SemiBold
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return $"{ListLeftText} ({Count})"
        End Function
    End Class

    Public NotInheritable Class ThreadProjectGroup
        Public Property Key As String = String.Empty
        Public Property HeaderLabel As String = String.Empty
        Public Property LatestActivitySortTimestamp As Long = Long.MinValue
        Public ReadOnly Property Threads As New List(Of ThreadListEntry)()
    End Class
End Namespace
