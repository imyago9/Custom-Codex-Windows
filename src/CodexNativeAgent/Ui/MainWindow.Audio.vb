Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Const TurnDoneSoundDuplicateSuppressWindowMs As Integer = 1200
        Private Const TurnFailedSoundDuplicateSuppressWindowMs As Integer = 1500
        Private Const GitCommitSoundDuplicateSuppressWindowMs As Integer = 900
        Private Const GitPushSoundDuplicateSuppressWindowMs As Integer = 900
        Private Const ApprovalNeededSoundDuplicateSuppressWindowMs As Integer = 1400
        Private Const GeneralErrorSoundDuplicateSuppressWindowMs As Integer = 900

        Private ReadOnly _activeUiSoundPlayers As New List(Of MediaPlayer)()
        Private _lastTurnDoneSoundKey As String = String.Empty
        Private _lastTurnDoneSoundPlayedUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastTurnFailedSoundKey As String = String.Empty
        Private _lastTurnFailedSoundPlayedUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastGitCommitSoundPlayedUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastGitPushSoundPlayedUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastApprovalNeededSoundKey As String = String.Empty
        Private _lastApprovalNeededSoundPlayedUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _lastGeneralErrorSoundKey As String = String.Empty
        Private _lastGeneralErrorSoundPlayedUtc As DateTimeOffset = DateTimeOffset.MinValue

        Private Sub PlayLoadThreadSoundIfEnabled()
            PlayUiSoundIfEnabled("load-thread.mp3")
        End Sub

        Private Sub PlayThreadSelectionSoundIfEnabled()
            PlayLoadThreadSoundIfEnabled()
        End Sub

        Private Sub PlayTurnDoneSoundIfEnabled(threadId As String, turnId As String)
            Dim soundKey = $"{If(threadId, String.Empty).Trim()}:{If(turnId, String.Empty).Trim()}"
            If StringComparer.Ordinal.Equals(soundKey, _lastTurnDoneSoundKey) AndAlso
               (DateTimeOffset.UtcNow - _lastTurnDoneSoundPlayedUtc).TotalMilliseconds < TurnDoneSoundDuplicateSuppressWindowMs Then
                Return
            End If

            If PlayUiSoundIfEnabled("turn-done.mp3") Then
                _lastTurnDoneSoundKey = soundKey
                _lastTurnDoneSoundPlayedUtc = DateTimeOffset.UtcNow
            End If
        End Sub

        Private Sub PlayTurnFailedSoundIfEnabled(threadId As String, turnId As String, status As String)
            Dim normalizedStatus = NormalizeStatusSoundToken(status)
            Dim soundKey = $"{If(threadId, String.Empty).Trim()}:{If(turnId, String.Empty).Trim()}:{normalizedStatus}"
            If StringComparer.Ordinal.Equals(soundKey, _lastTurnFailedSoundKey) AndAlso
               (DateTimeOffset.UtcNow - _lastTurnFailedSoundPlayedUtc).TotalMilliseconds < TurnFailedSoundDuplicateSuppressWindowMs Then
                Return
            End If

            If PlayUiSoundIfEnabled("turn-failed.mp3") Then
                _lastTurnFailedSoundKey = soundKey
                _lastTurnFailedSoundPlayedUtc = DateTimeOffset.UtcNow
            End If
        End Sub

        Private Sub PlayGitCommitSoundIfEnabled()
            If (DateTimeOffset.UtcNow - _lastGitCommitSoundPlayedUtc).TotalMilliseconds < GitCommitSoundDuplicateSuppressWindowMs Then
                Return
            End If

            If PlayUiSoundIfEnabled("git-commit.mp3") Then
                _lastGitCommitSoundPlayedUtc = DateTimeOffset.UtcNow
            End If
        End Sub

        Private Sub PlayGitPushSoundIfEnabled()
            If (DateTimeOffset.UtcNow - _lastGitPushSoundPlayedUtc).TotalMilliseconds < GitPushSoundDuplicateSuppressWindowMs Then
                Return
            End If

            If PlayUiSoundIfEnabled("git-push.mp3") Then
                _lastGitPushSoundPlayedUtc = DateTimeOffset.UtcNow
            End If
        End Sub

        Private Sub PlayApprovalNeededSoundIfEnabled(statusMessage As String)
            Dim soundKey = BuildStatusSoundKey(statusMessage)
            If StringComparer.Ordinal.Equals(soundKey, _lastApprovalNeededSoundKey) AndAlso
               (DateTimeOffset.UtcNow - _lastApprovalNeededSoundPlayedUtc).TotalMilliseconds < ApprovalNeededSoundDuplicateSuppressWindowMs Then
                Return
            End If

            If PlayUiSoundIfEnabled("approval-needed.mp3") Then
                _lastApprovalNeededSoundKey = soundKey
                _lastApprovalNeededSoundPlayedUtc = DateTimeOffset.UtcNow
            End If
        End Sub

        Private Sub PlayGeneralErrorSoundIfEnabled(statusMessage As String)
            Dim soundKey = BuildStatusSoundKey(statusMessage)
            If StringComparer.Ordinal.Equals(soundKey, _lastGeneralErrorSoundKey) AndAlso
               (DateTimeOffset.UtcNow - _lastGeneralErrorSoundPlayedUtc).TotalMilliseconds < GeneralErrorSoundDuplicateSuppressWindowMs Then
                Return
            End If

            If PlayUiSoundIfEnabled("general-error.mp3") Then
                _lastGeneralErrorSoundKey = soundKey
                _lastGeneralErrorSoundPlayedUtc = DateTimeOffset.UtcNow
            End If
        End Sub

        Private Shared Function IsApprovalNeededStatusMessage(message As String) As Boolean
            Dim normalized = If(message, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            If normalized.StartsWith("Approval queued:", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Return normalized.IndexOf("approval required", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Shared Function IsTurnFailureStatusMessage(message As String) As Boolean
            Dim normalized = NormalizeStatusSoundToken(message)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            Return normalized.StartsWith("turn failed", StringComparison.Ordinal) OrElse
                   normalized.StartsWith("turn interrupted", StringComparison.Ordinal) OrElse
                   normalized.StartsWith("turn canceled", StringComparison.Ordinal) OrElse
                   normalized.StartsWith("turn cancelled", StringComparison.Ordinal) OrElse
                   normalized.StartsWith("turn aborted", StringComparison.Ordinal)
        End Function

        Private Shared Function ShouldPlayGeneralErrorStatusSound(message As String, isError As Boolean) As Boolean
            If isError Then
                Return True
            End If

            Dim normalized = NormalizeStatusSoundToken(message)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            If IsTurnFailureStatusMessage(normalized) Then
                Return False
            End If

            If normalized.StartsWith("warning", StringComparison.Ordinal) OrElse
               normalized.StartsWith("warn:", StringComparison.Ordinal) OrElse
               normalized.IndexOf("warning", StringComparison.Ordinal) >= 0 Then
                Return True
            End If

            If normalized.StartsWith("error", StringComparison.Ordinal) OrElse
               normalized.Contains(" error", StringComparison.Ordinal) OrElse
               normalized.StartsWith("failed", StringComparison.Ordinal) OrElse
               normalized.Contains(" failed", StringComparison.Ordinal) OrElse
               normalized.Contains("could not", StringComparison.Ordinal) OrElse
               normalized.Contains("unable to", StringComparison.Ordinal) Then
                Return True
            End If

            Return False
        End Function

        Private Shared Function BuildStatusSoundKey(message As String) As String
            Dim normalized = NormalizeStatusSoundToken(message)
            If normalized.Length > 220 Then
                normalized = normalized.Substring(0, 220)
            End If

            Return normalized
        End Function

        Private Shared Function NormalizeStatusSoundToken(value As String) As String
            Return If(value, String.Empty).Trim().ToLowerInvariant()
        End Function

        Private Function PlayUiSoundIfEnabled(soundFileName As String) As Boolean
            If Not Dispatcher.CheckAccess() Then
                Dispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    New Action(
                        Sub()
                            PlayUiSoundIfEnabled(soundFileName)
                        End Sub))
                Return False
            End If

            If _viewModel Is Nothing OrElse _viewModel.SettingsPanel Is Nothing Then
                Return False
            End If

            If Not _viewModel.SettingsPanel.PlayUiSounds Then
                Return False
            End If

            Dim soundPath = ResolveUiSoundFilePath(soundFileName)
            If String.IsNullOrWhiteSpace(soundPath) OrElse Not File.Exists(soundPath) Then
                AppendProtocol("debug", $"sound_playback_skipped missing_file={If(soundFileName, String.Empty)}")
                Return False
            End If

            Try
                Dim volumePercent = _viewModel.SettingsPanel.UiSoundVolumePercent
                If Double.IsNaN(volumePercent) OrElse Double.IsInfinity(volumePercent) Then
                    volumePercent = 100.0R
                End If

                volumePercent = Math.Max(0.0R, Math.Min(100.0R, volumePercent))
                Dim player As New MediaPlayer() With {
                    .Volume = volumePercent / 100.0R
                }

                AddHandler player.MediaEnded,
                    Sub(sender, e)
                        CleanupUiSoundPlayer(TryCast(sender, MediaPlayer))
                    End Sub
                AddHandler player.MediaFailed,
                    Sub(sender, e)
                        Dim failureMessage = If(e?.ErrorException?.Message, "Unknown media playback error.")
                        AppendProtocol("debug", $"sound_playback_failed file={If(soundFileName, String.Empty)} error={failureMessage}")
                        CleanupUiSoundPlayer(TryCast(sender, MediaPlayer))
                    End Sub

                _activeUiSoundPlayers.Add(player)
                player.Open(New Uri(soundPath, UriKind.Absolute))
                player.Play()
                Return True
            Catch ex As Exception
                AppendProtocol("debug", $"sound_playback_failed file={If(soundFileName, String.Empty)} error={ex.Message}")
                Return False
            End Try
        End Function

        Private Shared Function ResolveUiSoundFilePath(soundFileName As String) As String
            Dim normalizedFileName = If(soundFileName, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedFileName) Then
                Return String.Empty
            End If

            Dim outputPath = Path.Combine(AppContext.BaseDirectory, "Assets", "sounds", normalizedFileName)
            If File.Exists(outputPath) Then
                Return outputPath
            End If

            Dim repoPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
                                                         "..",
                                                         "..",
                                                         "..",
                                                         "..",
                                                         "..",
                                                         "assets",
                                                         "sounds",
                                                         normalizedFileName))
            If File.Exists(repoPath) Then
                Return repoPath
            End If

            Return outputPath
        End Function

        Private Sub CleanupUiSoundPlayer(player As MediaPlayer)
            If player Is Nothing Then
                Return
            End If

            _activeUiSoundPlayers.Remove(player)

            Try
                player.Stop()
            Catch
            End Try

            Try
                player.Close()
            Catch
            End Try
        End Sub

        Private Sub StopAndDisposeUiSoundPlayers()
            If Not Dispatcher.CheckAccess() Then
                Dispatcher.Invoke(
                    DispatcherPriority.Send,
                    New Action(
                        Sub()
                            StopAndDisposeUiSoundPlayers()
                        End Sub))
                Return
            End If

            If _activeUiSoundPlayers.Count = 0 Then
                Return
            End If

            Dim players = _activeUiSoundPlayers.ToArray()
            _activeUiSoundPlayers.Clear()
            For Each player In players
                If player Is Nothing Then
                    Continue For
                End If

                Try
                    player.Stop()
                Catch
                End Try

                Try
                    player.Close()
                Catch
                End Try
            Next
        End Sub
    End Class
End Namespace
