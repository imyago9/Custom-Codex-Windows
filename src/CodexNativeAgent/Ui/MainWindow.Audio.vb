Imports System.Collections.Generic
Imports System.IO
Imports System.Windows.Media
Imports System.Windows.Threading

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private Const TurnDoneSoundDuplicateSuppressWindowMs As Integer = 1200

        Private ReadOnly _activeUiSoundPlayers As New List(Of MediaPlayer)()
        Private _lastTurnDoneSoundKey As String = String.Empty
        Private _lastTurnDoneSoundPlayedUtc As DateTimeOffset = DateTimeOffset.MinValue

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
