Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Services
    Public NotInheritable Class CodexConnectionService
        Implements IConnectionService

        Public Function DetectCodexExecutablePath() As String Implements IConnectionService.DetectCodexExecutablePath
            Dim cmdPreferred = DetectCodexCmdPath()
            If Not String.IsNullOrWhiteSpace(cmdPreferred) Then
                Return cmdPreferred
            End If

            Dim fromWhere = ResolveCommandFromWhere("codex")
            If Not String.IsNullOrWhiteSpace(fromWhere) Then
                Return fromWhere
            End If

            Dim userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)

            Dim candidates As New List(Of String) From {
                Path.Combine(appData, "npm", "codex.cmd"),
                Path.Combine(appData, "npm", "codex.exe"),
                Path.Combine(appData, "npm", "codex.ps1"),
                Path.Combine(localAppData, "Programs", "nodejs", "codex.cmd"),
                Path.Combine(localAppData, "Programs", "nodejs", "codex.exe"),
                Path.Combine(userProfile, "AppData", "Roaming", "npm", "codex.cmd"),
                Path.Combine(userProfile, "AppData", "Roaming", "npm", "codex.exe"),
                Path.Combine(userProfile, ".npm-global", "bin", "codex.cmd"),
                Path.Combine(userProfile, ".npm-global", "bin", "codex.exe")
            }

            For Each candidate In candidates
                If String.IsNullOrWhiteSpace(candidate) Then
                    Continue For
                End If

                If File.Exists(candidate) Then
                    Return candidate
                End If
            Next

            Return String.Empty
        End Function

        Public Function ResolveWindowsCodexExecutable(executable As String) As String Implements IConnectionService.ResolveWindowsCodexExecutable
            Dim candidate = If(executable, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(candidate) Then
                candidate = "codex"
            End If

            Dim prefersCmd = StringComparer.OrdinalIgnoreCase.Equals(candidate, "codex") OrElse
                             StringComparer.OrdinalIgnoreCase.Equals(candidate, "codex.cmd") OrElse
                             StringComparer.OrdinalIgnoreCase.Equals(candidate, "codex.exe")

            If prefersCmd Then
                Dim detectedCmd = DetectCodexCmdPath()
                If Not String.IsNullOrWhiteSpace(detectedCmd) Then
                    Return detectedCmd
                End If
            End If

            If IsPathLike(candidate) Then
                If Not Path.HasExtension(candidate) Then
                    Dim cmdCandidate = candidate & ".cmd"
                    If File.Exists(cmdCandidate) Then
                        Return cmdCandidate
                    End If
                End If

                Return candidate
            End If

            Dim resolved = ResolveCommandFromWhere(candidate)
            If Not String.IsNullOrWhiteSpace(resolved) Then
                Return resolved
            End If

            Return candidate
        End Function

        Public Function BuildLaunchEnvironment(windowsCodexHome As String) As IDictionary(Of String, String) Implements IConnectionService.BuildLaunchEnvironment
            Dim value = If(windowsCodexHome, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(value) Then
                Return Nothing
            End If

            Return New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"CODEX_HOME", value}
            }
        End Function

        Public Function CreateClient() As CodexAppServerClient Implements IConnectionService.CreateClient
            Return New CodexAppServerClient() With {
                .DefaultRequestTimeout = TimeSpan.FromSeconds(45)
            }
        End Function

        Public Async Function StartAndInitializeAsync(client As CodexAppServerClient,
                                                      executablePath As String,
                                                      arguments As String,
                                                      workingDirectory As String,
                                                      environmentVariables As IDictionary(Of String, String),
                                                      cancellationToken As CancellationToken) As Task Implements IConnectionService.StartAndInitializeAsync
            If client Is Nothing Then
                Throw New ArgumentNullException(NameOf(client))
            End If

            Await client.StartAsync(executablePath,
                                    arguments,
                                    workingDirectory,
                                    cancellationToken,
                                    environmentVariables)
            Await client.InitializeAsync("codex-native-agent-vb", "1.0.0", cancellationToken)
        End Function

        Public Async Function StopAsync(client As CodexAppServerClient,
                                        reason As String,
                                        cancellationToken As CancellationToken) As Task Implements IConnectionService.StopAsync
            If client Is Nothing Then
                Return
            End If

            Await client.StopAsync(reason)
        End Function

        Private Shared Function DetectCodexCmdPath() As String
            Dim userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)

            Dim cmdCandidates As New List(Of String) From {
                Path.Combine(appData, "npm", "codex.cmd"),
                Path.Combine(userProfile, "AppData", "Roaming", "npm", "codex.cmd"),
                Path.Combine(localAppData, "Programs", "nodejs", "codex.cmd"),
                Path.Combine(userProfile, ".npm-global", "bin", "codex.cmd")
            }

            For Each candidate In cmdCandidates
                If String.IsNullOrWhiteSpace(candidate) Then
                    Continue For
                End If

                If File.Exists(candidate) Then
                    Return candidate
                End If
            Next

            Dim fromWhereCmd = ResolveCommandFromWhere("codex.cmd")
            If Not String.IsNullOrWhiteSpace(fromWhereCmd) Then
                Return fromWhereCmd
            End If

            Return String.Empty
        End Function

        Private Shared Function ResolveCommandFromWhere(commandName As String) As String
            If String.IsNullOrWhiteSpace(commandName) Then
                Return String.Empty
            End If

            Try
                Dim startInfo As New ProcessStartInfo("where.exe", commandName) With {
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True
                }

                Dim whereProcess As Process = Process.Start(startInfo)
                If whereProcess Is Nothing Then
                    Return String.Empty
                End If

                Using whereProcess
                    Dim output As String = whereProcess.StandardOutput.ReadToEnd()
                    whereProcess.WaitForExit(1500)

                    If whereProcess.ExitCode <> 0 Then
                        Return String.Empty
                    End If

                    Dim separators As String() = {ControlChars.CrLf, ControlChars.Lf, ControlChars.Cr}
                    Dim lines As String() = output.Split(separators, StringSplitOptions.RemoveEmptyEntries)
                    Dim fileMatches As New List(Of String)()

                    For Each rawLine As String In lines
                        Dim line As String = rawLine.Trim()
                        If String.IsNullOrWhiteSpace(line) Then
                            Continue For
                        End If

                        If IsVsCodeExtensionCodexPath(line) Then
                            Continue For
                        End If

                        If File.Exists(line) Then
                            fileMatches.Add(line)
                        End If
                    Next

                    For Each match As String In fileMatches
                        If IsCmdPath(match) Then
                            Return match
                        End If
                    Next

                    If fileMatches.Count > 0 Then
                        Return fileMatches(0)
                    End If
                End Using
            Catch
            End Try

            Return String.Empty
        End Function

        Private Shared Function IsPathLike(value As String) As Boolean
            If String.IsNullOrWhiteSpace(value) Then
                Return False
            End If

            Return value.Contains(Path.DirectorySeparatorChar) OrElse
                   value.Contains(Path.AltDirectorySeparatorChar) OrElse
                   value.Contains(":"c)
        End Function

        Private Shared Function IsCmdPath(path As String) As Boolean
            If String.IsNullOrWhiteSpace(path) Then
                Return False
            End If

            Return path.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function IsVsCodeExtensionCodexPath(path As String) As Boolean
            If String.IsNullOrWhiteSpace(path) Then
                Return False
            End If

            Dim normalized = path.Replace("/"c, "\"c)
            Return normalized.IndexOf("\.vscode\extensions\openai.chatgpt-", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   normalized.IndexOf("\.vscode-insiders\extensions\openai.chatgpt-", StringComparison.OrdinalIgnoreCase) >= 0
        End Function
    End Class
End Namespace
