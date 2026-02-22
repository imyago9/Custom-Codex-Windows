Imports System.Collections.Generic
Imports System.Threading
Imports System.Threading.Tasks
Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Services
    Public Interface IConnectionService
        Function DetectCodexExecutablePath() As String

        Function ResolveWindowsCodexExecutable(executable As String) As String

        Function BuildLaunchEnvironment(windowsCodexHome As String) As IDictionary(Of String, String)

        Function CreateClient() As CodexAppServerClient

        Function StartAndInitializeAsync(client As CodexAppServerClient,
                                         executablePath As String,
                                         arguments As String,
                                         workingDirectory As String,
                                         environmentVariables As IDictionary(Of String, String),
                                         cancellationToken As CancellationToken) As Task

        Function StopAsync(client As CodexAppServerClient,
                           reason As String,
                           cancellationToken As CancellationToken) As Task
    End Interface
End Namespace
