Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks

Namespace CodexNativeAgent.Services
    Public Interface IAccountService
        Function ReadAccountAsync(refreshToken As Boolean,
                                  cancellationToken As CancellationToken) As Task(Of AccountReadResult)

        Function ReadRateLimitsAsync(cancellationToken As CancellationToken) As Task(Of JsonNode)

        Function StartApiKeyLoginAsync(apiKey As String,
                                       cancellationToken As CancellationToken) As Task

        Function StartChatGptLoginAsync(cancellationToken As CancellationToken) As Task(Of ChatGptLoginStartResult)

        Function CancelLoginAsync(loginId As String,
                                  cancellationToken As CancellationToken) As Task

        Function LogoutAsync(cancellationToken As CancellationToken) As Task

        Function StartExternalTokenLoginAsync(idToken As String,
                                              accessToken As String,
                                              cancellationToken As CancellationToken) As Task
    End Interface
End Namespace
