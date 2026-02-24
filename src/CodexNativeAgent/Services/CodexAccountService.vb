Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Services
    Public NotInheritable Class CodexAccountService
        Implements IAccountService

        Private ReadOnly _clientProvider As Func(Of CodexAppServerClient)

        Public Sub New(clientProvider As Func(Of CodexAppServerClient))
            If clientProvider Is Nothing Then
                Throw New ArgumentNullException(NameOf(clientProvider))
            End If

            _clientProvider = clientProvider
        End Sub

        Public Async Function ReadAccountAsync(refreshToken As Boolean,
                                               cancellationToken As CancellationToken) As Task(Of AccountReadResult) Implements IAccountService.ReadAccountAsync
            Dim paramsNode As New JsonObject()
            paramsNode("refreshToken") = refreshToken

            Dim responseNode = Await CurrentClient().SendRequestAsync("account/read",
                                                                      paramsNode,
                                                                      cancellationToken:=cancellationToken)
            Dim responseObject = AsObject(responseNode)
            Dim result As New AccountReadResult()

            If responseObject IsNot Nothing Then
                result.RequiresOpenAiAuth = GetPropertyBoolean(responseObject, "requiresOpenaiAuth", False)
                result.Account = GetPropertyObject(responseObject, "account")
            End If

            Return result
        End Function

        Public Async Function ReadRateLimitsAsync(cancellationToken As CancellationToken) As Task(Of JsonNode) Implements IAccountService.ReadRateLimitsAsync
            Return Await CurrentClient().SendRequestAsync("account/rateLimits/read",
                                                          cancellationToken:=cancellationToken)
        End Function

        Public Async Function StartApiKeyLoginAsync(apiKey As String,
                                                    cancellationToken As CancellationToken) As Task Implements IAccountService.StartApiKeyLoginAsync
            Dim paramsNode As New JsonObject()
            paramsNode("type") = "apiKey"
            paramsNode("apiKey") = apiKey

            Await CurrentClient().SendRequestAsync("account/login/start",
                                                   paramsNode,
                                                   cancellationToken:=cancellationToken)
        End Function

        Public Async Function StartChatGptLoginAsync(cancellationToken As CancellationToken) As Task(Of ChatGptLoginStartResult) Implements IAccountService.StartChatGptLoginAsync
            Dim paramsNode As New JsonObject()
            paramsNode("type") = "chatgpt"

            Dim responseNode = Await CurrentClient().SendRequestAsync("account/login/start",
                                                                      paramsNode,
                                                                      cancellationToken:=cancellationToken)
            Dim responseObject = AsObject(responseNode)
            If responseObject Is Nothing Then
                Throw New InvalidOperationException("ChatGPT login response payload was empty.")
            End If

            Return New ChatGptLoginStartResult() With {
                .LoginId = GetPropertyString(responseObject, "loginId"),
                .AuthUrl = GetPropertyString(responseObject, "authUrl")
            }
        End Function

        Public Async Function CancelLoginAsync(loginId As String,
                                               cancellationToken As CancellationToken) As Task Implements IAccountService.CancelLoginAsync
            Dim paramsNode As New JsonObject()
            paramsNode("loginId") = loginId

            Await CurrentClient().SendRequestAsync("account/login/cancel",
                                                   paramsNode,
                                                   cancellationToken:=cancellationToken)
        End Function

        Public Async Function LogoutAsync(cancellationToken As CancellationToken) As Task Implements IAccountService.LogoutAsync
            Await CurrentClient().SendRequestAsync("account/logout",
                                                   cancellationToken:=cancellationToken)
        End Function

        Public Async Function StartExternalTokenLoginAsync(idToken As String,
                                                           accessToken As String,
                                                           cancellationToken As CancellationToken) As Task Implements IAccountService.StartExternalTokenLoginAsync
            Dim paramsNode As New JsonObject()
            paramsNode("type") = "chatgptAuthTokens"
            paramsNode("idToken") = idToken
            paramsNode("accessToken") = accessToken

            Await CurrentClient().SendRequestAsync("account/login/start",
                                                   paramsNode,
                                                   cancellationToken:=cancellationToken)
        End Function

        Private Function CurrentClient() As CodexAppServerClient
            Dim client = _clientProvider()
            If client Is Nothing OrElse Not client.IsRunning Then
                Throw New InvalidOperationException("Not connected to Codex App Server.")
            End If

            Return client
        End Function
    End Class
End Namespace
