Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Services
    Public NotInheritable Class CodexTurnService
        Implements ITurnService

        Private ReadOnly _clientProvider As Func(Of CodexAppServerClient)

        Public Sub New(clientProvider As Func(Of CodexAppServerClient))
            If clientProvider Is Nothing Then
                Throw New ArgumentNullException(NameOf(clientProvider))
            End If

            _clientProvider = clientProvider
        End Sub

        Public Async Function StartTurnAsync(threadId As String,
                                             inputText As String,
                                             modelId As String,
                                             effort As String,
                                             approvalPolicy As String,
                                             cancellationToken As CancellationToken) As Task(Of TurnStartResult) Implements ITurnService.StartTurnAsync
            Dim paramsNode As New JsonObject()
            paramsNode("threadId") = threadId
            paramsNode("input") = BuildTextInput(inputText)

            If Not String.IsNullOrWhiteSpace(modelId) Then
                paramsNode("model") = modelId
            End If

            If Not String.IsNullOrWhiteSpace(effort) Then
                paramsNode("effort") = effort
            End If

            If Not String.IsNullOrWhiteSpace(approvalPolicy) Then
                paramsNode("approvalPolicy") = approvalPolicy
            End If

            Dim responseNode = Await CurrentClient().SendRequestAsync("turn/start",
                                                                      paramsNode,
                                                                      cancellationToken:=cancellationToken)
            Dim responseObject = AsObject(responseNode)
            Dim turnObject = GetPropertyObject(responseObject, "turn")

            Return New TurnStartResult() With {
                .TurnId = GetPropertyString(turnObject, "id")
            }
        End Function

        Public Async Function SteerTurnAsync(threadId As String,
                                             expectedTurnId As String,
                                             inputText As String,
                                             cancellationToken As CancellationToken) As Task(Of String) Implements ITurnService.SteerTurnAsync
            Dim paramsNode As New JsonObject()
            paramsNode("threadId") = threadId
            paramsNode("expectedTurnId") = expectedTurnId
            paramsNode("input") = BuildTextInput(inputText)

            Dim responseNode = Await CurrentClient().SendRequestAsync("turn/steer",
                                                                      paramsNode,
                                                                      cancellationToken:=cancellationToken)
            Dim responseObject = AsObject(responseNode)
            Return GetPropertyString(responseObject, "turnId")
        End Function

        Public Async Function InterruptTurnAsync(threadId As String,
                                                 turnId As String,
                                                 cancellationToken As CancellationToken) As Task Implements ITurnService.InterruptTurnAsync
            Dim paramsNode As New JsonObject()
            paramsNode("threadId") = threadId
            paramsNode("turnId") = turnId

            Await CurrentClient().SendRequestAsync("turn/interrupt",
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

        Private Shared Function BuildTextInput(text As String) As JsonArray
            Dim list As New JsonArray()
            Dim item As New JsonObject()
            item("type") = "text"
            item("text") = text
            list.Add(item)
            Return list
        End Function
    End Class
End Namespace
