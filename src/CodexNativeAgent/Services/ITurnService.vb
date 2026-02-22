Imports System.Threading
Imports System.Threading.Tasks

Namespace CodexNativeAgent.Services
    Public Interface ITurnService
        Function StartTurnAsync(threadId As String,
                                inputText As String,
                                modelId As String,
                                effort As String,
                                approvalPolicy As String,
                                cancellationToken As CancellationToken) As Task(Of TurnStartResult)

        Function SteerTurnAsync(threadId As String,
                                expectedTurnId As String,
                                inputText As String,
                                cancellationToken As CancellationToken) As Task(Of String)

        Function InterruptTurnAsync(threadId As String,
                                    turnId As String,
                                    cancellationToken As CancellationToken) As Task
    End Interface
End Namespace
