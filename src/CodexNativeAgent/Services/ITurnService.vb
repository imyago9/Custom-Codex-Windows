Imports System.Collections.Generic
Imports System.Threading
Imports System.Threading.Tasks

Namespace CodexNativeAgent.Services
    Public Interface ITurnService
        Function StartTurnAsync(threadId As String,
                                inputItems As IReadOnlyList(Of TurnInputItem),
                                modelId As String,
                                effort As String,
                                approvalPolicy As String,
                                cancellationToken As CancellationToken) As Task(Of TurnStartResult)

        Function SteerTurnAsync(threadId As String,
                                expectedTurnId As String,
                                inputItems As IReadOnlyList(Of TurnInputItem),
                                cancellationToken As CancellationToken) As Task(Of String)

        Function InterruptTurnAsync(threadId As String,
                                    turnId As String,
                                    cancellationToken As CancellationToken) As Task
    End Interface
End Namespace
