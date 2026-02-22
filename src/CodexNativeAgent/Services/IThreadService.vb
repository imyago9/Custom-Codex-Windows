Imports System.Collections.Generic
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks

Namespace CodexNativeAgent.Services
    Public Interface IThreadService
        Function ListModelsAsync(cancellationToken As CancellationToken) As Task(Of IReadOnlyList(Of ModelSummary))

        Function StartThreadAsync(options As ThreadRequestOptions,
                                  cancellationToken As CancellationToken) As Task(Of JsonObject)

        Function ListThreadsAsync(includeArchived As Boolean,
                                  cwd As String,
                                  cancellationToken As CancellationToken) As Task(Of IReadOnlyList(Of ThreadSummary))

        Function ResumeThreadAsync(threadId As String,
                                   options As ThreadRequestOptions,
                                   cancellationToken As CancellationToken) As Task(Of JsonObject)

        Function ReadThreadAsync(threadId As String,
                                 includeTurns As Boolean,
                                 cancellationToken As CancellationToken) As Task(Of JsonObject)

        Function ForkThreadAsync(threadId As String,
                                 options As ThreadRequestOptions,
                                 cancellationToken As CancellationToken) As Task(Of JsonObject)

        Function ArchiveThreadAsync(threadId As String,
                                    cancellationToken As CancellationToken) As Task

        Function UnarchiveThreadAsync(threadId As String,
                                      cancellationToken As CancellationToken) As Task
    End Interface
End Namespace
