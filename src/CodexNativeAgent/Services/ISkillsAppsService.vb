Imports System.Collections.Generic
Imports System.Threading
Imports System.Threading.Tasks

Namespace CodexNativeAgent.Services
    Public Interface ISkillsAppsService
        Function ListSkillsAsync(cwds As IReadOnlyList(Of String),
                                 forceReload As Boolean,
                                 cancellationToken As CancellationToken) As Task(Of IReadOnlyList(Of SkillSummary))

        Function ListAppsAsync(threadId As String,
                               forceRefetch As Boolean,
                               cancellationToken As CancellationToken) As Task(Of IReadOnlyList(Of AppSummary))

        Function SetSkillEnabledAsync(path As String,
                                      enabled As Boolean,
                                      cancellationToken As CancellationToken) As Task
    End Interface
End Namespace
