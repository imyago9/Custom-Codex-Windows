Imports System.Collections.Generic
Imports System.Text.Json.Nodes
Imports System.Threading
Imports System.Threading.Tasks
Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Services
    Public NotInheritable Class CodexSkillsAppsService
        Implements ISkillsAppsService

        Private ReadOnly _clientProvider As Func(Of CodexAppServerClient)

        Public Sub New(clientProvider As Func(Of CodexAppServerClient))
            If clientProvider Is Nothing Then
                Throw New ArgumentNullException(NameOf(clientProvider))
            End If

            _clientProvider = clientProvider
        End Sub

        Public Async Function ListSkillsAsync(cwds As IReadOnlyList(Of String),
                                              forceReload As Boolean,
                                              cancellationToken As CancellationToken) As Task(Of IReadOnlyList(Of SkillSummary)) Implements ISkillsAppsService.ListSkillsAsync
            Dim normalizedCwds = BuildNormalizedCwdList(cwds)
            If normalizedCwds.Count = 0 Then
                Return Array.Empty(Of SkillSummary)()
            End If

            Dim paramsNode As New JsonObject()
            Dim cwdsArray As New JsonArray()
            For Each cwd In normalizedCwds
                cwdsArray.Add(cwd)
            Next

            paramsNode("cwds") = cwdsArray
            paramsNode("forceReload") = forceReload

            Dim responseNode = Await CurrentClient().SendRequestAsync("skills/list",
                                                                      paramsNode,
                                                                      cancellationToken:=cancellationToken)
            Dim responseObject = AsObject(responseNode)
            Dim dataArray = GetPropertyArray(responseObject, "data")
            If dataArray Is Nothing OrElse dataArray.Count = 0 Then
                Return Array.Empty(Of SkillSummary)()
            End If

            Dim skills As New List(Of SkillSummary)()
            For Each cwdEntryNode In dataArray
                Dim cwdEntry = AsObject(cwdEntryNode)
                If cwdEntry Is Nothing Then
                    Continue For
                End If

                Dim cwdValue = GetPropertyString(cwdEntry, "cwd")
                Dim skillsArray = GetPropertyArray(cwdEntry, "skills")
                If skillsArray Is Nothing Then
                    Continue For
                End If

                For Each skillNode In skillsArray
                    Dim skillObject = AsObject(skillNode)
                    If skillObject Is Nothing Then
                        Continue For
                    End If

                    Dim skillName = GetPropertyString(skillObject, "name").Trim()
                    If String.IsNullOrWhiteSpace(skillName) Then
                        Continue For
                    End If

                    skills.Add(New SkillSummary() With {
                        .Cwd = If(cwdValue, String.Empty).Trim(),
                        .Name = skillName,
                        .Description = GetPropertyString(skillObject, "description"),
                        .Path = GetPropertyString(skillObject, "path"),
                        .Enabled = GetPropertyBoolean(skillObject, "enabled", True)
                    })
                Next
            Next

            Return skills
        End Function

        Public Async Function ListAppsAsync(threadId As String,
                                            forceRefetch As Boolean,
                                            cancellationToken As CancellationToken) As Task(Of IReadOnlyList(Of AppSummary)) Implements ISkillsAppsService.ListAppsAsync
            Dim apps As New List(Of AppSummary)()
            Dim seenIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim seenCursors As New HashSet(Of String)(StringComparer.Ordinal)
            Dim cursor As String = String.Empty
            Dim normalizedThreadId = If(threadId, String.Empty).Trim()

            Do
                Dim paramsNode As New JsonObject()
                paramsNode("limit") = 100
                paramsNode("forceRefetch") = forceRefetch

                If Not String.IsNullOrWhiteSpace(normalizedThreadId) Then
                    paramsNode("threadId") = normalizedThreadId
                End If

                If Not String.IsNullOrWhiteSpace(cursor) Then
                    paramsNode("cursor") = cursor
                End If

                Dim responseNode = Await CurrentClient().SendRequestAsync("app/list",
                                                                          paramsNode,
                                                                          cancellationToken:=cancellationToken)
                Dim responseObject = AsObject(responseNode)
                Dim dataArray = GetPropertyArray(responseObject, "data")
                If dataArray IsNot Nothing Then
                    For Each appNode In dataArray
                        Dim appObject = AsObject(appNode)
                        If appObject Is Nothing Then
                            Continue For
                        End If

                        Dim appId = GetPropertyString(appObject, "id").Trim()
                        If String.IsNullOrWhiteSpace(appId) OrElse Not seenIds.Add(appId) Then
                            Continue For
                        End If

                        apps.Add(New AppSummary() With {
                            .Id = appId,
                            .Name = GetPropertyString(appObject, "name"),
                            .Description = GetPropertyString(appObject, "description"),
                            .InstallUrl = GetPropertyString(appObject, "installUrl"),
                            .IsAccessible = GetPropertyBoolean(appObject, "isAccessible", False),
                            .IsEnabled = GetPropertyBoolean(appObject, "isEnabled", False)
                        })
                    Next
                End If

                Dim nextCursor = GetPropertyString(responseObject, "nextCursor")
                If String.IsNullOrWhiteSpace(nextCursor) OrElse seenCursors.Contains(nextCursor) Then
                    Exit Do
                End If

                seenCursors.Add(nextCursor)
                cursor = nextCursor
            Loop

            Return apps
        End Function

        Public Async Function SetSkillEnabledAsync(path As String,
                                                   enabled As Boolean,
                                                   cancellationToken As CancellationToken) As Task Implements ISkillsAppsService.SetSkillEnabledAsync
            Dim normalizedPath = If(path, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedPath) Then
                Throw New InvalidOperationException("Skill path is required.")
            End If

            Dim paramsNode As New JsonObject()
            paramsNode("path") = normalizedPath
            paramsNode("enabled") = enabled

            Await CurrentClient().SendRequestAsync("skills/config/write",
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

        Private Shared Function BuildNormalizedCwdList(cwds As IReadOnlyList(Of String)) As IReadOnlyList(Of String)
            Dim output As New List(Of String)()
            If cwds Is Nothing Then
                Return output
            End If

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each cwd In cwds
                Dim normalized = If(cwd, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(normalized) Then
                    Continue For
                End If

                If seen.Add(normalized) Then
                    output.Add(normalized)
                End If
            Next

            Return output
        End Function
    End Class
End Namespace
