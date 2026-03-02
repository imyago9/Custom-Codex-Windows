Imports System.Collections.Generic
Imports System.IO
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
                                             inputItems As IReadOnlyList(Of TurnInputItem),
                                             modelId As String,
                                             effort As String,
                                             approvalPolicy As String,
                                             cancellationToken As CancellationToken) As Task(Of TurnStartResult) Implements ITurnService.StartTurnAsync
            Dim paramsNode As New JsonObject()
            paramsNode("threadId") = threadId
            paramsNode("input") = BuildInputItems(inputItems)

            If Not String.IsNullOrWhiteSpace(modelId) Then
                paramsNode("model") = modelId
            End If

            If Not String.IsNullOrWhiteSpace(effort) Then
                paramsNode("effort") = effort
            End If

            Dim normalizedApprovalPolicy = NormalizeApprovalPolicy(approvalPolicy)
            If Not String.IsNullOrWhiteSpace(normalizedApprovalPolicy) Then
                paramsNode("approvalPolicy") = normalizedApprovalPolicy
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
                                             inputItems As IReadOnlyList(Of TurnInputItem),
                                             cancellationToken As CancellationToken) As Task(Of String) Implements ITurnService.SteerTurnAsync
            Dim paramsNode As New JsonObject()
            paramsNode("threadId") = threadId
            paramsNode("expectedTurnId") = expectedTurnId
            paramsNode("input") = BuildInputItems(inputItems)

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

        Private Shared Function BuildInputItems(inputItems As IReadOnlyList(Of TurnInputItem)) As JsonArray
            Dim list As New JsonArray()
            Dim effectiveItems = inputItems
            If effectiveItems Is Nothing OrElse effectiveItems.Count = 0 Then
                effectiveItems = New List(Of TurnInputItem) From {
                    TurnInputItem.TextItem(String.Empty)
                }
            End If

            For Each inputItem In effectiveItems
                Dim item = NormalizeInputItem(inputItem)
                If item IsNot Nothing Then
                    list.Add(item)
                End If
            Next

            If list.Count = 0 Then
                list.Add(NormalizeInputItem(TurnInputItem.TextItem(String.Empty)))
            End If

            Return list
        End Function

        Private Shared Function NormalizeInputItem(inputItem As TurnInputItem) As JsonObject
            Dim safeItem = If(inputItem, TurnInputItem.TextItem(String.Empty))
            Dim itemType = If(safeItem.Type, String.Empty).Trim().ToLowerInvariant()
            If String.IsNullOrWhiteSpace(itemType) Then
                itemType = "text"
            End If

            Select Case itemType
                Case "text"
                    Dim item As New JsonObject()
                    item("type") = "text"
                    item("text") = If(safeItem.Text, String.Empty)
                    Return item

                Case "skill"
                    Dim skillPath = If(safeItem.Path, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(skillPath) Then
                        Return Nothing
                    End If

                    Dim skillName = If(safeItem.Name, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(skillName) Then
                        skillName = Path.GetFileNameWithoutExtension(skillPath)
                    End If

                    If String.IsNullOrWhiteSpace(skillName) Then
                        Return Nothing
                    End If

                    Dim item As New JsonObject()
                    item("type") = "skill"
                    item("name") = skillName
                    item("path") = skillPath
                    Return item

                Case "mention"
                    Dim mentionPath = If(safeItem.Path, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(mentionPath) Then
                        Return Nothing
                    End If

                    Dim mentionName = If(safeItem.Name, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(mentionName) Then
                        mentionName = mentionPath
                    End If

                    Dim item As New JsonObject()
                    item("type") = "mention"
                    item("name") = mentionName
                    item("path") = mentionPath
                    Return item

                Case Else
                    Throw New InvalidOperationException($"Unsupported turn input item type '{safeItem.Type}'.")
            End Select
        End Function
    End Class
End Namespace
