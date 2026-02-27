Imports System.Globalization
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.Json.Nodes
Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Ui.Coordinators
    Friend Module TurnFlowEventModelHelpers
        Friend Function NormalizeIdentifier(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Return value.Trim()
        End Function

        Friend Function CloneJsonObject(obj As JsonObject) As JsonObject
            Return TryCast(CloneJsonNode(obj), JsonObject)
        End Function

        Friend Function CloneJsonNode(node As JsonNode) As JsonNode
            If node Is Nothing Then
                Return Nothing
            End If

            Return JsonNode.Parse(node.ToJsonString())
        End Function
    End Module

    Public Enum ItemDeltaKind
        AgentMessage
        Plan
        ReasoningSummaryText
        ReasoningSummaryPartAdded
        ReasoningText
        CommandExecutionOutput
        FileChangeOutput
    End Enum

    Public Enum ApprovalRequestKind
        CommandExecution
        FileChange
    End Enum

    Public MustInherit Class TurnFlowEvent
        Protected Sub New(methodName As String, rawParams As JsonObject)
            Me.MethodName = If(methodName, String.Empty)
            Me.RawParams = CloneJsonObject(rawParams)
        End Sub

        Public ReadOnly Property MethodName As String
        Public ReadOnly Property RawParams As JsonObject
    End Class

    Public MustInherit Class TurnScopedEvent
        Inherits TurnFlowEvent

        Protected Sub New(methodName As String,
                          rawParams As JsonObject,
                          threadId As String,
                          turnId As String)
            MyBase.New(methodName, rawParams)
            Me.ThreadId = NormalizeIdentifier(threadId)
            Me.TurnId = NormalizeIdentifier(turnId)
        End Sub

        Public ReadOnly Property ThreadId As String
        Public ReadOnly Property TurnId As String
    End Class

    Public NotInheritable Class TurnStartedEvent
        Inherits TurnScopedEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       turnId As String)
            MyBase.New(methodName, rawParams, threadId, turnId)
        End Sub
    End Class

    Public NotInheritable Class TurnCompletedEvent
        Inherits TurnScopedEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       turnId As String,
                       status As String,
                       errorMessage As String)
            MyBase.New(methodName, rawParams, threadId, turnId)
            Me.Status = NormalizeIdentifier(status)
            Me.ErrorMessage = If(errorMessage, String.Empty).Trim()
        End Sub

        Public ReadOnly Property Status As String
        Public ReadOnly Property ErrorMessage As String
    End Class

    Public NotInheritable Class TurnDiffUpdatedEvent
        Inherits TurnScopedEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       turnId As String,
                       summaryText As String)
            MyBase.New(methodName, rawParams, threadId, turnId)
            Me.SummaryText = If(summaryText, String.Empty)
        End Sub

        Public ReadOnly Property SummaryText As String
    End Class

    Public NotInheritable Class TurnPlanUpdatedEvent
        Inherits TurnScopedEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       turnId As String,
                       summaryText As String)
            MyBase.New(methodName, rawParams, threadId, turnId)
            Me.SummaryText = If(summaryText, String.Empty)
        End Sub

        Public ReadOnly Property SummaryText As String
    End Class

    Public NotInheritable Class ThreadTokenUsageUpdatedEvent
        Inherits TurnScopedEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       turnId As String,
                       tokenUsage As JsonObject)
            MyBase.New(methodName, rawParams, threadId, turnId)
            Me.TokenUsage = CloneJsonObject(tokenUsage)
        End Sub

        Public ReadOnly Property TokenUsage As JsonObject
    End Class

    Public NotInheritable Class ItemLifecycleEvent
        Inherits TurnScopedEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       turnId As String,
                       itemId As String,
                       itemObject As JsonObject,
                       isCompleted As Boolean)
            MyBase.New(methodName, rawParams, threadId, turnId)
            Me.ItemId = NormalizeIdentifier(itemId)
            Me.ItemObject = CloneJsonObject(itemObject)
            Me.IsCompleted = isCompleted
        End Sub

        Public ReadOnly Property ItemId As String
        Public ReadOnly Property ItemObject As JsonObject
        Public ReadOnly Property IsCompleted As Boolean
    End Class

    Public NotInheritable Class ItemDeltaEvent
        Inherits TurnScopedEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       turnId As String,
                       itemId As String,
                       deltaKind As ItemDeltaKind,
                       deltaText As String,
                       summaryPartText As String)
            MyBase.New(methodName, rawParams, threadId, turnId)
            Me.ItemId = NormalizeIdentifier(itemId)
            Me.DeltaKind = deltaKind
            Me.DeltaText = If(deltaText, String.Empty)
            Me.SummaryPartText = If(summaryPartText, String.Empty)
        End Sub

        Public ReadOnly Property ItemId As String
        Public ReadOnly Property DeltaKind As ItemDeltaKind
        Public ReadOnly Property DeltaText As String
        Public ReadOnly Property SummaryPartText As String
    End Class

    Public NotInheritable Class ApprovalRequestEvent
        Inherits TurnScopedEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       requestId As JsonNode,
                       threadId As String,
                       turnId As String,
                       itemId As String,
                       requestKind As ApprovalRequestKind)
            MyBase.New(methodName, rawParams, threadId, turnId)
            Me.RequestId = CloneJsonNode(requestId)
            Me.ItemId = NormalizeIdentifier(itemId)
            Me.RequestKind = requestKind
        End Sub

        Public ReadOnly Property RequestId As JsonNode
        Public ReadOnly Property ItemId As String
        Public ReadOnly Property RequestKind As ApprovalRequestKind
    End Class

    Public NotInheritable Class ErrorEvent
        Inherits TurnScopedEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       turnId As String,
                       message As String,
                       willRetry As Boolean?)
            MyBase.New(methodName, rawParams, threadId, turnId)
            Me.Message = If(message, String.Empty)
            Me.WillRetry = willRetry
        End Sub

        Public ReadOnly Property Message As String
        Public ReadOnly Property WillRetry As Boolean?
    End Class

    Public NotInheritable Class ThreadStartedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       threadObject As JsonObject)
            MyBase.New(methodName, rawParams)
            Me.ThreadId = NormalizeIdentifier(threadId)
            Me.ThreadObject = CloneJsonObject(threadObject)
        End Sub

        Public ReadOnly Property ThreadId As String
        Public ReadOnly Property ThreadObject As JsonObject
    End Class

    Public NotInheritable Class ThreadArchivedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String, rawParams As JsonObject, threadId As String)
            MyBase.New(methodName, rawParams)
            Me.ThreadId = NormalizeIdentifier(threadId)
        End Sub

        Public ReadOnly Property ThreadId As String
    End Class

    Public NotInheritable Class ThreadUnarchivedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String, rawParams As JsonObject, threadId As String)
            MyBase.New(methodName, rawParams)
            Me.ThreadId = NormalizeIdentifier(threadId)
        End Sub

        Public ReadOnly Property ThreadId As String
    End Class

    Public NotInheritable Class ThreadStatusChangedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       threadId As String,
                       statusSummary As String)
            MyBase.New(methodName, rawParams)
            Me.ThreadId = NormalizeIdentifier(threadId)
            Me.StatusSummary = If(statusSummary, String.Empty)
        End Sub

        Public ReadOnly Property ThreadId As String
        Public ReadOnly Property StatusSummary As String
    End Class

    Public NotInheritable Class AccountLoginCompletedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       loginId As String,
                       success As Boolean,
                       errorMessage As String)
            MyBase.New(methodName, rawParams)
            Me.LoginId = NormalizeIdentifier(loginId)
            Me.Success = success
            Me.ErrorMessage = If(errorMessage, String.Empty)
        End Sub

        Public ReadOnly Property LoginId As String
        Public ReadOnly Property Success As Boolean
        Public ReadOnly Property ErrorMessage As String
    End Class

    Public NotInheritable Class AccountUpdatedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String, rawParams As JsonObject)
            MyBase.New(methodName, rawParams)
        End Sub
    End Class

    Public NotInheritable Class AccountRateLimitsUpdatedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String, rawParams As JsonObject, payload As JsonObject)
            MyBase.New(methodName, rawParams)
            Me.Payload = CloneJsonObject(payload)
        End Sub

        Public ReadOnly Property Payload As JsonObject
    End Class

    Public NotInheritable Class ModelReroutedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       fromModel As String,
                       toModel As String,
                       reason As String)
            MyBase.New(methodName, rawParams)
            Me.FromModel = If(fromModel, String.Empty)
            Me.ToModel = If(toModel, String.Empty)
            Me.Reason = If(reason, String.Empty)
        End Sub

        Public ReadOnly Property FromModel As String
        Public ReadOnly Property ToModel As String
        Public ReadOnly Property Reason As String
    End Class

    Public NotInheritable Class AppListUpdatedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String, rawParams As JsonObject, appCount As Integer)
            MyBase.New(methodName, rawParams)
            Me.AppCount = Math.Max(0, appCount)
        End Sub

        Public ReadOnly Property AppCount As Integer
    End Class

    Public NotInheritable Class McpOauthLoginCompletedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       name As String,
                       success As Boolean,
                       errorMessage As String)
            MyBase.New(methodName, rawParams)
            Me.Name = If(name, String.Empty)
            Me.Success = success
            Me.ErrorMessage = If(errorMessage, String.Empty)
        End Sub

        Public ReadOnly Property Name As String
        Public ReadOnly Property Success As Boolean
        Public ReadOnly Property ErrorMessage As String
    End Class

    Public NotInheritable Class WindowsSandboxSetupCompletedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       mode As String,
                       success As Boolean,
                       errorMessage As String)
            MyBase.New(methodName, rawParams)
            Me.Mode = If(mode, String.Empty)
            Me.Success = success
            Me.ErrorMessage = If(errorMessage, String.Empty)
        End Sub

        Public ReadOnly Property Mode As String
        Public ReadOnly Property Success As Boolean
        Public ReadOnly Property ErrorMessage As String
    End Class

    Public NotInheritable Class FuzzyFileSearchSessionUpdatedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String,
                       rawParams As JsonObject,
                       sessionId As String,
                       query As String,
                       fileCount As Integer)
            MyBase.New(methodName, rawParams)
            Me.SessionId = NormalizeIdentifier(sessionId)
            Me.Query = If(query, String.Empty)
            Me.FileCount = Math.Max(0, fileCount)
        End Sub

        Public ReadOnly Property SessionId As String
        Public ReadOnly Property Query As String
        Public ReadOnly Property FileCount As Integer
    End Class

    Public NotInheritable Class FuzzyFileSearchSessionCompletedNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String, rawParams As JsonObject, sessionId As String)
            MyBase.New(methodName, rawParams)
            Me.SessionId = NormalizeIdentifier(sessionId)
        End Sub

        Public ReadOnly Property SessionId As String
    End Class

    Public NotInheritable Class GenericServerNotificationEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String, rawParams As JsonObject)
            MyBase.New(methodName, rawParams)
        End Sub
    End Class

    Public NotInheritable Class GenericServerRequestEvent
        Inherits TurnFlowEvent

        Public Sub New(methodName As String, rawParams As JsonObject, requestId As JsonNode)
            MyBase.New(methodName, rawParams)
            Me.RequestId = CloneJsonNode(requestId)
        End Sub

        Public ReadOnly Property RequestId As JsonNode
    End Class

    Public NotInheritable Class TurnFlowEventParser
        Private Sub New()
        End Sub

        Public Shared Function ParseNotification(methodName As String, paramsNode As JsonNode) As TurnFlowEvent
            Dim paramsObject = AsJsonObject(paramsNode)
            Dim normalizedMethodName = NormalizeIdentifier(methodName)
            NormalizeLegacyNotificationEnvelope(normalizedMethodName, paramsObject)

            Select Case normalizedMethodName
                Case "thread/started"
                    Dim threadObject = ReadObject(paramsObject, "thread")
                    Return New ThreadStartedNotificationEvent(normalizedMethodName,
                                                              paramsObject,
                                                              ReadThreadId(paramsObject),
                                                              threadObject)

                Case "thread/archived"
                    Return New ThreadArchivedNotificationEvent(normalizedMethodName,
                                                               paramsObject,
                                                               ReadThreadId(paramsObject))

                Case "thread/unarchived"
                    Return New ThreadUnarchivedNotificationEvent(normalizedMethodName,
                                                                 paramsObject,
                                                                 ReadThreadId(paramsObject))

                Case "thread/status/changed"
                    Return New ThreadStatusChangedNotificationEvent(normalizedMethodName,
                                                                    paramsObject,
                                                                    ReadThreadId(paramsObject),
                                                                    ReadThreadStatusSummary(paramsObject))

                Case "account/login/completed"
                    Return New AccountLoginCompletedNotificationEvent(normalizedMethodName,
                                                                      paramsObject,
                                                                      ReadStringFirst(paramsObject, "loginId", "login_id"),
                                                                      ReadBoolean(paramsObject, "success", False),
                                                                      ReadStringFirst(paramsObject, "error"))

                Case "account/updated"
                    Return New AccountUpdatedNotificationEvent(normalizedMethodName, paramsObject)

                Case "account/rateLimits/updated"
                    Return New AccountRateLimitsUpdatedNotificationEvent(normalizedMethodName,
                                                                         paramsObject,
                                                                         paramsObject)

                Case "model/rerouted"
                    Return New ModelReroutedNotificationEvent(normalizedMethodName,
                                                              paramsObject,
                                                              ReadStringFirst(paramsObject, "fromModel", "from_model"),
                                                              ReadStringFirst(paramsObject, "toModel", "to_model"),
                                                              ReadStringFirst(paramsObject, "reason"))

                Case "app/list/updated"
                    Return New AppListUpdatedNotificationEvent(normalizedMethodName,
                                                               paramsObject,
                                                               CountApps(paramsObject))

                Case "mcpServer/oauthLogin/completed"
                    Return New McpOauthLoginCompletedNotificationEvent(normalizedMethodName,
                                                                       paramsObject,
                                                                       ReadStringFirst(paramsObject, "name"),
                                                                       ReadBoolean(paramsObject, "success", False),
                                                                       ReadStringFirst(paramsObject, "error"))

                Case "windowsSandbox/setupCompleted"
                    Return New WindowsSandboxSetupCompletedNotificationEvent(normalizedMethodName,
                                                                             paramsObject,
                                                                             ReadStringFirst(paramsObject, "mode"),
                                                                             ReadBoolean(paramsObject, "success", False),
                                                                             ReadStringFirst(paramsObject, "error"))

                Case "fuzzyFileSearch/sessionUpdated"
                    Return New FuzzyFileSearchSessionUpdatedNotificationEvent(normalizedMethodName,
                                                                              paramsObject,
                                                                              ReadStringFirst(paramsObject, "sessionId", "session_id"),
                                                                              ReadStringFirst(paramsObject, "query"),
                                                                              CountFiles(paramsObject))

                Case "fuzzyFileSearch/sessionCompleted"
                    Return New FuzzyFileSearchSessionCompletedNotificationEvent(normalizedMethodName,
                                                                                paramsObject,
                                                                                ReadStringFirst(paramsObject, "sessionId", "session_id"))

                Case "turn/started"
                    Return New TurnStartedEvent(normalizedMethodName,
                                                paramsObject,
                                                ReadThreadId(paramsObject),
                                                ReadTurnId(paramsObject))

                Case "turn/completed"
                    Return New TurnCompletedEvent(normalizedMethodName,
                                                  paramsObject,
                                                  ReadThreadId(paramsObject),
                                                  ReadTurnId(paramsObject),
                                                  ReadTurnCompletionStatus(paramsObject),
                                                  ReadErrorMessage(paramsObject))

                Case "turn/diff/updated"
                    Return New TurnDiffUpdatedEvent(normalizedMethodName,
                                                    paramsObject,
                                                    ReadThreadId(paramsObject),
                                                    ReadTurnId(paramsObject),
                                                    ReadTurnDiffSummary(paramsObject))

                Case "turn/plan/updated"
                    Return New TurnPlanUpdatedEvent(normalizedMethodName,
                                                    paramsObject,
                                                    ReadThreadId(paramsObject),
                                                    ReadTurnId(paramsObject),
                                                    ReadTurnPlanSummary(paramsObject))

                Case "thread/tokenUsage/updated"
                    Return New ThreadTokenUsageUpdatedEvent(normalizedMethodName,
                                                            paramsObject,
                                                            ReadThreadId(paramsObject),
                                                            ReadTurnId(paramsObject),
                                                            ReadTokenUsageObject(paramsObject))

                Case "item/started"
                    Dim startedItem = ReadItemObject(paramsObject)
                    Return New ItemLifecycleEvent(normalizedMethodName,
                                                  paramsObject,
                                                  ReadThreadId(paramsObject, startedItem),
                                                  ReadTurnId(paramsObject, startedItem),
                                                  ReadItemIdFromItem(startedItem),
                                                  startedItem,
                                                  isCompleted:=False)

                Case "item/completed"
                    Dim completedItem = ReadItemObject(paramsObject)
                    Return New ItemLifecycleEvent(normalizedMethodName,
                                                  paramsObject,
                                                  ReadThreadId(paramsObject, completedItem),
                                                  ReadTurnId(paramsObject, completedItem),
                                                  ReadItemIdFromItem(completedItem),
                                                  completedItem,
                                                  isCompleted:=True)

                Case "item/agentMessage/delta"
                    Return New ItemDeltaEvent(normalizedMethodName,
                                              paramsObject,
                                              ReadThreadId(paramsObject),
                                              ReadTurnId(paramsObject),
                                              ReadItemIdFromParams(paramsObject),
                                              ItemDeltaKind.AgentMessage,
                                              ReadDeltaFromParams(paramsObject),
                                              String.Empty)

                Case "item/plan/delta"
                    Return New ItemDeltaEvent(normalizedMethodName,
                                              paramsObject,
                                              ReadThreadId(paramsObject),
                                              ReadTurnId(paramsObject),
                                              ReadItemIdFromParams(paramsObject),
                                              ItemDeltaKind.Plan,
                                              ReadDeltaFromParams(paramsObject),
                                              String.Empty)

                Case "item/reasoning/summaryTextDelta"
                    Return New ItemDeltaEvent(normalizedMethodName,
                                              paramsObject,
                                              ReadThreadId(paramsObject),
                                              ReadTurnId(paramsObject),
                                              ReadItemIdFromParams(paramsObject),
                                              ItemDeltaKind.ReasoningSummaryText,
                                              ReadDeltaFromParams(paramsObject),
                                              String.Empty)

                Case "item/reasoning/summaryPartAdded"
                    Return New ItemDeltaEvent(normalizedMethodName,
                                              paramsObject,
                                              ReadThreadId(paramsObject),
                                              ReadTurnId(paramsObject),
                                              ReadItemIdFromParams(paramsObject),
                                              ItemDeltaKind.ReasoningSummaryPartAdded,
                                              String.Empty,
                                              ReadReasoningSummaryPartText(paramsObject))

                Case "item/reasoning/textDelta"
                    Return New ItemDeltaEvent(normalizedMethodName,
                                              paramsObject,
                                              ReadThreadId(paramsObject),
                                              ReadTurnId(paramsObject),
                                              ReadItemIdFromParams(paramsObject),
                                              ItemDeltaKind.ReasoningText,
                                              ReadDeltaFromParams(paramsObject),
                                              String.Empty)

                Case "item/commandExecution/outputDelta"
                    Return New ItemDeltaEvent(normalizedMethodName,
                                              paramsObject,
                                              ReadThreadId(paramsObject),
                                              ReadTurnId(paramsObject),
                                              ReadItemIdFromParams(paramsObject),
                                              ItemDeltaKind.CommandExecutionOutput,
                                              ReadDeltaFromParams(paramsObject),
                                              String.Empty)

                Case "item/fileChange/outputDelta"
                    Return New ItemDeltaEvent(normalizedMethodName,
                                              paramsObject,
                                              ReadThreadId(paramsObject),
                                              ReadTurnId(paramsObject),
                                              ReadItemIdFromParams(paramsObject),
                                              ItemDeltaKind.FileChangeOutput,
                                              ReadDeltaFromParams(paramsObject),
                                              String.Empty)

                Case "error"
                    Return New ErrorEvent(normalizedMethodName,
                                          paramsObject,
                                          ReadThreadId(paramsObject),
                                          ReadTurnId(paramsObject),
                                          ReadErrorMessage(paramsObject),
                                          ReadWillRetry(paramsObject))

                Case Else
                    Return New GenericServerNotificationEvent(normalizedMethodName, paramsObject)
            End Select
        End Function

        Public Shared Function ParseServerRequest(request As RpcServerRequest) As TurnFlowEvent
            If request Is Nothing Then
                Return New GenericServerRequestEvent(String.Empty, New JsonObject(), Nothing)
            End If

            Dim paramsObject = AsJsonObject(request.ParamsNode)
            Dim normalizedMethodName = NormalizeIdentifier(request.MethodName)

            Select Case normalizedMethodName
                Case "item/commandExecution/requestApproval"
                    Return New ApprovalRequestEvent(normalizedMethodName,
                                                    paramsObject,
                                                    request.Id,
                                                    ReadThreadId(paramsObject),
                                                    ReadTurnId(paramsObject),
                                                    ReadItemIdFromParams(paramsObject),
                                                    ApprovalRequestKind.CommandExecution)

                Case "item/fileChange/requestApproval"
                    Return New ApprovalRequestEvent(normalizedMethodName,
                                                    paramsObject,
                                                    request.Id,
                                                    ReadThreadId(paramsObject),
                                                    ReadTurnId(paramsObject),
                                                    ReadItemIdFromParams(paramsObject),
                                                    ApprovalRequestKind.FileChange)

                Case Else
                    Return New GenericServerRequestEvent(normalizedMethodName, paramsObject, request.Id)
            End Select
        End Function

        Private Shared Function AsJsonObject(node As JsonNode) As JsonObject
            Dim obj = TryCast(node, JsonObject)
            If obj IsNot Nothing Then
                Return obj
            End If

            Return New JsonObject()
        End Function

        Private Shared Sub NormalizeLegacyNotificationEnvelope(ByRef methodName As String,
                                                               ByRef paramsObject As JsonObject)
            Dim normalizedMethodName = NormalizeIdentifier(methodName)
            If String.IsNullOrWhiteSpace(normalizedMethodName) Then
                methodName = String.Empty
                paramsObject = AsJsonObject(paramsObject)
                Return
            End If

            If Not normalizedMethodName.StartsWith("codex/event/", StringComparison.OrdinalIgnoreCase) Then
                methodName = normalizedMethodName
                paramsObject = AsJsonObject(paramsObject)
                Return
            End If

            Dim envelopeParams = AsJsonObject(paramsObject)
            Dim msgObject = ReadObject(envelopeParams, "msg")
            Dim legacyEventName = normalizedMethodName.Substring("codex/event/".Length)
            Dim msgType = ReadStringFirst(msgObject, "type")
            If Not String.IsNullOrWhiteSpace(msgType) Then
                legacyEventName = msgType
            End If

            Dim canonicalMethod = MapLegacyCodexEventMethod(legacyEventName)
            If String.IsNullOrWhiteSpace(canonicalMethod) Then
                methodName = normalizedMethodName
                paramsObject = envelopeParams
                Return
            End If

            Dim normalizedParams = If(CloneJsonObject(msgObject), New JsonObject())

            Dim conversationId = ReadStringFirst(envelopeParams, "conversationId", "conversation_id")
            If Not String.IsNullOrWhiteSpace(conversationId) AndAlso
               String.IsNullOrWhiteSpace(ReadStringFirst(normalizedParams, "threadId", "thread_id")) Then
                normalizedParams("thread_id") = conversationId
            End If

            Dim envelopeId = ReadStringFirst(envelopeParams, "id")
            If Not String.IsNullOrWhiteSpace(envelopeId) AndAlso
               String.IsNullOrWhiteSpace(ReadStringFirst(normalizedParams, "turnId", "turn_id")) Then
                normalizedParams("turn_id") = envelopeId
            End If

            methodName = canonicalMethod
            paramsObject = normalizedParams
        End Sub

        Private Shared Function MapLegacyCodexEventMethod(legacyEventName As String) As String
            Dim normalized = NormalizeIdentifier(legacyEventName)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            Dim lookup = normalized.Replace("."c, "_"c).ToLowerInvariant()
            Select Case lookup
                Case "turn_started"
                    Return "turn/started"
                Case "turn_completed"
                    Return "turn/completed"
                Case "item_started"
                    Return "item/started"
                Case "item_completed"
                    Return "item/completed"
                Case Else
                    Return String.Empty
            End Select
        End Function

        Private Shared Function ReadItemObject(paramsObject As JsonObject) As JsonObject
            Return ReadObject(paramsObject, "item")
        End Function

        Private Shared Function ReadItemIdFromParams(paramsObject As JsonObject) As String
            Dim itemId = ReadStringFirst(paramsObject, "itemId", "item_id")
            If Not String.IsNullOrWhiteSpace(itemId) Then
                Return itemId
            End If

            Return ReadNestedStringFirst(paramsObject,
                                         New String() {"item", "id"})
        End Function

        Private Shared Function ReadItemIdFromItem(itemObject As JsonObject) As String
            Return ReadStringFirst(itemObject, "id")
        End Function

        Private Shared Function ReadThreadId(paramsObject As JsonObject,
                                             Optional itemObject As JsonObject = Nothing) As String
            Dim threadId = ReadStringFirst(paramsObject, "threadId", "thread_id")
            If Not String.IsNullOrWhiteSpace(threadId) Then
                Return threadId
            End If

            threadId = ReadNestedStringFirst(paramsObject,
                                             New String() {"turn", "threadId"},
                                             New String() {"turn", "thread_id"},
                                             New String() {"thread", "id"})
            If Not String.IsNullOrWhiteSpace(threadId) Then
                Return threadId
            End If

            Return ReadStringFirst(itemObject, "threadId", "thread_id")
        End Function

        Private Shared Function ReadTurnId(paramsObject As JsonObject,
                                           Optional itemObject As JsonObject = Nothing) As String
            Dim turnId = ReadStringFirst(paramsObject, "turnId", "turn_id")
            If Not String.IsNullOrWhiteSpace(turnId) Then
                Return turnId
            End If

            turnId = ReadNestedStringFirst(paramsObject,
                                           New String() {"turn", "id"},
                                           New String() {"item", "turnId"},
                                           New String() {"item", "turn_id"})
            If Not String.IsNullOrWhiteSpace(turnId) Then
                Return turnId
            End If

            Return ReadStringFirst(itemObject, "turnId", "turn_id")
        End Function

        Private Shared Function ReadThreadStatusSummary(paramsObject As JsonObject) As String
            Dim statusNode = ReadNestedNode(paramsObject, New String() {"status"})
            If statusNode Is Nothing Then
                statusNode = ReadNestedNode(paramsObject, New String() {"thread", "status"})
            End If

            If statusNode Is Nothing Then
                Return String.Empty
            End If

            Dim statusObject = TryCast(statusNode, JsonObject)
            If statusObject IsNot Nothing Then
                Dim statusType = ReadStringFirst(statusObject, "type")
                If String.IsNullOrWhiteSpace(statusType) Then
                    Return statusObject.ToJsonString()
                End If

                Dim activeFlags = ReadArray(statusObject, "activeFlags")
                If activeFlags Is Nothing OrElse activeFlags.Count = 0 Then
                    Return statusType
                End If

                Return $"{statusType}: {activeFlags.ToJsonString()}"
            End If

            Dim value As String = Nothing
            If TryGetStringValue(statusNode, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then
                Return value.Trim()
            End If

            Return statusNode.ToJsonString()
        End Function

        Private Shared Function CountApps(paramsObject As JsonObject) As Integer
            Dim appsArray = ReadArray(paramsObject, "apps")
            If appsArray IsNot Nothing Then
                Return appsArray.Count
            End If

            appsArray = ReadNestedArray(paramsObject, New String() {"result", "apps"})
            If appsArray IsNot Nothing Then
                Return appsArray.Count
            End If

            Return 0
        End Function

        Private Shared Function CountFiles(paramsObject As JsonObject) As Integer
            Dim filesArray = ReadArray(paramsObject, "files")
            If filesArray Is Nothing Then
                Return 0
            End If

            Return filesArray.Count
        End Function

        Private Shared Function ReadTurnCompletionStatus(paramsObject As JsonObject) As String
            Dim status = ReadNestedStringFirst(paramsObject,
                                               New String() {"turn", "status"})
            If Not String.IsNullOrWhiteSpace(status) Then
                Return status
            End If

            status = ReadStringFirst(paramsObject, "status")
            If Not String.IsNullOrWhiteSpace(status) Then
                Return status
            End If

            Return "completed"
        End Function

        Private Shared Function ReadTurnDiffSummary(paramsObject As JsonObject) As String
            Dim diff = ReadStringFirst(paramsObject, "diff")
            If Not String.IsNullOrWhiteSpace(diff) Then
                Return diff
            End If

            Dim summary = ReadStringFirst(paramsObject, "summary", "text")
            If Not String.IsNullOrWhiteSpace(summary) Then
                Return summary
            End If

            Dim diffObject = ReadObject(paramsObject, "diff")
            If diffObject IsNot Nothing Then
                Return diffObject.ToJsonString()
            End If

            Return String.Empty
        End Function

        Private Shared Function ReadTurnPlanSummary(paramsObject As JsonObject) As String
            Dim explanation = ReadStringFirst(paramsObject, "explanation")
            Dim planArray = ReadArray(paramsObject, "plan")

            Dim lines As New List(Of String)()
            If Not String.IsNullOrWhiteSpace(explanation) Then
                lines.Add(explanation)
            End If

            If planArray IsNot Nothing AndAlso planArray.Count > 0 Then
                For Each planNode In planArray
                    Dim planStep = TryCast(planNode, JsonObject)
                    If planStep Is Nothing Then
                        Continue For
                    End If

                    Dim stepText = ReadStringFirst(planStep, "step")
                    Dim stepStatus = ReadStringFirst(planStep, "status")
                    If String.IsNullOrWhiteSpace(stepText) Then
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(stepStatus) Then
                        lines.Add(stepText)
                    Else
                        lines.Add($"[{stepStatus}] {stepText}")
                    End If
                Next
            End If

            If lines.Count > 0 Then
                Return String.Join(Environment.NewLine, lines)
            End If

            Dim summary = ReadStringFirst(paramsObject, "summary", "text")
            If Not String.IsNullOrWhiteSpace(summary) Then
                Return summary
            End If

            If planArray IsNot Nothing Then
                Return planArray.ToJsonString()
            End If

            Return String.Empty
        End Function

        Private Shared Function ReadTokenUsageObject(paramsObject As JsonObject) As JsonObject
            Dim usageObject = ReadObject(paramsObject, "tokenUsage")
            If usageObject Is Nothing Then
                usageObject = ReadObject(paramsObject, "token_usage")
            End If

            Dim result = If(CloneJsonObject(usageObject), New JsonObject())
            If paramsObject Is Nothing Then
                Return result
            End If

            For Each fieldName In {
                "contextWindow",
                "context_window",
                "contextUsage",
                "context_usage",
                "context",
                "modelContextWindow",
                "model_context_window",
                "total",
                "last",
                "totalUsage",
                "total_usage",
                "totalTokenUsage",
                "total_token_usage",
                "lastUsage",
                "last_usage",
                "lastTokenUsage",
                "last_token_usage"
            }
                CopyTokenUsageFieldIfMissing(paramsObject, result, fieldName)
            Next

            For Each fieldName In {
                "contextPercent",
                "context_percent",
                "contextUsagePercent",
                "context_usage_percent",
                "contextUsagePercentage",
                "context_usage_percentage",
                "contextRemainingPercent",
                "context_remaining_percent",
                "remainingPercent",
                "remaining_percent",
                "remainingPercentage",
                "remaining_percentage",
                "contextUsedTokens",
                "context_used_tokens",
                "usedContextTokens",
                "used_context_tokens",
                "usedTokens",
                "used_tokens",
                "contextWindowTokens",
                "context_window_tokens",
                "windowTokens",
                "window_tokens",
                "contextWindowSize",
                "context_window_size",
                "windowSize",
                "window_size",
                "maxContextTokens",
                "max_context_tokens",
                "maxTokens",
                "max_tokens",
                "tokenLimit",
                "token_limit",
                "inputTokens",
                "input_tokens",
                "outputTokens",
                "output_tokens",
                "reasoningTokens",
                "reasoning_tokens",
                "cachedInputTokens",
                "cached_input_tokens",
                "totalTokens",
                "total_tokens",
                "total"
            }
                CopyTokenUsageFieldIfMissing(paramsObject, result, fieldName)
            Next

            If result.Count = 0 Then
                For Each kvp In paramsObject
                    Dim key = If(kvp.Key, String.Empty)
                    If IsTokenUsageMetaKey(key) Then
                        Continue For
                    End If

                    result(key) = CloneJsonNode(kvp.Value)
                Next
            End If

            Return result
        End Function

        Private Shared Sub CopyTokenUsageFieldIfMissing(source As JsonObject,
                                                         target As JsonObject,
                                                         fieldName As String)
            If source Is Nothing OrElse target Is Nothing OrElse String.IsNullOrWhiteSpace(fieldName) Then
                Return
            End If

            If target.ContainsKey(fieldName) Then
                Return
            End If

            Dim node As JsonNode = Nothing
            If Not source.TryGetPropertyValue(fieldName, node) OrElse node Is Nothing Then
                Return
            End If

            target(fieldName) = CloneJsonNode(node)
        End Sub

        Private Shared Function IsTokenUsageMetaKey(fieldName As String) As Boolean
            Dim normalized = NormalizeIdentifier(fieldName)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return True
            End If

            Select Case normalized.ToLowerInvariant()
                Case "threadid", "thread_id", "turnid", "turn_id", "method", "id"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function ReadDeltaFromParams(paramsObject As JsonObject) As String
            If paramsObject Is Nothing Then
                Return String.Empty
            End If

            Dim deltaNode As JsonNode = Nothing
            If paramsObject.TryGetPropertyValue("delta", deltaNode) AndAlso deltaNode IsNot Nothing Then
                Dim extracted = ExtractDeltaText(deltaNode)
                If Not String.IsNullOrWhiteSpace(extracted) Then
                    Return extracted
                End If
            End If

            Return ReadStringFirst(paramsObject, "delta_text")
        End Function

        Private Shared Function ExtractDeltaText(node As JsonNode) As String
            If node Is Nothing Then
                Return String.Empty
            End If

            Dim obj = TryCast(node, JsonObject)
            If obj IsNot Nothing Then
                Dim preferredKeys = {"text", "delta", "value", "summary", "content", "parts", "items"}
                For Each key In preferredKeys
                    Dim child As JsonNode = Nothing
                    If Not obj.TryGetPropertyValue(key, child) OrElse child Is Nothing Then
                        Continue For
                    End If

                    Dim childText = ExtractDeltaText(child)
                    If Not String.IsNullOrWhiteSpace(childText) Then
                        Return childText
                    End If
                Next

                Dim builder As New StringBuilder()
                AppendDeltaTextFragments(obj, builder)
                Return builder.ToString()
            End If

            Dim arr = TryCast(node, JsonArray)
            If arr IsNot Nothing Then
                Dim builder As New StringBuilder()
                AppendDeltaTextFragments(arr, builder)
                Return builder.ToString()
            End If

            Dim scalarValue As String = Nothing
            If TryGetStringValue(node, scalarValue) Then
                Return If(scalarValue, String.Empty)
            End If

            Return node.ToString()
        End Function

        Private Shared Sub AppendDeltaTextFragments(node As JsonNode, builder As StringBuilder)
            If node Is Nothing OrElse builder Is Nothing Then
                Return
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue IsNot Nothing Then
                Dim value As String = Nothing
                If jsonValue.TryGetValue(Of String)(value) AndAlso value IsNot Nothing Then
                    builder.Append(value)
                End If
                Return
            End If

            Dim arr = TryCast(node, JsonArray)
            If arr IsNot Nothing Then
                For Each child In arr
                    AppendDeltaTextFragments(child, builder)
                Next
                Return
            End If

            Dim obj = TryCast(node, JsonObject)
            If obj Is Nothing Then
                Return
            End If

            Dim preferredKeys = {"text", "delta", "value", "summary", "content", "parts", "items"}
            Dim foundPreferred = False

            For Each key In preferredKeys
                Dim child As JsonNode = Nothing
                If obj.TryGetPropertyValue(key, child) AndAlso child IsNot Nothing Then
                    Dim before = builder.Length
                    AppendDeltaTextFragments(child, builder)
                    If builder.Length > before Then
                        foundPreferred = True
                    End If
                End If
            Next

            If foundPreferred Then
                Return
            End If

            For Each pair In obj
                Select Case If(pair.Key, String.Empty).Trim().ToLowerInvariant()
                    Case "type", "id", "index", "role", "status"
                        Continue For
                End Select

                AppendDeltaTextFragments(pair.Value, builder)
            Next
        End Sub

        Private Shared Function ReadReasoningSummaryPartText(paramsObject As JsonObject) As String
            Dim text = ReadNestedStringFirst(paramsObject,
                                             New String() {"part", "text"},
                                             New String() {"summaryPart", "text"})
            If Not String.IsNullOrWhiteSpace(text) Then
                Return text
            End If

            Dim partNode As JsonNode = Nothing
            If paramsObject IsNot Nothing AndAlso paramsObject.TryGetPropertyValue("part", partNode) Then
                Dim value As String = Nothing
                If TryGetStringValue(partNode, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then
                    Return value
                End If
            End If

            Return String.Empty
        End Function

        Private Shared Function ReadErrorMessage(paramsObject As JsonObject) As String
            Dim message = ReadNestedStringFirst(paramsObject,
                                                New String() {"turn", "error", "message"},
                                                New String() {"error", "message"},
                                                New String() {"message"},
                                                New String() {"error"})
            If String.IsNullOrWhiteSpace(message) Then
                Return String.Empty
            End If

            Return message
        End Function

        Private Shared Function ReadWillRetry(paramsObject As JsonObject) As Boolean?
            Dim raw = ReadStringFirst(paramsObject, "willRetry", "will_retry")
            If String.IsNullOrWhiteSpace(raw) Then
                raw = ReadNestedStringFirst(paramsObject,
                                            New String() {"error", "willRetry"},
                                            New String() {"error", "will_retry"})
            End If

            If String.IsNullOrWhiteSpace(raw) Then
                Return Nothing
            End If

            Dim parsed As Boolean
            If Boolean.TryParse(raw, parsed) Then
                Return parsed
            End If

            Return Nothing
        End Function

        Private Shared Function ReadStringFirst(obj As JsonObject, ParamArray propertyNames() As String) As String
            If obj Is Nothing OrElse propertyNames Is Nothing Then
                Return String.Empty
            End If

            For Each propertyName In propertyNames
                If String.IsNullOrWhiteSpace(propertyName) Then
                    Continue For
                End If

                Dim node As JsonNode = Nothing
                If obj.TryGetPropertyValue(propertyName, node) AndAlso node IsNot Nothing Then
                    Dim value As String = Nothing
                    If TryGetStringValue(node, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then
                        Return value.Trim()
                    End If
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Function ReadNestedStringFirst(obj As JsonObject,
                                                      ParamArray candidatePaths()() As String) As String
            If obj Is Nothing OrElse candidatePaths Is Nothing Then
                Return String.Empty
            End If

            For Each path In candidatePaths
                If path Is Nothing OrElse path.Length = 0 Then
                    Continue For
                End If

                Dim node = ReadNestedNode(obj, path)
                Dim value As String = Nothing
                If TryGetStringValue(node, value) AndAlso Not String.IsNullOrWhiteSpace(value) Then
                    Return value.Trim()
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Function ReadObject(obj As JsonObject, propertyName As String) As JsonObject
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(propertyName, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonObject)
        End Function

        Private Shared Function ReadArray(obj As JsonObject, propertyName As String) As JsonArray
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(propertyName, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonArray)
        End Function

        Private Shared Function ReadNestedArray(obj As JsonObject, path As String()) As JsonArray
            Return TryCast(ReadNestedNode(obj, path), JsonArray)
        End Function

        Private Shared Function ReadNestedNode(obj As JsonObject, path As String()) As JsonNode
            If obj Is Nothing OrElse path Is Nothing OrElse path.Length = 0 Then
                Return Nothing
            End If

            Dim current As JsonNode = obj
            For Each segment In path
                Dim currentObject = TryCast(current, JsonObject)
                If currentObject Is Nothing OrElse String.IsNullOrWhiteSpace(segment) Then
                    Return Nothing
                End If

                If Not currentObject.TryGetPropertyValue(segment, current) Then
                    Return Nothing
                End If
            Next

            Return current
        End Function

        Private Shared Function ReadBoolean(obj As JsonObject,
                                            key As String,
                                            fallback As Boolean) As Boolean
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return fallback
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) OrElse node Is Nothing Then
                Return fallback
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue Is Nothing Then
                Return fallback
            End If

            Dim boolValue As Boolean
            If jsonValue.TryGetValue(Of Boolean)(boolValue) Then
                Return boolValue
            End If

            Dim stringValue As String = Nothing
            If jsonValue.TryGetValue(Of String)(stringValue) Then
                Dim parsed As Boolean
                If Boolean.TryParse(If(stringValue, String.Empty).Trim(), parsed) Then
                    Return parsed
                End If
            End If

            Return fallback
        End Function

        Private Shared Function TryGetStringValue(node As JsonNode, ByRef value As String) As Boolean
            value = String.Empty
            If node Is Nothing Then
                Return False
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue IsNot Nothing Then
                Dim stringValue As String = Nothing
                If jsonValue.TryGetValue(Of String)(stringValue) Then
                    value = If(stringValue, String.Empty)
                    Return True
                End If

                Dim boolValue As Boolean
                If jsonValue.TryGetValue(Of Boolean)(boolValue) Then
                    value = boolValue.ToString(CultureInfo.InvariantCulture)
                    Return True
                End If

                Dim longValue As Long
                If jsonValue.TryGetValue(Of Long)(longValue) Then
                    value = longValue.ToString(CultureInfo.InvariantCulture)
                    Return True
                End If

                value = node.ToJsonString().Trim(""""c)
                Return True
            End If

            value = node.ToString()
            Return Not String.IsNullOrWhiteSpace(value)
        End Function
    End Class
End Namespace
