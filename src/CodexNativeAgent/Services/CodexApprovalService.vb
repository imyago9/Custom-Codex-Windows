Imports System.Globalization
Imports System.Text
Imports System.Text.Json.Nodes
Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Services
    Public NotInheritable Class CodexApprovalService
        Implements IApprovalService

        Public Function TryCreateApproval(request As RpcServerRequest,
                                          ByRef approvalInfo As PendingApprovalInfo) As Boolean Implements IApprovalService.TryCreateApproval
            approvalInfo = Nothing

            If request Is Nothing Then
                Return False
            End If

            Select Case request.MethodName
                Case "item/commandExecution/requestApproval"
                    approvalInfo = BuildApproval(request,
                                                header:="Command execution approval requested.",
                                                isCommandApproval:=True,
                                                isFileChangeApproval:=False,
                                                acceptDecision:="accept",
                                                acceptSessionDecision:="acceptForSession",
                                                declineDecision:="decline",
                                                cancelDecision:="cancel")

                Case "item/fileChange/requestApproval"
                    approvalInfo = BuildApproval(request,
                                                header:="File change approval requested.",
                                                isCommandApproval:=False,
                                                isFileChangeApproval:=True,
                                                acceptDecision:="accept",
                                                acceptSessionDecision:="acceptForSession",
                                                declineDecision:="decline",
                                                cancelDecision:="cancel")

                Case Else
                    Return False
            End Select

            Return approvalInfo IsNot Nothing
        End Function

        Public Function ResolveDecisionPayload(approvalInfo As PendingApprovalInfo,
                                               action As String,
                                               ByRef decisionLabel As String) As JsonNode Implements IApprovalService.ResolveDecisionPayload
            If approvalInfo Is Nothing Then
                Throw New InvalidOperationException("No active approval is available.")
            End If

            decisionLabel = String.Empty

            Select Case action
                Case "accept"
                    decisionLabel = approvalInfo.AcceptDecision
                    Return JsonValue.Create(decisionLabel)
                Case "accept_session"
                    decisionLabel = approvalInfo.AcceptSessionDecision
                    Return JsonValue.Create(decisionLabel)
                Case "accept_amended"
                    If Not approvalInfo.SupportsExecpolicyAmendment OrElse approvalInfo.ProposedExecpolicyAmendment Is Nothing Then
                        Throw New InvalidOperationException("No execpolicy amendment is available for this approval.")
                    End If

                    Dim amendmentPayload As New JsonObject()
                    amendmentPayload("execpolicy_amendment") = CloneJson(approvalInfo.ProposedExecpolicyAmendment)

                    Dim decisionPayload As New JsonObject()
                    decisionPayload("acceptWithExecpolicyAmendment") = amendmentPayload

                    decisionLabel = "acceptWithExecpolicyAmendment"
                    Return decisionPayload
                Case "decline"
                    decisionLabel = approvalInfo.DeclineDecision
                    Return JsonValue.Create(decisionLabel)
                Case "cancel"
                    decisionLabel = approvalInfo.CancelDecision
                    Return JsonValue.Create(decisionLabel)
                Case Else
                    Throw New InvalidOperationException($"Unknown approval action '{action}'.")
            End Select
        End Function

        Private Shared Function BuildApproval(request As RpcServerRequest,
                                              header As String,
                                              isCommandApproval As Boolean,
                                              isFileChangeApproval As Boolean,
                                              acceptDecision As String,
                                              acceptSessionDecision As String,
                                              declineDecision As String,
                                              cancelDecision As String) As PendingApprovalInfo
            Dim paramsObject = AsObject(request.ParamsNode)
            Dim networkContext = ReadObject(paramsObject, "networkApprovalContext")
            Dim proposedExecpolicy = ReadArray(paramsObject, "proposedExecpolicyAmendment")
            Dim threadId = ReadString(paramsObject, "threadId", "thread_id")
            Dim turnId = ReadString(paramsObject, "turnId", "turn_id")
            Dim itemId = ReadString(paramsObject, "itemId", "item_id")
            Dim networkPort As Integer? = ReadNullableInt32(networkContext, "port")

            Return New PendingApprovalInfo() With {
                .RequestId = CloneJson(request.Id),
                .MethodName = request.MethodName,
                .Summary = BuildSummaryText(header,
                                            request.MethodName,
                                            paramsObject,
                                            networkContext,
                                            threadId,
                                            turnId,
                                            itemId,
                                            proposedExecpolicy),
                .ThreadId = threadId,
                .TurnId = turnId,
                .ItemId = itemId,
                .IsCommandApproval = isCommandApproval,
                .IsFileChangeApproval = isFileChangeApproval,
                .IsNetworkApproval = networkContext IsNot Nothing,
                .NetworkHost = ReadString(networkContext, "host"),
                .NetworkProtocol = ReadString(networkContext, "protocol"),
                .NetworkPort = networkPort,
                .ProposedExecpolicyAmendment = TryCast(CloneJson(proposedExecpolicy), JsonArray),
                .SupportsExecpolicyAmendment = isCommandApproval AndAlso proposedExecpolicy IsNot Nothing AndAlso proposedExecpolicy.Count > 0,
                .AcceptDecision = acceptDecision,
                .AcceptSessionDecision = acceptSessionDecision,
                .DeclineDecision = declineDecision,
                .CancelDecision = cancelDecision
            }
        End Function

        Private Shared Function BuildSummaryText(defaultHeader As String,
                                                 methodName As String,
                                                 paramsObject As JsonObject,
                                                 networkContext As JsonObject,
                                                 threadId As String,
                                                 turnId As String,
                                                 itemId As String,
                                                 proposedExecpolicy As JsonArray) As String
            Dim sb As New StringBuilder()

            If networkContext IsNot Nothing Then
                sb.AppendLine("Managed network access approval requested.")
            Else
                sb.AppendLine(defaultHeader)
            End If

            If Not String.IsNullOrWhiteSpace(methodName) Then
                sb.AppendLine($"Method: {methodName}")
            End If

            If Not String.IsNullOrWhiteSpace(threadId) OrElse
               Not String.IsNullOrWhiteSpace(turnId) OrElse
               Not String.IsNullOrWhiteSpace(itemId) Then
                sb.Append("Scope:")
                If Not String.IsNullOrWhiteSpace(threadId) Then
                    sb.Append($" thread={threadId}")
                End If
                If Not String.IsNullOrWhiteSpace(turnId) Then
                    sb.Append($" turn={turnId}")
                End If
                If Not String.IsNullOrWhiteSpace(itemId) Then
                    sb.Append($" item={itemId}")
                End If
                sb.AppendLine()
            End If

            If networkContext IsNot Nothing Then
                Dim host = ReadString(networkContext, "host")
                Dim protocol = ReadString(networkContext, "protocol")
                Dim port = ReadNullableInt32(networkContext, "port")

                If Not String.IsNullOrWhiteSpace(host) Then
                    sb.AppendLine($"Network host: {host}")
                End If

                If Not String.IsNullOrWhiteSpace(protocol) Then
                    sb.AppendLine($"Network protocol: {protocol}")
                End If

                If port.HasValue Then
                    sb.AppendLine($"Network port: {port.Value.ToString(CultureInfo.InvariantCulture)}")
                End If
            End If

            AppendIfPresent(sb, "Reason", ReadString(paramsObject, "reason"))
            AppendIfPresent(sb, "Command", ReadString(paramsObject, "command"))
            AppendIfPresent(sb, "Working directory", ReadString(paramsObject, "cwd"))
            AppendIfPresent(sb, "Grant root", ReadString(paramsObject, "grantRoot"))

            If proposedExecpolicy IsNot Nothing AndAlso proposedExecpolicy.Count > 0 Then
                sb.AppendLine("Proposed execpolicy amendment:")
                For Each node In proposedExecpolicy
                    Dim part = ReadScalarString(node)
                    If String.IsNullOrWhiteSpace(part) Then
                        part = If(node?.ToJsonString(), String.Empty)
                    End If
                    sb.AppendLine($"  - {part}")
                Next
            End If

            If paramsObject IsNot Nothing Then
                sb.AppendLine()
                sb.Append("Raw params:")
                sb.AppendLine()
                sb.Append(PrettyJson(paramsObject))
            End If

            Return sb.ToString().TrimEnd()
        End Function

        Private Shared Sub AppendIfPresent(sb As StringBuilder, label As String, value As String)
            If sb Is Nothing Then
                Return
            End If

            If String.IsNullOrWhiteSpace(value) Then
                Return
            End If

            sb.AppendLine($"{label}: {value}")
        End Sub

        Private Shared Function ReadObject(obj As JsonObject, key As String) As JsonObject
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonObject)
        End Function

        Private Shared Function ReadArray(obj As JsonObject, key As String) As JsonArray
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonArray)
        End Function

        Private Shared Function ReadString(obj As JsonObject, ParamArray keys() As String) As String
            If obj Is Nothing OrElse keys Is Nothing Then
                Return String.Empty
            End If

            For Each key In keys
                If String.IsNullOrWhiteSpace(key) Then
                    Continue For
                End If

                Dim node As JsonNode = Nothing
                If Not obj.TryGetPropertyValue(key, node) Then
                    Continue For
                End If

                Dim value As String = Nothing
                If TryGetStringValue(node, value) Then
                    Return If(value, String.Empty)
                End If
            Next

            Return String.Empty
        End Function

        Private Shared Function ReadNullableInt32(obj As JsonObject, key As String) As Integer?
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) OrElse node Is Nothing Then
                Return Nothing
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue IsNot Nothing Then
                Dim intValue As Integer
                If jsonValue.TryGetValue(Of Integer)(intValue) Then
                    Return intValue
                End If

                Dim longValue As Long
                If jsonValue.TryGetValue(Of Long)(longValue) Then
                    If longValue >= Integer.MinValue AndAlso longValue <= Integer.MaxValue Then
                        Return CInt(longValue)
                    End If
                End If

                Dim text As String = Nothing
                If jsonValue.TryGetValue(Of String)(text) Then
                    Dim parsed As Integer
                    If Integer.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
                        Return parsed
                    End If
                End If
            End If

            Return Nothing
        End Function

        Private Shared Function ReadScalarString(node As JsonNode) As String
            Dim value As String = Nothing
            If TryGetStringValue(node, value) Then
                Return If(value, String.Empty)
            End If

            Return String.Empty
        End Function
    End Class
End Namespace
