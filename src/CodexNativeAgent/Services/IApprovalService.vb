Imports System.Text.Json.Nodes
Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Services
    Public Interface IApprovalService
        Function TryCreateApproval(request As RpcServerRequest,
                                   ByRef approvalInfo As PendingApprovalInfo) As Boolean

        Function ResolveDecisionPayload(approvalInfo As PendingApprovalInfo,
                                        action As String,
                                        ByRef decisionLabel As String) As JsonNode
    End Interface
End Namespace
