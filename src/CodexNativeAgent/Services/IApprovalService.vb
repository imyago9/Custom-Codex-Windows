Imports CodexNativeAgent.AppServer

Namespace CodexNativeAgent.Services
    Public Interface IApprovalService
        Function TryCreateApproval(request As RpcServerRequest,
                                   ByRef approvalInfo As PendingApprovalInfo) As Boolean

        Function ResolveDecision(approvalInfo As PendingApprovalInfo,
                                 action As String) As String
    End Interface
End Namespace
