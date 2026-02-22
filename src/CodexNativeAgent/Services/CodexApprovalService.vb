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
                                                "Command execution approval requested.",
                                                "accept",
                                                "acceptForSession",
                                                "decline",
                                                "cancel")

                Case "item/fileChange/requestApproval"
                    approvalInfo = BuildApproval(request,
                                                "File change approval requested.",
                                                "accept",
                                                "acceptForSession",
                                                "decline",
                                                "cancel")

                Case "execCommandApproval"
                    approvalInfo = BuildApproval(request,
                                                "Legacy command approval requested.",
                                                "approved",
                                                "approved_for_session",
                                                "denied",
                                                "abort")

                Case "applyPatchApproval"
                    approvalInfo = BuildApproval(request,
                                                "Legacy patch approval requested.",
                                                "approved",
                                                "approved_for_session",
                                                "denied",
                                                "abort")

                Case Else
                    Return False
            End Select

            Return approvalInfo IsNot Nothing
        End Function

        Public Function ResolveDecision(approvalInfo As PendingApprovalInfo,
                                        action As String) As String Implements IApprovalService.ResolveDecision
            If approvalInfo Is Nothing Then
                Throw New InvalidOperationException("No active approval is available.")
            End If

            Select Case action
                Case "accept"
                    Return approvalInfo.AcceptDecision
                Case "accept_session"
                    Return approvalInfo.AcceptSessionDecision
                Case "decline"
                    Return approvalInfo.DeclineDecision
                Case "cancel"
                    Return approvalInfo.CancelDecision
                Case Else
                    Throw New InvalidOperationException($"Unknown approval action '{action}'.")
            End Select
        End Function

        Private Shared Function BuildApproval(request As RpcServerRequest,
                                              header As String,
                                              acceptDecision As String,
                                              acceptSessionDecision As String,
                                              declineDecision As String,
                                              cancelDecision As String) As PendingApprovalInfo
            Return New PendingApprovalInfo() With {
                .RequestId = CloneJson(request.Id),
                .MethodName = request.MethodName,
                .Summary = header & Environment.NewLine & PrettyJson(request.ParamsNode),
                .AcceptDecision = acceptDecision,
                .AcceptSessionDecision = acceptSessionDecision,
                .DeclineDecision = declineDecision,
                .CancelDecision = cancelDecision
            }
        End Function
    End Class
End Namespace
