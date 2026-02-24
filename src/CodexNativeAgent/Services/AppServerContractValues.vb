Namespace CodexNativeAgent.Services
    Friend Module AppServerContractValues
        Friend Const ApprovalPolicyUntrusted As String = "untrusted"
        Friend Const ApprovalPolicyOnFailure As String = "on-failure"
        Friend Const ApprovalPolicyOnRequest As String = "on-request"
        Friend Const ApprovalPolicyNever As String = "never"

        Friend Const SandboxWorkspaceWrite As String = "workspace-write"
        Friend Const SandboxReadOnly As String = "read-only"
        Friend Const SandboxDangerFullAccess As String = "danger-full-access"

        Friend Const ToolRequestUserInputMethod As String = "tool/requestUserInput"

        Friend Function NormalizeApprovalPolicy(value As String) As String
            Dim normalized = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            Select Case NormalizeToken(normalized)
                Case "untrusted", "unlesstrusted"
                    Return ApprovalPolicyUntrusted
                Case "onfailure"
                    Return ApprovalPolicyOnFailure
                Case "onrequest"
                    Return ApprovalPolicyOnRequest
                Case "never"
                    Return ApprovalPolicyNever
            End Select

            Throw New InvalidOperationException(
                $"Unsupported approvalPolicy '{normalized}'. Expected one of: {ApprovalPolicyUntrusted}, {ApprovalPolicyOnFailure}, {ApprovalPolicyOnRequest}, {ApprovalPolicyNever}.")
        End Function

        Friend Function NormalizeSandboxMode(value As String) As String
            Dim normalized = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            Select Case NormalizeToken(normalized)
                Case "workspacewrite"
                    Return SandboxWorkspaceWrite
                Case "readonly"
                    Return SandboxReadOnly
                Case "dangerfullaccess", "fullaccess"
                    Return SandboxDangerFullAccess
            End Select

            Throw New InvalidOperationException(
                $"Unsupported sandbox '{normalized}'. Expected one of: {SandboxWorkspaceWrite}, {SandboxReadOnly}, {SandboxDangerFullAccess}.")
        End Function

        Private Function NormalizeToken(value As String) As String
            Dim compact = If(value, String.Empty).
                Replace("-", String.Empty, StringComparison.Ordinal).
                Replace("_", String.Empty, StringComparison.Ordinal).
                Replace(" ", String.Empty, StringComparison.Ordinal)

            Return compact.Trim().ToLowerInvariant()
        End Function
    End Module
End Namespace
