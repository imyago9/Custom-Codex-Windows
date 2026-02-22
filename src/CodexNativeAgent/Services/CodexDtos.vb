Imports System.Text.Json.Nodes

Namespace CodexNativeAgent.Services
    Public NotInheritable Class ModelSummary
        Public Property Id As String = String.Empty
        Public Property DisplayName As String = String.Empty
        Public Property IsDefault As Boolean
    End Class

    Public NotInheritable Class ThreadSummary
        Public Property Id As String = String.Empty
        Public Property Preview As String = String.Empty
        Public Property UpdatedAtText As String = String.Empty
        Public Property UpdatedSortValue As Long
        Public Property LastActiveText As String = String.Empty
        Public Property LastActiveSortValue As Long
        Public Property Cwd As String = String.Empty
        Public Property IsArchived As Boolean
    End Class

    Public NotInheritable Class ThreadRequestOptions
        Public Property ApprovalPolicy As String = String.Empty
        Public Property Sandbox As String = String.Empty
        Public Property Cwd As String = String.Empty
        Public Property Model As String = String.Empty
    End Class

    Public NotInheritable Class AccountReadResult
        Public Property RequiresOpenAiAuth As Boolean
        Public Property Account As JsonObject
    End Class

    Public NotInheritable Class ChatGptLoginStartResult
        Public Property LoginId As String = String.Empty
        Public Property AuthUrl As String = String.Empty
    End Class

    Public NotInheritable Class TurnStartResult
        Public Property TurnId As String = String.Empty
    End Class

    Public NotInheritable Class PendingApprovalInfo
        Public Property RequestId As JsonNode
        Public Property MethodName As String = String.Empty
        Public Property Summary As String = String.Empty
        Public Property AcceptDecision As String = String.Empty
        Public Property AcceptSessionDecision As String = String.Empty
        Public Property DeclineDecision As String = String.Empty
        Public Property CancelDecision As String = String.Empty
    End Class
End Namespace
