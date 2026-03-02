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

    Public NotInheritable Class SkillSummary
        Public Property Cwd As String = String.Empty
        Public Property Name As String = String.Empty
        Public Property Description As String = String.Empty
        Public Property Path As String = String.Empty
        Public Property Enabled As Boolean = True
    End Class

    Public NotInheritable Class AppSummary
        Public Property Id As String = String.Empty
        Public Property Name As String = String.Empty
        Public Property Description As String = String.Empty
        Public Property InstallUrl As String = String.Empty
        Public Property IsAccessible As Boolean
        Public Property IsEnabled As Boolean
    End Class

    Public NotInheritable Class TurnInputItem
        Public Property Type As String = "text"
        Public Property Text As String = String.Empty
        Public Property Name As String = String.Empty
        Public Property Path As String = String.Empty

        Public Shared Function TextItem(value As String) As TurnInputItem
            Return New TurnInputItem() With {
                .Type = "text",
                .Text = If(value, String.Empty)
            }
        End Function

        Public Shared Function SkillItem(name As String, path As String) As TurnInputItem
            Return New TurnInputItem() With {
                .Type = "skill",
                .Name = If(name, String.Empty),
                .Path = If(path, String.Empty)
            }
        End Function

        Public Shared Function MentionItem(name As String, path As String) As TurnInputItem
            Return New TurnInputItem() With {
                .Type = "mention",
                .Name = If(name, String.Empty),
                .Path = If(path, String.Empty)
            }
        End Function
    End Class

    Public NotInheritable Class PendingApprovalInfo
        Public Property RequestId As JsonNode
        Public Property MethodName As String = String.Empty
        Public Property Summary As String = String.Empty
        Public Property ThreadId As String = String.Empty
        Public Property TurnId As String = String.Empty
        Public Property ItemId As String = String.Empty
        Public Property IsCommandApproval As Boolean
        Public Property IsFileChangeApproval As Boolean
        Public Property IsNetworkApproval As Boolean
        Public Property NetworkHost As String = String.Empty
        Public Property NetworkProtocol As String = String.Empty
        Public Property NetworkPort As Integer?
        Public Property ProposedExecpolicyAmendment As JsonArray
        Public Property SupportsExecpolicyAmendment As Boolean
        Public Property AcceptDecision As String = String.Empty
        Public Property AcceptSessionDecision As String = String.Empty
        Public Property DeclineDecision As String = String.Empty
        Public Property CancelDecision As String = String.Empty
    End Class
End Namespace
