Imports System.Text.Json.Nodes

Namespace CodexNativeAgent.Ui
    Friend Module JsonNodeHelpers
        Friend Function AsObject(node As JsonNode) As JsonObject
            Return TryCast(node, JsonObject)
        End Function

        Friend Function GetPropertyObject(obj As JsonObject, propertyName As String) As JsonObject
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(propertyName, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonObject)
        End Function

        Friend Function GetPropertyArray(obj As JsonObject, propertyName As String) As JsonArray
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(propertyName, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonArray)
        End Function

        Friend Function GetPropertyString(obj As JsonObject,
                                          propertyName As String,
                                          Optional fallback As String = "") As String
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return fallback
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(propertyName, node) Then
                Return fallback
            End If

            Dim value As String = String.Empty
            If TryGetStringValue(node, value) Then
                Return value
            End If

            Return fallback
        End Function

        Friend Function GetPropertyBoolean(obj As JsonObject,
                                           propertyName As String,
                                           Optional fallback As Boolean = False) As Boolean
            If obj Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return fallback
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(propertyName, node) OrElse node Is Nothing Then
                Return fallback
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue IsNot Nothing Then
                Dim boolValue As Boolean
                If jsonValue.TryGetValue(Of Boolean)(boolValue) Then
                    Return boolValue
                End If

                Dim stringValue As String = Nothing
                If jsonValue.TryGetValue(Of String)(stringValue) Then
                    Dim parsed As Boolean
                    If Boolean.TryParse(stringValue, parsed) Then
                        Return parsed
                    End If
                End If
            End If

            Return fallback
        End Function

        Friend Function GetAgentMessageChannel(itemObject As JsonObject) As String
            If itemObject Is Nothing Then
                Return String.Empty
            End If

            Dim directChannel = GetPropertyString(itemObject, "channel")
            If Not String.IsNullOrWhiteSpace(directChannel) Then
                Return directChannel.Trim()
            End If

            Dim nestedChannelNode = GetNestedProperty(itemObject, "metadata", "channel")
            If nestedChannelNode Is Nothing Then
                nestedChannelNode = GetNestedProperty(itemObject, "meta", "channel")
            End If

            If nestedChannelNode Is Nothing Then
                Return String.Empty
            End If

            Dim nestedChannel As String = Nothing
            If TryGetStringValue(nestedChannelNode, nestedChannel) Then
                Return If(nestedChannel, String.Empty).Trim()
            End If

            Return String.Empty
        End Function

        Friend Function IsCommentaryAgentMessage(itemObject As JsonObject) As Boolean
            Dim channel = GetAgentMessageChannel(itemObject)
            If String.IsNullOrWhiteSpace(channel) Then
                Return False
            End If

            Return StringComparer.OrdinalIgnoreCase.Equals(channel, "commentary") OrElse
                   StringComparer.OrdinalIgnoreCase.Equals(channel, "comment")
        End Function

        Friend Function GetNestedProperty(obj As JsonObject, ParamArray path() As String) As JsonNode
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

        Friend Function TryGetStringValue(node As JsonNode, ByRef value As String) As Boolean
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

                value = node.ToJsonString().Trim(""""c)
                Return True
            End If

            value = node.ToString()
            Return Not String.IsNullOrWhiteSpace(value)
        End Function
    End Module
End Namespace
