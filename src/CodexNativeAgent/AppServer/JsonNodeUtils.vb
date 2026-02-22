Imports System.Globalization
Imports System.Text.Json
Imports System.Text.Json.Nodes

Namespace CodexNativeAgent.AppServer
    Friend Module JsonNodeUtils
        Private ReadOnly _prettyOptions As New JsonSerializerOptions With {
            .WriteIndented = True
        }

        Friend Function CloneJson(node As JsonNode) As JsonNode
            If node Is Nothing Then
                Return Nothing
            End If

            Return JsonNode.Parse(node.ToJsonString())
        End Function

        Friend Function AsObject(node As JsonNode) As JsonObject
            Return TryCast(node, JsonObject)
        End Function

        Friend Function AsArray(node As JsonNode) As JsonArray
            Return TryCast(node, JsonArray)
        End Function

        Friend Function PrettyJson(node As JsonNode) As String
            If node Is Nothing Then
                Return String.Empty
            End If

            Return node.ToJsonString(_prettyOptions)
        End Function

        Friend Function TryGetStringValue(node As JsonNode, ByRef value As String) As Boolean
            value = String.Empty

            If node Is Nothing Then
                Return False
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue Is Nothing Then
                Return False
            End If

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

            Dim doubleValue As Double
            If jsonValue.TryGetValue(Of Double)(doubleValue) Then
                value = doubleValue.ToString(CultureInfo.InvariantCulture)
                Return True
            End If

            Return False
        End Function

        Friend Function TryGetBooleanValue(node As JsonNode, ByRef value As Boolean) As Boolean
            value = False

            If node Is Nothing Then
                Return False
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue Is Nothing Then
                Return False
            End If

            If jsonValue.TryGetValue(Of Boolean)(value) Then
                Return True
            End If

            Dim stringValue As String = Nothing
            If jsonValue.TryGetValue(Of String)(stringValue) Then
                Return Boolean.TryParse(stringValue, value)
            End If

            Return False
        End Function

        Friend Function TryGetInt64Value(node As JsonNode, ByRef value As Long) As Boolean
            value = 0

            If node Is Nothing Then
                Return False
            End If

            Dim jsonValue = TryCast(node, JsonValue)
            If jsonValue Is Nothing Then
                Return False
            End If

            If jsonValue.TryGetValue(Of Long)(value) Then
                Return True
            End If

            Dim intValue As Integer
            If jsonValue.TryGetValue(Of Integer)(intValue) Then
                value = intValue
                Return True
            End If

            Dim stringValue As String = Nothing
            If jsonValue.TryGetValue(Of String)(stringValue) Then
                Return Long.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, value)
            End If

            Return False
        End Function

        Friend Function GetPropertyString(obj As JsonObject, key As String, Optional fallback As String = "") As String
            If obj Is Nothing Then
                Return fallback
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) OrElse node Is Nothing Then
                Return fallback
            End If

            Dim value As String = Nothing
            If TryGetStringValue(node, value) Then
                Return value
            End If

            Return fallback
        End Function

        Friend Function GetPropertyBoolean(obj As JsonObject, key As String, Optional fallback As Boolean = False) As Boolean
            If obj Is Nothing Then
                Return fallback
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) OrElse node Is Nothing Then
                Return fallback
            End If

            Dim value As Boolean
            If TryGetBooleanValue(node, value) Then
                Return value
            End If

            Return fallback
        End Function

        Friend Function GetPropertyObject(obj As JsonObject, key As String) As JsonObject
            If obj Is Nothing Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonObject)
        End Function

        Friend Function GetPropertyArray(obj As JsonObject, key As String) As JsonArray
            If obj Is Nothing Then
                Return Nothing
            End If

            Dim node As JsonNode = Nothing
            If Not obj.TryGetPropertyValue(key, node) Then
                Return Nothing
            End If

            Return TryCast(node, JsonArray)
        End Function

        Friend Function RequestIdToKey(idNode As JsonNode) As String
            If idNode Is Nothing Then
                Return String.Empty
            End If

            Dim stringValue As String = Nothing
            If TryGetStringValue(idNode, stringValue) Then
                Return stringValue
            End If

            Dim longValue As Long
            If TryGetInt64Value(idNode, longValue) Then
                Return longValue.ToString(CultureInfo.InvariantCulture)
            End If

            Return String.Empty
        End Function

        Friend Function GetNestedProperty(node As JsonNode, ParamArray keys() As String) As JsonNode
            Dim current As JsonNode = node

            For Each key In keys
                Dim currentObject = TryCast(current, JsonObject)
                If currentObject Is Nothing Then
                    Return Nothing
                End If

                Dim nextNode As JsonNode = Nothing
                If Not currentObject.TryGetPropertyValue(key, nextNode) Then
                    Return Nothing
                End If

                current = nextNode
            Next

            Return current
        End Function
    End Module
End Namespace
