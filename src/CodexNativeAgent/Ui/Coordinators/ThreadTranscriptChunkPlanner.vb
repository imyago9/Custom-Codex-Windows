Imports System.Collections.Generic
Imports CodexNativeAgent.Ui.ViewModels.Transcript

Namespace CodexNativeAgent.Ui.Coordinators
    Public NotInheritable Class ThreadTranscriptDisplayChunkPlan
        Public Property TotalEntryCount As Integer
        Public Property LoadedRangeStart As Integer
        Public Property LoadedRangeEnd As Integer = -1
        Public Property TotalSelectedRenderWeight As Integer
        Public Property HasMoreOlderEntries As Boolean

        Public ReadOnly Property DisplayEntries As New List(Of TranscriptEntryDescriptor)()

        Public ReadOnly Property SelectedEntryCount As Integer
            Get
                Return DisplayEntries.Count
            End Get
        End Property
    End Class

    Public NotInheritable Class ThreadTranscriptChunkPlanner
        Private Sub New()
        End Sub

        Public Shared Function BuildLatestDisplayChunk(entries As IReadOnlyList(Of TranscriptEntryDescriptor),
                                                       Optional maxRowsPerChunk As Integer = 140,
                                                       Optional maxRenderWeightPerChunk As Integer = 280) As ThreadTranscriptDisplayChunkPlan
            Dim plan As New ThreadTranscriptDisplayChunkPlan()
            If entries Is Nothing OrElse entries.Count = 0 Then
                Return plan
            End If

            Dim normalizedMaxRows = Math.Max(1, maxRowsPerChunk)
            Dim normalizedMaxRenderWeight = Math.Max(1, maxRenderWeightPerChunk)
            Dim totalCount = entries.Count

            Dim startIndex = totalCount - 1
            Dim selectedCount = 0
            Dim selectedWeight = 0

            For i = totalCount - 1 To 0 Step -1
                Dim descriptor = entries(i)
                selectedCount += 1
                selectedWeight += EstimateRenderWeight(descriptor)
                startIndex = i

                Dim hitRowCap = selectedCount >= normalizedMaxRows
                Dim hitWeightCap = selectedWeight >= normalizedMaxRenderWeight
                If i > 0 AndAlso (hitRowCap OrElse hitWeightCap) Then
                    Exit For
                End If
            Next

            plan.TotalEntryCount = totalCount
            plan.LoadedRangeStart = startIndex
            plan.LoadedRangeEnd = totalCount - 1
            plan.TotalSelectedRenderWeight = selectedWeight
            plan.HasMoreOlderEntries = startIndex > 0

            For i = startIndex To totalCount - 1
                plan.DisplayEntries.Add(entries(i))
            Next

            Return plan
        End Function

        Private Shared Function EstimateRenderWeight(descriptor As TranscriptEntryDescriptor) As Integer
            If descriptor Is Nothing Then
                Return 1
            End If

            Dim weight = 1
            Dim kind = If(descriptor.Kind, String.Empty).Trim()
            Dim bodyLength = If(descriptor.BodyText, String.Empty).Length
            Dim detailsLength = If(descriptor.DetailsText, String.Empty).Length
            Dim fileChangeCount = If(descriptor.FileChangeItems Is Nothing, 0, descriptor.FileChangeItems.Count)

            If bodyLength > 0 Then
                weight += 1
            End If
            If bodyLength > 280 Then
                weight += 1
            End If
            If bodyLength > 1200 Then
                weight += 2
            End If

            If detailsLength > 0 Then
                weight += 2
            End If
            If detailsLength > 800 Then
                weight += 2
            End If
            If detailsLength > 4000 Then
                weight += 3
            End If

            If descriptor.IsMonospaceBody Then
                weight += 1
            End If
            If descriptor.IsCommandLike Then
                weight += 3
            End If
            If descriptor.IsReasoning OrElse descriptor.UseRawReasoningLayout Then
                weight += 4
            End If
            If descriptor.IsError Then
                weight += 1
            End If
            If descriptor.IsStreaming Then
                weight += 2
            End If

            If fileChangeCount > 0 Then
                weight += 2
                weight += Math.Min(6, fileChangeCount)
            End If

            If StringComparer.OrdinalIgnoreCase.Equals(kind, "turn") OrElse
               StringComparer.OrdinalIgnoreCase.Equals(kind, "turnLifecycle") Then
                weight += 1
            End If

            Return Math.Max(1, weight)
        End Function
    End Class
End Namespace
