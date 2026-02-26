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

        Public Shared Function PrependPreviousDisplayChunk(entries As IReadOnlyList(Of TranscriptEntryDescriptor),
                                                           loadedRangeStart As Integer,
                                                           loadedRangeEnd As Integer,
                                                           Optional maxRowsPerChunk As Integer = 140,
                                                           Optional maxRenderWeightPerChunk As Integer = 280) As ThreadTranscriptDisplayChunkPlan
            Dim plan As New ThreadTranscriptDisplayChunkPlan()
            If entries Is Nothing OrElse entries.Count = 0 Then
                Return plan
            End If

            Dim totalCount = entries.Count
            Dim normalizedLoadedStart = Math.Max(0, Math.Min(loadedRangeStart, totalCount - 1))
            Dim normalizedLoadedEnd = Math.Max(normalizedLoadedStart, Math.Min(loadedRangeEnd, totalCount - 1))

            If normalizedLoadedStart <= 0 Then
                plan.TotalEntryCount = totalCount
                plan.LoadedRangeStart = 0
                plan.LoadedRangeEnd = normalizedLoadedEnd
                plan.HasMoreOlderEntries = False
                For i = 0 To normalizedLoadedEnd
                    plan.DisplayEntries.Add(entries(i))
                Next

                plan.TotalSelectedRenderWeight = CalculateTotalRenderWeight(plan.DisplayEntries)
                Return plan
            End If

            Dim normalizedMaxRows = Math.Max(1, maxRowsPerChunk)
            Dim normalizedMaxRenderWeight = Math.Max(1, maxRenderWeightPerChunk)

            Dim prependStartIndex = normalizedLoadedStart - 1
            Dim prependedCount = 0
            Dim prependedWeight = 0

            For i = normalizedLoadedStart - 1 To 0 Step -1
                prependedCount += 1
                prependedWeight += EstimateRenderWeight(entries(i))
                prependStartIndex = i

                Dim hitRowCap = prependedCount >= normalizedMaxRows
                Dim hitWeightCap = prependedWeight >= normalizedMaxRenderWeight
                If i > 0 AndAlso (hitRowCap OrElse hitWeightCap) Then
                    Exit For
                End If
            Next

            plan.TotalEntryCount = totalCount
            plan.LoadedRangeStart = prependStartIndex
            plan.LoadedRangeEnd = normalizedLoadedEnd
            plan.HasMoreOlderEntries = prependStartIndex > 0

            For i = prependStartIndex To normalizedLoadedEnd
                plan.DisplayEntries.Add(entries(i))
            Next

            plan.TotalSelectedRenderWeight = CalculateTotalRenderWeight(plan.DisplayEntries)
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

        Private Shared Function CalculateTotalRenderWeight(entries As IEnumerable(Of TranscriptEntryDescriptor)) As Integer
            If entries Is Nothing Then
                Return 0
            End If

            Dim total = 0
            For Each entry In entries
                total += EstimateRenderWeight(entry)
            Next

            Return total
        End Function
    End Class
End Namespace
