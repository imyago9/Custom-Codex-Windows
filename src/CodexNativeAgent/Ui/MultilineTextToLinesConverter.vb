Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows.Data

Namespace CodexNativeAgent.Ui
    Public NotInheritable Class MultilineTextToLinesConverter
        Implements IValueConverter

        Public Function Convert(value As Object,
                                targetType As Type,
                                parameter As Object,
                                culture As CultureInfo) As Object Implements IValueConverter.Convert
            Dim text = If(value, String.Empty).ToString()
            If String.IsNullOrWhiteSpace(text) Then
                Return New List(Of String)()
            End If

            Dim normalized = text.Replace(ControlChars.CrLf, ControlChars.Lf).
                                  Replace(ControlChars.Cr, ControlChars.Lf)
            Dim segments = normalized.Split({ControlChars.Lf}, StringSplitOptions.None)
            Dim lines As New List(Of String)()
            For Each segment In segments
                Dim line = If(segment, String.Empty).Trim()
                If line.Length > 0 Then
                    lines.Add(line)
                End If
            Next

            Return lines
        End Function

        Public Function ConvertBack(value As Object,
                                    targetType As Type,
                                    parameter As Object,
                                    culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotSupportedException()
        End Function
    End Class
End Namespace
