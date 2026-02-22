Imports System.Windows
Imports System.Windows.Media
Imports CodexNativeAgent.Ui.Mvvm

Namespace CodexNativeAgent.Ui.ViewModels.Transcript
    Public NotInheritable Class TranscriptEntryDescriptor
        Public Property Kind As String = String.Empty
        Public Property TimestampText As String = String.Empty
        Public Property RoleText As String = String.Empty
        Public Property BodyText As String = String.Empty
        Public Property SecondaryText As String = String.Empty
        Public Property DetailsText As String = String.Empty
        Public Property AddedLineCount As Integer?
        Public Property RemovedLineCount As Integer?
        Public Property IsMuted As Boolean
        Public Property IsMonospaceBody As Boolean
        Public Property IsCommandLike As Boolean
        Public Property IsReasoning As Boolean
        Public Property IsError As Boolean
        Public Property IsStreaming As Boolean
    End Class

    Public NotInheritable Class TranscriptEntryViewModel
        Inherits ViewModelBase

        Private _kind As String = String.Empty
        Private _timestampText As String = String.Empty
        Private _roleText As String = String.Empty
        Private _bodyText As String = String.Empty
        Private _secondaryText As String = String.Empty
        Private _detailsText As String = String.Empty
        Private _rowOpacity As Double = 1.0R
        Private _rowBackground As Brush = Brushes.Transparent
        Private _rowBorderBrush As Brush = Brushes.Transparent
        Private _roleBadgeBackground As Brush = Brushes.Transparent
        Private _roleBadgeForeground As Brush = Brushes.Black
        Private _bodyForeground As Brush = Brushes.Black
        Private _secondaryForeground As Brush = Brushes.Gray
        Private _detailsForeground As Brush = Brushes.Black
        Private _detailsBackground As Brush = Brushes.Transparent
        Private _bodyFontFamily As FontFamily = New FontFamily("Segoe UI")
        Private _detailsFontFamily As FontFamily = New FontFamily("Cascadia Code")
        Private _roleVisibility As Visibility = Visibility.Visible
        Private _secondaryVisibility As Visibility = Visibility.Collapsed
        Private _detailsVisibility As Visibility = Visibility.Collapsed
        Private _timestampVisibility As Visibility = Visibility.Collapsed
        Private _streamingIndicatorVisibility As Visibility = Visibility.Collapsed
        Private _streamingIndicatorText As String = "in progress"
        Private _changeStatsVisibility As Visibility = Visibility.Collapsed
        Private _addedLinesText As String = String.Empty
        Private _removedLinesText As String = String.Empty
        Private _addedLinesVisibility As Visibility = Visibility.Collapsed
        Private _removedLinesVisibility As Visibility = Visibility.Collapsed

        Public Property Kind As String
            Get
                Return _kind
            End Get
            Set(value As String)
                SetProperty(_kind, If(value, String.Empty))
            End Set
        End Property

        Public Property TimestampText As String
            Get
                Return _timestampText
            End Get
            Set(value As String)
                SetProperty(_timestampText, If(value, String.Empty))
                TimestampVisibility = If(String.IsNullOrWhiteSpace(_timestampText), Visibility.Collapsed, Visibility.Visible)
            End Set
        End Property

        Public Property RoleText As String
            Get
                Return _roleText
            End Get
            Set(value As String)
                SetProperty(_roleText, If(value, String.Empty))
                RoleVisibility = If(String.IsNullOrWhiteSpace(_roleText), Visibility.Collapsed, Visibility.Visible)
            End Set
        End Property

        Public Property BodyText As String
            Get
                Return _bodyText
            End Get
            Set(value As String)
                SetProperty(_bodyText, If(value, String.Empty))
            End Set
        End Property

        Public Property SecondaryText As String
            Get
                Return _secondaryText
            End Get
            Set(value As String)
                SetProperty(_secondaryText, If(value, String.Empty))
                SecondaryVisibility = If(String.IsNullOrWhiteSpace(_secondaryText), Visibility.Collapsed, Visibility.Visible)
            End Set
        End Property

        Public Property DetailsText As String
            Get
                Return _detailsText
            End Get
            Set(value As String)
                SetProperty(_detailsText, If(value, String.Empty))
                DetailsVisibility = If(String.IsNullOrWhiteSpace(_detailsText), Visibility.Collapsed, Visibility.Visible)
            End Set
        End Property

        Public Property RowOpacity As Double
            Get
                Return _rowOpacity
            End Get
            Set(value As Double)
                SetProperty(_rowOpacity, value)
            End Set
        End Property

        Public Property RowBackground As Brush
            Get
                Return _rowBackground
            End Get
            Set(value As Brush)
                SetProperty(_rowBackground, If(value, Brushes.Transparent))
            End Set
        End Property

        Public Property RowBorderBrush As Brush
            Get
                Return _rowBorderBrush
            End Get
            Set(value As Brush)
                SetProperty(_rowBorderBrush, If(value, Brushes.Transparent))
            End Set
        End Property

        Public Property RoleBadgeBackground As Brush
            Get
                Return _roleBadgeBackground
            End Get
            Set(value As Brush)
                SetProperty(_roleBadgeBackground, If(value, Brushes.Transparent))
            End Set
        End Property

        Public Property RoleBadgeForeground As Brush
            Get
                Return _roleBadgeForeground
            End Get
            Set(value As Brush)
                SetProperty(_roleBadgeForeground, If(value, Brushes.Black))
            End Set
        End Property

        Public Property BodyForeground As Brush
            Get
                Return _bodyForeground
            End Get
            Set(value As Brush)
                SetProperty(_bodyForeground, If(value, Brushes.Black))
            End Set
        End Property

        Public Property SecondaryForeground As Brush
            Get
                Return _secondaryForeground
            End Get
            Set(value As Brush)
                SetProperty(_secondaryForeground, If(value, Brushes.Gray))
            End Set
        End Property

        Public Property DetailsForeground As Brush
            Get
                Return _detailsForeground
            End Get
            Set(value As Brush)
                SetProperty(_detailsForeground, If(value, Brushes.Black))
            End Set
        End Property

        Public Property DetailsBackground As Brush
            Get
                Return _detailsBackground
            End Get
            Set(value As Brush)
                SetProperty(_detailsBackground, If(value, Brushes.Transparent))
            End Set
        End Property

        Public Property BodyFontFamily As FontFamily
            Get
                Return _bodyFontFamily
            End Get
            Set(value As FontFamily)
                SetProperty(_bodyFontFamily, If(value, New FontFamily("Segoe UI")))
            End Set
        End Property

        Public Property DetailsFontFamily As FontFamily
            Get
                Return _detailsFontFamily
            End Get
            Set(value As FontFamily)
                SetProperty(_detailsFontFamily, If(value, New FontFamily("Cascadia Code")))
            End Set
        End Property

        Public Property RoleVisibility As Visibility
            Get
                Return _roleVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_roleVisibility, value)
            End Set
        End Property

        Public Property SecondaryVisibility As Visibility
            Get
                Return _secondaryVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_secondaryVisibility, value)
            End Set
        End Property

        Public Property DetailsVisibility As Visibility
            Get
                Return _detailsVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_detailsVisibility, value)
            End Set
        End Property

        Public Property TimestampVisibility As Visibility
            Get
                Return _timestampVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_timestampVisibility, value)
            End Set
        End Property

        Public Property StreamingIndicatorVisibility As Visibility
            Get
                Return _streamingIndicatorVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_streamingIndicatorVisibility, value)
            End Set
        End Property

        Public Property StreamingIndicatorText As String
            Get
                Return _streamingIndicatorText
            End Get
            Set(value As String)
                SetProperty(_streamingIndicatorText, If(value, String.Empty))
            End Set
        End Property

        Public Property ChangeStatsVisibility As Visibility
            Get
                Return _changeStatsVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_changeStatsVisibility, value)
            End Set
        End Property

        Public Property AddedLinesText As String
            Get
                Return _addedLinesText
            End Get
            Set(value As String)
                SetProperty(_addedLinesText, If(value, String.Empty))
                AddedLinesVisibility = If(String.IsNullOrWhiteSpace(_addedLinesText), Visibility.Collapsed, Visibility.Visible)
                UpdateChangeStatsVisibility()
            End Set
        End Property

        Public Property RemovedLinesText As String
            Get
                Return _removedLinesText
            End Get
            Set(value As String)
                SetProperty(_removedLinesText, If(value, String.Empty))
                RemovedLinesVisibility = If(String.IsNullOrWhiteSpace(_removedLinesText), Visibility.Collapsed, Visibility.Visible)
                UpdateChangeStatsVisibility()
            End Set
        End Property

        Public Property AddedLinesVisibility As Visibility
            Get
                Return _addedLinesVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_addedLinesVisibility, value)
                UpdateChangeStatsVisibility()
            End Set
        End Property

        Public Property RemovedLinesVisibility As Visibility
            Get
                Return _removedLinesVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_removedLinesVisibility, value)
                UpdateChangeStatsVisibility()
            End Set
        End Property

        Public Sub AppendBodyChunk(chunk As String)
            If String.IsNullOrEmpty(chunk) Then
                Return
            End If

            BodyText = _bodyText & chunk
        End Sub

        Private Sub UpdateChangeStatsVisibility()
            Dim shouldShow = _addedLinesVisibility = Visibility.Visible OrElse _removedLinesVisibility = Visibility.Visible
            ChangeStatsVisibility = If(shouldShow, Visibility.Visible, Visibility.Collapsed)
        End Sub
    End Class
End Namespace
