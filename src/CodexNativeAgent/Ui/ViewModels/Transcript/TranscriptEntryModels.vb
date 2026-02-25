Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Input
Imports System.Windows.Interop
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports CodexNativeAgent.Ui.Mvvm

Namespace CodexNativeAgent.Ui.ViewModels.Transcript
    Public NotInheritable Class TranscriptEntryDescriptor
        Public Property Kind As String = String.Empty
        Public Property TimestampText As String = String.Empty
        Public Property RoleText As String = String.Empty
        Public Property BodyText As String = String.Empty
        Public Property StatusText As String = String.Empty
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
        Public Property UseRawReasoningLayout As Boolean
        Public Property FileChangeItems As List(Of TranscriptFileChangeListItemViewModel)
    End Class

    Public NotInheritable Class TranscriptFileChangeListItemViewModel
        Private Shared ReadOnly _fileIconCache As New Dictionary(Of String, ImageSource)(StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly _fileIconCacheLock As New Object()

        Public Property FullPathText As String = String.Empty
        Public Property DisplayPathPrefixText As String = String.Empty
        Public Property DisplayPathFileNameText As String = String.Empty
        Public Property OverflowText As String = String.Empty
        Public Property IsOverflow As Boolean
        Public Property AddedLinesText As String = String.Empty
        Public Property RemovedLinesText As String = String.Empty
        Public Property FileIconSource As ImageSource

        Public ReadOnly Property PathVisibility As Visibility
            Get
                Return If(IsOverflow, Visibility.Collapsed, Visibility.Visible)
            End Get
        End Property

        Public ReadOnly Property OverflowVisibility As Visibility
            Get
                Return If(IsOverflow, Visibility.Visible, Visibility.Collapsed)
            End Get
        End Property

        Public ReadOnly Property AddedLinesVisibility As Visibility
            Get
                Return If(String.IsNullOrWhiteSpace(AddedLinesText), Visibility.Collapsed, Visibility.Visible)
            End Get
        End Property

        Public ReadOnly Property RemovedLinesVisibility As Visibility
            Get
                Return If(String.IsNullOrWhiteSpace(RemovedLinesText), Visibility.Collapsed, Visibility.Visible)
            End Get
        End Property

        Public ReadOnly Property FileIconVisibility As Visibility
            Get
                Return If(FileIconSource Is Nothing, Visibility.Collapsed, Visibility.Visible)
            End Get
        End Property

        Public ReadOnly Property FileIconFallbackVisibility As Visibility
            Get
                Return If(FileIconSource Is Nothing, Visibility.Visible, Visibility.Collapsed)
            End Get
        End Property

        Public Shared Function CreatePathItem(pathText As String,
                                              Optional addedLineCount As Integer? = Nothing,
                                              Optional removedLineCount As Integer? = Nothing) As TranscriptFileChangeListItemViewModel
            Dim text = If(pathText, String.Empty).Trim()
            Dim parts = BuildDisplayPathParts(text)
            Dim iconPath = ResolveTargetPathForIcon(text)
            Return New TranscriptFileChangeListItemViewModel() With {
                .FullPathText = text,
                .DisplayPathPrefixText = parts.Prefix,
                .DisplayPathFileNameText = parts.FileName,
                .AddedLinesText = If(addedLineCount.HasValue AndAlso addedLineCount.Value > 0,
                                     $"+{addedLineCount.Value}",
                                     String.Empty),
                .RemovedLinesText = If(removedLineCount.HasValue AndAlso removedLineCount.Value > 0,
                                       $"-{removedLineCount.Value}",
                                       String.Empty),
                .FileIconSource = GetFileIconSource(iconPath)
            }
        End Function

        Public Shared Function CreateOverflowItem(summaryText As String) As TranscriptFileChangeListItemViewModel
            Return New TranscriptFileChangeListItemViewModel() With {
                .IsOverflow = True,
                .OverflowText = If(summaryText, String.Empty).Trim()
            }
        End Function

        Private Shared Function BuildDisplayPathParts(pathText As String) As (Prefix As String, FileName As String)
            Dim text = If(pathText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return (String.Empty, String.Empty)
            End If

            Dim renamePrefix As String = String.Empty
            Dim targetPath = text
            Dim renameArrowIndex = text.IndexOf(" -> ", StringComparison.Ordinal)
            If renameArrowIndex >= 0 Then
                renamePrefix = text.Substring(0, renameArrowIndex + 4)
                targetPath = text.Substring(renameArrowIndex + 4)
            End If

            Dim lastSlash = Math.Max(targetPath.LastIndexOf("/"c), targetPath.LastIndexOf("\"c))
            If lastSlash < 0 Then
                Return (renamePrefix, CompactFileNamePreservingEnd(targetPath))
            End If

            Dim prefix = renamePrefix & targetPath.Substring(0, lastSlash + 1)
            Dim fileName = targetPath.Substring(lastSlash + 1)
            If String.IsNullOrWhiteSpace(fileName) Then
                Return (String.Empty, CompactFileNamePreservingEnd(text))
            End If

            Return (prefix, CompactFileNamePreservingEnd(fileName))
        End Function

        Private Shared Function ResolveTargetPathForIcon(pathText As String) As String
            Dim text = If(pathText, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If

            Dim renameArrowIndex = text.IndexOf(" -> ", StringComparison.Ordinal)
            If renameArrowIndex >= 0 AndAlso renameArrowIndex + 4 < text.Length Then
                Return text.Substring(renameArrowIndex + 4).Trim()
            End If

            Return text
        End Function

        Private Shared Function CompactFileNamePreservingEnd(fileName As String) As String
            Dim text = If(fileName, String.Empty)
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If

            Const maxLen As Integer = 34
            If text.Length <= maxLen Then
                Return text
            End If

            Dim ext = Path.GetExtension(text)
            Dim stem = text
            If Not String.IsNullOrWhiteSpace(ext) AndAlso ext.Length < text.Length Then
                stem = text.Substring(0, text.Length - ext.Length)
            Else
                ext = String.Empty
            End If

            If String.IsNullOrEmpty(stem) Then
                Return text
            End If

            Dim tailStemLen = Math.Min(8, Math.Max(3, stem.Length \ 3))
            Dim headStemLen = Math.Max(5, maxLen - ext.Length - 3 - tailStemLen)
            If headStemLen + tailStemLen >= stem.Length Then
                Return text
            End If

            Return stem.Substring(0, headStemLen) &
                   "..." &
                   stem.Substring(stem.Length - tailStemLen) &
                   ext
        End Function

        Private Shared Function GetFileIconSource(pathText As String) As ImageSource
            Dim relativePath = If(pathText, String.Empty).Trim()
            Dim isDirectory = relativePath.EndsWith("/", StringComparison.Ordinal) OrElse
                              relativePath.EndsWith("\", StringComparison.Ordinal)

            Dim extension = String.Empty
            If Not isDirectory Then
                extension = Path.GetExtension(relativePath)
            End If

            Dim cacheKey As String
            If isDirectory Then
                cacheKey = "dir"
            ElseIf Not String.IsNullOrWhiteSpace(extension) Then
                cacheKey = "ext:" & extension.Trim().ToLowerInvariant()
            Else
                cacheKey = "file"
            End If

            SyncLock _fileIconCacheLock
                Dim cached As ImageSource = Nothing
                If _fileIconCache.TryGetValue(cacheKey, cached) Then
                    Return cached
                End If
            End SyncLock

            Dim iconSource = CreateShellAssociatedIconSource(relativePath, isDirectory)

            SyncLock _fileIconCacheLock
                If Not _fileIconCache.ContainsKey(cacheKey) Then
                    _fileIconCache(cacheKey) = iconSource
                End If
                Return _fileIconCache(cacheKey)
            End SyncLock
        End Function

        Private Shared Function CreateShellAssociatedIconSource(relativePath As String, isDirectory As Boolean) As ImageSource
            Try
                Dim queryPath As String
                If isDirectory Then
                    queryPath = "folder"
                Else
                    Dim ext = Path.GetExtension(If(relativePath, String.Empty))
                    If String.IsNullOrWhiteSpace(ext) Then
                        queryPath = "file.bin"
                    Else
                        queryPath = "file" & ext.Trim()
                    End If
                End If

                Dim fileAttrs = If(isDirectory, FILE_ATTRIBUTE_DIRECTORY, FILE_ATTRIBUTE_NORMAL)
                Dim flags = SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES

                Dim info As New SHFILEINFO()
                Dim result = SHGetFileInfo(queryPath, fileAttrs, info, CUInt(Marshal.SizeOf(GetType(SHFILEINFO))), flags)
                If result = IntPtr.Zero OrElse info.hIcon = IntPtr.Zero Then
                    Return Nothing
                End If

                Try
                    Dim iconImage = Imaging.CreateBitmapSourceFromHIcon(info.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(16, 16))
                    If iconImage IsNot Nothing AndAlso iconImage.CanFreeze Then
                        iconImage.Freeze()
                    End If
                    Return iconImage
                Finally
                    DestroyIcon(info.hIcon)
                End Try
            Catch
                Return Nothing
            End Try
        End Function

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Private Structure SHFILEINFO
            Public hIcon As IntPtr
            Public iIcon As Integer
            Public dwAttributes As UInteger
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
            Public szDisplayName As String
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
            Public szTypeName As String
        End Structure

        <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function SHGetFileInfo(pszPath As String,
                                              dwFileAttributes As UInteger,
                                              ByRef psfi As SHFILEINFO,
                                              cbFileInfo As UInteger,
                                              uFlags As UInteger) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function DestroyIcon(hIcon As IntPtr) As Boolean
        End Function

        Private Const FILE_ATTRIBUTE_DIRECTORY As UInteger = &H10UI
        Private Const FILE_ATTRIBUTE_NORMAL As UInteger = &H80UI
        Private Const SHGFI_ICON As UInteger = &H100UI
        Private Const SHGFI_SMALLICON As UInteger = &H1UI
        Private Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10UI
    End Class

    Public NotInheritable Class TranscriptEntryViewModel
        Inherits ViewModelBase

        Private _kind As String = String.Empty
        Private _timestampText As String = String.Empty
        Private _roleText As String = String.Empty
        Private _bodyText As String = String.Empty
        Private _statusText As String = String.Empty
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
        Private _rowVisibility As Visibility = Visibility.Visible
        Private _roleVisibility As Visibility = Visibility.Visible
        Private _secondaryVisibility As Visibility = Visibility.Collapsed
        Private _detailsVisibility As Visibility = Visibility.Collapsed
        Private _timestampVisibility As Visibility = Visibility.Collapsed
        Private _streamingIndicatorVisibility As Visibility = Visibility.Collapsed
        Private _streamingIndicatorText As String = "in progress"
        Private _changeStatsVisibility As Visibility = Visibility.Collapsed
        Private _fileChangeListVisibility As Visibility = Visibility.Collapsed
        Private _activityBodyChipVisibility As Visibility = Visibility.Visible
        Private _addedLinesText As String = String.Empty
        Private _removedLinesText As String = String.Empty
        Private _addedLinesVisibility As Visibility = Visibility.Collapsed
        Private _removedLinesVisibility As Visibility = Visibility.Collapsed
        Private _allowDetailsCollapse As Boolean
        Private _isDetailsExpanded As Boolean = True
        Private _detailsToggleVisibility As Visibility = Visibility.Collapsed
        Private _detailsToggleText As String = "Show details"
        Private ReadOnly _toggleDetailsCommand As ICommand
        Private ReadOnly _fileChangeItems As New ObservableCollection(Of TranscriptFileChangeListItemViewModel)()

        Public Sub New()
            _toggleDetailsCommand = New RelayCommand(Sub() ToggleDetails())
        End Sub

        Public Property Kind As String
            Get
                Return _kind
            End Get
            Set(value As String)
                If SetProperty(_kind, If(value, String.Empty)) Then
                    RefreshDetailsPresentationState()
                End If
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

        Public Property StatusText As String
            Get
                Return _statusText
            End Get
            Set(value As String)
                SetProperty(_statusText, If(value, String.Empty))
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
                If SetProperty(_detailsText, If(value, String.Empty)) Then
                    RefreshDetailsPresentationState()
                End If
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

        Public Property RowVisibility As Visibility
            Get
                Return _rowVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_rowVisibility, value)
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

        Public Property FileChangeListVisibility As Visibility
            Get
                Return _fileChangeListVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_fileChangeListVisibility, value)
            End Set
        End Property

        Public Property ActivityBodyChipVisibility As Visibility
            Get
                Return _activityBodyChipVisibility
            End Get
            Set(value As Visibility)
                SetProperty(_activityBodyChipVisibility, value)
            End Set
        End Property

        Public ReadOnly Property FileChangeItems As ObservableCollection(Of TranscriptFileChangeListItemViewModel)
            Get
                Return _fileChangeItems
            End Get
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

        Public Property AllowDetailsCollapse As Boolean
            Get
                Return _allowDetailsCollapse
            End Get
            Set(value As Boolean)
                If SetProperty(_allowDetailsCollapse, value) Then
                    RefreshDetailsPresentationState()
                End If
            End Set
        End Property

        Public Property IsDetailsExpanded As Boolean
            Get
                Return _isDetailsExpanded
            End Get
            Set(value As Boolean)
                If SetProperty(_isDetailsExpanded, value) Then
                    RefreshDetailsPresentationState()
                End If
            End Set
        End Property

        Public Property DetailsToggleVisibility As Visibility
            Get
                Return _detailsToggleVisibility
            End Get
            Private Set(value As Visibility)
                SetProperty(_detailsToggleVisibility, value)
            End Set
        End Property

        Public Property DetailsToggleText As String
            Get
                Return _detailsToggleText
            End Get
            Private Set(value As String)
                SetProperty(_detailsToggleText, If(value, String.Empty))
            End Set
        End Property

        Public ReadOnly Property ToggleDetailsCommand As ICommand
            Get
                Return _toggleDetailsCommand
            End Get
        End Property

        Public Sub AppendBodyChunk(chunk As String)
            If String.IsNullOrEmpty(chunk) Then
                Return
            End If

            BodyText = _bodyText & chunk
        End Sub

        Public Sub SetFileChangeItems(items As IEnumerable(Of TranscriptFileChangeListItemViewModel))
            _fileChangeItems.Clear()

            If items IsNot Nothing Then
                For Each item In items
                    If item Is Nothing Then
                        Continue For
                    End If

                    _fileChangeItems.Add(item)
                Next
            End If

            FileChangeListVisibility = If(_fileChangeItems.Count > 0, Visibility.Visible, Visibility.Collapsed)
            ActivityBodyChipVisibility = If(_fileChangeItems.Count > 0, Visibility.Collapsed, Visibility.Visible)
        End Sub

        Private Sub ToggleDetails()
            If Not _allowDetailsCollapse OrElse String.IsNullOrWhiteSpace(_detailsText) Then
                Return
            End If

            IsDetailsExpanded = Not _isDetailsExpanded
        End Sub

        Private Sub UpdateChangeStatsVisibility()
            Dim shouldShow = _addedLinesVisibility = Visibility.Visible OrElse _removedLinesVisibility = Visibility.Visible
            ChangeStatsVisibility = If(shouldShow, Visibility.Visible, Visibility.Collapsed)
        End Sub

        Private Sub RefreshDetailsPresentationState()
            Dim hasDetails = Not String.IsNullOrWhiteSpace(_detailsText)
            Dim showDetails = hasDetails AndAlso (Not _allowDetailsCollapse OrElse _isDetailsExpanded)

            DetailsVisibility = If(showDetails, Visibility.Visible, Visibility.Collapsed)
            DetailsToggleVisibility = If(_allowDetailsCollapse AndAlso hasDetails, Visibility.Visible, Visibility.Collapsed)
            DetailsToggleText = BuildDetailsToggleText()
        End Sub

        Private Function BuildDetailsToggleText() As String
            Dim noun = If(StringComparer.OrdinalIgnoreCase.Equals(_kind, "command"), "result", "details")
            Return If(_isDetailsExpanded, $"Hide {noun}", $"Show {noun}")
        End Function
    End Class
End Namespace
