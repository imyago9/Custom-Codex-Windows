Imports System.Collections.Generic
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports CodexNativeAgent.Services

Namespace CodexNativeAgent.Ui
    Public NotInheritable Partial Class MainWindow
        Private NotInheritable Class SkillCatalogEntry
            Public Property Name As String = String.Empty
            Public Property Description As String = String.Empty
            Public Property Path As String = String.Empty
        End Class

        Private NotInheritable Class AppCatalogEntry
            Public Property Id As String = String.Empty
            Public Property Name As String = String.Empty
            Public Property Description As String = String.Empty
        End Class

        Private NotInheritable Class TurnComposerTokenSuggestionEntry
            Public Property Kind As String = String.Empty
            Public Property KindLabel As String = String.Empty
            Public Property InsertToken As String = String.Empty
            Public Property DisplayName As String = String.Empty
            Public Property MetaText As String = String.Empty
            Public Property Description As String = String.Empty
            Public Property SortScore As Integer = Integer.MaxValue
        End Class

        Private ReadOnly _skillsAppsService As ISkillsAppsService
        Private ReadOnly _skillsCatalogByToken As New Dictionary(Of String, SkillCatalogEntry)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _appsCatalogByToken As New Dictionary(Of String, AppCatalogEntry)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _skillsCatalogRefreshGate As New SemaphoreSlim(1, 1)
        Private ReadOnly _appsCatalogRefreshGate As New SemaphoreSlim(1, 1)

        Private _skillsCatalogCwdFingerprint As String = String.Empty
        Private _skillsCatalogLoadedAtUtc As DateTimeOffset = DateTimeOffset.MinValue
        Private _appsCatalogThreadId As String = String.Empty
        Private _appsCatalogLoadedAtUtc As DateTimeOffset = DateTimeOffset.MinValue

        Private Shared ReadOnly SkillsCatalogRefreshInterval As TimeSpan = TimeSpan.FromSeconds(30)
        Private Shared ReadOnly AppsCatalogRefreshInterval As TimeSpan = TimeSpan.FromSeconds(30)

        Private Async Function EnsureSkillAndAppCatalogFreshForTurnAsync(threadId As String,
                                                                          Optional forceRefresh As Boolean = False) As Task
            If _skillsAppsService Is Nothing OrElse Not IsClientRunning() Then
                Return
            End If

            Await RefreshSkillsCatalogAsync(forceReload:=forceRefresh,
                                            preferredThreadId:=threadId,
                                            quietErrors:=True).ConfigureAwait(True)
            Await RefreshAppsCatalogAsync(preferredThreadId:=threadId,
                                          forceRefetch:=forceRefresh,
                                          quietErrors:=True,
                                          forceRefresh:=forceRefresh).ConfigureAwait(True)
        End Function

        Private Async Function RefreshSkillsCatalogAsync(forceReload As Boolean,
                                                         Optional preferredThreadId As String = Nothing,
                                                         Optional quietErrors As Boolean = True) As Task
            If _skillsAppsService Is Nothing OrElse Not IsClientRunning() Then
                Return
            End If

            Dim cwds = BuildSkillCatalogCwdCandidates(preferredThreadId)
            If cwds.Count = 0 Then
                Return
            End If

            Dim cwdFingerprint = BuildCwdFingerprint(cwds)
            If Not forceReload AndAlso ShouldReuseSkillsCatalog(cwdFingerprint) Then
                Return
            End If

            Await _skillsCatalogRefreshGate.WaitAsync().ConfigureAwait(True)
            Try
                If Not forceReload AndAlso ShouldReuseSkillsCatalog(cwdFingerprint) Then
                    Return
                End If

                Dim skills = Await _skillsAppsService.ListSkillsAsync(cwds, forceReload, CancellationToken.None).ConfigureAwait(True)
                ApplySkillsCatalog(skills, cwdFingerprint)
            Catch ex As Exception
                If quietErrors Then
                    AppendProtocol("debug", $"skills/list failed: {ex.Message}")
                Else
                    AppendAndShowSystemMessage($"Could not refresh skills: {ex.Message}",
                                               $"Could not refresh skills: {ex.Message}",
                                               isError:=True)
                End If
            Finally
                _skillsCatalogRefreshGate.Release()
            End Try
        End Function

        Private Async Function RefreshAppsCatalogAsync(Optional preferredThreadId As String = Nothing,
                                                       Optional forceRefetch As Boolean = False,
                                                       Optional quietErrors As Boolean = True,
                                                       Optional forceRefresh As Boolean = False) As Task
            If _skillsAppsService Is Nothing OrElse Not IsClientRunning() Then
                Return
            End If

            Dim normalizedThreadId = NormalizeCatalogIdentifier(preferredThreadId)
            If String.IsNullOrWhiteSpace(normalizedThreadId) Then
                normalizedThreadId = NormalizeCatalogIdentifier(GetVisibleThreadId())
            End If

            If Not forceRefresh AndAlso
               Not forceRefetch AndAlso
               ShouldReuseAppsCatalog(normalizedThreadId) Then
                Return
            End If

            Await _appsCatalogRefreshGate.WaitAsync().ConfigureAwait(True)
            Try
                If Not forceRefresh AndAlso
                   Not forceRefetch AndAlso
                   ShouldReuseAppsCatalog(normalizedThreadId) Then
                    Return
                End If

                Dim apps = Await _skillsAppsService.ListAppsAsync(normalizedThreadId,
                                                                  forceRefetch,
                                                                  CancellationToken.None).ConfigureAwait(True)
                ApplyAppsCatalog(apps, normalizedThreadId)
            Catch ex As Exception
                If quietErrors Then
                    AppendProtocol("debug", $"app/list failed: {ex.Message}")
                Else
                    AppendAndShowSystemMessage($"Could not refresh apps: {ex.Message}",
                                               $"Could not refresh apps: {ex.Message}",
                                               isError:=True)
                End If
            Finally
                _appsCatalogRefreshGate.Release()
            End Try
        End Function

        Private Function BuildTurnInputItemsForSubmission(inputText As String) As IReadOnlyList(Of TurnInputItem)
            Dim text = If(inputText, String.Empty)
            Dim items As New List(Of TurnInputItem) From {
                TurnInputItem.TextItem(text)
            }

            Dim tokens = ExtractDollarTokens(text)
            If tokens.Count = 0 Then
                Return items
            End If

            Dim addedSkillPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim addedAppIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each token In tokens
                Dim normalizedToken = NormalizeLookupToken(token)
                If String.IsNullOrWhiteSpace(normalizedToken) Then
                    Continue For
                End If

                Dim matchedSkill As SkillCatalogEntry = Nothing
                If _skillsCatalogByToken.TryGetValue(normalizedToken, matchedSkill) AndAlso matchedSkill IsNot Nothing Then
                    Dim skillPath = If(matchedSkill.Path, String.Empty).Trim()
                    If Not String.IsNullOrWhiteSpace(skillPath) AndAlso addedSkillPaths.Add(skillPath) Then
                        items.Add(TurnInputItem.SkillItem(matchedSkill.Name, skillPath))
                    End If

                    Continue For
                End If

                Dim matchedApp As AppCatalogEntry = Nothing
                If _appsCatalogByToken.TryGetValue(normalizedToken, matchedApp) AndAlso matchedApp IsNot Nothing Then
                    Dim appId = If(matchedApp.Id, String.Empty).Trim()
                    If Not String.IsNullOrWhiteSpace(appId) AndAlso addedAppIds.Add(appId) Then
                        Dim mentionPath = $"app://{appId}"
                        items.Add(TurnInputItem.MentionItem(matchedApp.Name, mentionPath))
                    End If
                End If
            Next

            Return items
        End Function

        Private Function BuildTurnComposerTokenSuggestions(query As String,
                                                           Optional maxItems As Integer = 12) As IReadOnlyList(Of TurnComposerTokenSuggestionEntry)
            Dim normalizedQuery = NormalizeLookupToken(query)
            Dim candidates As New List(Of TurnComposerTokenSuggestionEntry)()

            Dim seenAppIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each appEntry In _appsCatalogByToken.Values
                If appEntry Is Nothing Then
                    Continue For
                End If

                Dim appId = NormalizeLookupToken(appEntry.Id)
                If String.IsNullOrWhiteSpace(appId) OrElse Not seenAppIds.Add(appId) Then
                    Continue For
                End If

                Dim appName = If(appEntry.Name, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(appName) Then
                    appName = appId
                End If

                Dim appDescription = NormalizeSuggestionDescription(appEntry.Description)
                Dim appMetaText = "$" & appId
                If StringComparer.OrdinalIgnoreCase.Equals(NormalizeLookupToken(appName), appId) Then
                    appMetaText = String.Empty
                End If

                Dim score = ScoreTurnComposerTokenSuggestion(normalizedQuery, appId, appName)
                If score < 0 Then
                    Continue For
                End If

                candidates.Add(New TurnComposerTokenSuggestionEntry() With {
                    .Kind = "app",
                    .KindLabel = "App",
                    .InsertToken = appId,
                    .DisplayName = appName,
                    .MetaText = appMetaText,
                    .Description = appDescription,
                    .SortScore = score
                })
            Next

            Dim seenSkillPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each skillEntry In _skillsCatalogByToken.Values
                If skillEntry Is Nothing Then
                    Continue For
                End If

                Dim skillPath = If(skillEntry.Path, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(skillPath) OrElse Not seenSkillPaths.Add(skillPath) Then
                    Continue For
                End If

                Dim skillName = If(skillEntry.Name, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(skillName) Then
                    Continue For
                End If

                Dim skillToken = NormalizeLookupToken(skillName)
                If String.IsNullOrWhiteSpace(skillToken) Then
                    skillToken = SlugifyLookupToken(skillName)
                End If

                If String.IsNullOrWhiteSpace(skillToken) Then
                    Continue For
                End If

                Dim skillDescription = NormalizeSuggestionDescription(skillEntry.Description)
                Dim skillMetaText = String.Empty
                If Not StringComparer.OrdinalIgnoreCase.Equals(NormalizeLookupToken(skillName), skillToken) Then
                    skillMetaText = "$" & skillToken
                End If

                Dim score = ScoreTurnComposerTokenSuggestion(normalizedQuery, skillToken, skillName)
                If score < 0 Then
                    Continue For
                End If

                candidates.Add(New TurnComposerTokenSuggestionEntry() With {
                    .Kind = "skill",
                    .KindLabel = "Skill",
                    .InsertToken = skillToken,
                    .DisplayName = skillName,
                    .MetaText = skillMetaText,
                    .Description = skillDescription,
                    .SortScore = score
                })
            Next

            candidates.Sort(
                Function(left, right)
                    If left Is Nothing AndAlso right Is Nothing Then
                        Return 0
                    End If

                    If left Is Nothing Then
                        Return 1
                    End If

                    If right Is Nothing Then
                        Return -1
                    End If

                    Dim scoreCompare = left.SortScore.CompareTo(right.SortScore)
                    If scoreCompare <> 0 Then
                        Return scoreCompare
                    End If

                    Dim kindCompare = StringComparer.OrdinalIgnoreCase.Compare(left.Kind, right.Kind)
                    If kindCompare <> 0 Then
                        Return kindCompare
                    End If

                    Dim tokenCompare = StringComparer.OrdinalIgnoreCase.Compare(left.InsertToken, right.InsertToken)
                    If tokenCompare <> 0 Then
                        Return tokenCompare
                    End If

                    Return StringComparer.OrdinalIgnoreCase.Compare(left.DisplayName, right.DisplayName)
                End Function)

            Dim limit = Math.Max(1, maxItems)
            If candidates.Count > limit Then
                candidates.RemoveRange(limit, candidates.Count - limit)
            End If

            Return candidates
        End Function

        Private Shared Function ScoreTurnComposerTokenSuggestion(query As String,
                                                                 token As String,
                                                                 displayName As String) As Integer
            Dim normalizedQuery = NormalizeLookupToken(query)
            Dim normalizedToken = NormalizeLookupToken(token)
            Dim normalizedName = NormalizeLookupToken(displayName)

            If String.IsNullOrWhiteSpace(normalizedToken) Then
                Return -1
            End If

            If String.IsNullOrWhiteSpace(normalizedQuery) Then
                Return 100
            End If

            If StringComparer.OrdinalIgnoreCase.Equals(normalizedToken, normalizedQuery) Then
                Return 0
            End If

            If normalizedToken.StartsWith(normalizedQuery, StringComparison.OrdinalIgnoreCase) Then
                Return 10 + Math.Max(0, normalizedToken.Length - normalizedQuery.Length)
            End If

            If Not String.IsNullOrWhiteSpace(normalizedName) AndAlso
               normalizedName.StartsWith(normalizedQuery, StringComparison.OrdinalIgnoreCase) Then
                Return 40 + Math.Max(0, normalizedName.Length - normalizedQuery.Length)
            End If

            If normalizedToken.Contains(normalizedQuery, StringComparison.OrdinalIgnoreCase) Then
                Return 70
            End If

            If Not String.IsNullOrWhiteSpace(normalizedName) AndAlso
               normalizedName.Contains(normalizedQuery, StringComparison.OrdinalIgnoreCase) Then
                Return 85
            End If

            Return -1
        End Function

        Private Shared Function ExtractDollarTokens(inputText As String) As IReadOnlyList(Of String)
            Dim tokens As New List(Of String)()
            If String.IsNullOrWhiteSpace(inputText) Then
                Return tokens
            End If

            Dim text = If(inputText, String.Empty)
            Dim i = 0

            While i < text.Length
                If text(i) <> "$"c Then
                    i += 1
                    Continue While
                End If

                If i > 0 AndAlso IsTokenChar(text(i - 1)) Then
                    i += 1
                    Continue While
                End If

                If i + 1 >= text.Length Then
                    i += 1
                    Continue While
                End If

                Dim start = i + 1
                Dim firstChar = text(start)
                If Not IsTokenStartChar(firstChar) Then
                    i += 1
                    Continue While
                End If

                Dim endIndex = start + 1
                While endIndex < text.Length AndAlso IsTokenChar(text(endIndex))
                    endIndex += 1
                End While

                If endIndex > start Then
                    Dim token = text.Substring(start, endIndex - start).TrimEnd("."c)
                    If Not String.IsNullOrWhiteSpace(token) Then
                        tokens.Add(token)
                    End If
                    i = endIndex
                Else
                    i += 1
                End If
            End While

            Return tokens
        End Function

        Private Shared Function IsTokenStartChar(ch As Char) As Boolean
            Return Char.IsLetter(ch) OrElse ch = "_"c
        End Function

        Private Shared Function IsTokenChar(ch As Char) As Boolean
            Return Char.IsLetterOrDigit(ch) OrElse
                   ch = "-"c OrElse
                   ch = "_"c OrElse
                   ch = "."c
        End Function

        Private Function BuildSkillCatalogCwdCandidates(preferredThreadId As String) As IReadOnlyList(Of String)
            Dim values As New List(Of String)()

            AddNormalizedCwdCandidate(values, EffectiveThreadWorkingDirectory())
            AddNormalizedCwdCandidate(values, _currentThreadCwd)
            AddNormalizedCwdCandidate(values, _newThreadTargetOverrideCwd)

            Dim sourceLabel As String = Nothing
            AddNormalizedCwdCandidate(values, ResolveNewThreadTargetCwd(sourceLabel))

            Dim preferredEntry = FindThreadListEntryById(preferredThreadId)
            If preferredEntry IsNot Nothing Then
                AddNormalizedCwdCandidate(values, preferredEntry.Cwd)
            End If

            If values.Count = 0 Then
                AddNormalizedCwdCandidate(values, Environment.CurrentDirectory)
            End If

            Return values
        End Function

        Private Shared Sub AddNormalizedCwdCandidate(target As IList(Of String), rawValue As String)
            If target Is Nothing Then
                Return
            End If

            Dim normalized = NormalizeProjectPath(rawValue)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return
            End If

            For Each existing In target
                If StringComparer.OrdinalIgnoreCase.Equals(existing, normalized) Then
                    Return
                End If
            Next

            target.Add(normalized)
        End Sub

        Private Shared Function BuildCwdFingerprint(cwds As IReadOnlyList(Of String)) As String
            If cwds Is Nothing OrElse cwds.Count = 0 Then
                Return String.Empty
            End If

            Dim ordered As New List(Of String)(cwds)
            ordered.Sort(StringComparer.OrdinalIgnoreCase)
            Return String.Join("|", ordered)
        End Function

        Private Function ShouldReuseSkillsCatalog(cwdFingerprint As String) As Boolean
            Dim normalizedFingerprint = If(cwdFingerprint, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(normalizedFingerprint) Then
                Return False
            End If

            If Not StringComparer.Ordinal.Equals(_skillsCatalogCwdFingerprint, normalizedFingerprint) Then
                Return False
            End If

            If _skillsCatalogLoadedAtUtc = DateTimeOffset.MinValue Then
                Return False
            End If

            Return DateTimeOffset.UtcNow - _skillsCatalogLoadedAtUtc < SkillsCatalogRefreshInterval
        End Function

        Private Function ShouldReuseAppsCatalog(threadId As String) As Boolean
            Dim normalizedThreadId = NormalizeCatalogIdentifier(threadId)
            If Not StringComparer.Ordinal.Equals(_appsCatalogThreadId, normalizedThreadId) Then
                Return False
            End If

            If _appsCatalogLoadedAtUtc = DateTimeOffset.MinValue Then
                Return False
            End If

            Return DateTimeOffset.UtcNow - _appsCatalogLoadedAtUtc < AppsCatalogRefreshInterval
        End Function

        Private Sub ApplySkillsCatalog(skills As IReadOnlyList(Of SkillSummary), cwdFingerprint As String)
            Dim nextByToken As New Dictionary(Of String, SkillCatalogEntry)(StringComparer.OrdinalIgnoreCase)
            If skills IsNot Nothing Then
                For Each skill In skills
                    If skill Is Nothing OrElse Not skill.Enabled Then
                        Continue For
                    End If

                    Dim skillPath = If(skill.Path, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(skillPath) Then
                        Continue For
                    End If

                    Dim skillName = If(skill.Name, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(skillName) Then
                        Continue For
                    End If

                    Dim entry As New SkillCatalogEntry() With {
                        .Name = skillName,
                        .Description = NormalizeSuggestionDescription(skill.Description),
                        .Path = skillPath
                    }

                    For Each tokenKey In BuildSkillLookupKeys(skillName)
                        If String.IsNullOrWhiteSpace(tokenKey) OrElse nextByToken.ContainsKey(tokenKey) Then
                            Continue For
                        End If

                        nextByToken(tokenKey) = entry
                    Next
                Next
            End If

            _skillsCatalogByToken.Clear()
            For Each kvp In nextByToken
                _skillsCatalogByToken(kvp.Key) = kvp.Value
            Next

            _skillsCatalogCwdFingerprint = If(cwdFingerprint, String.Empty).Trim()
            _skillsCatalogLoadedAtUtc = DateTimeOffset.UtcNow
            If Not _suppressTurnComposerSuggestionRefresh Then
                RefreshTurnComposerTokenSuggestionsPopup(triggerCatalogWarmup:=False)
            End If
        End Sub

        Private Sub ApplyAppsCatalog(apps As IReadOnlyList(Of AppSummary), threadId As String)
            Dim nextByToken As New Dictionary(Of String, AppCatalogEntry)(StringComparer.OrdinalIgnoreCase)
            If apps IsNot Nothing Then
                For Each app In apps
                    If app Is Nothing OrElse
                       Not app.IsAccessible OrElse
                       Not app.IsEnabled Then
                        Continue For
                    End If

                    Dim appId = If(app.Id, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(appId) Then
                        Continue For
                    End If

                    Dim appName = If(app.Name, String.Empty).Trim()
                    If String.IsNullOrWhiteSpace(appName) Then
                        appName = appId
                    End If

                    Dim entry As New AppCatalogEntry() With {
                        .Id = appId,
                        .Name = appName,
                        .Description = NormalizeSuggestionDescription(app.Description)
                    }

                    For Each tokenKey In BuildAppLookupKeys(app)
                        If String.IsNullOrWhiteSpace(tokenKey) OrElse nextByToken.ContainsKey(tokenKey) Then
                            Continue For
                        End If

                        nextByToken(tokenKey) = entry
                    Next
                Next
            End If

            _appsCatalogByToken.Clear()
            For Each kvp In nextByToken
                _appsCatalogByToken(kvp.Key) = kvp.Value
            Next

            _appsCatalogThreadId = NormalizeCatalogIdentifier(threadId)
            _appsCatalogLoadedAtUtc = DateTimeOffset.UtcNow
            If Not _suppressTurnComposerSuggestionRefresh Then
                RefreshTurnComposerTokenSuggestionsPopup(triggerCatalogWarmup:=False)
            End If
        End Sub

        Private Shared Function BuildSkillLookupKeys(skillName As String) As IReadOnlyList(Of String)
            Dim keys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            AddLookupKey(keys, skillName)
            AddLookupKey(keys, SlugifyLookupToken(skillName))
            Return New List(Of String)(keys)
        End Function

        Private Shared Function BuildAppLookupKeys(app As AppSummary) As IReadOnlyList(Of String)
            Dim keys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If app Is Nothing Then
                Return Array.Empty(Of String)()
            End If

            AddLookupKey(keys, app.Id)
            AddLookupKey(keys, SlugifyLookupToken(app.Id))
            AddLookupKey(keys, app.Name)
            AddLookupKey(keys, SlugifyLookupToken(app.Name))

            Dim installTokens = ExtractInstallUrlTokens(app.InstallUrl)
            For Each token In installTokens
                AddLookupKey(keys, token)
            Next

            Return New List(Of String)(keys)
        End Function

        Private Shared Sub AddLookupKey(target As HashSet(Of String), rawValue As String)
            If target Is Nothing Then
                Return
            End If

            Dim normalized = NormalizeLookupToken(rawValue)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return
            End If

            target.Add(normalized)
        End Sub

        Private Shared Function NormalizeLookupToken(rawValue As String) As String
            Dim normalized = If(rawValue, String.Empty).Trim()
            If normalized.StartsWith("$", StringComparison.Ordinal) Then
                normalized = normalized.Substring(1)
            End If

            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            Return normalized.Trim().ToLowerInvariant()
        End Function

        Private Shared Function SlugifyLookupToken(rawValue As String) As String
            Dim value = If(rawValue, String.Empty)
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder()
            Dim previousWasSeparator = False

            For Each ch In value
                If Char.IsLetterOrDigit(ch) Then
                    builder.Append(Char.ToLowerInvariant(ch))
                    previousWasSeparator = False
                    Continue For
                End If

                If ch = "-"c OrElse ch = "_"c OrElse ch = "."c OrElse ch = "/"c OrElse ch = "\"c OrElse Char.IsWhiteSpace(ch) Then
                    If Not previousWasSeparator AndAlso builder.Length > 0 Then
                        builder.Append("-"c)
                        previousWasSeparator = True
                    End If
                End If
            Next

            Dim slug = builder.ToString().Trim("-"c)
            Return NormalizeLookupToken(slug)
        End Function

        Private Shared Function ExtractInstallUrlTokens(installUrl As String) As IReadOnlyList(Of String)
            Dim tokens As New List(Of String)()
            Dim raw = If(installUrl, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(raw) Then
                Return tokens
            End If

            Dim uri As Uri = Nothing
            If Not Uri.TryCreate(raw, UriKind.Absolute, uri) OrElse uri Is Nothing Then
                Return tokens
            End If

            Dim segments = uri.AbsolutePath.Split({"/"c}, StringSplitOptions.RemoveEmptyEntries)
            If segments.Length = 0 Then
                Return tokens
            End If

            tokens.Add(segments(segments.Length - 1))
            If segments.Length >= 2 Then
                tokens.Add(segments(segments.Length - 2))
            End If

            Return tokens
        End Function

        Private Shared Function NormalizeSuggestionDescription(rawValue As String) As String
            Dim text = If(rawValue, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(text) Then
                Return String.Empty
            End If

            Return text.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        End Function

        Private Sub ClearSkillsAndAppsCatalogCache()
            _skillsCatalogByToken.Clear()
            _appsCatalogByToken.Clear()
            _skillsCatalogCwdFingerprint = String.Empty
            _skillsCatalogLoadedAtUtc = DateTimeOffset.MinValue
            _appsCatalogThreadId = String.Empty
            _appsCatalogLoadedAtUtc = DateTimeOffset.MinValue
        End Sub

        Private Shared Function NormalizeCatalogIdentifier(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Return value.Trim()
        End Function
    End Class
End Namespace
