Attribute VB_Name = "VefeyatEklenti0341"
Option Explicit

' Pass UTF-16 string pointers directly; VBA MsgBox uses the ANSI code page.
#If VBA7 Then
Private Declare PtrSafe Function VfMessageBoxW Lib "user32" Alias "MessageBoxW" (ByVal owner As LongPtr, ByVal message As LongPtr, ByVal title As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function VfMessageBoxW Lib "user32" Alias "MessageBoxW" (ByVal owner As Long, ByVal message As Long, ByVal title As Long, ByVal flags As Long) As Long
#End If

' Local UTF-8 CSV catalog. No network or installer.
' Keep this source ASCII so it imports consistently into the VBA editor.
Private mVfRows As Collection
Private mVfSearchText() As String
Private mVfNameText() As String
Private mVfNameKeys() As String
Private mVfFoldedKeys() As String
Private mVfIndexReady As Boolean
Private mVfUniqueCounter As Long
Private mVfQueryOwner As Collection
Private mVfQueryKeys(0 To 127) As String
Private mVfQueryResults(0 To 127) As Collection
Private mVfQueryCount As Long
Private mVfQueryNext As Long

Private Const VfpPrefixVariable As String = "VefeyatPrefix"
Private Const VfpScopeVariable As String = "VefeyatScope"
Private Const VfpAuditVariable As String = "VefeyatAuditMode"
Private Const VfpInvalidSetting As Long = vbObjectError + 2301
Private Const VfpInvalidDocument As Long = vbObjectError + 2302
Private Const VfpCheckFailure As Long = vbObjectError + 2303

' Only an explicit leading death marker is changed. The rest of a catalog
' expression, including spaces, uncertainty, ranges, and punctuation, is kept.

Private Const VfaTagPrefix As String = "Vefeyat.Date."
Private Const VfaVariablePrefix As String = "Vefeyat.Binding."
Private Const VfaNextVariable As String = "Vefeyat.NextBinding"
Private Const VfaNamespaceVariable As String = "Vefeyat.BindingNamespace"
Private Const VfaManualBookmarkPrefix As String = "_VfM_"
Private Const VfaInvalidData As Long = vbObjectError + 2401
Private Const VfaCheckFailure As Long = vbObjectError + 2402
Private Const VfaMaxName As Long = 200
Private Const VfaMaxDate As Long = 500
Private mVfaRandomized As Boolean
Private mVfaCommandChanged As Boolean

' No catalog access or document writes are allowed in this enabled callback.

' Normal use leaves Word's field display preference untouched. The selected
' managed date is deselected after a successful explicit insertion instead.
Private mVfvAnchor As Word.Range
Private mVfvCaret As Word.Range
Private mVfvViews As Collection

' Idle insertion is deferred. Never pay its document-wide tracking cost on
' explicit commands, and never schedule its unverified Word timer.
Private Const VflEnabled As Boolean = False
Private mVflPending As Boolean
Private mVflBusy As Boolean
Private mVflSuspended As Boolean
Private mVflStates As Collection
Private mVflStatus As String
Private mVflDueAt As Date
Private mVflLastTick As Date
Private mVflTicks As Long
Private mVflPhase As String
Private mVflProbePath As String

' A source spelling is a proposal, never a claim that all other forms are wrong.
' The selected occurrence is edited after choosing the preferred name form.

' Open only a catalog's official DIA article, after an explicit user command.
' No document text is sent to a search service and no URL is invented.

' Deterministic local context, not a language model. Inspect at most 800 UTF-16
' units beside a discovery hit; never read Word ranges or call the catalog here.

' Citation islands are classified once per rendered footnote, independently
' of the names later offered for review. All offsets are UTF-16 string offsets.

Private mVfzDoc As Word.Document
Private mVfzNotes As Collection
Private mVfzNoteItems As Variant
Private mVfzNoteCount As Long
Private mVfzPatterns As Object, mVfzHeads As Object, mVfzCache As Object
Private mVfzEpoch As Long

' Indexed, bounded text matching; no per-record full-document Find loop.
' All decisions and an editable review list precede the first document write.

Private Const VfbCancelError As Long = vbObjectError + 2450
Private mVfbRunning As Boolean
Private mVfbExecuting As Boolean
Private mVfbExecuted As Boolean
Private mVfbIndexBuilding As Boolean
Private mVfbProgress As VefeyatKisiSec0341
Private mVfbLastPulse As Single
Private mVfbLastStage As String
Private mVfbHasPainted As Boolean
Private mVfbLastAllowCancel As Boolean
Private mVfbParentStage As String
Private mVfbParentCompleted As Long
Private mVfbParentTotal As Long
Private mVfbCachedRows As Collection
Private mVfbCachedPatterns As Object
Private mVfbCachedHeads As Object

' Document-first shortlist. Prefixes only prune work; the existing exact/context
' matcher still decides whether a complete name is an admissible discovery.
Private mVfxOwner As Collection
Private mVfxCodes() As String
Private mVfxPatternOwner As Collection
Private mVfxWanted As Object, mVfxPatterns As Object, mVfxHeads As Object
Private mVfxBusy As Boolean
Private mVfxLastRows As Long, mVfxLastHeads As Long, mVfxLastReused As Boolean

' Per-open-document RAM only: no bookmarks, file cache or document variables.
' The inverted word index points to bounded plain-text blocks, not whole stories.
Private mVfmStates As Collection
Private mVfmBusy As Boolean
Private mVfmNeedsPrune As Boolean
Private mVfmWorkDoc As Word.Document
Private mVfmWorkState As Object
Private Const VfmBlockSize As Long = 4096
Private Const VfmLookupSeconds As Single = 2!

' A modeless, occurrence-by-occurrence review. Live Word ranges move with edits;
' every write revalidates both the name and the context that was displayed.
Private mVfwDoc As Word.Document
Private mVfwHits As Collection
Private mVfwRows As Collection
Private mVfwCatalog As Collection
Private mVfwPatterns As Object
Private mVfwPanel As VefeyatKisiSec0341
Private mVfwContext As Word.Range
Private mVfwContextText As String
Private mVfwIndex As Long
Private mVfwAdded As Long
Private mVfwSkipped As Long
Private mVfwBusy As Boolean
Private mVfwRendering As Boolean
Private mVfwPrepared As Boolean
Private mVfwClosing As Boolean
Private mVfwSession As Long
Private mVfwPaused As Boolean
Private mVfwPausing As Boolean
Private mVfwFaulted As Boolean
Private mVfwExclusions As Object
Private mVfwExclusionStamp As String
Private mVfwIgnoredSkipped As Long
Private mVfwVariables As Object
Private mVfwVariableCount As Long
Private mVfwSkippedIndices As Object
Private mVfwAppliedIndices As Object
Private mVfwChosenPeople As Object
Private mVfwConfirmed As Object
Private mVfwCandidateCache As Object
Private mVfwVisited As Object
Private mVfwAutomatic As Object
Private mVfwFirstVisited As Long
Private mVfwSkipDated As Boolean
Private mVfwFirstOnly As Boolean
Private mVfwDatedSkipped As Long
Private mVfwRepeatedSkipped As Long
Private mVfwState As Long
Private mVfwLeft As Single, mVfwTop As Single, mVfwWidth As Single, mVfwHeight As Single
Private mVfwReserved As Single
Private mVfwAppliedLeft As Single, mVfwAppliedTop As Single, mVfwAppliedHeight As Single
Private mVfwLayout As Boolean

' Fixed operation keys and numbers only. Never store text, names, paths or URLs.
Private mVfgStats As Object

Private Const VfiVariable As String = "VefeyatReviewIgnores"
Private Const VfiInvalid As Long = vbObjectError + 2701

' Match literal spellings, never accent-fold a common word into a person's name.

' Decode the existing length-prefixed identity, without guessing from today's CSV.

Private mVfqCatalogOwner As Collection
Private mVfqCatalogIndex As Object

' A canonical official article link is a durable catalog identity. Never use
' name similarity to merge people, and never replace recorded dates on import.

Private Const VfuModeVariable As String = "VefeyatNameForm"
Private Const VfuNamesVariable As String = "VefeyatPreferredNames"
Private Const VfuInvalid As Long = vbObjectError + 2801

' Bounded length-prefixed UTF-16 data. No executable text, JSON evaluation,
' catalog rows, document body, file path or macro name is accepted in a profile.

Private Const VdiPrefix As String = "Vefeyat.Context.C"
Private Const VdiRevisionName As String = "Vefeyat.Context.Revision"
Private Const VdiLimit As Long = 200
Private Const VdiInvalid As Long = vbObjectError + 2851

' Evidence comes only from the supplied six-column catalog. Scores order
' proposals; they are not calibrated probabilities or permission to edit.

Private Const VrgMode As String = "Vefeyat.Registry.Enabled"
Private Const VrgPrefix As String = "Vefeyat.Registry.R"
Private Const VrgLimit As Long = 2000
Private Const VrgInvalid As Long = vbObjectError + 2871
Private mVrgPaused As Boolean
Private mVrgCatalogOwner As Collection
Private mVrgCatalogIndex As Object, mVrgLegacyIndex As Object
Private mVrgCatalogCount As Long

' Bookmarks and variables travel with the Word document; no external path or
' machine identity is part of a record. No Content Controls/API declaration.

Private Const VbpInvalid As Long = vbObjectError + 2861
Private mVbpDoc As Word.Document
Private mVbpSnapshots As Collection
Private mVbpPrepared As Collection
Private mVbpReady As Boolean

' Reuse discovery, identity and name forms. No text is written during planning.

Private mVftPanel As VefeyatKisiSec0341
Private mVftDoc As Word.Document
Private mVftItems As Collection
Private mVftHistory As Collection
Private mVftEpoch As Long
Private mVftBusy As Boolean
Private mVftClosing As Boolean
Private mVftMode As String

' Inventory discovery uses the shared strict whole-document matcher.

Private Const VfjPageSize As Long = 250
Private mVfjDoc As Word.Document
Private mVfjCatalog As Collection
Private mVfjPanel As VefeyatKisiSec0341
Private mVfjItems As Variant
Private mVfjCount As Long, mVfjPage As Long, mVfjEpoch As Long
Private mVfjBusy As Boolean, mVfjClosing As Boolean, mVfjPreparing As Boolean
Private mVfjReady As Boolean, mVfjFaulted As Boolean
Private mVfjCache As Object, mVfjFirst As Object

' Explicit Windows acceptance harness. It uses one disposable synthetic document,
' never an open author's document. Registry acceptance saves a disposable DOCX
' in TEMP, reopens it and deletes it. No catalog is imported.

Private Const VfrVersion As String = "0.3.41"
Private mVfrRibbon As Object
Private mVfrEvents As VefeyatOlaylar0341
Private mVfrInsertBusy As Boolean

' Explicit, user-run setup inside a separate, user-authored Word template.
' Never copy code from Normal or access the VBProject model.
' All ZIP and XML operations are local Windows COM calls, without shell commands.
Private mVfsBusy As Boolean
Private mVfcBusy As Boolean
Private mVdiReason As String
Private mVfaSkipReason As String
Private Const VfsSourceTemplateName As String = "VefeyatHazirlik.dotm"
Private Const VfsInstalledModule As String = "VefeyatEklenti0341"
Private Const VfsTemplateName As String = "VefeyatEklenti.dotm"
Private Const VfsRibbonName As String = "VefeyatRibbon.xml"
Private Const VfsUiNamespace As String = "http://schemas.microsoft.com/office/2009/07/customui"
Private Const VfsUiRelationship As String = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
Private Const VfsRelationshipsNamespace As String = "http://schemas.openxmlformats.org/package/2006/relationships"
Private Const VfsContentTypesNamespace As String = "http://schemas.openxmlformats.org/package/2006/content-types"

Private Function VfMessage(ByVal message As String, Optional ByVal flags As Long = 0, Optional ByVal title As String = "Vefeyat") As Long
    VfMessage = VfMessageBoxW(0, StrPtr(message), StrPtr(title), flags Or &H2000&)
End Function

Private Function VfText(ByVal escaped As String) As String
    Dim i As Long, j As Long, code As Long, digit As Long
    Dim part As String, result As String, valid As Boolean
    i = 1
    Do While i <= Len(escaped)
        If Mid$(escaped, i, 2) = "\u" And i + 5 <= Len(escaped) Then
            part = Mid$(escaped, i + 2, 4)
            valid = True
            code = 0
            For j = 1 To 4
                digit = InStr(1, "0123456789abcdef", LCase$(Mid$(part, j, 1)), vbBinaryCompare) - 1
                If digit < 0 Then
                    valid = False
                    Exit For
                End If
                code = code * 16 + digit
            Next j
            If valid Then
                If code > 32767 Then code = code - 65536
                result = result & ChrW(code)
                i = i + 6
            Else
                result = result & Mid$(escaped, i, 1)
                i = i + 1
            End If
        Else
            result = result & Mid$(escaped, i, 1)
            i = i + 1
        End If
    Loop
    VfText = result
End Function

Private Function VefeyatDataPath() As String
    Dim root As String
    root = Environ$("LOCALAPPDATA")
    If Len(root) = 0 Then VfFail "Windows LOCALAPPDATA folder is unavailable."
    If Right$(root, 1) = "\" Then root = Left$(root, Len(root) - 1)
    VfCheckLocalPath root
    VefeyatDataPath = root & "\Vefeyat\WordVBA\people.csv"
End Function

Private Function VefeyatEnsureData() As Boolean
    Dim fso As Object, catalogPath As String
    On Error GoTo Failed
    If Not mVfRows Is Nothing Then
        VefeyatEnsureData = True
        Exit Function
    End If
    catalogPath = VefeyatDataPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(catalogPath) Then
        If Not VfImportCatalog(False) Then Exit Function
    End If
    If mVfRows Is Nothing Then VfLoadCache catalogPath
    VefeyatEnsureData = True
    Exit Function
Failed:
    VfMessage VfText("Veriler a\u00e7\u0131lamad\u0131: ") & Err.Description, vbExclamation, "Vefeyat"
End Function

Private Sub VefeyatChooseData()
    On Error GoTo Failed
    If VfImportCatalog(True) Then
        ' A new import is loaded lazily; no old search index survives it.
        Set mVfRows = Nothing
        mVfIndexReady = False
    End If
    Exit Sub
Failed:
    VfMessage VfText("CSV al\u0131namad\u0131: ") & Err.Description, vbExclamation, "Vefeyat"
End Sub

Private Function VefeyatRecordCount() As Long
    If VefeyatEnsureData() Then VefeyatRecordCount = mVfRows.Count
End Function

' Strip only a recognized Turkish inflection after the last apostrophe.
' Internal name punctuation such as Kivamu's-sunne is never a suffix.
Private Function VfNameQuery(ByVal value As String) As String
    Dim plain As String, at As Long, ending As String
    VfNameQuery = value
    plain = Replace(Replace(Replace(value, ChrW(&H2019), "'"), ChrW(&H2018), "'"), ChrW(&H2BC), "'")
    at = InStrRev(plain, "'")
    If at <= 1 Then Exit Function
    ending = Mid$(plain, at + 1)
    ' Circumflex vowels belong to a transliterated name, not a Turkish ending.
    If InStr(1, ending, VfText("\u00e2"), vbTextCompare) > 0 Or _
       InStr(1, ending, VfText("\u00ee"), vbTextCompare) > 0 Or _
       InStr(1, ending, VfText("\u00fb"), vbTextCompare) > 0 Then Exit Function
    ending = VefeyatNormalize(ending)
    Select Case ending
        Case "in", "nin", "un", "nun", "a", "e", "ya", "ye", "na", "ne", "i", "u", "yi", "yu", _
             "da", "de", "ta", "te", "nda", "nde", "dan", "den", "tan", "ten", "ndan", "nden", _
             "la", "le", "yla", "yle", "daki", "deki", "taki", "teki", "ndaki", "ndeki", _
             "ca", "ce", "casina", "cesine", "inki", "ninki", "unki", "nunki"
            VfNameQuery = Left$(value, at - 1)
    End Select
End Function

Private Function VfFollowingSuffix(ByVal value As String) As String
    Dim i As Long, ch As String, part As String
    If Len(value) < 2 Then Exit Function
    ch = Left$(value, 1)
    If ch <> "'" And ch <> ChrW(&H2019) And ch <> ChrW(&H2018) And ch <> ChrW(&H2BC) Then Exit Function
    For i = 2 To Len(value)
        ch = VefeyatNormalize(Mid$(value, i, 1))
        If Len(ch) <> 1 Then Exit For
        If ch < "a" Or ch > "z" Then Exit For
    Next i
    part = Left$(value, i - 1)
    If VfNameQuery("X" & part) = "X" Then VfFollowingSuffix = part
End Function

Private Function VfLookupTokens(ByVal value As String) As String
    Dim part As Variant, result As String
    For Each part In Split(value, " ")
        Select Case CStr(part)
            Case "", "el", "er", "es", "et", "ez", "en", "ed", "al"
            Case Else
                If Len(result) > 0 Then result = result & " "
                result = result & CStr(part)
        End Select
    Next part
    VfLookupTokens = result
End Function

' All meaningful query parts must match distinct words in ONE source name.
' This fallback is used only after the regular lookup found no candidates.
Private Function VfLooseNameMatch(ByVal name As String, ByVal query As String) As Long
    Dim wanted As Variant, words As Variant, used() As Boolean, matched() As Boolean
    Dim i As Long, j As Long, limit As Long, found As Boolean, approximate As Boolean
    query = VfLookupTokens(query)
    If InStr(query, " ") = 0 Then Exit Function
    If VfContainsTokens(name, query) Then
        VfLooseNameMatch = 2
        Exit Function
    End If
    wanted = Split(query, " ")
    words = Split(VfLookupTokens(name), " ")
    If UBound(wanted) > UBound(words) Then Exit Function
    ReDim used(0 To UBound(words))
    ReDim matched(0 To UBound(wanted))
    For i = 0 To UBound(wanted)
        For j = 0 To UBound(words)
            If Not used(j) And CStr(wanted(i)) = CStr(words(j)) Then
                used(j) = True: matched(i) = True
                Exit For
            End If
        Next j
    Next i
    For i = 0 To UBound(wanted)
        If Not matched(i) Then
            If Len(CStr(wanted(i))) < 5 Then Exit Function
            limit = 1
            If Len(CStr(wanted(i))) >= 9 Then limit = 2
            found = False
            For j = 0 To UBound(words)
                If Not used(j) Then
                    If VfNearWord(CStr(words(j)), CStr(wanted(i)), limit) Then
                        used(j) = True: found = True: approximate = True
                        Exit For
                    End If
                End If
            Next j
            If Not found Then Exit Function
        End If
    Next i
    VfLooseNameMatch = 2
    If approximate Then VfLooseNameMatch = 1
End Function

Private Function VefeyatFind(ByVal query As String, Optional ByVal includeDescriptions As Boolean = True) As Collection
    Dim result As New Collection
    Dim order() As Long, ranks() As Long, found As Long, position As Long, rank As Long
    Dim term As String, compact As String, folded As String, cacheKey As String, row As Variant
    Dim i As Long, j As Long, score As Long, best As Long
    Set VefeyatFind = result
    If Not VefeyatEnsureData() Then Exit Function
    term = VefeyatNormalize(VfNameQuery(query))
    If Len(term) = 0 Then Exit Function
    compact = Replace(term, " ", "")
    If Not mVfIndexReady Then VfBuildIndex
    If mVfRows.Count = 0 Then Exit Function
    If Not (mVfQueryOwner Is mVfRows) Then
        Set mVfQueryOwner = mVfRows
        For i = 0 To 127
            Set mVfQueryResults(i) = Nothing
        Next i
        mVfQueryCount = 0
        mVfQueryNext = 0
    End If
    cacheKey = CStr(includeDescriptions) & ":" & term
    For i = 0 To mVfQueryCount - 1
        If mVfQueryKeys(i) = cacheKey Then
            Set VefeyatFind = VfCopyResults(mVfQueryResults(i))
            Exit Function
        End If
    Next i
    folded = VfCollapseLetters(compact)
    ReDim order(1 To mVfRows.Count)
    ReDim ranks(1 To mVfRows.Count)
    For i = 1 To mVfRows.Count
        best = 0
        For j = 0 To 2
            score = VfNameMatch(mVfNameText(i, j), mVfNameKeys(i, j), term, compact, mVfFoldedKeys(i, j), folded)
            If score > best Then best = score
        Next j
        row = mVfRows.Item(i)
        rank = 0
        If best > 0 Then
            rank = best * 1000 + VfProminence(row, compact)
        ElseIf includeDescriptions Then
            If VfContainsTokens(mVfSearchText(i), term) Then rank = 1
        End If
        If rank > 0 Then
            found = found + 1
            position = found
            Do While position > 1
                If ranks(position - 1) >= rank Then Exit Do
                ranks(position) = ranks(position - 1)
                order(position) = order(position - 1)
                position = position - 1
            Loop
            ranks(position) = rank
            order(position) = i
        End If
    Next i
    If found = 0 And InStr(VfLookupTokens(term), " ") > 0 Then
        For i = 1 To mVfRows.Count
            best = 0
            For j = 0 To 2
                score = VfLooseNameMatch(mVfNameText(i, j), term)
                If score > best Then best = score
            Next j
            If best > 0 Then
                row = mVfRows(i)
                rank = best * 1000 + VfProminence(row, compact)
                found = found + 1
                position = found
                Do While position > 1
                    If ranks(position - 1) >= rank Then Exit Do
                    ranks(position) = ranks(position - 1)
                    order(position) = order(position - 1)
                    position = position - 1
                Loop
                ranks(position) = rank
                order(position) = i
            End If
        Next i
    End If
    For i = 1 To found
        result.Add mVfRows.Item(order(i))
    Next i
    ' Bounded, session-local cache; callers cannot mutate the saved collection.
    mVfQueryKeys(mVfQueryNext) = cacheKey
    Set mVfQueryResults(mVfQueryNext) = VfCopyResults(result)
    mVfQueryNext = (mVfQueryNext + 1) Mod 128
    If mVfQueryCount < 128 Then mVfQueryCount = mVfQueryCount + 1
End Function

Private Function VfCopyResults(ByVal source As Collection) As Collection
    Dim result As New Collection, row As Variant
    For Each row In source
        result.Add row
    Next row
    Set VfCopyResults = result
End Function

' Compare each name separately. A compact key also joins words separated by
' apostrophes, dashes or spaces, without changing the selected document text.
Private Function VefeyatNameMatch(ByVal name As String, ByVal query As String) As Long
    Dim term As String, normalizedName As String
    term = VefeyatNormalize(VfNameQuery(query))
    normalizedName = VefeyatNormalize(name)
    VefeyatNameMatch = VfNameMatch(normalizedName, Replace(normalizedName, " ", ""), _
                                  term, Replace(term, " ", ""))
    If VefeyatNameMatch = 0 Then VefeyatNameMatch = VfLooseNameMatch(normalizedName, term)
End Function

Private Function VfNameMatch(ByVal name As String, ByVal compactName As String, _
                             ByVal term As String, ByVal compactTerm As String, _
                             Optional ByVal foldedName As String = "", Optional ByVal foldedTerm As String = "") As Long
    If Len(term) = 0 Or Len(name) = 0 Or Len(compactTerm) = 0 Then Exit Function
    If name = term Then
        VfNameMatch = 4
    ElseIf compactName = compactTerm Then
        VfNameMatch = 3
    ElseIf VfContainsTokens(name, term) Then
        VfNameMatch = 2
    ElseIf InStr(1, compactName, compactTerm, vbBinaryCompare) > 0 Then
        VfNameMatch = 1
    ElseIf VfSpellingMatch(name, compactName, compactTerm, foldedName, foldedTerm) Then
        VfNameMatch = 1
    End If
End Function

Private Function VfCollapseLetters(ByVal value As String) As String
    Dim i As Long, previous As String, ch As String, result As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch <> previous Then result = result & ch
        previous = ch
    Next i
    VfCollapseLetters = result
End Function

Private Function VfSpellingMatch(ByVal name As String, ByVal compactName As String, ByVal term As String, _
                                 Optional ByVal foldedName As String = "", Optional ByVal foldedTerm As String = "") As Boolean
    Dim token As Variant, limit As Long
    If Len(term) < 5 Or Len(term) > 80 Then Exit Function
    If Len(foldedTerm) = 0 Then foldedTerm = VfCollapseLetters(term)
    If Len(foldedName) = 0 Then foldedName = VfCollapseLetters(compactName)
    If InStr(1, foldedName, foldedTerm, vbBinaryCompare) > 0 Then
        VfSpellingMatch = True
        Exit Function
    End If
    ' At lengths >=5 (distance 1) or >=9 (distance 2), a possible edit
    ' match must retain a bigram. Tokens are substrings of compactName too.
    ' This only rules out impossible matches; the original scorer decides.
    If Not VfSharesBigram(compactName, term) Then Exit Function
    limit = 1
    If Len(term) >= 9 Then limit = 2
    If VfNearWord(compactName, term, limit) Then
        VfSpellingMatch = True
        Exit Function
    End If
    For Each token In Split(name, " ")
        If VfNearWord(CStr(token), term, limit) Then
            VfSpellingMatch = True
            Exit Function
        End If
    Next token
End Function

Private Function VfSharesBigram(ByVal name As String, ByVal term As String) As Boolean
    Dim i As Long
    For i = 1 To Len(term) - 1
        If InStr(1, name, Mid$(term, i, 2), vbBinaryCompare) > 0 Then
            VfSharesBigram = True
            Exit Function
        End If
    Next i
End Function

Private Function VfNearWord(ByVal word As String, ByVal term As String, ByVal limit As Long) As Boolean
    Dim previous() As Long, current() As Long, i As Long, j As Long
    Dim best As Long, cost As Long, value As Long
    If Abs(Len(word) - Len(term)) > limit Or Len(word) = 0 Then Exit Function
    ReDim previous(0 To Len(term))
    ReDim current(0 To Len(term))
    For j = 0 To Len(term): previous(j) = j: Next j
    For i = 1 To Len(word)
        current(0) = i
        best = i
        For j = 1 To Len(term)
            cost = 1
            If Mid$(word, i, 1) = Mid$(term, j, 1) Then cost = 0
            value = previous(j) + 1
            If current(j - 1) + 1 < value Then value = current(j - 1) + 1
            If previous(j - 1) + cost < value Then value = previous(j - 1) + cost
            current(j) = value
            If value < best Then best = value
        Next j
        If best > limit Then Exit Function
        For j = 0 To Len(term): previous(j) = current(j): Next j
    Next i
    VfNearWord = (current(Len(term)) <= limit)
End Function

' Explicit common-name priorities, not measured global fame statistics.
' They apply only to the ambiguous short query, never to an explicit full name.
Private Function VfProminence(ByRef row As Variant, ByVal query As String) As Long
    Dim shortName As String, preferred As String
    shortName = VefeyatNormalize(CStr(row(0)))
    Select Case VfCollapseLetters(query)
        Case "razi": preferred = "fahreddin er razi"
        Case "gazali": preferred = "gazzali"
        Case "buhari": preferred = "buhari muhammed b ismail"
        Case "muslim": preferred = "muslim b haccac"
        Case "taberi": preferred = "taberi muhammed b cerir"
    End Select
    If Len(preferred) > 0 And shortName = preferred Then VfProminence = 5000
    If shortName = query Then VfProminence = VfProminence + 100
End Function

Private Function VfContainsTokens(ByVal name As String, ByVal term As String) As Boolean
    Dim token As Variant
    If Len(term) = 0 Then Exit Function
    For Each token In Split(term, " ")
        If InStr(1, name, CStr(token), vbBinaryCompare) = 0 Then Exit Function
    Next token
    VfContainsTokens = True
End Function

Private Function VefeyatNormalize(ByVal text As String) As String
    Dim i As Long, code As Long, ch As String, result As String
    Dim pendingSpace As Boolean
    ' Fold Turkish I before lowercasing to avoid locale-specific results.
    text = Replace(text, ChrW(&H130), "i")
    text = Replace(text, ChrW(&H131), "i")
    text = Replace(text, "I", "i")
    text = LCase$(text)
    text = VfFoldLetters(text, "\u00e0\u00e1\u00e2\u00e3\u00e4\u00e5\u0101\u0103\u0105\u01ce", "a")
    text = VfFoldLetters(text, "\u00e7\u0107\u0109\u010b\u010d", "c")
    text = VfFoldLetters(text, "\u010f\u0111\u1e0d", "d")
    text = VfFoldLetters(text, "\u00e8\u00e9\u00ea\u00eb\u0113\u0115\u0117\u0119\u011b", "e")
    text = VfFoldLetters(text, "\u011d\u011f\u0121\u0123\u1e21", "g")
    text = VfFoldLetters(text, "\u0125\u0127\u1e25", "h")
    text = VfFoldLetters(text, "\u00ec\u00ed\u00ee\u00ef\u0129\u012b\u012d\u012f\u01d0", "i")
    text = VfFoldLetters(text, "\u0135", "j")
    text = VfFoldLetters(text, "\u0137\u1e33", "k")
    text = VfFoldLetters(text, "\u013a\u013c\u013e\u0142", "l")
    text = VfFoldLetters(text, "\u00f1\u0144\u0146\u0148\u1e45\u1e47", "n")
    text = VfFoldLetters(text, "\u00f2\u00f3\u00f4\u00f5\u00f6\u00f8\u014d\u014f\u0151\u01d2", "o")
    text = VfFoldLetters(text, "\u0155\u0157\u0159\u1e5b", "r")
    text = VfFoldLetters(text, "\u015b\u015d\u015f\u0161\u1e63", "s")
    text = VfFoldLetters(text, "\u0163\u0165\u0167\u1e6d", "t")
    text = VfFoldLetters(text, "\u00f9\u00fa\u00fb\u00fc\u0169\u016b\u016d\u016f\u0171\u0173\u01d4", "u")
    text = VfFoldLetters(text, "\u00fd\u00ff\u0177", "y")
    text = VfFoldLetters(text, "\u017a\u017c\u017e\u1e93", "z")
    text = Replace(text, ChrW(&HE6), "ae")
    text = Replace(text, ChrW(&H153), "oe")
    text = Replace(text, ChrW(&HDF), "ss")
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If (code >= &H300 And code <= &H36F) Or (code >= &H64B And code <= &H65F) Then
            ' Combining accents and Arabic vocalization marks are not search terms.
        ElseIf VfIsSeparator(code) Then
            If Len(result) > 0 Then pendingSpace = True
        Else
            If pendingSpace Then result = result & " "
            result = result & ch
            pendingSpace = False
        End If
    Next i
    result = " " & result & " "
    Do While InStr(1, result, " ibni ", vbBinaryCompare) > 0
        result = Replace(result, " ibni ", " ibn ")
    Loop
    VefeyatNormalize = Trim$(result)
End Function

Private Function VfFoldLetters(ByVal text As String, ByVal escaped As String, ByVal plain As String) As String
    Static decoded(0 To 25) As String
    Static sources(0 To 25) As String
    Dim letters As String, i As Long, slot As Long
    slot = -1
    If Len(plain) = 1 Then slot = AscW(plain) - 97
    If slot >= 0 And slot <= 25 Then
        If sources(slot) <> escaped Then
            decoded(slot) = VfText(escaped)
            sources(slot) = escaped
        End If
        letters = decoded(slot)
    Else
        letters = VfText(escaped)
    End If
    For i = 1 To Len(letters)
        text = Replace(text, Mid$(letters, i, 1), plain)
    Next i
    VfFoldLetters = text
End Function

Private Function VfIsSeparator(ByVal code As Long) As Boolean
    If code <= 32 Or code = 127 Or code = &HA0 Or code = &HAD Then
        VfIsSeparator = True
    ElseIf code < 128 Then
        VfIsSeparator = Not ((code >= 48 And code <= 57) Or (code >= 97 And code <= 122))
    ElseIf (code >= &H2000 And code <= &H206F) Or (code >= &H2E00 And code <= &H2E7F) Then
        VfIsSeparator = True
    ElseIf code = &H60C Or code = &H61B Or code = &H61F Or code = &H640 Then
        VfIsSeparator = True
    ElseIf code = &HB4 Or code = &H2BB Or code = &H2BC Or code = &H2BE Or code = &H2BF Then
        VfIsSeparator = True
    ElseIf code = &HFEFF& Or code = &H3000 Or code = &HFF0C& Or code = &HFF1B& Then
        VfIsSeparator = True
    End If
End Function

Private Sub VfBuildIndex()
    Dim i As Long, j As Long, row As Variant, combined As String
    Set mVfQueryOwner = Nothing
    If mVfRows.Count = 0 Then
        mVfIndexReady = True
        Exit Sub
    End If
    ReDim mVfSearchText(1 To mVfRows.Count)
    ReDim mVfNameText(1 To mVfRows.Count, 0 To 2)
    ReDim mVfNameKeys(1 To mVfRows.Count, 0 To 2)
    ReDim mVfFoldedKeys(1 To mVfRows.Count, 0 To 2)
    For i = 1 To mVfRows.Count
        row = mVfRows.Item(i)
        For j = 0 To 2
            mVfNameText(i, j) = VefeyatNormalize(CStr(row(j)))
            mVfNameKeys(i, j) = Replace(mVfNameText(i, j), " ", "")
            mVfFoldedKeys(i, j) = VfCollapseLetters(mVfNameKeys(i, j))
        Next j
        combined = CStr(row(0)) & " " & CStr(row(1)) & " " & CStr(row(2))
        combined = combined & " " & CStr(row(3)) & " " & CStr(row(4))
        mVfSearchText(i) = VefeyatNormalize(combined)
    Next i
    mVfIndexReady = True
End Sub

Private Sub VfLoadCache(ByVal path As String)
    Dim loaded As Collection
    Set loaded = VfReadCatalog(path)
    Set mVfRows = loaded
    mVfIndexReady = False
End Sub

Private Function VfImportCatalog(ByVal showSuccess As Boolean) As Boolean
    Dim dialog As Object, fso As Object, selectedPath As String
    Dim currentPath As String, folderPath As String, stagePath As String, backupPath As String
    Dim parsed As Collection, stagedRows As Collection, previousRows As Collection, summary As String
    Dim existed As Boolean, removedCurrent As Boolean, committed As Boolean
    Dim errorText As String, errorNumber As Long, restored As Boolean
    Set dialog = Application.FileDialog(3)
    With dialog
        .Title = VfText("Vefey\u00e2t: bilgisayar\u0131n\u0131zdaki CSV dosyas\u0131n\u0131 se\u00e7in")
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV", "*.csv"
        If .Show <> -1 Then Exit Function
        selectedPath = CStr(.SelectedItems(1))
    End With
    VfCheckLocalPath selectedPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    currentPath = VefeyatDataPath()
    Set parsed = VfReadCatalog(selectedPath)
    If StrComp(fso.GetAbsolutePathName(selectedPath), currentPath, vbTextCompare) = 0 Then
        Set mVfRows = Nothing
        mVfIndexReady = False
        VfImportCatalog = True
        Exit Function
    End If
    existed = fso.FileExists(currentPath)
    If existed Then
        Set previousRows = VfReadCatalog(currentPath)
        summary = VfqCatalogSummary(previousRows, parsed) & vbCrLf & vbCrLf
        If VfMessage(summary & VfText("Yerel veri kopyas\u0131 se\u00e7ti\u011finiz CSV ile de\u011fi\u015ftirilsin mi?") & _
            vbCrLf & VfText("Mevcut kopya ayn\u0131 klas\u00f6re yedeklenecek. Kaynak CSV de\u011fi\u015fmez."), _
            vbQuestion Or vbYesNo Or vbDefaultButton2, "Vefeyat") <> vbYes Then Exit Function
    End If
    On Error GoTo Failed
    folderPath = fso.GetParentFolderName(currentPath)
    VfEnsureFolder folderPath, fso
    stagePath = VfUniquePath(folderPath, "incoming-", ".tmp", fso)
    fso.CopyFile selectedPath, stagePath, False
    ' Validate the actual staged copy as well, before touching current data.
    Set stagedRows = VfReadCatalog(stagePath)
    If Not VfqCatalogSame(parsed, stagedRows) Then VfFail "CSV changed after its preview; current catalog preserved."
    If existed Then
        backupPath = VfUniquePath(folderPath, "people-backup-", ".csv", fso)
        fso.CopyFile currentPath, backupPath, False
        If Not VfqCatalogSame(previousRows, VfReadCatalog(backupPath)) Then VfFail "Current catalog changed during import; current catalog preserved."
        Kill currentPath
        removedCurrent = True
    End If
    fso.MoveFile stagePath, currentPath
    committed = True
    Set mVfRows = Nothing
    mVfIndexReady = False
    VfImportCatalog = True
    If showSuccess Then
        VfMessage CStr(stagedRows.Count) & VfText(" kay\u0131t al\u0131nd\u0131.") & vbCrLf & currentPath, _
            vbInformation, "Vefeyat"
    End If
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If removedCurrent And Not committed Then
        Err.Clear
        fso.CopyFile backupPath, currentPath, True
        restored = (Err.Number = 0)
    End If
    If Len(stagePath) > 0 Then
        If fso.FileExists(stagePath) Then fso.DeleteFile stagePath, True
    End If
    On Error GoTo 0
    If removedCurrent And Not committed And Not restored Then
        errorText = errorText & vbCrLf & "Restore the retained backup: " & backupPath
    End If
    Err.Raise errorNumber, "VefeyatData", errorText
End Function

Private Sub VfCheckLocalPath(ByVal path As String)
    Dim fso As Object, drive As Object
    ' Reject network/URL paths. Imported source data must be on a local drive.
    If Len(path) < 3 Then VfFail "Choose a CSV file on a local Windows drive."
    If Mid$(path, 2, 2) <> ":\" Then VfFail "Choose a CSV file on a local Windows drive."
    If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase$(Left$(path, 1)), vbBinaryCompare) = 0 Then
        VfFail "Choose a CSV file on a local Windows drive."
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set drive = fso.GetDrive(Left$(path, 2))
    If drive.DriveType = 3 Then VfFail "A network drive cannot hold the local catalog."
End Sub

Private Sub VfEnsureFolder(ByVal path As String, ByVal fso As Object)
    Dim parent As String
    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) = 0 Or parent = path Then VfFail "Cannot create the local data folder."
    If Not fso.FolderExists(parent) Then VfEnsureFolder parent, fso
    MkDir path
End Sub

Private Function VfUniquePath(ByVal folder As String, ByVal prefix As String, ByVal suffix As String, ByVal fso As Object) As String
    Dim candidate As String
    Do
        mVfUniqueCounter = mVfUniqueCounter + 1
        candidate = folder & "\" & prefix & Format$(Now, "yyyymmdd-hhnnss") & "-" & CStr(mVfUniqueCounter) & suffix
    Loop While fso.FileExists(candidate) Or fso.FolderExists(candidate)
    VfUniquePath = candidate
End Function

Private Function VfReadCatalog(ByVal path As String) As Collection
    Dim stream As Object, content As String, errorText As String, errorNumber As Long
    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    content = stream.ReadText(-1)
    stream.Close
    Set stream = Nothing
    If Len(content) > 0 Then
        If Left$(content, 1) = ChrW(&HFEFF) Then content = Mid$(content, 2)
    End If
    Set VfReadCatalog = VfParseCsv(content)
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise errorNumber, "VefeyatData", errorText
End Function

Private Function VfParseCsv(ByVal content As String) As Collection
    Dim rows As New Collection, fields As New Collection
    Dim fieldValue As String, ch As String, state As Long
    Dim i As Long, lineNumber As Long, recordLine As Long
    Dim recordStarted As Boolean, haveHeader As Boolean
    lineNumber = 1
    recordLine = 1
    i = 1
    Do While i <= Len(content)
        ch = Mid$(content, i, 1)
        If state = 1 Then
            If ch = """" Then
                If Mid$(content, i + 1, 1) = """" Then
                    fieldValue = fieldValue & """"
                    i = i + 1
                Else
                    state = 2
                End If
            Else
                fieldValue = fieldValue & ch
                If ch = vbLf Then lineNumber = lineNumber + 1
            End If
        Else
            If state = 2 And ch <> ";" And ch <> vbCr And ch <> vbLf Then
                VfCsvFail lineNumber, "Unexpected text after a closing quote."
            End If
            Select Case ch
                Case """"
                    If Len(fieldValue) <> 0 Then VfCsvFail lineNumber, "Quote inside an unquoted field."
                    state = 1
                    recordStarted = True
                Case ";"
                    fields.Add fieldValue
                    fieldValue = ""
                    state = 0
                    recordStarted = True
                Case vbCr, vbLf
                    If recordStarted Or fields.Count > 0 Then
                        fields.Add fieldValue
                        VfAppendCsvRecord rows, fields, haveHeader, recordLine
                    End If
                    Set fields = New Collection
                    fieldValue = ""
                    state = 0
                    recordStarted = False
                    If ch = vbCr And Mid$(content, i + 1, 1) = vbLf Then i = i + 1
                    lineNumber = lineNumber + 1
                    recordLine = lineNumber
                Case Else
                    fieldValue = fieldValue & ch
                    recordStarted = True
            End Select
        End If
        i = i + 1
    Loop
    If state = 1 Then VfCsvFail recordLine, "Unclosed quoted field."
    If recordStarted Or fields.Count > 0 Then
        fields.Add fieldValue
        VfAppendCsvRecord rows, fields, haveHeader, recordLine
    End If
    If Not haveHeader Then VfFail "The CSV file has no header."
    If rows.Count = 0 Then VfFail "The CSV file contains no people."
    VfValidateIdentities rows
    Set VfParseCsv = rows
End Function


' Validate one complete in-memory import before replacing any cached/file data.
' Conflicting dates for the same legacy identity must never become first-row wins.
' Shared article links alone do not merge distinct people; all name fields count.
Private Sub VfValidateIdentities(ByVal rows As Collection)
    Dim dates As Object, row As Variant, part As Variant, key As String
    Set dates = CreateObject("Scripting.Dictionary")
    For Each row In rows
        key = "Vfp1:"
        For Each part In Array(5, 0, 1, 2)
            key = key & CStr(Len(CStr(row(CLng(part))))) & ":" & CStr(row(CLng(part)))
        Next part
        If dates.Exists(key) Then
            If StrComp(CStr(dates(key)), CStr(row(3)), vbBinaryCompare) <> 0 Then
                VfFail "The catalog has conflicting dates for the same person. Review the CSV; no catalog was loaded."
            End If
        Else
            dates.Add key, CStr(row(3))
        End If
    Next row
End Sub

Private Sub VfAppendCsvRecord(ByVal rows As Collection, ByVal fields As Collection, ByRef haveHeader As Boolean, ByVal lineNumber As Long)
    Dim expected As Variant, row As Variant, i As Long
    If fields.Count <> 6 Then VfCsvFail lineNumber, "Expected exactly six semicolon-separated columns."
    If Not haveHeader Then
        expected = Array(VfText("K\u0131sa Ad"), "Tam Ad", "Orijinal Ad", "Vefat Tarihi", VfText("Tan\u0131t\u0131m"), "Kaynak Link")
        For i = 0 To 5
            If StrComp(CStr(fields.Item(i + 1)), CStr(expected(i)), vbBinaryCompare) <> 0 Then
                VfCsvFail lineNumber, "The six CSV headers do not match the expected catalog format."
            End If
        Next i
        haveHeader = True
        Exit Sub
    End If
    row = Array("", "", "", "", "", "")
    For i = 0 To 5
        row(i) = Trim$(CStr(fields.Item(i + 1)))
        If row(i) = "---" Then row(i) = ""
    Next i
    If Len(CStr(row(0))) = 0 Then VfCsvFail lineNumber, "A short name is required."
    rows.Add row
End Sub

Private Sub VfCsvFail(ByVal lineNumber As Long, ByVal message As String)
    VfFail "CSV line " & CStr(lineNumber) & ": " & message
End Sub

Private Sub VfFail(ByVal message As String)
    Err.Raise vbObjectError + 2141, "VefeyatData", message
End Sub

Private Function VfpFormatDate(ByVal raw As String, ByVal prefix As String) As String
    Dim key As String
    Dim markerStart As Long
    Dim markerLength As Long
    Dim replacement As String

    key = VfpRequirePrefix(prefix)
    VfpFormatDate = raw
    If key = "original" Or Len(raw) = 0 Then Exit Function

    markerStart = VfpSkipSpaces(raw, 1)
    If markerStart > Len(raw) Then Exit Function
    If Mid$(raw, markerStart, 1) = "(" Then
        markerStart = VfpSkipSpaces(raw, markerStart + 1)
    End If
    markerLength = VfpMarkerLength(raw, markerStart)
    If markerLength = 0 Then Exit Function

    Select Case key
        Case "o": replacement = VfText("\u00f6.")
        Case "ol": replacement = VfText("\u00f6l.")
        Case "v": replacement = "v."
    End Select
    VfpFormatDate = Left$(raw, markerStart - 1) & replacement & _
                    Mid$(raw, markerStart + markerLength)
End Function

' Settings belong to this Word document. Reading an absent or old invalid
' value returns a default without writing to the document or the registry.
Private Function VfpGetPrefix(ByVal doc As Word.Document) As String
    Dim key As String
    key = VfpNormalizeKey(VfpReadVariable(doc, VfpPrefixVariable))
    If Not VfpKnownPrefix(key) Then key = "original"
    VfpGetPrefix = key
End Function

Private Sub VfpSetPrefix(ByVal doc As Word.Document, ByVal value As String)
    Dim key As String
    key = VfpRequirePrefix(value)
    VfpWriteVariable doc, VfpPrefixVariable, key
End Sub

Private Function VfpGetScope(ByVal doc As Word.Document) As String
    Dim key As String
    key = VfpNormalizeKey(VfpReadVariable(doc, VfpScopeVariable))
    If Not VfpKnownScope(key) Then key = "first"
    VfpGetScope = key
End Function

Private Function VfpDateAllowed(ByVal doc As Word.Document) As Boolean
    VfpDateAllowed = (VfpGetScope(doc) <> "none")
End Function

Private Sub VfpSetScope(ByVal doc As Word.Document, ByVal value As String)
    Dim key As String
    key = VfpNormalizeKey(value)
    If Not VfpKnownScope(key) Then
        Err.Raise VfpInvalidSetting, "VefeyatPolicy", "Unknown occurrence scope."
    End If
    VfpWriteVariable doc, VfpScopeVariable, key
End Sub

Private Function VfpGetAuditMode(ByVal doc As Word.Document) As String
    Dim key As String
    key = VfpNormalizeKey(VfpReadVariable(doc, VfpAuditVariable))
    If key <> "audit_all" And key <> "audit_undated" Then key = "audit_all"
    VfpGetAuditMode = key
End Function

Private Sub VfpSetAuditMode(ByVal doc As Word.Document, ByVal value As String)
    Dim key As String
    key = VfpNormalizeKey(value)
    If key <> "audit_all" And key <> "audit_undated" Then
        Err.Raise VfpInvalidSetting, "VefeyatPolicy", "Unknown audit filter."
    End If
    VfpWriteVariable doc, VfpAuditVariable, key
End Sub

' A length-prefixed identity avoids delimiter and hash collisions. The CSV
' has no ID column; rows with the same URL and all three names share an ID.
' Dates and descriptions are deliberately excluded so a correction to them
' does not change the identity. This key is not used as a variable name.
Private Function VfpPersonKey(ByRef row As Variant) As String
    VfpPersonKey = "Vfp1:" & VfpKeyPart(CStr(row(5))) & _
                   VfpKeyPart(CStr(row(0))) & VfpKeyPart(CStr(row(1))) & _
                   VfpKeyPart(CStr(row(2)))
End Function

Private Function VfpKeyPart(ByVal value As String) As String
    VfpKeyPart = CStr(Len(value)) & ":" & value
End Function

Private Function VfpRequirePrefix(ByVal value As String) As String
    Dim key As String
    key = VfpNormalizeKey(value)
    If Not VfpKnownPrefix(key) Then
        Err.Raise VfpInvalidSetting, "VefeyatPolicy", "Unknown death-date prefix."
    End If
    VfpRequirePrefix = key
End Function

Private Function VfpNormalizeKey(ByVal value As String) As String
    VfpNormalizeKey = LCase$(Trim$(value))
End Function

Private Function VfpKnownPrefix(ByVal key As String) As Boolean
    Select Case key
        Case "original", "o", "ol", "v": VfpKnownPrefix = True
    End Select
End Function

Private Function VfpKnownScope(ByVal key As String) As Boolean
    Select Case key
        Case "first", "all", "none": VfpKnownScope = True
    End Select
End Function

Private Function VfpReadVariable(ByVal doc As Word.Document, ByVal name As String) As String
    Dim item As Word.Variable
    If doc Is Nothing Then
        Err.Raise VfpInvalidDocument, "VefeyatPolicy", "A Word document is required."
    End If
    ' A missing Word document variable can return a proxy whose Value raises
    ' 5825 (object deleted). Read only items present in the collection.
    For Each item In doc.Variables
        If StrComp(item.Name, name, vbTextCompare) = 0 Then
            VfpReadVariable = item.Value
            Exit Function
        End If
    Next item
End Function

Private Sub VfpWriteVariable(ByVal doc As Word.Document, ByVal name As String, _
                             ByVal value As String)
    Dim item As Word.Variable
    If doc Is Nothing Then
        Err.Raise VfpInvalidDocument, "VefeyatPolicy", "A Word document is required."
    End If
    For Each item In doc.Variables
        If StrComp(item.Name, name, vbTextCompare) = 0 Then
            item.Value = value
            Exit Sub
        End If
    Next item
    doc.Variables.Add Name:=name, Value:=value
End Sub

Private Function VfpSkipSpaces(ByVal value As String, ByVal start As Long) As Long
    Do While start <= Len(value)
        If Not VfpIsSpace(Mid$(value, start, 1)) Then Exit Do
        start = start + 1
    Loop
    VfpSkipSpaces = start
End Function

Private Function VfpIsSpace(ByVal value As String) As Boolean
    If Len(value) <> 1 Then Exit Function
    Select Case AscW(value)
        Case 9, 10, 13, 32, 160, 8199, 8239: VfpIsSpace = True
    End Select
End Function

Private Function VfpMarkerLength(ByVal value As String, ByVal start As Long) As Long
    Dim candidate As String
    Dim length As Long
    If start < 1 Or start > Len(value) Then Exit Function
    candidate = LCase$(Mid$(value, start, 3))
    candidate = Replace(candidate, VfText("\u00d6"), VfText("\u00f6"))
    If Left$(candidate, 3) = VfText("\u00f6l.") Then
        length = 3
    ElseIf Left$(candidate, 2) = VfText("\u00f6.") Or Left$(candidate, 2) = "v." Then
        length = 2
    Else
        Exit Function
    End If
    If VfpMarkerBoundary(value, start + length) Then VfpMarkerLength = length
End Function

Private Function VfpMarkerBoundary(ByVal value As String, ByVal nextIndex As Long) As Boolean
    Dim nextChar As String
    If nextIndex > Len(value) Then Exit Function
    nextChar = Mid$(value, nextIndex, 1)
    If VfpIsSpace(nextChar) Then
        VfpMarkerBoundary = True
    ElseIf nextChar >= "0" And nextChar <= "9" Then
        VfpMarkerBoundary = True
    ElseIf nextChar = "?" Then
        VfpMarkerBoundary = True
    End If
End Function

' Explicitly invoked only. These checks never create or modify a document.
' The return value is the number of successful pure helper assertions.
Private Function VfpRunPolicyChecks() As Long
    Dim count As Long
    Dim raw As String
    Dim rowA As Variant
    Dim rowB As Variant
    Dim formatted As String
    Dim errorNumber As Long

    raw = VfText("  (\u00f6.  680/1282) ")
    VfpCheck VfpFormatDate(raw, "original") = raw, "Original bytes", count
    VfpCheck VfpFormatDate(raw, "ol") = VfText("  (\u00f6l.  680/1282) "), "Spaces", count
    VfpCheck VfpFormatDate(raw, "v") = "  (v.  680/1282) ", "V prefix", count
    VfpCheck VfpFormatDate(VfText("(\u00f6l. 680/1282)"), "o") = _
             VfText("(\u00f6. 680/1282)"), "Long marker", count
    VfpCheck VfpFormatDate("(v. 680/1282)", "o") = _
             VfText("(\u00f6. 680/1282)"), "V marker", count
    VfpCheck VfpFormatDate(VfText("(\u00d6L. 680/1282)"), "v") = _
             "(v. 680/1282)", "Uppercase marker", count
    VfpCheck VfpFormatDate(VfText("\u00f6.680/1282"), "v") = "v.680/1282", _
             "Adjacent year", count
    VfpCheck VfpFormatDate("(v. ?)", "ol") = VfText("(\u00f6l. ?)"), "Uncertainty", count
    VfpCheck VfpFormatDate("(1844-1921)", "ol") = "(1844-1921)", "Lifespan", count
    raw = VfText("(m.\u00f6. 384-322)")
    VfpCheck VfpFormatDate(raw, "v") = raw, "BC expression", count
    raw = VfText("(M.\u00d6. 384-322)")
    VfpCheck VfpFormatDate(raw, "o") = raw, "Uppercase BC", count
    raw = "(gaybeti 260/874)"
    VfpCheck VfpFormatDate(raw, "ol") = raw, "Other event", count
    raw = "(tarih bilinmiyor)"
    VfpCheck VfpFormatDate(raw, "o") = raw, "Unknown expression", count
    raw = VfText("\u00d6zt\u00fcrk")
    VfpCheck VfpFormatDate(raw, "v") = raw, "Name substring", count
    raw = VfText("(\u00f6.zt\u00fcrk)")
    VfpCheck VfpFormatDate(raw, "v") = raw, "Marker boundary", count
    raw = VfText("(not: \u00f6. 680/1282)")
    VfpCheck VfpFormatDate(raw, "v") = raw, "Nonleading marker", count
    raw = VfText("(\u00f6. 680/1282; kaynak: \u00f6. 681)")
    VfpCheck VfpFormatDate(raw, "v") = _
             VfText("(v. 680/1282; kaynak: \u00f6. 681)"), "Only leading marker", count
    raw = VfText("\u00a0(\u00a0\u00f6.\u00a0680/1282)")
    VfpCheck VfpFormatDate(raw, "v") = _
             VfText("\u00a0(\u00a0v.\u00a0680/1282)"), "Unicode spaces", count
    VfpCheck VfpFormatDate("", "o") = "", "Empty expression", count
    VfpCheck VfpFormatDate("v.", "ol") = "v.", "Bare marker", count
    VfpCheck VfpFormatDate("(V. 680/1282)", " OL ") = _
             VfText("(\u00f6l. 680/1282)"), "Normalized setting", count
    On Error Resume Next
    formatted = VfpFormatDate("(v. 680)", "")
    errorNumber = Err.Number
    On Error GoTo 0
    VfpCheck errorNumber = VfpInvalidSetting, "Blank prefix rejected", count
    On Error Resume Next
    formatted = VfpFormatDate("(v. 680)", "invalid")
    errorNumber = Err.Number
    On Error GoTo 0
    VfpCheck errorNumber = VfpInvalidSetting, "Unknown prefix rejected", count
    VfpCheck VfpKnownScope("first") And VfpKnownScope("all"), "Known scopes", count
    VfpCheck Not VfpKnownScope("") And Not VfpKnownScope("selection"), "Unknown scopes", count
    rowA = Array("A:B", "C", "", "old date", "description", "url")
    rowB = Array("A", "B:C", "", "old date", "description", "url")
    VfpCheck VfpPersonKey(rowA) <> VfpPersonKey(rowB), "Identity delimiters", count
    rowB = Array("A:B", "C", "", "new date", "new description", "url")
    VfpCheck VfpPersonKey(rowA) = VfpPersonKey(rowB), "Stable metadata identity", count
    rowB = Array("A:B", "C", "", "old date", "description", "other url")
    VfpCheck VfpPersonKey(rowA) <> VfpPersonKey(rowB), "Source identity", count
    VfpRunPolicyChecks = count
    Debug.Print "Vefeyat policy checks passed: " & CStr(count)
End Function

Private Sub VfpCheck(ByVal passed As Boolean, ByVal description As String, ByRef count As Long)
    If Not passed Then
        Err.Raise VfpCheckFailure, "VefeyatPolicy", "Policy check failed: " & description
    End If
    count = count + 1
End Sub

Private Function VfaCanInsert() As Boolean
    Dim targetDoc As Word.Document
    Dim selectedText As String, aliasText As String
    On Error GoTo Unavailable
    If Application.Documents.Count = 0 Then Exit Function
    Set targetDoc = ActiveDocument
    If targetDoc.ReadOnly Then Exit Function
    If targetDoc.ProtectionType <> wdNoProtection Then Exit Function
    If Selection.Type <> wdSelectionNormal Then Exit Function
    If Not VfpDateAllowed(targetDoc) Then Exit Function
    If Not VfaSupportedRange(targetDoc, Selection.Range) Then Exit Function
    selectedText = Selection.Range.Text
    If Not VfaSelectionAlias(selectedText, aliasText) Then Exit Function
    VfaCanInsert = True
Unavailable:
End Function

Private Sub VfaInsertSelected()
    mVfaSkipReason = ""
    Dim targetDoc As Word.Document
    Dim targetRange As Word.Range
    Dim originalText As String, aliasText As String, reason As String
    Dim originalStart As Long, originalEnd As Long, originalStory As Long
    Dim candidates As Collection
    Dim row As Variant, chosen As Long, bindingId As String
    Dim named As Word.Range, offset As Long, firstOnly As Boolean, forceHere As Boolean, prompted As Boolean
    Dim undoObject As Object, undoStarted As Boolean, changed As Boolean
    Dim inserted As Long, skipped As Long, failure As String, contextual As Boolean, identityProof As Object

    On Error GoTo Failed
    mVfaCommandChanged = False
    If Not VfaCapture(targetDoc, targetRange, originalText, aliasText, _
                      originalStart, originalEnd, originalStory) Then Exit Sub
    If Not VfpDateAllowed(targetDoc) Then Exit Sub
    firstOnly = (VfpGetScope(targetDoc) = "first")
    If firstOnly Then
        Set named = targetRange.Duplicate
        offset = InStr(1, originalText, aliasText, vbBinaryCompare) - 1
        If offset < 0 Then Exit Sub
        named.SetRange originalStart + offset, originalStart + offset + Len(aliasText)
        ' A repeated click at an already dated occurrence needs no lookup,
        ' previous-mention search, memory repair, metadata write or Undo record.
        If VfmDatePresent(targetDoc, named) Then
            VfaWarn VfText("Bu ge\u00e7i\u015fte \u00f6l\u00fcm tarihi zaten var; yeniden eklenmedi.")
            Exit Sub
        End If
    End If
    If Not VefeyatEnsureData() Then Exit Sub
    Set candidates = VefeyatFind(aliasText, False)
    Set candidates = VfaConfirmationCandidates(candidates, aliasText)
    If candidates.Count = 0 Then
        VfaWarn VfText("Se\u00e7ili adla e\u015fle\u015fen bir \u015fah\u0131s bulunamad\u0131. Ad\u0131n yaz\u0131m\u0131n\u0131 ve se\u00e7ili CSV dosyas\u0131n\u0131 kontrol edin.")
        Exit Sub
    End If
    ' A single catalog identity needs no selection dialog.
    If VfqSingleSafe(candidates, aliasText) Then
        chosen = 1
    Else
        chosen = VfePickPerson(candidates, aliasText)
    End If
    If chosen = 0 Then Exit Sub
    If chosen < 1 Or chosen > candidates.Count Then Exit Sub
    row = candidates.Item(chosen)
    contextual = Not VfqSingleSafe(candidates, aliasText)
    If contextual Then
        If Not VdiRowCurrent(row) Then Exit Sub
        Set identityProof = VdiBuild(candidates, aliasText, targetRange)
    End If
    If Not VfaValidDate(CStr(row(3))) Then
        VfaWarn VfText("Bu kay\u0131ttaki \u00f6l\u00fcm tarihi bo\u015f veya tek sat\u0131rl\u0131k tarih olarak kullan\u0131lam\u0131yor. Belgeye eklenmedi.")
        Exit Sub
    End If
    If Not VfaTargetCurrent(targetDoc, targetRange, originalText, originalStart, _
                            originalEnd, originalStory, reason) Then
        VfaWarn reason
        Exit Sub
    End If

    firstOnly = (VfpGetScope(targetDoc) = "first")
    If firstOnly Then
        Set named = targetRange.Duplicate
        offset = InStr(1, originalText, aliasText, vbBinaryCompare) - 1
        If offset < 0 Then Exit Sub
        named.SetRange originalStart + offset, originalStart + offset + Len(aliasText)
        If Not VfmCheckEarlier(targetDoc, row, aliasText, named, forceHere, False, Nothing, prompted) Then Exit Sub
        If prompted Then
            If (VfpGetScope(targetDoc) = "first") <> firstOnly Then
                VfaWarn VfText("Belge ayar\u0131 onay s\u0131ras\u0131nda de\u011fi\u015fti. Tarih eklenmedi; ad\u0131 yeniden se\u00e7in.")
                Exit Sub
            End If
        End If
        If Not VfaTargetCurrent(targetDoc, targetRange, originalText, originalStart, originalEnd, originalStory, reason) Then
            VfaWarn reason
            Exit Sub
        End If
    End If
    If contextual Then
        If Not VdiCurrent(identityProof) Then VfaWarn "Tarih eklenmedi: " & mVdiReason: Exit Sub
    End If
    If Not VfpDateAllowed(targetDoc) Then Exit Sub
    undoStarted = VfaBeginUndo(undoObject)
    If contextual Then
        VdiInsertOne targetDoc, targetRange, aliasText, row, inserted, skipped, changed
    ElseIf firstOnly Then
        VfmInsertFirst targetDoc, targetRange, originalText, aliasText, row, inserted, skipped, changed, forceHere
    Else
    VfaEnsureNamespace targetDoc, changed
    bindingId = VfaGetBinding(targetDoc, row, aliasText, changed)
    ' An explicit new choice for this spelling replaces its previous person.
    VfaRetireOtherBindings targetDoc, bindingId, aliasText, changed
    VfaApplyBinding targetDoc, bindingId, aliasText, CStr(row(3)), _
                    inserted, skipped, changed
    End If
    VfaEndUndo undoObject, undoStarted
    mVfaCommandChanged = changed
    If inserted = 0 And skipped > 0 And Len(mVfaSkipReason) > 0 Then
        VfaWarn mVfaSkipReason
    ElseIf inserted = 0 And skipped > 0 Then
        VfaWarn VfText("Bu yaz\u0131m\u0131n uygulanacak yerinde \u00f6l\u00fcm tarihi zaten var. Elle yaz\u0131lm\u0131\u015f tarihe dokunulmad\u0131.")
    ElseIf inserted = 0 Then
        VfaWarn VfText("Se\u00e7ili yaz\u0131m ana metinde veya dipnotlarda tam s\u00f6zc\u00fck olarak bulunamad\u0131. Ad\u0131n tamam\u0131n\u0131 se\u00e7ip yeniden deneyin.")
    Else
        VftLog targetDoc, VfText("\u00d6l\u00fcm tarihi (") & CStr(inserted) & ")", targetRange, aliasText, aliasText & " " & CStr(row(3))
        Application.StatusBar = CStr(inserted) & VfText(" Vefey\u00e2t tarihi eklendi.")
    End If
    Exit Sub
Failed:
    failure = Err.Description
    mVfaCommandChanged = changed
    VfaEndUndo undoObject, undoStarted
    VfaReportFailure failure, changed
End Sub

' Called only by an explicit user command or a Ribbon setting change.
' Stored original dates make this operation independent of the CSV file.
Private Sub VfaRefreshDocument()
    Dim targetDoc As Word.Document
    Dim bindings As Collection, bindingId As Variant
    Dim aliasText As String, rawDate As String
    Dim undoObject As Object, undoStarted As Boolean, changed As Boolean
    Dim inserted As Long, skipped As Long, failure As String

    On Error GoTo Failed
    If Application.Documents.Count = 0 Then Exit Sub
    Set targetDoc = ActiveDocument
    If Not VfaEditable(targetDoc) Or Not VfpDateAllowed(targetDoc) Then Exit Sub
    Set bindings = VfaActiveBindingIds(targetDoc)
    If bindings.Count = 0 Then Exit Sub
    undoStarted = VfaBeginUndo(undoObject)
    For Each bindingId In bindings
        aliasText = VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "Alias"))
        rawDate = VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "Raw"))
        If Not VfaValidAlias(aliasText) Or Not VfaValidDate(rawDate) Then
            Err.Raise VfaInvalidData, "VefeyatActions", "Stored date binding is incomplete."
        End If
        VfaApplyBinding targetDoc, CStr(bindingId), aliasText, rawDate, _
                        inserted, skipped, changed
    Next bindingId
    VfaEndUndo undoObject, undoStarted
    Application.StatusBar = VfText("Vefey\u00e2t tarihleri belge ayarlar\u0131na g\u00f6re g\u00fcncellendi.")
    Exit Sub
Failed:
    failure = Err.Description
    VfaEndUndo undoObject, undoStarted
    VfaReportFailure failure, changed
End Sub

Private Function VfaCapture(ByRef targetDoc As Word.Document, ByRef targetRange As Word.Range, _
                            ByRef originalText As String, ByRef aliasText As String, _
                            ByRef originalStart As Long, ByRef originalEnd As Long, _
                            ByRef originalStory As Long) As Boolean
    Dim reason As String
    If Application.Documents.Count = 0 Then
        VfaWarn VfText("Bir Word belgesi a\u00e7\u0131p metindeki \u015fah\u0131s ad\u0131n\u0131 se\u00e7in.")
        Exit Function
    End If
    If Selection.Type <> wdSelectionNormal Then
        VfaWarn VfText("Metindeki tek bir \u015fah\u0131s ad\u0131n\u0131 se\u00e7in.")
        Exit Function
    End If
    Set targetDoc = ActiveDocument
    Set targetRange = Selection.Range.Duplicate
    VfaExpandSuffix targetRange
    originalText = targetRange.Text
    originalStart = targetRange.Start
    originalEnd = targetRange.End
    originalStory = targetRange.StoryType
    If Not VfaSupportedRange(targetDoc, targetRange) Then
        VfaWarn VfText("Ana metin veya dipnottaki \u015fah\u0131s ad\u0131n\u0131 se\u00e7in. \u00dcstbilgi, sonnot ve metin kutular\u0131 bu i\u015fleme dahil de\u011fildir.")
        Exit Function
    End If
    If Not VfaSelectionAlias(originalText, aliasText) Then
        VfaWarn VfText("Yaln\u0131zca tek sat\u0131rdaki ad\u0131 se\u00e7in (en fazla 200 karakter). Birden \u00e7ok paragraf\u0131, sekme veya h\u00fccre sonunu se\u00e7ime katmay\u0131n.")
        Exit Function
    End If
    If Not VfaTargetCurrent(targetDoc, targetRange, originalText, originalStart, _
                            originalEnd, originalStory, reason) Then
        VfaWarn reason
        Exit Function
    End If
    VfaCapture = True
End Function

Private Sub VfaExpandSuffix(ByVal named As Word.Range)
    Dim following As Word.Range, ending As String
    Set following = named.Duplicate
    following.Collapse wdCollapseEnd
    following.MoveEnd wdCharacter, 32
    ending = VfFollowingSuffix(following.Text)
    If Len(ending) = 0 Then Exit Sub
    named.SetRange named.Start, named.End + Len(ending)
    named.Select
End Sub

Private Function VfaTargetCurrent(ByVal targetDoc As Word.Document, ByVal targetRange As Word.Range, _
                                  ByVal originalText As String, ByVal originalStart As Long, _
                                  ByVal originalEnd As Long, ByVal originalStory As Long, _
                                  ByRef reason As String) As Boolean
    Dim currentDoc As Word.Document, rangeDoc As Word.Document
    Dim currentRange As Word.Range, aliasText As String
    reason = VfText("Belge veya se\u00e7ili ad de\u011fi\u015fti. Tarih eklenmedi; ad\u0131 yeniden se\u00e7in.")
    If Application.Documents.Count = 0 Then Exit Function
    Set currentDoc = ActiveDocument
    If Not currentDoc Is targetDoc Then Exit Function
    Set rangeDoc = targetRange.Document
    If Not rangeDoc Is targetDoc Then Exit Function
    If Selection.Type <> wdSelectionNormal Then Exit Function
    Set currentRange = Selection.Range.Duplicate
    If Not currentRange.IsEqual(targetRange) Then Exit Function
    If targetRange.Start <> originalStart Or targetRange.End <> originalEnd Then Exit Function
    If Not VfaSupportedRange(targetDoc, targetRange) Then Exit Function
    If targetRange.StoryType <> originalStory Or currentRange.StoryType <> originalStory Then Exit Function
    If StrComp(targetRange.Text, originalText, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(currentRange.Text, originalText, vbBinaryCompare) <> 0 Then Exit Function
    If Not VfaSelectionAlias(originalText, aliasText) Then Exit Function
    If targetDoc.ReadOnly Then
        reason = VfText("Bu belge salt okunur. Tarih eklenmedi.")
        Exit Function
    End If
    If targetDoc.ProtectionType <> wdNoProtection Then
        reason = VfText("Bu belge d\u00fczenlemeye kar\u015f\u0131 korumal\u0131. Tarih eklenmedi.")
        Exit Function
    End If
    If VfpGetScope(targetDoc) = "first" Then
        If Not VfmTargetExcluded(targetDoc, targetRange, aliasText) Then
            VfaTargetCurrent = True
            Exit Function
        End If
    ElseIf Not VfaOverlapsExcluded(targetRange, VfaExcludedRanges(targetDoc, aliasText)) Then
        VfaTargetCurrent = True
        Exit Function
    End If
    reason = VfText("Se\u00e7im bir Word alan\u0131, i\u00e7erik denetimi veya Vefey\u00e2t'\u0131n tarih ekledi\u011fi ba\u015fka bir adla \u00e7ak\u0131\u015f\u0131yor. Zotero at\u0131flar\u0131 ve y\u00f6netilen adlar korunur; ana metin veya dipnottaki ad\u0131n tamam\u0131n\u0131 se\u00e7in.")
End Function

Private Function VfaEditable(ByVal targetDoc As Word.Document) As Boolean
    If targetDoc.ReadOnly Then
        VfaWarn VfText("Bu belge salt okunur. Tarihler g\u00fcncellenmedi.")
        Exit Function
    End If
    If targetDoc.ProtectionType <> wdNoProtection Then
        VfaWarn VfText("Bu belge d\u00fczenlemeye kar\u015f\u0131 korumal\u0131. Tarihler g\u00fcncellenmedi.")
        Exit Function
    End If
    VfaEditable = True
End Function

Private Sub VfaNameCandidates(ByVal candidates As Collection, ByVal query As String, _
                              ByRef exact As Collection, ByRef partial As Collection)
    Dim row As Variant, personKey As String
    Dim i As Long, score As Long, isExact As Boolean, isPartial As Boolean
    Dim exactKeys As New Collection, partialKeys As New Collection
    Set exact = New Collection
    Set partial = New Collection
    For Each row In candidates
        isExact = False
        isPartial = False
        For i = 0 To 2
            score = VefeyatNameMatch(CStr(row(i)), query)
            If score >= 3 Then isExact = True
            If score > 0 Then isPartial = True
        Next i
        personKey = VfpPersonKey(row)
        If isExact Then
            If Not VfaCollectionContains(exactKeys, personKey) Then
                exact.Add row
                exactKeys.Add personKey
            End If
        ElseIf isPartial Then
            If Not VfaCollectionContains(partialKeys, personKey) Then
                partial.Add row
                partialKeys.Add personKey
            End If
        End If
    Next row
End Sub

Private Function VfaConfirmationCandidates(ByVal candidates As Collection, ByVal query As String) As Collection
    Dim result As New Collection, keys As New Collection
    Dim row As Variant, key As String, i As Long, score As Long
    For Each row In candidates
        score = 0
        For i = 0 To 2
            If VefeyatNameMatch(CStr(row(i)), query) > 0 Then score = 1
        Next i
        key = VfpPersonKey(row)
        If score > 0 And Not VfaCollectionContains(keys, key) Then
            result.Add row
            keys.Add key
        End If
    Next row
    Set VfaConfirmationCandidates = result
End Function

Private Function VfaCollectionContains(ByVal values As Collection, ByVal value As String) As Boolean
    Dim item As Variant
    For Each item In values
        If StrComp(CStr(item), value, vbBinaryCompare) = 0 Then
            VfaCollectionContains = True
            Exit Function
        End If
    Next item
End Function

' Pure lookup checks: no Word document, file, catalog, or user interface access.
Private Function VfaRunLookupChecks() As Long
    Dim candidates As Collection, exact As Collection, partial As Collection, confirmed As Collection
    Dim row As Variant, count As Long
    Dim aliasText As String, literalText As String
    Set candidates = New Collection
    candidates.Add Array("Ali", "Ali b. Hasan", "", "(o. 10)", "", "one")
    candidates.Add Array("Alim", "Alim b. Hasan", "", "(o. 20)", "", "two")
    candidates.Add Array("Zeyd", "Zeyd b. Hasan", "", "(o. 30)", "Ali", "three")
    VfaNameCandidates candidates, "Ali", exact, partial
    VfaLookupAssert exact.Count = 1, "Unique exact name", count
    row = exact.Item(1)
    VfaLookupAssert CStr(row(0)) = "Ali", "Exact identity", count
    VfaLookupAssert partial.Count = 1, "Description-only candidate excluded", count
    Set confirmed = VfaConfirmationCandidates(candidates, "Ali")
    VfaLookupAssert confirmed.Count = 2, "Confirmation retains exact and partial alternatives", count
    VfaLookupAssert VefeyatNameMatch("Abaka", "Abaka") = 4, "Exact name rank", count
    VfaLookupAssert VefeyatNameMatch(VfText("K\u0131v\u00e2m\u00fc\u2019s-s\u00fcnne"), _
                                    VfText("K\u0131v\u00e2m\u00fcss\u00fcnne")) = 3, "Joined alias rank", count
    VfaLookupAssert VefeyatNameMatch(VfText("Ebu'l-Kasim K\u0131v\u00e2m\u00fc\u2019s-s\u00fcnne Ismail"), _
                                    VfText("K\u0131v\u00e2m\u00fcss\u00fcnne")) > 0, "Joined alias within full name", count

    Set candidates = New Collection
    candidates.Add Array("Ali", "Ali b. Hasan", "", "(o. 10)", "", "one")
    candidates.Add Array("Ali", "Ali b. Huseyin", "", "(o. 20)", "", "two")
    candidates.Add Array("Ali", "Ali b. Hasan", "", "(o. 11)", "Updated date", "one")
    VfaNameCandidates candidates, "Ali", exact, partial
    VfaLookupAssert exact.Count = 2, "Distinct homonyms remain ambiguous", count
    VfaLookupAssert partial.Count = 0, "Duplicate identity removed", count

    Set candidates = New Collection
    candidates.Add Array(VfText("Gazz\u00e2l\u00ee"), "Abu Hamid", _
                         VfText("\u0627\u0644\u063a\u0632\u0627\u0644\u064a"), "(o. 505)", "", "arabic")
    VfaNameCandidates candidates, "Gazzali", exact, partial
    VfaLookupAssert exact.Count = 1, "Accent-folded name is exact", count
    VfaNameCandidates candidates, VfText("\u0627\u0644\u063a\u0632\u0627\u0644\u064a"), exact, partial
    VfaLookupAssert exact.Count = 1, "Original Arabic name is exact", count

    Set candidates = New Collection
    candidates.Add Array("Ali b. Hasan", "Ali b. Hasan", "", "(o. 10)", "", "one")
    VfaNameCandidates candidates, "Ali", exact, partial
    VfaLookupAssert exact.Count = 0 And partial.Count = 1, "Substring requires a chooser", count
    VfaLookupAssert VfaBindingFromTag("Vefeyat.Date.D0123456789ABCDEF.P1", _
                                    "D0123456789ABCDEF") = "P1", "Local namespace", count
    VfaLookupAssert Len(VfaBindingFromTag("Vefeyat.Date.D0123456789ABCDEF.P1", _
                                        "D1123456789ABCDEF")) = 0, "Foreign namespace rejected", count
    VfaLookupAssert Len(VfaBindingFromTag("Vefeyat.Date.P1", _
                                        "D0123456789ABCDEF")) = 0, "Unnamespaced tag rejected", count
    VfaLookupAssert VfaBindingFromFieldCode(" ADDIN Vefeyat.Date.D0123456789ABCDEF.P1 ", _
                                          "D0123456789ABCDEF") = "P1", "Owned ADDIN field", count
    VfaLookupAssert Len(VfaBindingFromFieldCode(" ADDIN ZOTERO_ITEM CSL_CITATION ", _
                                              "D0123456789ABCDEF")) = 0, "Other add-in field excluded", count
    VfaLookupAssert Len(VfaBindingFromFieldCode(" DOCVARIABLE Vefeyat.Date.D0123456789ABCDEF.P1 ", _
                                              "D0123456789ABCDEF")) = 0, "Other field type excluded", count
    VfaLookupAssert Len(VfaBindingFromFieldCode(" ADDIN Vefeyat.Date.D0123456789ABCDEF.P1 extra ", _
                                              "D0123456789ABCDEF")) = 0, "Extra field code rejected", count
    VfaLookupAssert VfaOrderBefore(10, 0, 0, 20, 1, 5), "Main before later footnote", count
    VfaLookupAssert VfaOrderBefore(20, 1, 5, 21, 0, 0), "Footnote before following main", count
    VfaLookupAssert VfaOrderBefore(20, 0, 0, 20, 1, 1), "Main reference precedes footnote", count
    VfaLookupAssert VfaOrderBefore(20, 1, 3, 20, 1, 9), "Within-footnote order", count
    VfaLookupAssert Not VfaOrderBefore(20, 1, 9, 20, 1, 3), "Reverse note order rejected", count
    VfaLookupAssert Not VfaOrderBefore(20, 1, 3, 20, 1, 3), "Equal order is stable", count
    VfaLookupAssert VfaSelectionAlias("Abaka" & vbCr, aliasText), "Trailing paragraph mark accepted", count
    VfaLookupAssert aliasText = "Abaka", "Paragraph mark excluded from alias", count
    literalText = "Ali" & ChrW$(160) & "  Hasan"
    VfaLookupAssert VfaSelectionAlias(ChrW$(8239) & literalText & ChrW$(8194) & vbCr & ChrW$(12288), _
                                    aliasText), "Unicode edge spaces with paragraph accepted", count
    VfaLookupAssert aliasText = literalText, "Internal Unicode and repeated spaces retained", count
    VfaLookupAssert VfaSelectionAlias(ChrW$(160) & "Abaka" & ChrW$(8239), _
                                    aliasText), "Unicode edge spaces without paragraph accepted", count
    VfaLookupAssert aliasText = "Abaka", "Unicode edge spaces trimmed", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & vbCr & "Hasan", aliasText), "Interior paragraph rejected", count
    VfaLookupAssert Len(aliasText) = 0, "Rejected selection returns no alias", count
    VfaLookupAssert Not VfaSelectionAlias(vbCr & "Ali", aliasText), "Leading paragraph rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & vbCr & vbCr, aliasText), "Multiple paragraphs rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & vbCr & " " & vbCr, aliasText), "Spaced paragraphs rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & vbTab, aliasText), "Trailing tab rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & vbLf, aliasText), "Line feed rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & vbCr & vbLf, aliasText), "CRLF rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & vbCr & Chr$(7), aliasText), "Word cell marker rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & Chr$(11), aliasText), "Manual line break rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & Chr$(27), aliasText), "Other C0 control rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & ChrW$(128), aliasText), "C1 control rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & ChrW$(133), aliasText), "Unicode next-line control rejected", count
    VfaLookupAssert Not VfaSelectionAlias("Ali" & ChrW$(8232), aliasText), "Unicode line separator rejected", count
    VfaLookupAssert Not VfaSelectionAlias(" " & vbCr & ChrW$(160), aliasText), "Empty paragraph rejected", count
    VfaLookupAssert VfaSelectionAlias(String$(VfaMaxName, "A") & " " & vbCr, _
                                    aliasText), "Literal maximum accepted despite selection extras", count
    VfaLookupAssert Len(aliasText) = VfaMaxName, "Literal maximum retained", count
    VfaLookupAssert Not VfaSelectionAlias(String$(VfaMaxName + 1, "A") & vbCr, _
                                        aliasText), "Overlong literal rejected", count
    VfaRunLookupChecks = count
End Function

Private Sub VfaLookupAssert(ByVal condition As Boolean, ByVal name As String, ByRef count As Long)
    If Not condition Then Err.Raise VfaCheckFailure, "VefeyatActions", name
    count = count + 1
End Sub

Private Function VfaGetBinding(ByVal targetDoc As Word.Document, ByRef row As Variant, _
                               ByVal aliasText As String, ByRef changed As Boolean) As String
    Dim bindings As Collection, bindingId As Variant
    Dim personKey As String, storedAlias As String, numberText As String
    Dim nextNumber As Long
    personKey = VfpPersonKey(row)
    Set bindings = VfaBindingIds(targetDoc)
    For Each bindingId In bindings
        storedAlias = VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "Alias"))
        If VfaAliasKey(storedAlias) = VfaAliasKey(aliasText) Then
            If VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "Key")) = personKey Then
                If VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "OccurrenceOnly")) = "1" Then GoTo NextBinding
                changed = True
                VfaWriteVariable targetDoc, VfaVariableName(CStr(bindingId), "Raw"), CStr(row(3))
                VfaGetBinding = CStr(bindingId)
                Exit Function
            End If
        End If
NextBinding:
    Next bindingId
    numberText = VfaReadVariable(targetDoc, VfaNextVariable)
    nextNumber = 1
    If Len(numberText) > 0 Then
        If IsNumeric(numberText) Then
            If CDbl(numberText) >= 1 And CDbl(numberText) < 2000000000# Then nextNumber = CLng(numberText)
        End If
    End If
    Do
        bindingId = "P" & CStr(nextNumber)
        nextNumber = nextNumber + 1
    Loop While VfaVariableExists(targetDoc, VfaVariableName(CStr(bindingId), "Key"))
    changed = True
    VfaWriteVariable targetDoc, VfaNextVariable, CStr(nextNumber)
    VfaWriteVariable targetDoc, VfaVariableName(CStr(bindingId), "Alias"), aliasText
    VfaWriteVariable targetDoc, VfaVariableName(CStr(bindingId), "Raw"), CStr(row(3))
    ' Key is the commit marker. Incomplete metadata is never scanned.
    VfaWriteVariable targetDoc, VfaVariableName(CStr(bindingId), "Key"), personKey
    VfaGetBinding = CStr(bindingId)
End Function

Private Sub VfaRetireOtherBindings(ByVal targetDoc As Word.Document, ByVal keepId As String, _
                                   ByVal aliasText As String, ByRef changed As Boolean)
    Dim bindings As Collection, bindingId As Variant, storedAlias As String
    Set bindings = VfaBindingIds(targetDoc)
    For Each bindingId In bindings
        If CStr(bindingId) <> keepId Then
            storedAlias = VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "Alias"))
            If VfaAliasKey(storedAlias) = VfaAliasKey(aliasText) And _
               VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "OccurrenceOnly")) <> "1" Then
                VfaDeleteManaged targetDoc, CStr(bindingId), changed
                VfaDeleteManualBookmark targetDoc, CStr(bindingId), changed
                changed = True
                targetDoc.Variables.Item(VfaVariableName(CStr(bindingId), "Key")).Delete
                VfaDeleteVariable targetDoc, VfaVariableName(CStr(bindingId), "Alias")
                VfaDeleteVariable targetDoc, VfaVariableName(CStr(bindingId), "Raw")
                VfaDeleteVariable targetDoc, VfaVariableName(CStr(bindingId), "Pinned")
            End If
        End If
    Next bindingId
End Sub

Private Sub VfaApplyBinding(ByVal targetDoc As Word.Document, ByVal bindingId As String, _
                            ByVal aliasText As String, ByVal rawDate As String, _
                            ByRef inserted As Long, ByRef skipped As Long, ByRef changed As Boolean)
    Dim occurrences As Collection, occurrence As Word.Range
    Dim i As Long, last As Long, dateText As String
    Dim previousInserted As Long, previousSkipped As Long
    If Not VfpDateAllowed(targetDoc) Then Exit Sub
    dateText = VfpFormatDate(rawDate, VfpGetPrefix(targetDoc))
    If VfaReadVariable(targetDoc, VfaVariableName(bindingId, "OccurrenceOnly")) = "1" Then
        VfaRefreshPinned targetDoc, bindingId, dateText, changed
        Exit Sub
    End If
    If VfaReadVariable(targetDoc, VfaVariableName(bindingId, "Pinned")) = "1" Then
        If VfpGetScope(targetDoc) = "first" Then
            VfaRefreshPinned targetDoc, bindingId, dateText, changed
            Exit Sub
        End If
        changed = True
        VfaDeleteVariable targetDoc, VfaVariableName(bindingId, "Pinned")
    End If
    previousInserted = inserted
    previousSkipped = skipped
    VfaDeleteManualBookmark targetDoc, bindingId, changed
    VfaDeleteManaged targetDoc, bindingId, changed
    Set occurrences = VfaFindOccurrences(targetDoc, aliasText)
    last = occurrences.Count
    If VfpGetScope(targetDoc) = "first" And last > 1 Then last = 1
    For i = last To 1 Step -1
        Set occurrence = occurrences.Item(i)
        If VfmDatePresent(targetDoc, occurrence) Or VfaAlreadyHasDate(occurrence, rawDate) Then
            skipped = skipped + 1
        Else
            VfaInsertManaged targetDoc, occurrence, dateText, bindingId, changed
            inserted = inserted + 1
        End If
    Next i
    If inserted = previousInserted And skipped > previousSkipped Then
        ' A manual first date can leave no owned fields after all -> first.
        ' Keep only a name bookmark for that explicit result, never own its date.
        Set occurrence = occurrences.Item(1)
        changed = True
        targetDoc.Bookmarks.Add Name:=VfaBookmarkName(targetDoc, bindingId), Range:=occurrence
    End If
End Sub

' Explicit first-occurrence exceptions keep their positions on a date refresh.
' Switching to every-occurrence mode intentionally reapplies the scope rule.
Private Sub VfaRefreshPinned(ByVal doc As Word.Document, ByVal bindingId As String, ByVal dateText As String, ByRef changed As Boolean)
    Dim region As Word.Range, field As Word.Field
    Dim number As Long, detail As String
    On Error GoTo Failed
    For Each region In VfaRegions(doc)
        For Each field In region.Fields
            If VfaFieldBinding(doc, field) = bindingId Then
                changed = True
                field.Locked = False
                field.Result.Text = " " & dateText
                field.Locked = True
                field.ShowCodes = False
            End If
        Next field
    Next region
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not field Is Nothing Then field.Locked = True
    On Error GoTo 0
    Err.Raise number, "Vefeyat", detail
End Sub

Private Sub VfaDeleteManualBookmark(ByVal targetDoc As Word.Document, ByVal bindingId As String, _
                                    ByRef changed As Boolean)
    Dim name As String
    name = VfaBookmarkName(targetDoc, bindingId)
    If Len(name) = 0 Then Exit Sub
    If targetDoc.Bookmarks.Exists(name) Then
        changed = True
        targetDoc.Bookmarks.Item(name).Delete
    End If
End Sub

Private Function VfaManualBookmarkActive(ByVal targetDoc As Word.Document, ByVal bindingId As String) As Boolean
    Dim target As Word.Range, aliasText As String, rawDate As String, name As String
    name = VfaBookmarkName(targetDoc, bindingId)
    If Len(name) = 0 Then Exit Function
    If Not targetDoc.Bookmarks.Exists(name) Then Exit Function
    Set target = targetDoc.Bookmarks.Item(name).Range
    If Not VfaSupportedRange(targetDoc, target) Then Exit Function
    aliasText = VfaReadVariable(targetDoc, VfaVariableName(bindingId, "Alias"))
    rawDate = VfaReadVariable(targetDoc, VfaVariableName(bindingId, "Raw"))
    If Not VfaValidAlias(aliasText) Or Not VfaValidDate(rawDate) Then Exit Function
    If StrComp(target.Text, aliasText, vbTextCompare) <> 0 Then Exit Function
    VfaManualBookmarkActive = VfaAlreadyHasDate(target, rawDate)
End Function

Private Sub VfaDeleteManaged(ByVal targetDoc As Word.Document, ByVal bindingId As String, _
                             ByRef changed As Boolean)
    Dim regions As Collection, region As Word.Range, field As Word.Field, i As Long
    Set regions = VfaRegions(targetDoc)
    For Each region In regions
        For i = region.Fields.Count To 1 Step -1
            Set field = region.Fields.Item(i)
            If VfaFieldBinding(targetDoc, field) = bindingId Then
                changed = True
                field.Locked = False
                ' Delete the field and its displayed date; never leave result text.
                field.Delete
            End If
        Next i
    Next region
End Sub

Private Function VfaFindOccurrences(ByVal targetDoc As Word.Document, ByVal aliasText As String) As Collection
    Dim result As New Collection, occurrence As Word.Range
    Dim excluded As Collection, literal As Collection
    Set VfaFindOccurrences = result
    Set excluded = VfaExcludedRanges(targetDoc, aliasText)
    Set literal = VfaLiteralOccurrences(targetDoc, aliasText)
    For Each occurrence In literal
        If Not VfaOverlapsExcluded(occurrence, excluded) Then result.Add occurrence
    Next occurrence
End Function

' Main text is followed at each reference by that footnote's text. Coordinates
' from different stories must never be compared without a main-text anchor.
Private Function VfaLiteralOccurrences(ByVal targetDoc As Word.Document, ByVal aliasText As String) As Collection
    Dim result As New Collection, main As Collection, notes As Collection
    Dim note As Word.Footnote, occurrence As Word.Range, nextMain As Long
    Set main = VfaRegionOccurrences(targetDoc.Content, aliasText)
    nextMain = 1
    For Each note In targetDoc.Footnotes
        Do While nextMain <= main.Count
            Set occurrence = main.Item(nextMain)
            If Not VfaOrderBefore(occurrence.Start, 0, 0, note.Reference.Start, 1, 0) Then Exit Do
            result.Add occurrence
            nextMain = nextMain + 1
        Loop
        Set notes = VfaRegionOccurrences(note.Range, aliasText)
        For Each occurrence In notes
            result.Add occurrence
        Next occurrence
    Next note
    Do While nextMain <= main.Count
        result.Add main.Item(nextMain)
        nextMain = nextMain + 1
    Loop
    Set VfaLiteralOccurrences = result
End Function

' Lexicographic order: main anchor, main/note, local position within a note.
' This helper is pure so the first-occurrence rule has deterministic checks.
Private Function VfaOrderBefore(ByVal anchorA As Long, ByVal partA As Long, ByVal localA As Long, _
                                 ByVal anchorB As Long, ByVal partB As Long, ByVal localB As Long) As Boolean
    If anchorA <> anchorB Then
        VfaOrderBefore = (anchorA < anchorB)
    ElseIf partA <> partB Then
        VfaOrderBefore = (partA < partB)
    Else
        VfaOrderBefore = (localA < localB)
    End If
End Function

' Raw, independent search used by both occurrence lookup and longer-name guards.
' It must never call the exclusion builder or managed occurrence finder.
Private Function VfaRegionOccurrences(ByVal region As Word.Range, ByVal aliasText As String) As Collection
    Dim result As New Collection, scan As Word.Range, nextPosition As Long
    Set VfaRegionOccurrences = result
    Set scan = region.Duplicate
    With scan.Find
        .ClearFormatting
        .Text = Replace(aliasText, "^", "^^")
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .IgnoreSpace = False
        .IgnorePunct = False
        .MatchPrefix = False
        .MatchSuffix = False
    End With
    Do While scan.Find.Execute
        If scan.End <= scan.Start Then Exit Do
        If Not VfaRangeInRegion(scan, region) Then Exit Do
        ' Independently reject any broadening by Word's language/find settings.
        If StrComp(scan.Text, aliasText, vbTextCompare) = 0 Then result.Add scan.Duplicate
        nextPosition = scan.End
        If nextPosition >= region.End Then Exit Do
        scan.SetRange nextPosition, region.End
    Loop
End Function

Private Function VfaRegions(ByVal targetDoc As Word.Document) As Collection
    Dim result As New Collection, note As Word.Footnote
    result.Add targetDoc.Content.Duplicate
    For Each note In targetDoc.Footnotes
        result.Add note.Range.Duplicate
    Next note
    Set VfaRegions = result
End Function

Private Function VfaRangeInRegion(ByVal target As Word.Range, ByVal region As Word.Range) As Boolean
    If Not target.InStory(region) Then Exit Function
    VfaRangeInRegion = target.InRange(region)
End Function

Private Function VfaSupportedRange(ByVal targetDoc As Word.Document, ByVal target As Word.Range) As Boolean
    Dim note As Word.Footnote
    If target.StoryType = wdMainTextStory Then
        VfaSupportedRange = VfaRangeInRegion(target, targetDoc.Content)
    ElseIf target.StoryType = wdFootnotesStory Then
        For Each note In targetDoc.Footnotes
            If VfaRangeInRegion(target, note.Range) Then
                VfaSupportedRange = True
                Exit Function
            End If
        Next note
    End If
End Function

Private Function VfaExcludedRanges(ByVal targetDoc As Word.Document, _
                                   Optional ByVal allowAlias As String = "") As Collection
    Dim result As New Collection, control As Word.ContentControl
    Dim field As Word.Field, protectedRange As Word.Range, region As Word.Range
    Dim storedAlias As String, nameRange As Word.Range, regions As Collection
    Dim bindings As Collection, bindingId As Variant, literal As Collection
    ' Store Range objects, not plain offsets: footnotes use another story.
    ' Never place dates in codes/results or any add-in's content controls.
    Set regions = VfaRegions(targetDoc)
    For Each region In regions
        For Each field In region.Fields
            Set protectedRange = field.Code.Duplicate
            protectedRange.SetRange field.Code.Start - 1, VfdFieldEnd(field)
            result.Add protectedRange
            bindingId = VfaFieldBinding(targetDoc, field)
            If Len(CStr(bindingId)) > 0 Then
                If VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "OccurrenceOnly")) = "1" Then
                    storedAlias = VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "Alias"))
                    If Len(storedAlias) > 0 And field.Code.Start - 1 - Len(storedAlias) >= region.Start Then
                        Set nameRange = field.Code.Duplicate
                        nameRange.SetRange field.Code.Start - 1 - Len(storedAlias), field.Code.Start - 1
                        If StrComp(nameRange.Text, storedAlias, vbTextCompare) = 0 Then result.Add nameRange
                    End If
                End If
            End If
        Next field
        For Each control In region.ContentControls
            result.Add control.Range.Duplicate
        Next control
    Next region
    Set bindings = VfaActiveBindingIds(targetDoc)
    For Each bindingId In bindings
        storedAlias = VfaReadVariable(targetDoc, VfaVariableName(CStr(bindingId), "Alias"))
        If Len(storedAlias) > Len(allowAlias) Then
            If InStr(1, storedAlias, allowAlias, vbTextCompare) > 0 Then
                ' Protect EVERY occurrence of an active longer alias, including
                ' occurrences without dates under the first-occurrence setting.
                Set literal = VfaLiteralOccurrences(targetDoc, storedAlias)
                For Each nameRange In literal
                    result.Add nameRange
                Next nameRange
            End If
        End If
    Next bindingId
    Set VfaExcludedRanges = result
End Function

Private Function VfaOverlapsExcluded(ByVal target As Word.Range, ByVal excluded As Collection) As Boolean
    Dim protectedRange As Word.Range
    For Each protectedRange In excluded
        If target.InStory(protectedRange) Then
            If target.Start < protectedRange.End And target.End > protectedRange.Start Then
                VfaOverlapsExcluded = True
                Exit Function
            End If
        End If
    Next protectedRange
End Function

Private Sub VfaInsertManaged(ByVal targetDoc As Word.Document, ByVal occurrence As Word.Range, _
                             ByVal dateText As String, ByVal bindingId As String, ByRef changed As Boolean, _
                             Optional ByVal moveCaret As Boolean = False)
    Dim insertion As Word.Range, field As Word.Field
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    If Not VfpDateAllowed(targetDoc) Then Err.Raise VfaInvalidData, "Vefeyat", "Review-only mode forbids date writes."
    Set insertion = occurrence.Duplicate
    insertion.Collapse wdCollapseEnd
    changed = True
    ' A directly created ADDIN has no result separator: assigning Result.Text
    ' would write an unmanaged date after it. QUOTE supplies a real result.
    ' The duplicate retains its story, including when the name is in a footnote.
    Set field = insertion.Fields.Add(Range:=insertion, Type:=wdFieldQuote, _
                                    Text:=Chr$(34) & " " & Chr$(34), PreserveFormatting:=False)
    field.Result.Text = " " & dateText
    field.Code.Text = " ADDIN " & VfaFieldTag(targetDoc, bindingId) & " "
    field.Locked = True
    field.ShowCodes = False
    If field.Type <> wdFieldAddin Or StrComp(field.Result.Text, " " & dateText, vbBinaryCompare) <> 0 Then
        Err.Raise VfaInvalidData, "VefeyatActions", "Word could not retain the managed date inside its field."
    End If
    VrgCaptureBinding targetDoc, occurrence, bindingId
    VfvRememberCaret occurrence, field
    If moveCaret Then
        ' Continue typing after the closing field character, in the same story.
        Set insertion = field.Result.Duplicate
        insertion.SetRange field.Result.End + 1, field.Result.End + 1
        insertion.Select
    End If
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    ' Roll back only the field created by this call. Never retry insertion.
    On Error Resume Next
    If Not field Is Nothing Then
        field.Locked = False
        field.Delete
    End If
    On Error GoTo 0
    Err.Raise errorNumber, "VefeyatActions", errorText
End Sub

Private Function VfaAlreadyHasDate(ByVal occurrence As Word.Range, ByVal rawDate As String) As Boolean
    Dim following As Word.Range, afterText As String, dateText As String
    Dim key As Variant, i As Long, character As String
    Set following = occurrence.Duplicate
    following.SetRange occurrence.End, occurrence.End
    Call following.MoveEnd(wdCharacter, Len(rawDate) * 3 + 512)
    afterText = following.Text
    For i = 1 To Len(afterText)
        character = Mid$(afterText, i, 1)
        If character = vbCr Or character = Chr$(7) Or character = Chr$(11) Then
            afterText = Left$(afterText, i - 1)
            Exit For
        End If
    Next i
    afterText = VfaNormalizeWhitespace(afterText)
    For Each key In Array("original", "o", "ol", "v")
        dateText = VfaNormalizeWhitespace(VfpFormatDate(rawDate, CStr(key)))
        If Len(dateText) > 0 And Len(afterText) >= Len(dateText) Then
            If StrComp(Left$(afterText, Len(dateText)), dateText, vbTextCompare) = 0 Then
                If Len(afterText) = Len(dateText) Then
                    VfaAlreadyHasDate = True
                Else
                    VfaAlreadyHasDate = VfaDateBoundary(Mid$(afterText, Len(dateText) + 1, 1))
                End If
                If VfaAlreadyHasDate Then Exit Function
            End If
        End If
    Next key
End Function

Private Function VfaDateBoundary(ByVal character As String) As Boolean
    VfaDateBoundary = (InStr(1, " .,;:!?)]}" & Chr$(34), character, vbBinaryCompare) > 0)
End Function

Private Function VfaBindingIds(ByVal targetDoc As Word.Document) As Collection
    Dim result As New Collection, item As Word.Variable, name As String, bindingId As String
    For Each item In targetDoc.Variables
        name = item.Name
        If Left$(name, Len(VfaVariablePrefix)) = VfaVariablePrefix Then
            If Right$(name, 4) = ".Key" Then
                bindingId = Mid$(name, Len(VfaVariablePrefix) + 1, Len(name) - Len(VfaVariablePrefix) - 4)
                If VfaValidBindingId(bindingId) Then result.Add bindingId
            End If
        End If
    Next item
    Set VfaBindingIds = result
End Function

Private Function VfaActiveBindingIds(ByVal targetDoc As Word.Document) As Collection
    Dim result As New Collection, field As Word.Field, bindingId As String
    Dim bindings As Collection, markerId As Variant, regions As Collection, region As Word.Range
    ' Undoing or deleting the last managed field must deactivate its metadata.
    ' Only an explicit insertion command may reuse an inactive binding.
    Set regions = VfaRegions(targetDoc)
    For Each region In regions
        For Each field In region.Fields
            bindingId = VfaFieldBinding(targetDoc, field)
            If Len(bindingId) > 0 Then
                If Not VfaCollectionContains(result, bindingId) Then result.Add bindingId
            End If
        Next field
    Next region
    Set bindings = VfaBindingIds(targetDoc)
    For Each markerId In bindings
        If Not VfaCollectionContains(result, CStr(markerId)) Then
            If VfaManualBookmarkActive(targetDoc, CStr(markerId)) Then result.Add CStr(markerId)
        End If
    Next markerId
    Set VfaActiveBindingIds = result
End Function

Private Function VfaValidBindingId(ByVal value As String) As Boolean
    Dim i As Long, character As String
    If Len(value) < 2 Or Len(value) > 12 Then Exit Function
    If Left$(value, 1) <> "P" Then Exit Function
    For i = 2 To Len(value)
        character = Mid$(value, i, 1)
        If character < "0" Or character > "9" Then Exit Function
    Next i
    VfaValidBindingId = True
End Function

Private Sub VfaEnsureNamespace(ByVal targetDoc As Word.Document, ByRef changed As Boolean)
    Dim value As String
    value = VfaReadVariable(targetDoc, VfaNamespaceVariable)
    If Len(value) > 0 Then
        If Not VfaValidNamespace(value) Then
            Err.Raise VfaInvalidData, "VefeyatActions", "Stored binding namespace is invalid."
        End If
        Exit Sub
    End If
    If Not mVfaRandomized Then
        Randomize
        mVfaRandomized = True
    End If
    value = "D" & Right$("00000000" & Hex$(CLng(Rnd() * 2147483646#)), 8) & _
                  Right$("00000000" & Hex$(CLng(Rnd() * 2147483646#)), 8)
    changed = True
    VfaWriteVariable targetDoc, VfaNamespaceVariable, value
End Sub

Private Function VfaNamespace(ByVal targetDoc As Word.Document) As String
    Dim value As String
    value = VfaReadVariable(targetDoc, VfaNamespaceVariable)
    If VfaValidNamespace(value) Then VfaNamespace = value
End Function

Private Function VfaValidNamespace(ByVal value As String) As Boolean
    Dim i As Long
    If Len(value) <> 17 Then Exit Function
    If Left$(value, 1) <> "D" Then Exit Function
    For i = 2 To Len(value)
        If InStr(1, "0123456789ABCDEF", Mid$(value, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    VfaValidNamespace = True
End Function

Private Function VfaBindingFromTag(ByVal tag As String, ByVal namespace As String) As String
    Dim prefix As String, bindingId As String
    If Not VfaValidNamespace(namespace) Then Exit Function
    prefix = VfaTagPrefix & namespace & "."
    If Left$(tag, Len(prefix)) <> prefix Then Exit Function
    bindingId = Mid$(tag, Len(prefix) + 1)
    If VfaValidBindingId(bindingId) Then VfaBindingFromTag = bindingId
End Function

Private Function VfaFieldBinding(ByVal targetDoc As Word.Document, ByVal field As Word.Field) As String
    If field.Type <> wdFieldAddin Then Exit Function
    VfaFieldBinding = VfaBindingFromFieldCode(field.Code.Text, VfaNamespace(targetDoc))
End Function

Private Function VfaBindingFromFieldCode(ByVal code As String, ByVal namespace As String) As String
    Dim normalized As String
    normalized = VfaNormalizeWhitespace(code)
    If StrComp(Left$(normalized, 6), "ADDIN ", vbTextCompare) <> 0 Then Exit Function
    VfaBindingFromFieldCode = VfaBindingFromTag(Mid$(normalized, 7), namespace)
End Function

Private Function VfaFieldTag(ByVal targetDoc As Word.Document, ByVal bindingId As String) As String
    Dim namespace As String
    namespace = VfaNamespace(targetDoc)
    If Len(namespace) = 0 Or Not VfaValidBindingId(bindingId) Then
        Err.Raise VfaInvalidData, "VefeyatActions", "A local binding namespace and ID are required."
    End If
    VfaFieldTag = VfaTagPrefix & namespace & "." & bindingId
End Function

Private Function VfaBookmarkName(ByVal targetDoc As Word.Document, ByVal bindingId As String) As String
    Dim namespace As String
    namespace = VfaNamespace(targetDoc)
    If Len(namespace) = 0 Or Not VfaValidBindingId(bindingId) Then Exit Function
    VfaBookmarkName = VfaManualBookmarkPrefix & namespace & "_" & bindingId
End Function

Private Function VfaVariableName(ByVal bindingId As String, ByVal field As String) As String
    VfaVariableName = VfaVariablePrefix & bindingId & "." & field
End Function

Private Function VfaVariableExists(ByVal targetDoc As Word.Document, ByVal name As String) As Boolean
    Dim item As Word.Variable
    For Each item In targetDoc.Variables
        If StrComp(item.Name, name, vbTextCompare) = 0 Then
            VfaVariableExists = True
            Exit Function
        End If
    Next item
End Function

Private Function VfaReadVariable(ByVal targetDoc As Word.Document, ByVal name As String) As String
    Dim item As Word.Variable
    ' Enumerating avoids treating an absent setting as a deleted Word object.
    ' Errors from real, existing objects are deliberately allowed to propagate.
    For Each item In targetDoc.Variables
        If StrComp(item.Name, name, vbTextCompare) = 0 Then
            VfaReadVariable = item.Value
            Exit Function
        End If
    Next item
End Function

Private Sub VfaWriteVariable(ByVal targetDoc As Word.Document, ByVal name As String, ByVal value As String)
    VfmMetadataChanged targetDoc
    If VfaVariableExists(targetDoc, name) Then
        targetDoc.Variables.Item(name).Value = value
    Else
        targetDoc.Variables.Add Name:=name, Value:=value
    End If
End Sub

Private Sub VfaDeleteVariable(ByVal targetDoc As Word.Document, ByVal name As String)
    VfmMetadataChanged targetDoc
    If VfaVariableExists(targetDoc, name) Then targetDoc.Variables.Item(name).Delete
End Sub

Private Function VfaAliasKey(ByVal value As String) As String
    ' Literal spelling is retained: accent folding belongs only to catalog lookup.
    ' Internal spaces are also literal for Word.Find, so do not collapse them.
    VfaAliasKey = LCase$(VfaTrimSpaces(value))
End Function

Private Function VfaValidAlias(ByVal value As String) As Boolean
    If Len(value) = 0 Or Len(value) > VfaMaxName Then Exit Function
    If Not VfaSingleLine(value) Then Exit Function
    VfaValidAlias = (Len(VefeyatNormalize(value)) > 0)
End Function

' Word can include the paragraph mark when a name occupies the whole paragraph.
' Normalize only the lookup spelling; the captured selection remains unchanged.
Private Function VfaSelectionAlias(ByVal originalText As String, ByRef aliasText As String) As Boolean
    Dim candidate As String, i As Long, code As Long, trimmed As String, ch As String
    aliasText = ""
    candidate = VfaTrimSpaces(originalText)
    If Len(candidate) = 0 Then Exit Function
    If Right$(candidate, 1) = vbCr Then
        candidate = VfaTrimSpaces(Left$(candidate, Len(candidate) - 1))
    End If
    If Len(candidate) = 0 Or Len(candidate) > VfaMaxName Then Exit Function
    For i = 1 To Len(candidate)
        code = AscW(Mid$(candidate, i, 1)) And &HFFFF&
        If code >= 127 And code <= 159 Then Exit Function
    Next i
    If Not VfaValidAlias(candidate) Then Exit Function
    Do While Len(candidate) > 1
        ch = Right$(candidate, 1)
        trimmed = Left$(candidate, Len(candidate) - 1)
        If InStr(1, ",;:!?", ch, vbBinaryCompare) > 0 Then
            candidate = trimmed
        ElseIf ch = "." And Len(VfNameQuery(trimmed)) < Len(trimmed) Then
            candidate = trimmed
        Else
            Exit Do
        End If
    Loop
    aliasText = candidate
    VfaSelectionAlias = True
End Function

Private Function VfaValidDate(ByVal value As String) As Boolean
    If Len(VfaTrimSpaces(value)) = 0 Or Len(value) > VfaMaxDate Then Exit Function
    VfaValidDate = VfaSingleLine(value)
End Function

Private Function VfaSingleLine(ByVal value As String) As Boolean
    Dim i As Long, code As Long
    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1)) And &HFFFF&
        If code < 32 Or code = 127 Or code = 133 Or code = 8232 Or code = 8233 Then Exit Function
    Next i
    VfaSingleLine = True
End Function

Private Function VfaTrimSpaces(ByVal value As String) As String
    Dim first As Long, last As Long
    first = 1
    last = Len(value)
    Do While first <= last
        If Not VfaIsSpace(Mid$(value, first, 1)) Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If Not VfaIsSpace(Mid$(value, last, 1)) Then Exit Do
        last = last - 1
    Loop
    If last >= first Then VfaTrimSpaces = Mid$(value, first, last - first + 1)
End Function

Private Function VfaIsSpace(ByVal value As String) As Boolean
    Dim code As Long
    code = AscW(value) And &HFFFF&
    VfaIsSpace = (code = 32 Or code = 160 Or code = 5760 Or _
                 (code >= 8192 And code <= 8202) Or code = 8239 Or code = 8287 Or code = 12288)
End Function

Private Function VfaNormalizeWhitespace(ByVal value As String) As String
    Dim i As Long, character As String, result As String, pending As Boolean
    For i = 1 To Len(value)
        character = Mid$(value, i, 1)
        If VfaIsSpace(character) Or character = vbTab Then
            If Len(result) > 0 Then pending = True
        Else
            If pending Then result = result & " "
            result = result & character
            pending = False
        End If
    Next i
    VfaNormalizeWhitespace = result
End Function

Private Function VfaBeginUndo(ByRef undoObject As Object) As Boolean
    Dim currentLevel As Long
    On Error Resume Next
    Err.Clear
    Set undoObject = CallByName(Application, "UndoRecord", VbGet)
    If Err.Number = 0 Then
        currentLevel = CLng(CallByName(undoObject, "CustomRecordLevel", VbGet))
        If Err.Number = 0 And currentLevel = 0 Then
            Err.Clear
            CallByName undoObject, "StartCustomRecord", VbMethod, VfText("Vefey\u00e2t tarihleri")
            VfaBeginUndo = (Err.Number = 0)
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub VfaEndUndo(ByRef undoObject As Object, ByRef undoStarted As Boolean)
    If undoStarted Then
        On Error Resume Next
        CallByName undoObject, "EndCustomRecord", VbMethod
        Err.Clear
        On Error GoTo 0
    End If
    undoStarted = False
    Set undoObject = Nothing
End Sub

Private Sub VfaReportFailure(ByVal failure As String, ByVal changed As Boolean)
    Dim message As String
    message = VfText("Word tarih i\u015flemini tamamlayamad\u0131. ")
    If changed Then
        message = message & VfText("Belgede k\u0131smi de\u011fi\u015fiklik olabilir. Yeniden denemeden \u00f6nce tarihleri kontrol edin; gerekirse Word'\u00fcn Geri Al komutunu kullan\u0131n. ")
    End If
    VfaWarn message & Left$(failure, 300) & vbCrLf & vbCrLf & "[Vefeyat " & VfrVersion & "]"
End Sub

Private Sub VfaWarn(ByVal message As String)
    VfMessage message, vbInformation, VfText("Vefey\u00e2t")
End Sub

Private Sub VfvCancelCommand()
    Set mVfvCaret = Nothing
    Set mVfvAnchor = Nothing
End Sub

Private Sub VfvBeginCommand()
    VfvCancelCommand
    On Error GoTo Unavailable
    If Documents.Count = 0 Then Exit Sub
    If Selection.Type <> wdSelectionNormal Then Exit Sub
    Set mVfvAnchor = Selection.Range.Duplicate
Unavailable:
End Sub

Private Sub VfvRememberCaret(ByVal occurrence As Word.Range, ByVal field As Word.Field)
    On Error GoTo Unavailable
    If mVfvAnchor Is Nothing Then Exit Sub
    If Not (mVfvAnchor.Document Is occurrence.Document) Then Exit Sub
    If Not mVfvAnchor.InStory(occurrence) Then Exit Sub
    If mVfvAnchor.Start > occurrence.End Or mVfvAnchor.End < occurrence.Start Then Exit Sub
    Set mVfvCaret = field.Result.Duplicate
    mVfvCaret.SetRange field.Result.End + 1, field.Result.End + 1
    Exit Sub
Unavailable:
    Set mVfvCaret = Nothing
End Sub

Private Sub VfvFinishCommand()
    On Error GoTo Clean
    If mVfvCaret Is Nothing Or mVfvAnchor Is Nothing Then GoTo Clean
    If Documents.Count = 0 Then GoTo Clean
    If Selection.Type <> wdSelectionNormal And Selection.Type <> wdSelectionIP Then GoTo Clean
    If Not (ActiveDocument Is mVfvAnchor.Document) Then GoTo Clean
    If Not Selection.Range.InStory(mVfvAnchor) Then GoTo Clean
    ' Do not steal a selection that the author or another add-in moved away.
    If Selection.Range.Start < mVfvAnchor.Start Or Selection.Range.End > mVfvCaret.Start Then GoTo Clean
    mVfvCaret.Select
Clean:
    VfvCancelCommand
End Sub

' Only an explicit Advanced-menu command changes field shading. Store a
' per-window previous setting so disabling the option restores that setting.
Private Sub VfvSetFlat(ByVal flat As Boolean)
    Dim i As Long, found As Long, state As Variant, previous As Long
    If mVfvViews Is Nothing Then Set mVfvViews = New Collection
    For i = 1 To mVfvViews.Count
        state = mVfvViews(i)
        If state(0) Is ActiveWindow Then found = i: Exit For
    Next i
    If flat Then
        If found = 0 And ActiveWindow.View.FieldShading <> wdFieldShadingNever Then
            If mVfvViews.Count >= 32 Then Err.Raise vbObjectError + 2901, "Vefeyat", "Too many open view preferences."
            mVfvViews.Add Array(ActiveWindow, ActiveDocument, ActiveWindow.View.FieldShading)
        End If
        ActiveWindow.View.FieldShading = wdFieldShadingNever
    Else
        previous = wdFieldShadingWhenSelected
        If found > 0 Then
            state = mVfvViews(found)
            previous = CLng(state(2))
            mVfvViews.Remove found
        End If
        ActiveWindow.View.FieldShading = previous
    End If
End Sub

Private Sub VfvForgetDocument(ByVal doc As Word.Document)
    Dim i As Long, state As Variant
    On Error Resume Next
    If Not mVfvAnchor Is Nothing Then
        If mVfvAnchor.Document Is doc Then VfvCancelCommand
    End If
    If mVfvViews Is Nothing Then Exit Sub
    For i = mVfvViews.Count To 1 Step -1
        state = mVfvViews(i)
        If state(1) Is doc Then mVfvViews.Remove i
    Next i
End Sub

Private Function VflMacroPath(ByVal entry As String) As String
    ' Use the versioned Public Sub name accepted by Word's macro dispatcher.
    ' Setup still verifies an acknowledgement in this exact module/host.
    VflMacroPath = entry
End Function

Private Sub VflArm()
    If Not VflEnabled Then Exit Sub
    On Error GoTo Failed
    If StrComp(ThisDocument.Name, VfsTemplateName, vbTextCompare) <> 0 Then Exit Sub
    If mVflSuspended Then Exit Sub
    If mVflPending Then
        ' Another Word add-in may replace the single shared OnTime callback.
        ' Expire our stale bookkeeping when a later UI event asks to re-arm.
        If Now <= DateAdd("s", 3, mVflDueAt) Then Exit Sub
        mVflPending = False
    End If
    mVflDueAt = DateAdd("s", 1, Now)
    Application.OnTime When:=mVflDueAt, Name:=VflMacroPath("Vefeyat0341UiAutoTick")
    mVflPending = True
    Exit Sub
Failed:
    mVflPending = False
    mVflStatus = VfText("Zamanlay\u0131c\u0131 ba\u015flat\u0131lamad\u0131: ") & Err.Description
End Sub

Private Function VflDiagnostic() As String
    Dim result As String, scope As String, count As Long
    If Not VflEnabled Then
        VflDiagnostic = VfText("Yazarken otomatik ekleme: ertelendi. Komutlar kullan\u0131labilir.")
        Exit Function
    End If
    On Error GoTo Failed
    If Documents.Count > 0 Then
        scope = VfText("\u0130lk ge\u00e7ti\u011fi yere")
        If VfpGetScope(ActiveDocument) = "all" Then scope = VfText("Her ge\u00e7ti\u011fi yere")
        count = VfaActiveBindingIds(ActiveDocument).Count
        result = VfText("Bu belgedeki ekleme ayar\u0131: ") & scope & vbCrLf & _
            VfText("Onayl\u0131 ad ba\u011f\u0131: ") & CStr(count) & vbCrLf
    End If
    If mVflTicks = 0 Then
        result = result & VfText("Son otomatik denetim: Hen\u00fcz \u00e7al\u0131\u015fmad\u0131.")
    Else
        result = result & VfText("Son otomatik denetim: ") & Format$(mVflLastTick, "hh:nn:ss") & _
            " (" & CStr(mVflTicks) & ")"
    End If
    result = result & vbCrLf & VfText("Durum: ") & mVflPhase
    If mVflPending Then
        If Now > DateAdd("s", 3, mVflDueAt) Then
            result = result & vbCrLf & VfText("Planlanan denetim gecikti veya iptal edildi.")
        Else
            result = result & vbCrLf & VfText("Sonraki denetim planland\u0131.")
        End If
    End If
    If Len(mVflStatus) > 0 Then result = result & vbCrLf & mVflStatus
    VflDiagnostic = result
    Exit Function
Failed:
    VflDiagnostic = result & vbCrLf & VfText("Tan\u0131 okunamad\u0131: ") & Err.Description
End Function

' State is isolated per open document; changing windows must not discard Undo guards.
Private Function VflState(ByVal doc As Word.Document) As Object
    Dim state As Object, owner As Word.Document, current As Word.Document, i As Long, opened As Boolean
    If mVflStates Is Nothing Then Set mVflStates = New Collection
    For i = mVflStates.Count To 1 Step -1
        Set state = mVflStates(i)
        Set owner = state("Document")
        opened = False
        For Each current In Documents
            If current Is owner Then opened = True
        Next current
        If Not opened Then
            mVflStates.Remove i
        ElseIf owner Is doc Then
            Set VflState = state
        End If
    Next i
    If Not VflState Is Nothing Then Exit Function
    Set state = CreateObject("Scripting.Dictionary")
    Set state("Document") = doc
    state("Text") = vbNullString
    state("Processed") = vbNullString
    state("ProcessedPosition") = -1&
    state("ProcessedStory") = -1&
    state("Position") = -1&
    state("Story") = -1&
    state("Observed") = Now
    state("Initialized") = False
    Set state("Tracked") = New Collection
    Set state("Suppressed") = New Collection
    mVflStates.Add state
    Set VflState = state
End Function

Private Function VflSignature(ByVal doc As Word.Document) As String
    Dim note As Word.Footnote, result As String
    result = doc.Content.Text
    For Each note In doc.Footnotes
        result = result & ChrW$(30) & note.Range.Text
    Next note
    VflSignature = result
End Function

' A new text/caret observation only starts the wait. An unchanged observation
' becomes eligible after a full interval, even when no trailing space is typed.
Private Function VflIdleReady(ByVal signature As String, ByVal position As Long, ByVal story As Long, _
                              ByVal stamp As Date, ByVal state As Object) As Boolean
    If Not CBool(state("Initialized")) Or signature <> CStr(state("Text")) Or _
       position <> CLng(state("Position")) Or story <> CLng(state("Story")) Then
        state("Text") = signature
        state("Position") = position
        state("Story") = story
        state("Observed") = stamp
        state("Initialized") = True
        Exit Function
    End If
    If signature = CStr(state("Processed")) And position = CLng(state("ProcessedPosition")) And _
       story = CLng(state("ProcessedStory")) Then Exit Function
    If DateDiff("s", CDate(state("Observed")), stamp) < 1 Then Exit Function
    VflIdleReady = True
End Function

Public Sub Vefeyat0341UiAutoTick()
    If Not VflEnabled Then Exit Sub
    Dim doc As Word.Document, state As Object, signature As String, inserted As Long
    On Error GoTo Failed
    mVflPending = False
    mVflLastTick = Now
    mVflTicks = mVflTicks + 1
    mVflPhase = VfText("Denetim ba\u015flad\u0131.")
    If StrComp(ThisDocument.Name, VfsTemplateName, vbTextCompare) <> 0 Then Exit Sub
    If mVflBusy Or mVflSuspended Then
        mVflPhase = VfText("Ba\u015fka bir Vefey\u00e2t i\u015flemi bekleniyor.")
        GoTo Again
    End If
    If Documents.Count = 0 Then
        Set mVflStates = Nothing
        mVflPhase = VfText("A\u00e7\u0131k belge yok.")
        GoTo Again
    End If
    Set doc = ActiveDocument
    Set state = VflState(doc)
    If doc.ReadOnly Or doc.ProtectionType <> wdNoProtection Then
        state("Initialized") = False
        mVflPhase = VfText("Belge salt okunur veya korumal\u0131.")
        GoTo Again
    End If
    If VfpGetScope(doc) <> "all" Then
        state("Initialized") = False
        mVflPhase = VfText("\u0130lk ge\u00e7i\u015f modu: otomatik ekleme kapal\u0131.")
        GoTo Again
    End If
    If Selection.Type <> wdSelectionIP Then
        state("Initialized") = False
        mVflPhase = VfText("Metin se\u00e7ili; yaz\u0131 imleci bekleniyor.")
        GoTo Again
    End If
    If Not VfaSupportedRange(doc, Selection.Range) Then
        mVflPhase = VfText("\u0130mle\u00e7 ana metin veya dipnotta de\u011fil.")
        GoTo Again
    End If
    If VfaActiveBindingIds(doc).Count = 0 Then
        mVflPhase = VfText("Bu belgede etkin onayl\u0131 ad yok.")
        GoTo Again
    End If
    signature = VflSignature(doc)
    If Not VflIdleReady(signature, Selection.Range.Start, Selection.Range.StoryType, Now, state) Then
        mVflPhase = VfText("Yaz\u0131n\u0131n durmas\u0131 veya yeni metin bekleniyor.")
        GoTo Again
    End If
    mVflBusy = True
    VflProtectRemoved doc, state
    inserted = VfaAutoFill(doc, state)
    VflTrackDates doc, state
    state("Processed") = VflSignature(doc)
    state("ProcessedPosition") = Selection.Range.Start
    state("ProcessedStory") = Selection.Range.StoryType
    ' Re-observe after our own insertion to keep the next user edit debounced.
    state("Initialized") = False
    mVflBusy = False
    mVflPhase = VfText("Denetim tamamland\u0131; yeni tarih say\u0131s\u0131: ") & CStr(inserted)
    If inserted > 0 Then mVflStatus = CStr(inserted) & VfText(" tarih otomatik eklendi.")
Again:
    VflArm
    Exit Sub
Failed:
    mVflBusy = False
    mVflPhase = VfText("Hata nedeniyle durdu.")
    mVflStatus = VfText("Otomatik ekleme durdu: ") & Err.Description
    Application.StatusBar = mVflStatus
End Sub

' Explicit commands may reset the wait, but must retain per-document exclusions.
Private Sub VflResetWait()
    Dim state As Object
    If mVflStates Is Nothing Then Exit Sub
    For Each state In mVflStates
        state("Initialized") = False
        state("Processed") = vbNullString
    Next state
End Sub

Private Sub VflCaptureCommand()
    Dim state As Object
    If Not VflEnabled Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    Set state = VflState(ActiveDocument)
    VflTrackDates ActiveDocument, state
    VflResetWait
End Sub

Private Function VflAutoReady(ByVal occurrence As Word.Range, ByVal caret As Word.Range) As Boolean
    Dim following As Word.Range, ch As String
    Set following = occurrence.Duplicate
    following.Collapse wdCollapseEnd
    following.MoveEnd wdCharacter, 1
    ch = Left$(following.Text, 1)
    ' Never split Gazali'nin or an apostrophe continuation while typing.
    If ch = "'" Or ch = ChrW(&H2019) Or ch = ChrW(&H2018) Then Exit Function
    VflAutoReady = True
    If occurrence.InStory(caret) Then
        ' The completed name's end is eligible after the idle check above.
        If caret.Start >= occurrence.Start And caret.Start < occurrence.End Then VflAutoReady = False
    End If
End Function

Private Function VflSuppressed(ByVal occurrence As Word.Range, ByVal state As Object) As Boolean
    Dim ranges As Collection, previous As Word.Range
    Set ranges = state("Suppressed")
    For Each previous In ranges
        If previous.InStory(occurrence) Then
            If previous.Start = occurrence.Start And previous.End = occurrence.End Then
                If StrComp(previous.Text, occurrence.Text, vbBinaryCompare) = 0 Then
                    VflSuppressed = True
                    Exit Function
                End If
            End If
        End If
    Next previous
End Function

' Remember the name range, not a Field object that becomes invalid on Undo.
Private Sub VflTrackDates(ByVal doc As Word.Document, ByVal state As Object)
    Dim tracked As New Collection, ids As Collection, id As Variant, occurrence As Word.Range
    Dim aliasText As String, raw As String, record As Collection
    Set ids = VfaActiveBindingIds(doc)
    For Each id In ids
        aliasText = VfaReadVariable(doc, VfaVariableName(CStr(id), "Alias"))
        raw = VfaReadVariable(doc, VfaVariableName(CStr(id), "Raw"))
        If VfaValidAlias(aliasText) And VfaValidDate(raw) Then
            For Each occurrence In VfaFindOccurrences(doc, aliasText)
                If VfaAlreadyHasDate(occurrence, raw) Then
                    Set record = New Collection
                    record.Add occurrence.Duplicate
                    record.Add occurrence.Text
                    record.Add raw
                    tracked.Add record
                End If
            Next occurrence
        End If
    Next id
    Set state("Tracked") = tracked
End Sub

Private Sub VflProtectRemoved(ByVal doc As Word.Document, ByVal state As Object)
    Dim tracked As Collection, suppressed As Collection, record As Collection, name As Word.Range
    Set tracked = state("Tracked")
    Set suppressed = state("Suppressed")
    For Each record In tracked
        Set name = record(1)
        If name.End > name.Start Then
            If StrComp(name.Text, CStr(record(2)), vbBinaryCompare) = 0 Then
                If Not VfaAlreadyHasDate(name, CStr(record(3))) Then
                    If Not VflSuppressed(name, state) Then suppressed.Add name.Duplicate
                End If
            End If
        End If
    Next record
End Sub

Private Function VfaAutoFill(ByVal doc As Word.Document, ByVal state As Object) As Long
    Dim ids As Collection, id As Variant, ranges As Collection, occurrence As Word.Range
    Dim caret As Word.Range, aliasText As String, raw As String, text As String, i As Long, atCaret As Boolean
    Dim undoObject As Object, started As Boolean, changed As Boolean, number As Long, detail As String
    On Error GoTo Failed
    If Not ActiveDocument Is doc Then Exit Function
    If doc.ReadOnly Or doc.ProtectionType <> wdNoProtection Or VfpGetScope(doc) <> "all" Then Exit Function
    If Selection.Type <> wdSelectionIP Then Exit Function
    Set caret = Selection.Range.Duplicate
    If Not VfaSupportedRange(doc, caret) Then Exit Function
    Set ids = VfaActiveBindingIds(doc)
    For Each id In ids
        aliasText = VfaReadVariable(doc, VfaVariableName(CStr(id), "Alias"))
        raw = VfaReadVariable(doc, VfaVariableName(CStr(id), "Raw"))
        If VfaValidAlias(aliasText) And VfaValidDate(raw) Then
            text = VfpFormatDate(raw, VfpGetPrefix(doc))
            Set ranges = VfaFindOccurrences(doc, aliasText)
            For i = ranges.Count To 1 Step -1
                Set occurrence = ranges.Item(i)
                If VflAutoReady(occurrence, caret) And Not VflSuppressed(occurrence, state) Then
                    If Not VfaAlreadyHasDate(occurrence, raw) Then
                        atCaret = False
                        If occurrence.InStory(caret) Then atCaret = (caret.Start = occurrence.End)
                        If Not started Then started = VfaBeginUndo(undoObject)
                        VfaInsertManaged doc, occurrence, text, CStr(id), changed, atCaret
                        VfaAutoFill = VfaAutoFill + 1
                    End If
                End If
            Next i
        End If
    Next id
    VfaEndUndo undoObject, started
    Exit Function
Failed:
    number = Err.Number
    detail = Err.Description
    VfaEndUndo undoObject, started
    Err.Raise number, "VefeyatLive", detail
End Function

Private Sub VfnCheckSelected()
    Dim doc As Word.Document, target As Word.Range, named As Word.Range, adjacent As Word.Field
    Dim original As String, aliasText As String, reason As String, proposal As String, preview As String
    Dim start As Long, finish As Long, story As Long, offset As Long, choice As Long
    Dim candidates As Collection, row As Variant, bindingId As String, rawDate As String
    Dim undoObject As Object, started As Boolean, changed As Boolean, failure As String
    Dim wasPinned As Boolean, wasLocal As Boolean, identityProof As Object, registrySlot As Long
    On Error GoTo Failed
    If Not VfaCapture(doc, target, original, aliasText, start, finish, story) Then Exit Sub
    If Not VfpDateAllowed(doc) Then Exit Sub
    If Not VefeyatEnsureData() Then Exit Sub
    Set candidates = VfaConfirmationCandidates(VefeyatFind(aliasText, False), aliasText)
    If candidates.Count = 0 Then
        VfaWarn VfText("Bu yaz\u0131m i\u00e7in kaynakta aday bulunamad\u0131. Metin de\u011fi\u015ftirilmedi.")
        Exit Sub
    End If
    If candidates.Count = 1 Then
        choice = 1
    Else
        choice = VfePickPerson(candidates, aliasText, True)
    End If
    If choice < 1 Or choice > candidates.Count Then Exit Sub
    row = candidates(choice)
    If candidates.Count > 1 Then
        Set identityProof = VdiBuild(candidates, aliasText, target)
    End If
    proposal = VfnChooseName(row, VfNameQuery(aliasText))
    If Len(proposal) > 0 Then proposal = proposal & Mid$(aliasText, Len(VfNameQuery(aliasText)) + 1)
    If Len(proposal) = 0 Then Exit Sub
    If Not VfaValidAlias(proposal) Then
        VfaWarn VfText("Se\u00e7ili ad\u0131n bi\u00e7imini koruyan yak\u0131n bir kaynak yaz\u0131m\u0131 bulunamad\u0131. Metin de\u011fi\u015ftirilmedi.")
        Exit Sub
    End If
    If StrComp(aliasText, proposal, vbBinaryCompare) = 0 Then
        If candidates.Count > 1 Then
            If Not VdiCurrent(identityProof) Then VfaWarn VfText("\u0130\u015flem yap\u0131lmad\u0131: ") & mVdiReason: Exit Sub
            Set named = VdiNamed(target, aliasText)
            VdiCheckStore named, row
            started = VfaBeginUndo(undoObject)
            VdiStore named, row
            VfaEndUndo undoObject, started
        End If
        VfMessage VfText("Se\u00e7ili ad kaynak yaz\u0131m\u0131yla ayn\u0131. De\u011fi\u015fiklik gerekmedi."), vbInformation, VfText("Vefey\u00e2t: isim yaz\u0131m\u0131")
        Exit Sub
    End If
    preview = VfText("Kaynak yaz\u0131m\u0131 \u00f6nerisi") & vbCrLf & vbCrLf & _
        VfText("Belgedeki: ") & aliasText & vbCrLf & VfText("\u00d6nerilen: ") & proposal & vbCrLf & vbCrLf & _
        CStr(row(4)) & vbCrLf & VfText("Kaynak: ") & CStr(row(5)) & vbCrLf & vbCrLf & _
        VfText("Ba\u015fka bir yaz\u0131m\u0131n yanl\u0131\u015f oldu\u011fu anlam\u0131na gelmez. Yaln\u0131z bu se\u00e7ili ad\u0131 de\u011fi\u015ftirelim mi?")
    If VfMessage(preview, vbYesNo Or vbQuestion Or vbDefaultButton2, VfText("Vefey\u00e2t: yaz\u0131m\u0131 onaylay\u0131n")) <> vbYes Then Exit Sub
    If Not VfaTargetCurrent(doc, target, original, start, finish, story, reason) Then
        VfaWarn reason
        Exit Sub
    End If
    offset = InStr(1, original, aliasText, vbBinaryCompare) - 1
    If offset < 0 Then Exit Sub
    Set named = target.Duplicate
    named.SetRange start + offset, start + offset + Len(aliasText)
    Set adjacent = VfnAdjacentField(doc, named)
    If Not adjacent Is Nothing Then
        bindingId = VfaFieldBinding(doc, adjacent)
        rawDate = VfaReadVariable(doc, VfaVariableName(bindingId, "Raw"))
        wasPinned = (VfaReadVariable(doc, VfaVariableName(bindingId, "Pinned")) = "1")
        wasLocal = (VfaReadVariable(doc, VfaVariableName(bindingId, "OccurrenceOnly")) = "1")
        If VfaReadVariable(doc, VfaVariableName(bindingId, "Key")) <> VfpPersonKey(row) Then
            VfaWarn VfText("Bu ad\u0131n tarihi ba\u015fka bir \u015fahsa ba\u011fl\u0131. \u00d6nce \u00d6l\u00fcm tarihi ekle ile \u015fahs\u0131 yeniden do\u011frulay\u0131n.")
            Exit Sub
        End If
        If Not VfnBindingAvailable(doc, proposal, VfpPersonKey(row)) Then
            VfaWarn VfText("\u00d6nerilen yaz\u0131m bu belgede ba\u015fka bir \u015fahsa onaylanm\u0131\u015f. Ad ve tarih de\u011fi\u015ftirilmedi; daha a\u00e7\u0131k bir ad kullan\u0131n.")
            Exit Sub
        End If
    End If
    If candidates.Count > 1 Then
        If Not VdiCurrent(identityProof) Then VfaWarn VfText("\u0130\u015flem yap\u0131lmad\u0131: ") & mVdiReason: Exit Sub
        VdiCheckStore named, row
    End If
    If Not VfpDateAllowed(doc) Then Exit Sub
    started = VfaBeginUndo(undoObject)
    registrySlot = VrgNameSlot(doc, named)
    changed = True
    named.Text = proposal
    If Not adjacent Is Nothing Then
        bindingId = VfaGetBinding(doc, row, proposal, changed)
        ' A name correction must not silently replace an existing source date.
        VfaWriteVariable doc, VfaVariableName(bindingId, "Raw"), rawDate
        If wasPinned Then VfaWriteVariable doc, VfaVariableName(bindingId, "Pinned"), "1"
        If wasLocal Then VfaWriteVariable doc, VfaVariableName(bindingId, "OccurrenceOnly"), "1"
        adjacent.Locked = False
        adjacent.Code.Text = " ADDIN " & VfaFieldTag(doc, bindingId) & " "
        adjacent.Locked = True
        adjacent.ShowCodes = False
    End If
    If candidates.Count > 1 Then VdiStore named, row
    named.SetRange start + offset, start + offset + Len(proposal)
    VrgCapture doc, named, row, "name", registrySlot
    VfaEndUndo undoObject, started
    VftLog doc, VfText("\u0130sim yaz\u0131m\u0131"), named, aliasText, proposal
    VflCaptureCommand
    Application.StatusBar = VfText("Se\u00e7ili ad kaynak yaz\u0131m\u0131yla g\u00fcncellendi. Ctrl+Z ile geri alabilirsiniz.")
    Exit Sub
Failed:
    failure = Err.Description
    On Error Resume Next
    If Not adjacent Is Nothing Then adjacent.Locked = True
    On Error GoTo 0
    VfaEndUndo undoObject, started
    VfaReportFailure failure, changed
End Sub

Private Function VfnBindingAvailable(ByVal doc As Word.Document, ByVal aliasText As String, ByVal key As String) As Boolean
    Dim id As Variant
    VfnBindingAvailable = True
    For Each id In VfaActiveBindingIds(doc)
        If VfaAliasKey(VfaReadVariable(doc, VfaVariableName(CStr(id), "Alias"))) = VfaAliasKey(aliasText) Then
            If VfaReadVariable(doc, VfaVariableName(CStr(id), "Key")) <> key Then
                VfnBindingAvailable = False
                Exit Function
            End If
        End If
    Next id
End Function

Private Function VfnTextName(ByVal value As String) As String
    Dim comma As Long, leftPart As String, rightPart As String, marker As Variant
    value = Trim$(value): VfnTextName = value
    comma = InStr(1, value, ",", vbBinaryCompare)
    If comma < 2 Then Exit Function
    If InStr(comma + 1, value, ",", vbBinaryCompare) > 0 Then Exit Function
    For Each marker In Array("(", ")", ":", ";", "/", vbCr, vbLf)
        If InStr(1, value, CStr(marker), vbBinaryCompare) > 0 Then Exit Function
    Next marker
    leftPart = Trim$(Left$(value, comma - 1)): rightPart = Trim$(Mid$(value, comma + 1))
    If Len(leftPart) > 0 And Len(rightPart) > 0 Then VfnTextName = rightPart & " " & leftPart
End Function

Private Function VfnSourceName(ByRef row As Variant, ByVal query As String) As String
    Dim i As Long, first As Long, last As Long, score As Long, best As Long
    Dim name As String, candidate As String, key As String
    key = Replace(VefeyatNormalize(query), " ", "")
    If Len(key) = 0 Then Exit Function
    ' Compare contiguous source-name spans. Keep a full name full, and recover
    ' embedded aliases such as Kivamussunne without replacing them with Teymi.
    For i = 0 To 2
        name = CStr(row(i))
        If i < 2 Then
            name = VfnTextName(name)
            If name <> CStr(row(i)) And VefeyatNormalize(query) = VefeyatNormalize(CStr(row(i))) Then
                If best < 4 Then best = 4: VfnSourceName = name
            End If
        End If
        If Len(name) <= 2000 Then
            For first = 1 To Len(name)
                If first = 1 Then
                    score = 1
                Else
                    score = Abs(CLng(VfnNameBoundary(Mid$(name, first - 1, 1))))
                End If
                If score > 0 And Not VfnNameBoundary(Mid$(name, first, 1)) Then
                    For last = first To Len(name)
                        If last - first + 1 > VfaMaxName Then Exit For
                        score = 0
                        If last = Len(name) Then
                            score = 1
                        ElseIf VfnNameBoundary(Mid$(name, last + 1, 1)) Then
                            score = 1
                        End If
                        If score > 0 Then
                            candidate = Mid$(name, first, last - first + 1)
                            score = VfnSpellingScore(candidate, query, key)
                            If score > best Then
                                best = score
                                VfnSourceName = candidate
                            End If
                        End If
                    Next last
                End If
            Next first
        End If
    Next i
End Function

Private Function VfnNameBoundary(ByVal ch As String) As Boolean
    VfnNameBoundary = VfaIsSpace(ch) Or ch = "-" Or ch = ChrW(&H2013) Or ch = ChrW(&H2014)
End Function

Private Function VfnSpellingScore(ByVal candidate As String, ByVal query As String, ByVal key As String) As Long
    Dim normalized As String, compact As String, limit As Long
    normalized = VefeyatNormalize(candidate)
    compact = Replace(normalized, " ", "")
    If normalized = VefeyatNormalize(query) Then
        VfnSpellingScore = 4
    ElseIf compact = key Then
        VfnSpellingScore = 3
    ElseIf VfCollapseLetters(compact) = VfCollapseLetters(key) Then
        VfnSpellingScore = 2
    ElseIf Len(key) >= 5 Then
        limit = 1
        If Len(key) >= 9 Then limit = 2
        If VfNearWord(compact, key, limit) Then VfnSpellingScore = 1
    End If
End Function

Private Function VfnHasSuffix(ByVal value As String) As Boolean
    VfnHasSuffix = (Len(VfNameQuery(value)) < Len(value))
End Function

Private Function VfnAdjacentField(ByVal doc As Word.Document, ByVal name As Word.Range) As Word.Field
    Dim region As Word.Range, field As Word.Field
    For Each region In VfaRegions(doc)
        If name.InStory(region) Then
            For Each field In region.Fields
                If field.Code.Start - 1 = name.End Then
                    If Len(VfaFieldBinding(doc, field)) > 0 Then
                        Set VfnAdjacentField = field
                        Exit Function
                    End If
                End If
            Next field
        End If
    Next region
End Function

Private Function VfnNameOptions(ByRef row As Variant, ByVal query As String) As Collection
    Dim result As New Collection, seen As Object, value As String, i As Long, labels As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    labels = Array(VfText("K\u0131sa ad"), "Tam / uzun ad", "Orijinal ad")
    value = VfnSourceName(row, query)
    VfnAddNameOption result, seen, value, VfText("Se\u00e7ili ada yak\u0131n yaz\u0131m"), row
    For i = 0 To 2
        value = Trim$(CStr(row(i)))
        If i < 2 Then value = VfnTextName(value)
        VfnAddNameOption result, seen, value, CStr(labels(i)), row
    Next i
    Set VfnNameOptions = result
End Function

Private Sub VfnAddNameOption(ByVal options As Collection, ByVal seen As Object, _
                             ByVal value As String, ByVal label As String, ByRef row As Variant)
    If value = "---" Or Not VfaValidAlias(value) Then Exit Sub
    If seen.Exists(value) Then Exit Sub
    seen.Add value, True
    options.Add Array(value, label, "", CStr(row(3)), CStr(row(4)), CStr(row(5)))
End Sub

Private Function VfnChooseName(ByRef row As Variant, ByVal query As String) As String
    Dim options As Collection, picker As VefeyatKisiSec0341, index As Long, item As Variant
    Dim number As Long, detail As String
    On Error GoTo Failed
    VfnChooseName = VfuPreferred(ActiveDocument, row)
    If Len(VfnChooseName) = 0 Then VfnChooseName = VfuProposal(row, query, VfuNameMode(ActiveDocument))
    If Len(VfnChooseName) > 0 Then Exit Function
    Set options = VfnNameOptions(row, query)
    If options.Count = 0 Then Exit Function
    Set picker = New VefeyatKisiSec0341
    picker.InitCandidates options, query
    picker.SetNameMode
    picker.Show vbModal
    index = picker.ChosenIndex
    If index > 0 And index <= options.Count Then
        item = options(index)
        VfnChooseName = CStr(item(0))
    End If
    Unload picker
    Exit Function
Failed:
    number = Err.Number
    detail = Err.Description
    On Error Resume Next
    Unload picker
    On Error GoTo 0
    Err.Raise number, "Vefeyat", detail
End Function

Private Function VfdArticleUrl(ByVal value As String) As String
    Dim root As String, path As String, i As Long, ch As String
    value = Trim$(value)
    root = "https://islamansiklopedisi.org.tr/"
    If Left$(LCase$(value), Len("https://www.islamansiklopedisi.org.tr/")) = "https://www.islamansiklopedisi.org.tr/" Then
        value = "https://" & Mid$(value, 13)
    End If
    If Len(value) > 2048 Then Exit Function
    If StrComp(Left$(value, Len(root)), root, vbTextCompare) <> 0 Then Exit Function
    path = Mid$(value, Len(root) + 1)
    If Len(path) = 0 Then Exit Function
    For i = 1 To Len(path)
        ch = LCase$(Mid$(path, i, 1))
        If InStr(1, "abcdefghijklmnopqrstuvwxyz0123456789-_/#!", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    VfdArticleUrl = root & path
End Function

Private Sub VfdOpenSelected()
    Dim doc As Word.Document, query As String, candidates As Collection, row As Variant
    Dim picker As VefeyatKisiSec0341, choice As Long, url As String, detail As String
    On Error GoTo Failed
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    If Not VfaSelectionAlias(Selection.Range.Text, query) Then
        VfaWarn VfText("D\u0130A maddesini a\u00e7mak i\u00e7in \u015fah\u0131s ad\u0131n\u0131 se\u00e7in.")
        Exit Sub
    End If
    If Not VefeyatEnsureData() Then Exit Sub
    Set candidates = VfaConfirmationCandidates(VefeyatFind(query, False), query)
    If candidates.Count = 0 Then
        VfaWarn VfText("Se\u00e7ili adla e\u015fle\u015fen D\u0130A kayd\u0131 bulunamad\u0131.")
        Exit Sub
    End If
    If candidates.Count = 1 Then
        choice = 1
    Else
        Set picker = New VefeyatKisiSec0341
        picker.InitCandidates candidates, query
        picker.SetDiaMode
        picker.Show vbModal
        choice = picker.ChosenIndex
        Unload picker
        Set picker = Nothing
    End If
    If choice < 1 Or choice > candidates.Count Then Exit Sub
    row = candidates(choice)
    url = VfdArticleUrl(CStr(row(5)))
    If Len(url) = 0 Then
        VfaWarn VfText("Bu kay\u0131tta ge\u00e7erli bir D\u0130A madde ba\u011flant\u0131s\u0131 yok.")
        Exit Sub
    End If
    doc.FollowHyperlink Address:=url, NewWindow:=True
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    Unload picker
    On Error GoTo 0
    VfaWarn VfText("D\u0130A sayfas\u0131 a\u00e7\u0131lamad\u0131: ") & detail
End Sub

Private Function VfcBoundary(ByVal text As String, ByVal at As Long) As Boolean
    Dim ch As String, leftPart As String, i As Long, rightPart As String
    ch = Mid$(text, at, 1)
    If ch = vbCr Or ch = vbLf Or ch = Chr$(11) Or ch = ";" Or ch = "!" Or ch = "?" Then
        VfcBoundary = True
    ElseIf ch = "." Then
        i = at - 1
        Do While i > 0 And at - i <= 6
            ch = Mid$(text, i, 1)
            If ch = " " Or ch = vbTab Or ch = "," Or ch = "(" Then Exit Do
            i = i - 1
        Loop
        leftPart = LCase$(Mid$(text, i + 1, at - i - 1))
        Select Case leftPart
            Case "b", "s", "ss", "c", "bk", "kr", "vd", "hz", "dr", "prof", "thk", "nsr", "trc", "ed", "vol", "pp", "p"
                Exit Function
        End Select
        i = at + 1
        Do While i <= Len(text) And i - at <= 8
            rightPart = Mid$(text, i, 1)
            If rightPart <> " " And rightPart <> vbTab Then Exit Do
            i = i + 1
        Loop
        If Len(leftPart) > 0 And Len(rightPart) > 0 Then
            ch = Right$(leftPart, 1)
            If ch >= "0" And ch <= "9" And LCase$(rightPart) = rightPart And UCase$(rightPart) <> rightPart Then Exit Function
        End If
        VfcBoundary = True
    End If
End Function

Private Function VfcWindow(ByVal text As String, ByVal first As Long, ByVal length As Long, ByRef offset As Long) As String
    Dim a As Long, b As Long, low As Long, high As Long, ch As String, depth As Long
    low = first - 320: If low < 0 Then low = 0
    high = first + length + 480: If high > Len(text) Then high = Len(text)
    a = first
    Do While a > low
        ch = Mid$(text, a, 1)
        If VfcBoundary(text, a) Or ch = "(" Or ch = "[" Or ch = ")" Or ch = "]" Then Exit Do
        a = a - 1
    Loop
    b = first + length
    Do While b < high
        ch = Mid$(text, b + 1, 1)
        If VfcBoundary(text, b + 1) Then Exit Do
        If ch = "(" Or ch = "[" Then depth = depth + 1
        If ch = ")" Or ch = "]" Then
            If depth = 0 Then Exit Do
            depth = depth - 1
        End If
        b = b + 1
    Loop
    offset = a
    VfcWindow = Mid$(text, a + 1, b - a)
End Function

Private Function VfcWords(ByVal value As String) As String
    Dim part As Variant, result As String
    For Each part In Split(VefeyatNormalize(value), " ")
        If Len(CStr(part)) > 0 Then
            If IsNumeric(CStr(part)) Then
                result = result & CStr(part) & " "
            Else
                result = result & VfCollapseLetters(CStr(part)) & " "
            End If
        End If
    Next part
    VfcWords = " " & result
End Function

Private Function VfcHasWord(ByVal words As String, ByVal value As String) As Boolean
    VfcHasWord = (InStr(1, words, " " & value & " ", vbBinaryCompare) > 0)
End Function

Private Function VfcNarrative(ByVal words As String) As Boolean
    Dim item As Variant
    For Each item In Split("gore|soyle|der|dedi|demistir|belirtir|belirtmistir|savunur|savunmustur|aktarir|aciklar|yazar|yazmistir|vefat|dogdu|dogmustur|yilinda|eserinde|eserinin|sayfasinda|gorusune|tarafindan|dusuncesi|oldugunu|oldugu|ise|fakat|ancak|tartisir|inceler|degerlendirir|kaydeder|nakleder|soylemistir|gorur|olmustur|yazdi|karsilastirir|kabul|rededer|tanimlar|tanimlamistir|zikreder|dusunur|vurgular|ifade", "|")
        If VfcHasWord(words, CStr(item)) Then
            VfcNarrative = True
            Exit Function
        End If
    Next item
End Function

Private Function VfcDigit(ByVal ch As String) As Boolean
    VfcDigit = (Len(ch) = 1 And ch >= "0" And ch <= "9")
End Function

Private Function VfcLocator(ByVal value As String) As Boolean
    Dim i As Long, a As Long, b As Long, ch As String, leftPart As String, rightPart As String, words As String, part As Variant, publication As Boolean
    For i = 2 To Len(value) - 1
        ch = Mid$(value, i, 1)
        If ch = "/" Or ch = ":" Then
            a = i - 1: b = i + 1
            Do While a > 0
                If Mid$(value, a, 1) <> " " Then Exit Do
                a = a - 1
            Loop
            Do While b <= Len(value)
                If Mid$(value, b, 1) <> " " Then Exit Do
                b = b + 1
            Loop
            If a > 0 And b <= Len(value) Then
                rightPart = Mid$(value, b, 1): leftPart = LCase$(Mid$(value, a, 1))
                If VfcDigit(rightPart) Then
                    If VfcDigit(leftPart) Or InStr(1, "ivxlcdm", leftPart, vbBinaryCompare) > 0 Then
                        VfcLocator = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    words = VfcWords(value)
    For i = 1 To Len(value)
        If VfcDigit(Mid$(value, i, 1)) Then
            If VfcHasWord(words, "s") Or VfcHasWord(words, "ss") Or VfcHasWord(words, "sayfa") Or VfcHasWord(words, "p") Or VfcHasWord(words, "pp") Then
                VfcLocator = True
                Exit Function
            End If
            Exit For
        End If
    Next i
    For Each part In Split("beyrut|kahire|istanbul|ankara|press|yayinlari|yayinevi|nesriyat", "|")
        If VfcHasWord(words, CStr(part)) Then publication = True
    Next part
    If publication Then
        For Each part In Split(words, " ")
            If Len(CStr(part)) = 4 And IsNumeric(CStr(part)) Then
                VfcLocator = True
                Exit Function
            End If
        Next part
    End If
End Function

Private Function VfcRoman(ByVal value As String) As Boolean
    Dim i As Long
    If Len(value) = 0 Then Exit Function
    For i = 1 To Len(value)
        If InStr(1, "ivxlcdm", Mid$(value, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i
    VfcRoman = True
End Function

Private Function VfcTitleSignal(ByVal after As String) As Boolean
    Dim part As Variant, words As String
    words = VfcWords(after)
    If VfcHasWord(words, "s") Or VfcHasWord(words, "ss") Or VfcHasWord(words, "sayfa") Or VfcHasWord(words, "p") Or VfcHasWord(words, "pp") Then
        VfcTitleSignal = True
        Exit Function
    End If
    For Each part In Split(words, " ")
        If Len(CStr(part)) >= 3 And Not IsNumeric(CStr(part)) Then
            If CStr(part) <> "cilt" And Not VfcRoman(CStr(part)) Then
                VfcTitleSignal = True
                Exit Function
            End If
        End If
    Next part
End Function

Private Function VfcAuthorComma(ByVal after As String) As Long
    Dim i As Long, ch As String, comma As Long, prefix As String
    i = 1
    Do While i <= Len(after)
        ch = Mid$(after, i, 1)
        If ch <> " " And ch <> vbTab And ch <> ChrW(160) Then Exit Do
        i = i + 1
    Loop
    If Mid$(after, i, 1) = "," Then
        VfcAuthorComma = i
        Exit Function
    End If
    comma = InStr(i, Left$(after, 120), ",", vbBinaryCompare)
    If comma = 0 Then Exit Function
    prefix = VfcWords(Mid$(after, i, comma - i))
    If Left$(prefix, 4) = " ve " Or Mid$(after, i, 1) = "&" Then VfcAuthorComma = comma
End Function

Private Function VfcCitation(ByVal window As String, ByVal localFirst As Long, ByVal length As Long) As Boolean
    Dim after As String, words As String, comma As Long
    after = Mid$(window, localFirst + length + 1)
    comma = VfcAuthorComma(after)
    If comma = 0 Then Exit Function
    after = Mid$(after, comma)
    If Not VfcLocator(after) Then Exit Function
    If Not VfcTitleSignal(after) Then Exit Function
    words = VfcWords(window)
    If VfcNarrative(words) Then Exit Function
    VfcCitation = True
End Function

Private Function VfcSchoolRoot(ByVal key As String) As String
    Dim root As Variant, suffix As String
    For Each root In Split("hanefi|safi|maliki|hanbeli|zahiri|caferi|esari|maturidi", "|")
        If Left$(key, Len(CStr(root))) = CStr(root) Then
            suffix = Mid$(key, Len(CStr(root)) + 1)
            Select Case suffix
                Case "", "ler", "lar", "lerde", "larda", "lerden", "lardan", "lerin", "larin", "lere", "lara", "lerce", "larca", "lerince", "larinca", "nin", "ni", "ye", "yi", "den", "de"
                    VfcSchoolRoot = CStr(root)
                    Exit Function
            End Select
        End If
    Next root
End Function

Private Function VfcTerm(ByVal window As String, ByVal localFirst As Long, ByVal length As Long, ByVal key As String) As Boolean
    Dim after As String, before As String, part As Variant, stem As Variant
    If key <> "usuli" And key <> "usul" Then Exit Function
    If VfnHasSuffix(Mid$(window, localFirst + 1, length)) Then Exit Function
    before = RTrim$(VfcWords(Left$(window, localFirst)))
    For Each stem In Split("sair|sairi|mutasavif", "|")
        If Right$(before, Len(CStr(stem)) + 1) = " " & CStr(stem) Then Exit Function
    Next stem
    after = Trim$(VfcWords(Mid$(window, localFirst + length + 1)))
    If Len(after) = 0 Then Exit Function
    part = Split(after, " ")
    after = CStr(part(0))
    VfcTerm = (Left$(after, 5) = "fikih" Or Left$(after, 4) = "fikh" Or Left$(after, 5) = "hadis" Or Left$(after, 6) = "tefsir" Or after = "din" Or Left$(after, 4) = "dine")
    If VfcTerm Then Exit Function
    For Each stem In Split("dusunce|yaklasim|ilke|mesele|tartisma|cerceve|yontem|anlayis|birikim|gelenek|geleneg|gorus|delil|kaide|degerlendirme|tahlil|inceleme", "|")
        If Left$(after, Len(CStr(stem))) = CStr(stem) Then
            VfcTerm = True
            Exit Function
        End If
    Next stem
End Function

Private Function VfcOrdinary(ByVal window As String, ByVal localFirst As Long, ByVal length As Long, ByVal key As String) As Boolean
    Dim literal As String, before As String, after As String, part As Variant, stem As Variant
    If key <> "ayni" Then Exit Function
    literal = Mid$(window, localFirst + 1, length)
    If VfnHasSuffix(literal) Then Exit Function
    before = RTrim$(VfcWords(Left$(window, localFirst)))
    For Each stem In Split("imam|alame|hafiz|alim|muhadis|sair", "|")
        If Right$(before, Len(CStr(stem)) + 1) = " " & CStr(stem) Then Exit Function
    Next stem
    ' Keep the original Turkish letters: dotless ayni (same) and marked
    ' Ayni (the scholar) intentionally share a fuzzy search key, not a sense.
    literal = LCase$(Replace(literal, "I", ChrW(&H131)))
    If InStr(literal, ChrW(&HEE)) > 0 Or InStr(literal, ChrW(&H12B)) > 0 Then Exit Function
    If literal = "ayn" & ChrW(&H131) Then
        VfcOrdinary = True
        Exit Function
    End If
    after = Trim$(VfcWords(Mid$(window, localFirst + length + 1)))
    If Len(after) = 0 Then Exit Function
    part = Split(after, " ")
    after = CStr(part(0))
    For Each stem In Split("muelif|yazar|eser|kitap|gorus|sekil|donem|zaman|konu|mesele|dusunce|kaynak|metin|kisi|sahis|nokta|sonuc|anlam|hukum|yol|tavir", "|")
        If Left$(after, Len(CStr(stem))) = CStr(stem) Then
            VfcOrdinary = True
            Exit Function
        End If
    Next stem
End Function

Private Function VfcPersonPredicate(ByVal after As String) As Boolean
    Dim part As Variant, count As Long, found As Boolean
    For Each part In Split(Trim$(after), " ")
        If Len(CStr(part)) > 0 Then
            count = count + 1
            If count > 6 Then Exit For
            If Len(VfcSchoolRoot(CStr(part))) > 0 Then Exit Function
            If Left$(CStr(part), 6) = "mezhep" Or Left$(CStr(part), 6) = "mezheb" Or Left$(CStr(part), 4) = "ekol" Then Exit Function
            Select Case CStr(part)
                Case "der", "dedi", "demistir", "belirtir", "savunur", "savunmustur", "aktarir", "aciklar", "yazar", "vefat", "dogdu", "dusunur", "tartisir", "inceler", "vurgular", "rededer", "tanimlar", "kaydeder", "nakleder", "soylemistir"
                    found = True
            End Select
        End If
    Next part
    VfcPersonPredicate = found
End Function

Private Function VfcSchool(ByVal window As String, ByVal localFirst As Long, ByVal length As Long, ByVal literal As String) As Boolean
    Dim key As String, root As String, before As String, after As String, words As String, item As Variant, part As Variant
    Dim prior As Variant, index As Long
    key = VfbKey(VfbBareAlias(literal))
    root = VfcSchoolRoot(key)
    If Len(root) = 0 Then Exit Function
    If key <> root Then
        VfcSchool = True
        Exit Function
    End If
    before = RTrim$(VfcWords(Left$(window, localFirst)))
    after = VfcWords(Mid$(window, localFirst + length + 1))
    For Each item In Split("imam|fakih|alim|muhadis", "|")
        If Right$(before, Len(CStr(item)) + 1) = " " & CStr(item) Then Exit Function
    Next item
    If Len(Trim$(before)) > 0 Then
        prior = Split(Trim$(before), " ")
        index = UBound(prior)
        If CStr(prior(index)) = "ve" Or CStr(prior(index)) = "ile" Then index = index - 1
        If index >= 0 Then
            If Len(VfcSchoolRoot(CStr(prior(index)))) > 0 Then
                VfcSchool = True
                Exit Function
            End If
        End If
    End If
    If root = "safi" And VfnHasSuffix(literal) Then Exit Function
    If root = "safi" Or root = "esari" Or root = "maturidi" Then
        If VfcPersonPredicate(after) Then Exit Function
    End If
    ' Direct person predicates win over a more distant school reference.
    part = Trim$(after)
    If InStr(CStr(part), " ") > 0 Then part = Left$(CStr(part), InStr(CStr(part), " ") - 1)
    If CStr(part) <> "gore" And VfcNarrative(" " & CStr(part) & " ") Then Exit Function
    words = VfcWords(window)
    For Each part In Split(words, " ")
        If Left$(CStr(part), 6) = "mezhep" Or Left$(CStr(part), 6) = "mezheb" Or Left$(CStr(part), 4) = "ekol" Then
            VfcSchool = True
            Exit Function
        End If
    Next part
    For Each item In Split("alim|alimleri|ulema|fakih|fakihleri|fikhi|fikhina|gelenegi|geleneginde|olan", "|")
        If Left$(after, Len(CStr(item)) + 2) = " " & CStr(item) & " " Then
            VfcSchool = True
            Exit Function
        End If
    Next item
    ' A bare school label is insufficient evidence for any of the catalog's
    ' many scholars sharing that nisba. Explicit personal names remain eligible.
    VfcSchool = True
End Function

Private Function VfcNeedsContext(ByVal text As String, ByVal first As Long, ByVal length As Long, ByVal key As String) As Boolean
    If VfcWorkUsage(Mid$(text, first + length + 1, 100)) Then VfcNeedsContext = True: Exit Function
    If Len(VfcSchoolRoot(key)) > 0 Or key = "usuli" Or key = "usul" Or key = "ayni" Then
        VfcNeedsContext = True
    Else
        VfcNeedsContext = (VfcAuthorComma(Mid$(text, first + length + 1, 160)) > 0)
    End If
End Function

' A person-like string can be the title of a work. Require an immediate explicit
' title marker, not merely the word "eser" somewhere else in the sentence.
Private Function VfcWorkUsage(ByVal after As String) As Boolean
    Dim words As Variant, key As String
    key = Trim$(VfcWords(after))
    If Len(key) = 0 Then Exit Function
    words = Split(key, " ")
    If UBound(words) < 1 Then Exit Function
    If CStr(words(0)) <> "adli" And CStr(words(0)) <> "isimli" Then Exit Function
    Select Case CStr(words(1))
        Case "eser", "eserinde", "eserin", "kitap", "kitabinda", "kitabin", "makale", "makalesinde", "risale", "risalesinde"
            VfcWorkUsage = True
    End Select
End Function

Private Function VfcExclude(ByVal text As String, ByVal first As Long, ByVal length As Long, Optional ByVal key As String = "", Optional ByVal citationCheck As Boolean = True) As Boolean
    Dim window As String, offset As Long, literal As String
    If Len(key) = 0 Then key = VfbKey(VfbBareAlias(Mid$(text, first + 1, length)))
    If Not VfcNeedsContext(text, first, length, key) Then Exit Function
    window = VfcWindow(text, first, length, offset)
    literal = Mid$(text, first + 1, length)
    If VfcWorkUsage(Mid$(text, first + length + 1, 100)) Then VfcExclude = True: Exit Function
    If citationCheck Then
        If VfcCitation(window, first - offset, length) Then
            VfcExclude = True
            Exit Function
        End If
    End If
    If VfcOrdinary(window, first - offset, length, key) Then
        VfcExclude = True
    ElseIf VfcTerm(window, first - offset, length, key) Then
        VfcExclude = True
    Else
        VfcExclude = VfcSchool(window, first - offset, length, literal)
    End If
End Function

' Geometry comes from Word, text offsets from concatenated visible strings.
' This bridges protected field/control gaps without reading their code or
' using string offsets as Word positions. Never change those protected ranges.
Private Function VfcRangeExcluded(ByVal named As Word.Range, Optional ByVal citationCheck As Boolean = True) As Boolean
    Dim paragraph As Word.Range, before As Word.Range, after As Word.Range
    Dim first As Long, finish As Long, prefix As String, literal As String, suffix As String
    Set paragraph = named.Paragraphs(1).Range
    first = named.Start - 320: If first < paragraph.Start Then first = paragraph.Start
    finish = named.End + 480: If finish > paragraph.End Then finish = paragraph.End
    Set before = named.Duplicate
    before.SetRange first, named.Start
    Set after = named.Duplicate
    after.SetRange named.End, finish
    before.TextRetrievalMode.IncludeFieldCodes = False
    after.TextRetrievalMode.IncludeFieldCodes = False
    prefix = before.Text: literal = named.Text: suffix = after.Text
    VfcRangeExcluded = VfcExclude(prefix & literal & suffix, Len(prefix), Len(literal), "", citationCheck)
End Function

Private Function VfxWhite(ByVal ch As String) As Boolean
    VfxWhite = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = Chr$(11) Or ch = ChrW(160))
End Function

Private Function VfxQuoteClose(ByVal text As String, ByVal at As Long) As Long
    Dim closing As String
    Select Case Mid$(text, at, 1)
        Case Chr$(34): closing = Chr$(34)
        Case ChrW(&H201C), ChrW(&H201E): closing = ChrW(&H201D)
        Case ChrW(&HAB): closing = ChrW(&HBB)
        Case Else: Exit Function
    End Select
    VfxQuoteClose = InStr(at + 1, text, closing, vbBinaryCompare)
End Function

Private Function VfxDotBoundary(ByVal text As String, ByVal at As Long) As Boolean
    Dim i As Long, token As String, nextChar As String, raw As String, part As Variant, initials As Boolean
    If Not VfcBoundary(text, at) Then Exit Function
    If at < Len(text) And at > 1 Then
        nextChar = Mid$(text, at + 1, 1)
        If Not VfbBreak(nextChar) And at + 2 <= Len(text) Then
            If Mid$(text, at + 2, 1) = "." Then Exit Function
        End If
    End If
    i = at - 1
    Do While i > 0 And at - i <= 16
        If VfxWhite(Mid$(text, i, 1)) Or InStr("(,;[", Mid$(text, i, 1)) > 0 Then Exit Do
        i = i - 1
    Loop
    raw = Mid$(text, i + 1, at - i - 1)
    If Len(raw) = 1 And UCase$(raw) = raw And LCase$(raw) <> raw Then Exit Function
    If InStr(raw, ".") > 0 Then
        initials = True
        For Each part In Split(raw, ".")
            If Len(CStr(part)) <> 1 Then initials = False
        Next part
        If initials Then Exit Function
    End If
    token = VfbKey(raw)
    Select Case token
        Case "age", "agm", "amlf", "bkz", "krs", "a", "i", "ii", "iii", "iv", "v", "dn", "nr", "no", "bs", "eds", "haz", "cf", "vgl", "b", "s", "ss", "c", "bk", "kr", "vd", "hz", "dr", "prof", "thk", "nsr", "trc", "ed", "vol", "pp", "p"
            Exit Function
    End Select
    VfxDotBoundary = True
End Function

Private Function VfxUnits(ByVal text As String) As Collection
    Dim result As New Collection, i As Long, start As Long, ch As String, boundary As Boolean, quotedUntil As Long
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        boundary = False
        If i > quotedUntil Then
            quotedUntil = VfxQuoteClose(text, i)
            If quotedUntil = 0 Then
                boundary = (ch = ";" Or ch = ChrW(&H61B) Or ch = "!" Or ch = "?")
                If ch = "." Then boundary = VfxDotBoundary(text, i)
            End If
        End If
        If boundary Then
            If i - 1 > start Then result.Add Array(start, i - 1)
            start = i
        End If
    Next i
    If Len(text) > start Then result.Add Array(start, Len(text))
    Set VfxUnits = result
End Function

Private Function VfxNarration(ByVal value As String) As Boolean
    Dim words As String, item As Variant
    words = VfcWords(value)
    For Each item In Split("der|dedi|demistir|soylemistir|soylemektedir|savunur|savunmustur|belirtir|belirtmistir|aktarir|aciklar|yazar|yazmistir|dusunur|tartisir|inceler|vurgular|zikreder|kaydeder|nakleder|rededer|vefat|dogdu|olmustur|yilinda|sayfasinda|eserinde|gore", "|")
        If VfcHasWord(words, CStr(item)) Then
            VfxNarration = True
            Exit Function
        End If
    Next item
End Function

Private Function VfxNarrationOutsideQuotes(ByVal value As String) As Boolean
    Dim i As Long, first As Long, closing As Long, outside As String
    i = 1: first = 1
    Do While i <= Len(value)
        closing = VfxQuoteClose(value, i)
        If closing > i Then
            outside = outside & Mid$(value, first, i - first) & " "
            i = closing + 1: first = i
        Else
            i = i + 1
        End If
    Loop
    VfxNarrationOutsideQuotes = VfxNarration(outside & Mid$(value, first))
End Function

Private Function VfxPrefixIsName(ByVal value As String) As Boolean
    Dim part As Variant, words As String
    If Len(value) > 300 Or VfxNarration(value) Then Exit Function
    words = VfcWords(value)
    For Each part In Split(words, " ")
        Select Case CStr(part)
            Case "bu", "su", "icin", "gibi", "hakkinda", "uzerine", "konuda", "konuyla", "gorus", "gorusu", "karsilastirma", "eserlerinde", "kitaplarinda", "dusuncesi"
                Exit Function
        End Select
    Next part
    VfxPrefixIsName = True
End Function

Private Function VfxPageEnd(ByVal text As String, ByVal at As Long) As Long
    Dim i As Long, ch As String
    i = at
    Do While i <= Len(text)
        ch = Mid$(text, i, 1)
        If Not VfcDigit(ch) And ch <> "-" And ch <> ChrW(&H2013) And ch <> ChrW(&H2014) Then Exit Do
        i = i + 1
    Loop
    VfxPageEnd = i - 1
End Function

Private Function VfxLocators(ByVal text As String) As Collection
    Dim result As New Collection, i As Long, finish As Long, a As Long, b As Long, start As Long
    Dim word As String, ch As String, words As String, publication As Boolean, part As Variant, quotedUntil As Long
    words = VfcWords(text)
    For Each part In Split("dimask|beyrut|kahire|istanbul|ankara|london|oxford|cambridge|press|matbaa|yayinlari|yayinevi|nesriyat|thk|nsr|trc|haz", "|")
        If VfcHasWord(words, CStr(part)) Then publication = True
    Next part
    i = 1
    Do While i <= Len(text)
        quotedUntil = VfxQuoteClose(text, i)
        If quotedUntil > i Then
            i = quotedUntil + 1
            GoTo NextNumber
        End If
        If VfcDigit(Mid$(text, i, 1)) Then
            start = i - 1
            finish = i
            Do While finish <= Len(text)
                If Not VfcDigit(Mid$(text, finish, 1)) Then Exit Do
                finish = finish + 1
            Loop
            b = finish
            Do While b <= Len(text)
                If Not VfxWhite(Mid$(text, b, 1)) Then Exit Do
                b = b + 1
            Loop
            ch = Mid$(text, b, 1)
            If ch = "/" Or ch = ":" Then
                b = b + 1
                Do While b <= Len(text)
                    If Not VfxWhite(Mid$(text, b, 1)) Then Exit Do
                    b = b + 1
                Loop
                If VfcDigit(Mid$(text, b, 1)) Then
                    finish = VfxPageEnd(text, b)
                    result.Add Array(start, finish, "page")
                    i = finish + 1
                    GoTo NextNumber
                End If
            End If
            a = i - 1
            Do While a > 0
                If Not VfxWhite(Mid$(text, a, 1)) Then Exit Do
                a = a - 1
            Loop
            If a > 0 Then
                ch = Mid$(text, a, 1)
                If ch = "/" Or ch = ":" Then
                    b = a - 1
                    Do While b > 0
                        If Not VfxWhite(Mid$(text, b, 1)) Then Exit Do
                        b = b - 1
                    Loop
                    a = b
                    Do While a > 0
                        If InStr(1, "IVXLCDMivxlcdm", Mid$(text, a, 1), vbBinaryCompare) = 0 Then Exit Do
                        a = a - 1
                    Loop
                    If b > a Then
                        finish = VfxPageEnd(text, i)
                        result.Add Array(a, finish, "page")
                        i = finish + 1
                        GoTo NextNumber
                    End If
                End If
                b = a
                If Mid$(text, b, 1) = "." Then b = b - 1
                a = b
                Do While a > 0
                    If VfbBreak(Mid$(text, a, 1)) Then Exit Do
                    a = a - 1
                Loop
                word = VfbKey(Mid$(text, a + 1, b - a))
                Select Case word
                    Case "s", "ss", "p", "pp", "sayfa"
                        finish = VfxPageEnd(text, i)
                        result.Add Array(a, finish, "marked")
                        i = finish + 1
                        GoTo NextNumber
                End Select
            End If
            If publication And finish - i = 4 Then
                result.Add Array(start, finish - 1, "year")
            Else
                a = i - 1
                Do While a > 0
                    If Not VfxWhite(Mid$(text, a, 1)) Then Exit Do
                    a = a - 1
                Loop
                If a > 0 Then
                    If Mid$(text, a, 1) = "," Then
                        b = finish
                        Do While b <= Len(text)
                            If Not VfxWhite(Mid$(text, b, 1)) Then Exit Do
                            b = b + 1
                        Loop
                        ch = Mid$(text, b, 1)
                        If ch = "-" Or ch = ChrW(&H2013) Or ch = ChrW(&H2014) Then
                            b = b + 1
                            Do While b <= Len(text)
                                If Not VfxWhite(Mid$(text, b, 1)) Then Exit Do
                                b = b + 1
                            Loop
                            If VfcDigit(Mid$(text, b, 1)) Then
                                finish = VfxPageEnd(text, b)
                                result.Add Array(start, finish, "bare")
                                i = finish + 1
                                GoTo NextNumber
                            End If
                        ElseIf ch = "" Or ch = "." Or ch = "," Or ch = ")" Then
                            result.Add Array(start, finish - 1, "bare")
                        End If
                    End If
                End If
            End If
            i = finish
        Else
            i = i + 1
        End If
NextNumber:
    Loop
    Set VfxLocators = result
End Function

Private Function VfxWorkLead(ByVal value As String) As Boolean
    Dim clean As String, part As String
    clean = Trim$(value)
    Do While Len(clean) > 0
        If Not VfxWhite(Left$(clean, 1)) Then Exit Do
        clean = Mid$(clean, 2)
    Loop
    If Len(clean) = 0 Then Exit Function
    part = LCase$(Left$(clean, 3))
    Select Case part
        Case "el-", "et-", "er-", "es-", "ez-", "en-", "ed-": VfxWorkLead = True
    End Select
    If VfbInitialUpper(clean) Then VfxWorkLead = True
    If VfxQuoteClose(clean, 1) > 1 Then VfxWorkLead = True
End Function

Private Function VfxAuthorStart(ByVal text As String, ByVal first As Long, ByVal comma As Long) As Long
    Dim prefix As String, tokens As Collection, cache As Object, token As Variant, raw As String, key As String
    Dim i As Long, start As Long, finish As Long, depth As Long, ch As String
    VfxAuthorStart = -1
    finish = comma
    Do While finish > first
        If Not VfxWhite(Mid$(text, finish, 1)) Then Exit Do
        finish = finish - 1
    Loop
    If finish - first > 400 Then first = finish - 400
    If Mid$(text, finish, 1) = ")" Then
        i = finish
        Do While i > first
            ch = Mid$(text, i, 1)
            If ch = ")" Then depth = depth + 1
            If ch = "(" Then
                depth = depth - 1
                If depth = 0 Then finish = i - 1: Exit Do
            End If
            i = i - 1
        Loop
    End If
    prefix = Mid$(text, first + 1, finish - first)
    Set cache = CreateObject("Scripting.Dictionary")
    Set tokens = VfbTokens(prefix, cache)
    start = -1
    For i = tokens.Count To 1 Step -1
        token = tokens(i)
        raw = Mid$(prefix, CLng(token(1)) + 1, CLng(token(2)))
        key = CStr(token(0))
        If VfbCommonWord(key) And key <> "ve" And key <> "ile" Then
            If key <> "yazar" Or Not VfbInitialUpper(raw) Then Exit For
        End If
        If Not VfbInitialUpper(raw) And Not VfbNameLink(key) And key <> "ve" And key <> "ile" Then Exit For
        start = first + CLng(token(1))
    Next i
    VfxAuthorStart = start
End Function

Private Function VfxProbe(ByVal text As String, ByRef unit As Variant, ByRef hit As Variant, ByVal names As Collection) As Variant
    Dim nameStart As Long, nameEnd As Long, titleStart As Long, finish As Long, start As Long, comma As Long
    Dim tail As String, title As String, gap As String, locator As Variant, locations As Collection, item As Variant
    Dim chosen As Boolean, hasTitle As Boolean, nextSubject As Boolean, localEnd As Long, prefix As String
    nameStart = CLng(hit(0)): nameEnd = nameStart + CLng(hit(1))
    tail = Mid$(text, nameEnd + 1, CLng(unit(1)) - nameEnd)
    comma = VfcAuthorComma(tail)
    If comma > 0 Then
        titleStart = nameEnd + comma
    Else
        prefix = VfcWords(Mid$(text, CLng(unit(0)) + 1, nameStart - CLng(unit(0))))
        If Not VfcHasWord(prefix, "bk") And Not VfcHasWord(prefix, "bkz") And Not VfcHasWord(prefix, "krs") Then Exit Function
        If Not VfxWorkLead(tail) Then Exit Function
        titleStart = nameEnd
    End If
    title = Mid$(text, titleStart + 1, CLng(unit(1)) - titleStart)
    Set locations = VfxLocators(title)
    If locations.Count = 0 Then
        If comma = 0 Or Not VfxWorkLead(title) Or VfxNarration(title) Then Exit Function
        If VfbWordCount(title) > 16 Or Not VfcTitleSignal(title) Then Exit Function
        finish = CLng(unit(1))
    Else
        For Each locator In locations
            If Not chosen Then
                gap = Left$(title, CLng(locator(0)))
                hasTitle = VfcTitleSignal(gap)
                If Not hasTitle Then
                    If CStr(locator(2)) <> "marked" Then Exit Function
                End If
                If CStr(locator(2)) = "bare" And Not VfxWorkLead(title) Then Exit Function
                If VfxNarrationOutsideQuotes(gap) Then Exit Function
                chosen = True
            Else
                gap = Mid$(title, localEnd + 1, CLng(locator(0)) - localEnd)
                If VfxNarration(gap) Then Exit For
            End If
            localEnd = CLng(locator(1))
        Next locator
        finish = titleStart + localEnd
    End If
    If finish <= titleStart Then Exit Function
    start = nameStart
    If VfxPrefixIsName(Mid$(text, CLng(unit(0)) + 1, nameStart - CLng(unit(0)))) Then start = CLng(unit(0))
    gap = Mid$(text, finish + 1, CLng(unit(1)) - finish)
    If VfxNarration(gap) Then
        For Each item In names
            If CLng(item(0)) >= finish And CLng(item(0)) < CLng(unit(1)) Then nextSubject = True: Exit For
        Next item
        If Not nextSubject Then start = titleStart
    End If
    If VfnHasSuffix(CStr(hit(2))) Then start = titleStart
    If finish > start Then VfxProbe = Array(start, finish)
End Function

Private Function VfxContains(ByVal spans As Collection, ByVal first As Long, ByVal length As Long) As Boolean
    Dim span As Variant
    For Each span In spans
        If CLng(span(0)) <= first And CLng(span(1)) >= first + length Then
            VfxContains = True
            Exit Function
        End If
    Next span
End Function

' Strict author-year-page shorthand. The slash of a death/lifespan is deliberately
' unsupported, and only the bibliographic island ends at the closing parenthesis.
Private Function VfxAuthorYearEnd(ByVal tail As String) As Long
    Dim i As Long, first As Long, closing As Long, page As Boolean, ch As String
    i = 1
    Do While i <= Len(tail)
        If Not VfxWhite(Mid$(tail, i, 1)) Then Exit Do
        i = i + 1
    Loop
    If Mid$(tail, i, 1) <> "(" Then Exit Function
    first = i + 1
    For i = first To first + 3
        If Not VfcDigit(Mid$(tail, i, 1)) Then Exit Function
    Next i
    Do While i <= Len(tail)
        If Not VfxWhite(Mid$(tail, i, 1)) Then Exit Do
        i = i + 1
    Loop
    ch = Mid$(tail, i, 1)
    If ch <> ":" And ch <> "," Then Exit Function
    closing = InStr(i + 1, tail, ")", vbBinaryCompare)
    If closing = 0 Or closing - i > 32 Then Exit Function
    For i = i + 1 To closing - 1
        ch = Mid$(tail, i, 1)
        If VfcDigit(ch) Then
            page = True
        ElseIf Not VfxWhite(ch) And ch <> "-" And ch <> ChrW(&H2013) Then
            Exit Function
        End If
    Next i
    If page Then VfxAuthorYearEnd = closing
End Function

Private Function VfxClassify(ByVal text As String, ByVal names As Collection) As Collection
    Dim result As New Collection, units As Collection, unit As Variant, hit As Variant, span As Variant, index As Long
    Dim i As Long, start As Long, pseudo As Variant, authorYear As Long
    Set units = VfxUnits(text)
    index = 1
    For Each hit In names
        Do While index <= units.Count
            unit = units(index)
            If CLng(hit(0)) < CLng(unit(1)) Then Exit Do
            index = index + 1
        Loop
        If index > units.Count Then Exit For
        If Not VfxContains(result, CLng(hit(0)), CLng(hit(1))) Then
            span = VfxProbe(text, unit, hit, names)
            If IsEmpty(span) Then
                authorYear = VfxAuthorYearEnd(Mid$(text, CLng(hit(0)) + CLng(hit(1)) + 1, 100))
                If authorYear > 0 Then span = Array(CLng(hit(0)), CLng(hit(0)) + CLng(hit(1)) + authorYear)
            End If
            If Not IsEmpty(span) Then result.Add span
        End If
    Next hit
    ' Structural author/work evidence also covers authors absent from the
    ' person catalog, and names occurring only inside titles/editor credits.
    For Each unit In units
        For i = CLng(unit(0)) + 1 To CLng(unit(1))
            If Mid$(text, i, 1) = "," Then
                If Not VfxContains(result, i - 1, 1) Then
                    start = VfxAuthorStart(text, CLng(unit(0)), i - 1)
                    If start >= CLng(unit(0)) And start < i - 1 Then
                        pseudo = Array(start, i - 1 - start, Mid$(text, start + 1, i - 1 - start))
                        span = VfxProbe(text, unit, pseudo, names)
                        If Not IsEmpty(span) Then result.Add span
                    End If
                End If
            End If
        Next i
    Next unit
    Set VfxClassify = result
End Function

Private Sub VfzClear()
    Set mVfzDoc = Nothing: Set mVfzNotes = Nothing
    mVfzNoteItems = Array(): mVfzNoteCount = 0
    Set mVfzPatterns = Nothing: Set mVfzHeads = Nothing: Set mVfzCache = Nothing
End Sub

Private Sub VfzNextPass()
    If mVfzEpoch >= 2000000000 Then mVfzEpoch = 0
    mVfzEpoch = mVfzEpoch + 1
End Sub

Private Sub VfzBegin(ByVal doc As Word.Document, ByVal patterns As Object, ByVal heads As Object, ByVal cache As Object)
    VfzClear
    Set mVfzDoc = doc
    Set mVfzNotes = New Collection
    VfzNextPass
    Set mVfzPatterns = patterns: Set mVfzHeads = heads: Set mVfzCache = cache
End Sub

Private Function VfzVisible(ByVal region As Word.Range) As String
    Dim copy As Word.Range
    Set copy = region.Duplicate
    copy.TextRetrievalMode.IncludeFieldCodes = False
    VfzVisible = copy.Text
End Function

Private Function VfzBuild(ByVal region As Word.Range, ByVal segments As Collection, ByVal patterns As Object, ByVal heads As Object, ByVal cache As Object) As Object
    Dim map As Object, offsets As Object, texts As Object, segment As Word.Range, gap As Word.Range
    Dim pieces() As String, cursor As Long, count As Long, length As Long, key As String, value As String, rendered As String
    Dim names As Collection
    Set map = CreateObject("Scripting.Dictionary")
    Set offsets = CreateObject("Scripting.Dictionary")
    Set texts = CreateObject("Scripting.Dictionary")
    ReDim pieces(0 To segments.Count * 2)
    cursor = region.Start
    For Each segment In segments
        If segment.Start > cursor Then
            Set gap = region.Duplicate
            gap.SetRange cursor, segment.Start
            value = " " & VfzVisible(gap) & " "
            pieces(count) = value: count = count + 1: length = length + Len(value)
        End If
        key = CStr(segment.Start)
        offsets(key) = length
        value = segment.Text
        texts(key) = value
        pieces(count) = value: count = count + 1: length = length + Len(value)
        cursor = segment.End
    Next segment
    If cursor < region.End Then
        Set gap = region.Duplicate
        gap.SetRange cursor, region.End
        pieces(count) = " " & VfzVisible(gap) & " "
    End If
    rendered = Join(pieces, "")
    Set names = VfbTextMatches(rendered, patterns, heads, cache, False)
    Set map("Spans") = VfxClassify(rendered, names)
    map("Text") = rendered
    map("Signature") = VfzVisible(region)
    map("WordLength") = region.End - region.Start
    map("WordStart") = region.Start
    map("Checked") = mVfzEpoch
    Set map("Range") = region.Duplicate
    Set map("Offsets") = offsets: Set map("Texts") = texts
    Set map("Segments") = segments
    Set VfzBuild = map
End Function

Private Sub VfzRegister(ByVal map As Object)
    Dim wrapper As Object, region As Word.Range
    If mVfzDoc Is Nothing Then Exit Sub
    Set region = map("Range")
    If Not (region.Document Is mVfzDoc) Then Exit Sub
    Set wrapper = CreateObject("Scripting.Dictionary")
    Set wrapper("Range") = region.Duplicate
    Set wrapper("Map") = map
    mVfzNotes.Add wrapper
End Sub

Private Function VfzContains(ByVal map As Object, ByVal segment As Word.Range, ByVal first As Long, ByVal length As Long) As Boolean
    Dim offset As Long, key As String
    key = CStr(segment.Start)
    If Not map("Offsets").Exists(key) Then Exit Function
    offset = CLng(map("Offsets")(key)) + first
    VfzContains = VfxContains(map("Spans"), offset, length)
End Function

Private Sub VfzFreezeNotes()
    Dim items() As Variant, wrapper As Object, i As Long
    If mVfzNotes Is Nothing Then Exit Sub
    If mVfzNoteCount = mVfzNotes.Count Then Exit Sub
    ReDim items(0 To mVfzNotes.Count)
    For Each wrapper In mVfzNotes
        i = i + 1
        Set items(i) = wrapper
    Next wrapper
    mVfzNoteItems = items
    mVfzNoteCount = mVfzNotes.Count
End Sub

Private Function VfzReferenceAt(ByVal named As Word.Range) As Boolean
    Dim low As Long, high As Long, middle As Long, found As Long, wrapper As Object, map As Object
    Dim region As Word.Range, segment As Word.Range, segments As Collection, emptyVariables As Object
    Dim current As String, literal As String, offset As Long
    On Error GoTo Unavailable
    If named.StoryType <> wdFootnotesStory Then Exit Function
    If mVfzDoc Is Nothing Or mVfzNotes Is Nothing Then Exit Function
    If Not (named.Document Is mVfzDoc) Then Exit Function
    VfzFreezeNotes
    low = 1: high = mVfzNoteCount
    Do While low <= high
        middle = (low + high) \ 2
        Set wrapper = mVfzNoteItems(middle)
        Set region = wrapper("Range")
        If region.Start <= named.Start Then
            found = middle: low = middle + 1
        Else
            high = middle - 1
        End If
    Loop
    If found = 0 Then Exit Function
    Set wrapper = mVfzNoteItems(found)
    Set region = wrapper("Range")
    If named.End > region.End Or Not named.InStory(region) Then Exit Function
    Set map = wrapper("Map")
    If CLng(map("Checked")) <> mVfzEpoch Then
        current = VfzVisible(region)
        If current <> CStr(map("Signature")) Or region.End - region.Start <> CLng(map("WordLength")) Or region.Start <> CLng(map("WordStart")) Then
            ' Refresh only the changed footnote, never the complete thesis.
            Set emptyVariables = CreateObject("Scripting.Dictionary")
            Set segments = VfbPlainSegments(region, emptyVariables, False)
            Set map = VfzBuild(region, segments, mVfzPatterns, mVfzHeads, mVfzCache)
            Set wrapper("Map") = map
        End If
        map("Checked") = mVfzEpoch
    End If
    literal = named.Text
    For Each segment In map("Segments")
        If named.Start >= segment.Start And named.End <= segment.End Then
            offset = named.Start - segment.Start
            If Mid$(CStr(map("Texts")(CStr(segment.Start))), offset + 1, Len(literal)) = literal Then
                VfzReferenceAt = VfzContains(map, segment, offset, Len(literal))
            End If
            Exit Function
        End If
    Next segment
Unavailable:
End Function

Private Function VfbKey(ByVal value As String) As String
    VfbKey = VfCollapseLetters(Replace(VefeyatNormalize(value), " ", ""))
End Function

Private Function VfbBreak(ByVal ch As String) As Boolean
    Dim code As Long
    If ch = "-" Or ch = "'" Or ch = ChrW(&H2019) Or ch = ChrW(&H2018) Or _
       ch = ChrW(&H2BB) Or ch = ChrW(&H2BC) Or ch = ChrW(&H2BF) Then Exit Function
    code = AscW(ch)
    If code < 0 Then code = code + 65536
    If code >= 65 And code <= 90 Then code = code + 32
    VfbBreak = VfIsSeparator(code)
End Function

Private Function VfbTokens(ByVal text As String, ByVal cache As Object) As Collection
    Dim result As New Collection, i As Long, first As Long, value As String, key As String
    Dim slotKey As String, snapshot As Variant
    slotKey = vbNullChar & "VfbTokenBuffer"
    If Len(text) <= 32768 And cache.Exists(slotKey) Then
        snapshot = cache(slotKey)
        If StrComp(CStr(snapshot(0)), text, vbBinaryCompare) = 0 Then
            Set VfbTokens = snapshot(1)
            Exit Function
        End If
    End If
    For i = 1 To Len(text) + 1
        If i Mod 4096 = 0 Then VfbPulse VfText("Metin b\u00f6l\u00fcm\u00fc okunuyor"), i, Len(text)
        If i > Len(text) Then
            value = " "
        Else
            value = Mid$(text, i, 1)
        End If
        If VfbBreak(value) Then
            If first > 0 Then
                value = Mid$(text, first, i - first)
                If Not cache.Exists(value) Then cache.Add value, VfbKey(value)
                key = CStr(cache(value))
                If Len(key) > 0 Then result.Add Array(key, first - 1, i - first)
                first = 0
            End If
        ElseIf first = 0 Then
            first = i
        End If
    Next i
    If Len(text) <= 32768 Then
        cache(slotKey) = Array(text, result)
    End If
    Set VfbTokens = result
End Function

Private Function VfbNoise(ByVal key As String) As Boolean
    Select Case key
        Case "ibn", "ibni", "bin", "bint", "ebu", "ebul", "el", "er", "es", "et", "ez", "en", "ed", _
             "bey", "han", "sah", "seyh", "imam", "haci", "seyyid", "sair"
            VfbNoise = True
    End Select
End Function

' Whole-document discovery is intentionally stricter than an explicit lookup.
' These words can occur inside a person's full name, but cannot nominate a
' person on their own, even at the start of a sentence or in a heading.
Private Function VfbCommonWord(ByVal key As String) As Boolean
    Static words As Object
    Dim item As Variant
    If words Is Nothing Then
        Set words = CreateObject("Scripting.Dictionary")
        For Each item In Split( _
            "tarif|tanim|ilim|bilim|isim|ad|bilgi|tarih|eser|kitap|usul|fikih|" & _
            "hukuk|hadis|tefsir|kelam|felsefe|mantik|delil|akil|nazar|mana|anlam|hak|" & _
            "hakikat|adalet|hayat|olum|dogum|zaman|yer|islam|kuran|sunet|tasavuf|ahlak|" & _
            "fikir|dusunce|mesele|bahis|bab|fasl|fasil|cilt|sayfa|baski|yayin|tercume|" & _
            "serh|hasiye|kavram|terim|kisi|sahis|yazar|arastirma|metin|deneme|konu|once|" & _
            "sonra|gore|icin|ile|ve|veya|yahut|ama|fakat|bu|su|bir|" & _
            "iki|uc|her|olan|olarak|oldu|dedi|der|gibi|uzere|hakinda", "|")
            words(CStr(item)) = True
        Next item
    End If
    VfbCommonWord = words.Exists(key)
End Function

Private Sub VfbAddPattern(ByVal patterns As Object, ByVal heads As Object, ByVal value As String, ByVal index As Long, ByVal cache As Object)
    Dim key As String, ids As Object
    If Not cache.Exists(value) Then cache.Add value, VfbKey(value)
    key = CStr(cache(value))
    If Len(key) < 3 Or Len(value) > VfaMaxName Or value = "---" Then Exit Sub
    If VfbNoise(key) Then Exit Sub
    If VfbCommonWord(key) Then Exit Sub
    If Not patterns.Exists(key) Then
        Set ids = CreateObject("Scripting.Dictionary")
        patterns.Add key, ids
        heads(Left$(key, 3)) = True
    End If
    Set ids = patterns(key)
    ids(CStr(index)) = True
    VfbRecordShape patterns, key, VfbWordCount(value)
End Sub

Private Function VfbWordCount(ByVal value As String) As Long
    Dim i As Long, inWord As Boolean
    For i = 1 To Len(value)
        If VfbBreak(Mid$(value, i, 1)) Then
            inWord = False
        ElseIf Not inWord Then
            VfbWordCount = VfbWordCount + 1
            inWord = True
        End If
    Next i
End Function

Private Sub VfbRecordShape(ByVal patterns As Object, ByVal key As String, ByVal count As Long)
    Dim shapes As Object, counts As Object
    If Not patterns.Exists("!word-counts") Then
        Set shapes = CreateObject("Scripting.Dictionary")
        patterns.Add "!word-counts", shapes
    End If
    Set shapes = patterns("!word-counts")
    If Not shapes.Exists(key) Then
        Set counts = CreateObject("Scripting.Dictionary")
        shapes.Add key, counts
    End If
    shapes(key)(CStr(count)) = True
End Sub

Private Function VfbShapeAllowed(ByVal patterns As Object, ByVal key As String, ByVal count As Long) As Boolean
    Dim shapes As Object
    If Not patterns.Exists("!word-counts") Then Exit Function
    Set shapes = patterns("!word-counts")
    If shapes.Exists(key) Then VfbShapeAllowed = shapes(key).Exists(CStr(count))
End Function

Private Function VfbIndependentAlias(ByVal value As String, ByVal shortName As String) As Boolean
    Dim key As String
    key = VfbKey(value)
    If Len(VfcSchoolRoot(key)) > 0 Then
        VfbIndependentAlias = (key = VfbKey(shortName))
    Else
        VfbIndependentAlias = True
    End If
End Function

Private Function VfbNaturalOrder(ByVal value As String) As String
    Dim comma As Long, first As String, last As String
    comma = InStr(1, value, ",", vbBinaryCompare)
    If comma = 0 Or InStr(comma + 1, value, ",", vbBinaryCompare) > 0 Then Exit Function
    first = Trim$(Mid$(value, comma + 1)): last = Trim$(Left$(value, comma - 1))
    If Len(first) > 0 And Len(last) > 0 Then VfbNaturalOrder = first & " " & last
End Function

Private Sub VfbBuildPatterns(ByVal patterns As Object, ByVal heads As Object, ByVal cache As Object, Optional ByVal rowIds As Object = Nothing)
    Dim i As Long, done As Long, row As Variant, id As Variant, owner As Collection
    Set owner = mVfRows
    If rowIds Is Nothing Then
        For Each row In owner
            i = i + 1
            VfbIndexRow patterns, heads, cache, row, i
            If i Mod 8 = 0 Then VfbPulse VfText("Ad dizini haz\u0131rlan\u0131yor"), i, owner.Count
        Next row
        VfbPulse VfText("Ad dizini haz\u0131r"), owner.Count, owner.Count
    Else
        For Each id In rowIds.Keys
            i = CLng(id): row = owner(i)
            VfbIndexRow patterns, heads, cache, row, i
            done = done + 1
            If done Mod 8 = 0 Then VfbPulse VfText("Yaln\u0131z ilgili adlar dizine al\u0131n\u0131yor"), done, rowIds.Count
        Next id
        VfbPulse VfText("Belge ad dizini haz\u0131r"), done, rowIds.Count
    End If
End Sub

Private Sub VfbIndexRow(ByVal patterns As Object, ByVal heads As Object, ByVal cache As Object, ByRef row As Variant, ByVal i As Long)
    Dim j As Long, token As Variant, tokens As Collection, value As String, part As Variant, aliasText As String
        For j = 0 To 2
            value = CStr(row(j))
            VfbAddPattern patterns, heads, value, i, cache
            VfbAddPattern patterns, heads, VfbWithoutArticles(value), i, cache
            If j = 0 Then
                aliasText = VfbNaturalOrder(value)
                If Len(aliasText) > 0 Then
                    VfbAddPattern patterns, heads, aliasText, i, cache
                    VfbAddPattern patterns, heads, VfbWithoutArticles(aliasText), i, cache
                End If
            End If
            Set tokens = VfbTokens(value, cache)
            For Each token In tokens
                aliasText = Mid$(value, CLng(token(1)) + 1, CLng(token(2)))
                ' Short catalog names supply independent aliases. Do not turn
                ' every ancestor/component of a long name into a batch alias.
                If j = 0 Or VfbCompoundAlias(aliasText) Then
                    If VfbIndependentAlias(aliasText, CStr(row(0))) Then
                        VfbAddPattern patterns, heads, aliasText, i, cache
                    End If
                End If
            Next token
            If j = 0 Then
                For Each part In Split(VefeyatNormalize(value), " ")
                    If VfbIndependentAlias(CStr(part), CStr(row(0))) Then
                        VfbAddPattern patterns, heads, CStr(part), i, cache
                    End If
                Next part
            End If
        Next j
End Sub

Private Function VfbCompoundAlias(ByVal value As String) As Boolean
    Dim key As String
    If InStr(value, "-") = 0 Then Exit Function
    If InStr(value, "'") = 0 And InStr(value, ChrW(&H2019)) = 0 Then Exit Function
    key = VfbKey(value)
    If Left$(key, 3) = "ebu" Or Left$(key, 3) = "abu" Then Exit Function
    VfbCompoundAlias = True
End Function

Private Function VfbNameLink(ByVal key As String) As Boolean
    Select Case key
        Case "b", "ibn", "ibni", "bin", "bint", "ebu", "ebul", "el", "er", "es", "et", "ez", "en", "ed"
            VfbNameLink = True
    End Select
End Function

Private Function VfbInitialUpper(ByVal value As String) As Boolean
    Dim ch As String, code As Long, i As Long
    Select Case LCase$(Left$(value, 3))
        Case "el-", "er-", "es-", "et-", "ez-", "en-", "ed-"
            value = Mid$(value, 4)
    End Select
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If UCase$(ch) <> LCase$(ch) Then
            VfbInitialUpper = (ch = UCase$(ch))
            Exit Function
        End If
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        ' Uncased scripts have no initial-capital convention.
        If code >= &H590 And code <= &H1FFF Or code >= &H2E80 Then
            VfbInitialUpper = True
            Exit Function
        End If
    Next i
End Function

' Consume a plausible complete name before considering a component match.
' Unknown capitalized surnames also extend it, so an unlisted person's given
' names cannot nominate unrelated historical people. Lowercase catalog parts
' participate in a compound, but standalone Latin discovery needs a capital.
Private Function VfbNameSpanEnd(ByVal text As String, ByRef tokens As Variant, ByVal first As Long, _
                                ByVal count As Long, ByVal patterns As Object) As Long
    Dim i As Long, token As Variant, previous As Variant, raw As String, key As String, gap As String
    VfbNameSpanEnd = first
    For i = first To count
        token = tokens(i)
        raw = Mid$(text, CLng(token(1)) + 1, CLng(token(2)))
        key = CStr(token(0))
        If VfnHasSuffix(raw) Then key = VfbKey(VfbBareAlias(raw))
        If VfbCommonWord(key) Then Exit For
        If VfbNoise(key) And Not VfbNameLink(key) Then Exit For
        If Not VfbNameLink(key) And Not patterns.Exists(key) And Not VfbInitialUpper(raw) Then Exit For
        If i > first Then
            previous = tokens(i - 1)
            gap = Mid$(text, CLng(previous(1)) + CLng(previous(2)) + 1, CLng(token(1)) - CLng(previous(1)) - CLng(previous(2)))
            If Not VfbJoinable(gap, CStr(previous(0))) Then Exit For
        End If
        If Not VfbNameLink(key) Then VfbNameSpanEnd = i
        If VfnHasSuffix(raw) Then Exit For
    Next i
End Function

Private Function VfbJoinable(ByVal gap As String, ByVal previous As String) As Boolean
    Dim i As Long, ch As String
    VfbJoinable = True
    For i = 1 To Len(gap)
        ch = Mid$(gap, i, 1)
        If ch = "." And (previous = "b" Or previous = "i" Or previous = "v") Then
        ElseIf ch <> " " And ch <> vbTab And ch <> ChrW(160) Then
            VfbJoinable = False
            Exit Function
        End If
    Next i
End Function

' Return zero-based UTF-16 offsets, literal text and indexed identity key.
' Longest non-overlapping names win. Never cross a paragraph or a sentence.
Private Function VfbTextMatches(ByVal text As String, ByVal patterns As Object, _
                                ByVal heads As Object, ByVal cache As Object, Optional ByVal strictPeople As Boolean = False, Optional ByVal citationCheck As Boolean = True) As Collection
    Dim result As New Collection, sourceTokens As Collection, tokens As Variant, token As Variant, previous As Variant, tokenCount As Long
    Dim i As Long, j As Long, first As Long, finish As Long, bestEnd As Long, bestIndex As Long
    Dim key As String, bestKey As String, gap As String, tokenKey As String, rawToken As String, inflected As Boolean
    Dim spanEnd As Long, accepted As Boolean
    Set sourceTokens = VfbTokens(text, cache)
    tokenCount = sourceTokens.Count
    tokens = VfbTokenArray(sourceTokens)
    Set sourceTokens = Nothing
    i = 1
    Do While i <= tokenCount
        If i Mod 64 = 1 Then VfbPulse VfText("Metindeki adlar aran\u0131yor"), i, tokenCount
        token = tokens(i)
        first = CLng(token(1))
        key = vbNullString
        bestEnd = 0
        bestIndex = i
        spanEnd = i
        If strictPeople Then spanEnd = VfbNameSpanEnd(text, tokens, i, tokenCount, patterns)
        For j = i To tokenCount
            token = tokens(j)
            finish = CLng(token(1)) + CLng(token(2))
            If finish - first > VfaMaxName Then Exit For
            If j > i Then
                previous = tokens(j - 1)
                gap = Mid$(text, CLng(previous(1)) + CLng(previous(2)) + 1, CLng(token(1)) - CLng(previous(1)) - CLng(previous(2)))
                If Not VfbJoinable(gap, CStr(previous(0))) Then Exit For
            End If
            tokenKey = CStr(token(0))
            rawToken = Mid$(text, CLng(token(1)) + 1, CLng(token(2)))
            inflected = False
            If VfnHasSuffix(rawToken) Then
                If Not patterns.Exists(VfCollapseLetters(key & tokenKey)) Then
                    tokenKey = VfbKey(VfbBareAlias(rawToken))
                    inflected = True
                End If
            End If
            key = VfCollapseLetters(key & tokenKey)
            If Len(key) >= 3 Then
                If Not heads.Exists(Left$(key, 3)) Then Exit For
            End If
            If patterns.Exists(key) Then
                bestEnd = finish
                bestIndex = j
                bestKey = key
            End If
            If inflected Then Exit For
        Next j
        If bestEnd > first Then
            accepted = True
            If strictPeople Then
                If Not VfbShapeAllowed(patterns, bestKey, bestIndex - i + 1) Then accepted = False
                If spanEnd > bestIndex Then accepted = False
                If bestIndex = i Then
                    rawToken = Mid$(text, first + 1, bestEnd - first)
                    If VfbCommonWord(VfbKey(VfbBareAlias(rawToken))) Or Not VfbInitialUpper(rawToken) Then accepted = False
                End If
            End If
            If accepted And strictPeople Then accepted = Not VfcExclude(text, first, bestEnd - first, bestKey, citationCheck)
            If accepted Then result.Add Array(first, bestEnd - first, Mid$(text, first + 1, bestEnd - first), bestKey)
            If spanEnd > bestIndex Then bestIndex = spanEnd
            i = bestIndex + 1
        Else
            i = spanEnd + 1
        End If
    Loop
    Set VfbTextMatches = result
End Function

Private Sub VfbSortIntervals(ByRef starts() As Long, ByRef ends() As Long, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Long, value As Long
    i = low: j = high: pivot = starts((low + high) \ 2)
    Do While i <= j
        Do While starts(i) < pivot: i = i + 1: Loop
        Do While starts(j) > pivot: j = j - 1: Loop
        If i <= j Then
            value = starts(i): starts(i) = starts(j): starts(j) = value
            value = ends(i): ends(i) = ends(j): ends(j) = value
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then VfbSortIntervals starts, ends, low, j
    If i < high Then VfbSortIntervals starts, ends, i, high
End Sub

' Split around fields/controls BEFORE reading text. Word field codes consume
' positions even when Range.Text omits them; whole-story string offsets are unsafe.
Private Function VfbPlainSegments(ByVal region As Word.Range, ByVal variables As Object, Optional ByVal protectOwnedNames As Boolean = True) As Collection
    Dim result As New Collection, field As Word.Field, control As Word.ContentControl
    Dim starts() As Long, ends() As Long, count As Long, i As Long, cursor As Long
    Dim segment As Word.Range, bindingId As String, aliasText As String
    count = region.Fields.Count + region.ContentControls.Count
    If count > 0 Then
        ReDim starts(1 To count): ReDim ends(1 To count)
        i = 0
        For Each field In region.Fields
            i = i + 1
            starts(i) = field.Code.Start - 1
            ends(i) = VfdFieldEnd(field)
            If i Mod 32 = 0 Then VfbPulse VfText("Mevcut alanlar korunuyor"), i, count
            If protectOwnedNames And variables.Exists(VfaNamespaceVariable) Then
                bindingId = VfaBindingFromFieldCode(field.Code.Text, CStr(variables(VfaNamespaceVariable)))
                If Len(bindingId) > 0 Then
                    If variables.Exists(VfaVariableName(bindingId, "Alias")) Then
                        aliasText = CStr(variables(VfaVariableName(bindingId, "Alias")))
                        If Len(aliasText) > 0 And starts(i) - Len(aliasText) >= region.Start Then
                            Set segment = region.Duplicate
                            segment.SetRange starts(i) - Len(aliasText), starts(i)
                            If StrComp(segment.Text, aliasText, vbTextCompare) = 0 Then starts(i) = segment.Start
                        End If
                    End If
                End If
            End If
        Next field
        For Each control In region.ContentControls
            i = i + 1
            starts(i) = control.Range.Start
            ends(i) = control.Range.End
            If i Mod 32 = 0 Then VfbPulse VfText("Mevcut alanlar korunuyor"), i, count
        Next control
        VfbSortIntervals starts, ends, 1, count
    End If
    cursor = region.Start
    For i = 1 To count
        If starts(i) > cursor Then
            Set segment = region.Duplicate
            segment.SetRange cursor, starts(i)
            result.Add segment
        End If
        If ends(i) > cursor Then cursor = ends(i)
    Next i
    If cursor < region.End Then
        Set segment = region.Duplicate
        segment.SetRange cursor, region.End
        result.Add segment
    End If
    Set VfbPlainSegments = result
End Function

Private Function VfbScanRegion(ByVal region As Word.Range, ByVal patterns As Object, _
                                ByVal heads As Object, ByVal cache As Object, ByVal snapshots As Collection, ByVal variables As Object, Optional ByVal includeOwnedNames As Boolean = False, Optional ByVal noteAnchor As Word.Range = Nothing) As Collection
    Dim result As New Collection, segment As Word.Range, matches As Collection, item As Variant
    Dim hit As Object, snapshot As Object, text As String, named As Word.Range, segments As Collection, citationMap As Object
    Set segments = VfbPlainSegments(region, variables, Not includeOwnedNames)
    If region.StoryType = wdFootnotesStory Then
        Set citationMap = VfzBuild(region, segments, patterns, heads, cache)
        VfzRegister citationMap
    End If
    For Each segment In segments
        If citationMap Is Nothing Then
            text = segment.Text
        Else
            text = CStr(citationMap("Texts")(CStr(segment.Start)))
        End If
        Set snapshot = CreateObject("Scripting.Dictionary")
        Set snapshot("Range") = segment.Duplicate
        snapshot("Text") = text
        If Not noteAnchor Is Nothing Then Set snapshot("Anchor") = noteAnchor.Duplicate
        snapshots.Add snapshot
        Set matches = VfbTextMatches(text, patterns, heads, cache, True, (citationMap Is Nothing))
        For Each item In matches
            If Not citationMap Is Nothing Then
                If VfzContains(citationMap, segment, CLng(item(0)), CLng(item(1))) Then GoTo NextHit
            End If
            Set named = segment.Duplicate
            named.SetRange segment.Start + CLng(item(0)), segment.Start + CLng(item(0)) + CLng(item(1))
            If StrComp(named.Text, CStr(item(2)), vbBinaryCompare) = 0 Then
                If VfcNeedsContext(text, CLng(item(0)), CLng(item(1)), CStr(item(3))) Then
                    If VfcRangeExcluded(named, (citationMap Is Nothing)) Then GoTo NextHit
                End If
                Set hit = CreateObject("Scripting.Dictionary")
                Set hit("Range") = named
                hit("Alias") = CStr(item(2))
                hit("Key") = CStr(item(3))
                result.Add hit
            End If
NextHit:
        Next item
    Next segment
    Set VfbScanRegion = result
End Function

Private Function VfbOrderedHits(ByVal doc As Word.Document, ByVal patterns As Object, ByVal heads As Object, _
                                ByVal cache As Object, ByVal snapshots As Collection, Optional ByVal includeOwnedNames As Boolean = False) As Collection
    Dim result As New Collection, main As Collection, notes As Collection, note As Word.Footnote
    Dim index As Long, hit As Object, named As Word.Range, variables As Object, variable As Word.Variable
    Dim number As Long, detail As String, source As String
    On Error GoTo Failed
    Set variables = CreateObject("Scripting.Dictionary")
    For Each variable In doc.Variables
        variables(variable.Name) = variable.Value
    Next variable
    Set main = VfbScanRegion(doc.Content, patterns, heads, cache, snapshots, variables, includeOwnedNames)
    index = 1
    VfbBeginNotes doc.Footnotes.Count
    For Each note In doc.Footnotes
        mVfbParentCompleted = note.Index - 1
        Do While index <= main.Count
            Set hit = main(index)
            Set named = hit("Range")
            If named.Start >= note.Reference.Start Then Exit Do
            result.Add hit
            index = index + 1
        Loop
        Set notes = VfbScanRegion(note.Range, patterns, heads, cache, snapshots, variables, includeOwnedNames, note.Reference)
        For Each hit In notes
            Set hit("Anchor") = note.Reference.Duplicate
            result.Add hit
        Next hit
        mVfbParentCompleted = note.Index
        VfbPulse VfText("Dipnotlar taran\u0131yor"), note.Index, doc.Footnotes.Count
    Next note
    VfbEndNotes
    Do While index <= main.Count
        result.Add main(index)
        index = index + 1
    Loop
    Set VfbOrderedHits = result
    Exit Function
Failed:
    number = Err.Number: detail = Err.Description: source = Err.Source
    VfbEndNotes
    Err.Raise number, source, detail
End Function

Private Function VfbRankRows(ByVal ids As Object, ByVal query As String) As Collection
    Dim result As New Collection, indices() As Long, ranks() As Long, i As Long, j As Long, index As Variant
    Dim row As Variant, rank As Long, best As Long, score As Long, count As Long, seen As Object, key As String
    ReDim indices(1 To ids.Count): ReDim ranks(1 To ids.Count)
    Set seen = CreateObject("Scripting.Dictionary")
    For Each index In ids.Keys
        row = mVfRows(CLng(index))
        key = VfpPersonKey(row)
        If Not seen.Exists(key) Then
            seen.Add key, True
            best = 0
            For j = 0 To 2
                score = VefeyatNameMatch(CStr(row(j)), query)
                If score > best Then best = score
            Next j
            rank = best * 1000 + VfProminence(row, VfbKey(query))
            count = count + 1
            i = count
            Do While i > 1
                If ranks(i - 1) >= rank Then Exit Do
                ranks(i) = ranks(i - 1): indices(i) = indices(i - 1): i = i - 1
            Loop
            ranks(i) = rank: indices(i) = CLng(index)
        End If
    Next index
    For i = 1 To count
        result.Add mVfRows(indices(i))
    Next i
    Set VfbRankRows = result
End Function

Private Function VfbReview(ByVal groups As Object, ByVal order As Collection, ByVal patterns As Object) As Boolean
    Dim key As Variant, group As Object, candidates As Collection, row As Variant, choice As Long
    Dim picker As VefeyatKisiSec0341, summary As New Collection, approved As New Collection, i As Long, stopped As Boolean
    Dim number As Long, detail As String
    On Error GoTo Failed
    For i = 1 To order.Count
        Set group = groups(CStr(order(i)))
        Set candidates = VfbRankRows(patterns(CStr(order(i))), CStr(group("Query")))
        choice = 0
        If candidates.Count = 1 Then
            choice = 1
        ElseIf candidates.Count > 1 Then
            Set picker = New VefeyatKisiSec0341
            picker.InitCandidates candidates, CStr(group("Query"))
            picker.SetScanMode i, order.Count
            picker.Show vbModal
            choice = picker.ChosenIndex
            stopped = picker.StopRequested
            Unload picker
            Set picker = Nothing
            If stopped Then Exit Function
        End If
        group("Enabled") = False
        If choice > 0 And choice <= candidates.Count Then
            row = candidates(choice)
            If VfaValidDate(CStr(row(3))) Then
                group("Row") = row
                group("Enabled") = True
                approved.Add group
                summary.Add Array(CStr(group("Alias")), CStr(row(0)), CStr(row(1)), CStr(row(3)), _
                    CStr(group("Count")) & VfText(" ge\u00e7i\u015f. ") & CStr(row(4)), CStr(row(5)))
            End If
        End If
    Next i
    If summary.Count = 0 Then Exit Function
    Set picker = New VefeyatKisiSec0341
    picker.InitCandidates summary, VfText("Belge taramas\u0131: son kontrol")
    picker.SetReviewMode
    picker.Show vbModal
    If picker.ChosenIndex > 0 Then
        For i = 1 To approved.Count
            Set group = approved(i)
            group("Enabled") = picker.ReviewIncludes(i)
        Next i
        VfbReview = True
    End If
    Unload picker
    Exit Function
Failed:
    number = Err.Number: detail = Err.Description
    On Error Resume Next
    Unload picker
    On Error GoTo 0
    Err.Raise number, "Vefeyat", detail
End Function

Private Sub VfbAuditDocument()
    Dim doc As Word.Document, patterns As Object, heads As Object, cache As Object, groups As Object
    Dim snapshots As New Collection, order As New Collection, hits As Collection, hit As Object, group As Object, snapshot As Object
    Dim plan As New Collection, seen As Object, bindings As Object, aliases As Object, id As Variant, person As String
    Dim key As String, row As Variant, named As Word.Range, source As Word.Range, previousScreen As Boolean
    Dim i As Long, inserted As Long, skipped As Long, changed As Boolean, undoObject As Object, started As Boolean
    Dim failure As String, failureNumber As Long, firstOnly As Boolean, literal As String, bindingId As String
    On Error GoTo Failed
    previousScreen = Application.ScreenUpdating
    If Not VfrEditableDocument() Then Exit Sub
    If Not VefeyatEnsureData() Then Exit Sub
    Set doc = ActiveDocument
    previousScreen = Application.ScreenUpdating
    Set patterns = CreateObject("Scripting.Dictionary")
    Set heads = CreateObject("Scripting.Dictionary")
    Set cache = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    VfxPrepareDocument doc, patterns, heads, cache, False
    Application.StatusBar = VfText("Belge ve dipnotlar taran\u0131yor...")
    Set hits = VfbOrderedHits(doc, patterns, heads, cache, snapshots)
    VfxVerifySnapshots snapshots, cache
    For Each hit In hits
        key = CStr(hit("Key"))
        If Not groups.Exists(key) Then
            Set group = CreateObject("Scripting.Dictionary")
            group("Alias") = CStr(hit("Alias"))
            group("Query") = VfbBareAlias(CStr(hit("Alias")))
            group("Count") = 0&
            groups.Add key, group
            order.Add key
        End If
        Set group = groups(key)
        group("Count") = CLng(group("Count")) + 1
    Next hit
    If order.Count = 0 Then
        VfMessage VfText("Toplu denetimde g\u00fcvenilir bir ad e\u015fle\u015fmesi bulunamad\u0131. Kavramlar ve b\u00fct\u00fcn\u00fc e\u015fle\u015fmeyen ad gruplar\u0131 par\u00e7alanarak \u015fah\u0131s \u00f6nerilmez. Ad\u0131 kendiniz se\u00e7ip \u00d6l\u00fcm tarihi ekle ile arayabilirsiniz."), vbInformation, "Vefeyat"
        Exit Sub
    End If
    If Not VfbReview(groups, order, patterns) Then Exit Sub
    VfbPulse VfText("Belge de\u011fi\u015fiklik i\u00e7in kontrol ediliyor"), 0, snapshots.Count
    If Not ActiveDocument Is doc Then Err.Raise VfaInvalidData, "Vefeyat", "Active document changed."
    If Not VfaEditable(doc) Then Exit Sub
    For Each snapshot In snapshots
        Set source = snapshot("Range")
        If StrComp(source.Text, CStr(snapshot("Text")), vbBinaryCompare) <> 0 Then
            Err.Raise VfaInvalidData, "Vefeyat", "Document changed during review; scan again."
        End If
    Next snapshot
    Set seen = CreateObject("Scripting.Dictionary")
    Set bindings = CreateObject("Scripting.Dictionary")
    Set aliases = CreateObject("Scripting.Dictionary")
    firstOnly = (VfpGetScope(doc) = "first")
    For Each id In VfaActiveBindingIds(doc)
        person = VfaReadVariable(doc, VfaVariableName(CStr(id), "Key"))
        aliases(VfaAliasKey(VfaReadVariable(doc, VfaVariableName(CStr(id), "Alias")))) = person
        If firstOnly Then seen(person) = True
    Next id
    For Each hit In hits
        Set group = groups(CStr(hit("Key")))
        If CBool(group("Enabled")) Then
            row = group("Row")
            person = VfpPersonKey(row)
            literal = CStr(hit("Alias"))
            If VfbShouldConsider(firstOnly, seen, person) Then
                Set named = hit("Range")
                If aliases.Exists(VfaAliasKey(literal)) Then
                    If CStr(aliases(VfaAliasKey(literal))) <> person Then
                        skipped = skipped + 1
                        GoTo NextHit
                    End If
                End If
                If VfaAlreadyHasDate(named, CStr(row(3))) Or VfbAnyDate(named) Then
                    skipped = skipped + 1
                Else
                    plan.Add hit
                End If
            End If
        End If
NextHit:
    Next hit
    If plan.Count = 0 Then
        VfMessage VfText("Eklenecek yeni tarih yok. Mevcut tarihler korundu."), vbInformation, "Vefeyat"
        Exit Sub
    End If
    VfbPulse VfText("Onaylanan tarihler uygulan\u0131yor"), 0, plan.Count, False
    started = VfaBeginUndo(undoObject)
    VfaEnsureNamespace doc, changed
    Application.ScreenUpdating = False
    For i = plan.Count To 1 Step -1
        Set hit = plan(i)
        Set group = groups(CStr(hit("Key")))
        row = group("Row")
        literal = CStr(hit("Alias"))
        key = VfaAliasKey(literal) & "|" & VfpPersonKey(row)
        If Not bindings.Exists(key) Then
            bindingId = VfaGetBinding(doc, row, literal, changed)
            bindings.Add key, bindingId
        End If
        bindingId = CStr(bindings(key))
        Set named = hit("Range")
        If StrComp(named.Text, literal, vbBinaryCompare) <> 0 Then Err.Raise VfaInvalidData, "Vefeyat", "Name changed during insertion."
        VfaInsertManaged doc, named, VfpFormatDate(CStr(row(3)), VfpGetPrefix(doc)), bindingId, changed
        inserted = inserted + 1
        VfbPulse VfText("Onaylanan tarihler uygulan\u0131yor"), inserted, plan.Count, False
    Next i
    VfbPulse VfText("Tarihler eklendi"), inserted, plan.Count, False
    Application.ScreenUpdating = previousScreen
    VfaEndUndo undoObject, started
    VflCaptureCommand
    VfMessage CStr(inserted) & VfText(" tarih eklendi. ") & CStr(skipped) & _
        VfText(" mevcut tarih veya \u00e7ak\u0131\u015fma korundu. Ctrl+Z ile geri alabilirsiniz."), vbInformation, "Vefeyat"
    Exit Sub
Failed:
    failureNumber = Err.Number
    failure = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = previousScreen
    On Error GoTo 0
    VfaEndUndo undoObject, started
    If failureNumber = VfbCancelError And Not changed Then
        Application.StatusBar = VfText("Tarama durduruldu; belge de\u011fi\u015ftirilmedi.")
        Exit Sub
    End If
    VfaReportFailure failure, changed
End Sub

' Preserve a recognizable existing death date even if it differs from the CSV.
Private Function VfbAnyDate(ByVal occurrence As Word.Range) As Boolean
    Dim following As Word.Range, text As String, field As Word.Field, key As Variant
    Set following = occurrence.Duplicate
    following.SetRange occurrence.End, occurrence.End
    Call following.MoveEnd(wdCharacter, 100)
    For Each field In following.Fields
        If field.Code.Start - 1 <= occurrence.End + 2 And Not VfdIsIndexMark(field) Then
            VfbAnyDate = True
            Exit Function
        End If
    Next field
    text = LCase$(VfaTrimSpaces(following.Text))
    For Each key In Array("(o.", VfText("(\u00f6."), VfText("(\u00f6l."), "(v.", "(d.")
        If Left$(text, Len(CStr(key))) = CStr(key) Then
            VfbAnyDate = True
            Exit Function
        End If
    Next key
End Function

Private Function VfbShouldConsider(ByVal firstOnly As Boolean, ByVal seen As Object, ByVal person As String) As Boolean
    If firstOnly And seen.Exists(person) Then Exit Function
    seen(person) = True
    VfbShouldConsider = True
End Function

Private Function VfbBareAlias(ByVal value As String) As String
    ' Use the same apostrophe normalization as suffix detection. Repeating
    ' only part of it made left/modifier quotes produce Left$(value, -1).
    VfbBareAlias = VfNameQuery(value)
End Function

' Collection.Item numeric lookup can be linear; materialize once for large texts.
Private Function VfbTokenArray(ByVal tokens As Collection) As Variant
    Dim result() As Variant, token As Variant, index As Long
    ReDim result(0 To tokens.Count)
    For Each token In tokens
        index = index + 1
        result(index) = token
        If index Mod 4096 = 0 Then VfbPulse VfText("Metin dizini haz\u0131rlan\u0131yor"), index, tokens.Count
    Next token
    VfbTokenArray = result
End Function

Private Function VfbWithoutArticles(ByVal value As String) As String
    Dim article As Variant
    value = " " & VefeyatNormalize(value) & " "
    For Each article In Array("el", "er", "es", "et", "ez", "en", "ed")
        value = Replace(value, " " & CStr(article) & " ", " ")
    Next article
    VfbWithoutArticles = Trim$(value)
End Function

Private Function VfbUiClock() As Single
    VfbUiClock = Timer
End Function

' Yield at most every 100 ms, not at every name/character. The owning modal
' progress form blocks document editing while still accepting cancellation.
Private Sub VfbPulse(ByVal stage As String, ByVal completed As Long, ByVal total As Long, _
                     Optional ByVal allowCancel As Boolean = True)
    Dim stamp As Single, elapsed As Single
    If mVfbProgress Is Nothing Then Exit Sub
    If mVfbProgress.StopRequested Then Err.Raise VfbCancelError, "Vefeyat", "Scan cancelled."
    stamp = VfbUiClock()
    elapsed = stamp - mVfbLastPulse
    If elapsed < 0 Then elapsed = elapsed + 86400!
    ' Short notes may alternate stages or complete immediately. Neither is a
    ' reason to repaint thousands of times. Cancellation-gate changes are immediate.
    mVfbLastStage = stage
    If mVfbHasPainted And allowCancel = mVfbLastAllowCancel And elapsed < 0.1 Then Exit Sub
    mVfbHasPainted = True
    mVfbLastAllowCancel = allowCancel
    mVfbLastStage = stage
    mVfbLastPulse = stamp
    If Len(mVfbParentStage) > 0 Then
        mVfbProgress.UpdateProgress mVfbParentStage, mVfbParentCompleted, mVfbParentTotal, allowCancel
    Else
        mVfbProgress.UpdateProgress stage, completed, total, allowCancel
    End If
    mVfbProgress.Repaint
    DoEvents
    If mVfbProgress.StopRequested Then Err.Raise VfbCancelError, "Vefeyat", "Scan cancelled."
End Sub

Private Sub VfbResetProgress()
    mVfbExecuting = False
    mVfbExecuted = False
    mVfbHasPainted = False
    mVfbLastStage = vbNullString
    mVfbLastPulse = 0!
    VfbEndNotes
End Sub

Private Sub VfbPrepareIndex(ByRef patterns As Object, ByRef heads As Object, ByVal cache As Object)
    Dim owner As Collection, number As Long, detail As String
    If mVfbIndexBuilding Then Err.Raise vbObjectError + 2451, "Vefeyat", "Name index preparation is already running."
    If Not mVfbCachedRows Is Nothing Then
        If mVfbCachedRows Is mVfRows Then
            Set patterns = mVfbCachedPatterns
            Set heads = mVfbCachedHeads
            VfbPulse VfText("Ad dizini haz\u0131r"), mVfRows.Count, mVfRows.Count
            Exit Sub
        End If
    End If
    Set owner = mVfRows
    Set mVfbCachedRows = Nothing
    mVfbIndexBuilding = True
    On Error GoTo Failed
    VfbPulse VfText("Ad dizini haz\u0131rlan\u0131yor"), 0, owner.Count
    VfbBuildPatterns patterns, heads, cache
    If Not (owner Is mVfRows) Then Err.Raise vbObjectError + 2452, "Vefeyat", "Catalog changed during preparation; start a new scan explicitly."
    ' Publish only a complete index for the captured, still-current catalog.
    Set mVfbCachedPatterns = patterns
    Set mVfbCachedHeads = heads
    Set mVfbCachedRows = owner
    mVfbIndexBuilding = False
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    mVfbIndexBuilding = False
    Set mVfbCachedRows = Nothing
    Err.Raise number, "Vefeyat", detail
End Sub

Private Sub VfbBeginNotes(ByVal total As Long)
    mVfbParentStage = VfText("Dipnotlar taran\u0131yor")
    mVfbParentCompleted = 0
    mVfbParentTotal = total
End Sub

Private Sub VfbEndNotes()
    mVfbParentStage = vbNullString
    mVfbParentCompleted = 0
    mVfbParentTotal = 0
End Sub

Private Function VfxCharacter(ByVal ch As String) As String
    Static folded As Object
    Dim code As Long
    code = AscW(ch)
    If code >= 65 And code <= 90 Then code = code + 32
    If code >= 0 And code < 128 Then
        If VfIsSeparator(code) Then VfxCharacter = " " Else VfxCharacter = ChrW(code)
        Exit Function
    End If
    If code < 0 Then code = code + 65536
    If VfIsSeparator(code) Then VfxCharacter = " ": Exit Function
    If folded Is Nothing Then Set folded = CreateObject("Scripting.Dictionary")
    If Not folded.Exists(ch) Then folded(ch) = VefeyatNormalize(ch)
    VfxCharacter = CStr(folded(ch))
End Function

' Every normalized lexical prefix plus the compact field prefix. The latter
' preserves matches spanning short connectors/articles. Repeated letters fold.
Private Function VfxPrefixCodes(ByVal text As String, Optional ByVal wanted As Object = Nothing) As String
    Dim i As Long, j As Long, ch As String, folded As String, prefix As String, whole As String
    Dim previous As String, previousWhole As String, result As String
    Dim raw As String, joined As String, previousJoined As String, carryOne As String, carryTwo As String
    For i = 1 To Len(text)
        raw = Mid$(text, i, 1)
        If VfbBreak(raw) Then joined = "": previousJoined = ""
        folded = VfxCharacter(raw)
        For j = 1 To Len(folded)
            ch = Mid$(folded, j, 1)
            If ch = " " Then
                If Len(prefix) > 0 And Len(prefix) < 3 Then
                    If prefix <> carryOne And prefix <> carryTwo Then
                        If Len(carryOne) = 0 Then carryOne = prefix Else carryTwo = prefix
                    End If
                End If
                prefix = "": previous = ""
            Else
                If Len(carryOne) > 0 And ch <> Right$(carryOne, 1) Then
                    carryOne = carryOne & ch
                    If Len(carryOne) = 3 Then
                        VfxEmit result, carryOne, wanted
                        carryOne = ""
                    End If
                End If
                If Len(carryTwo) > 0 And ch <> Right$(carryTwo, 1) Then
                    carryTwo = carryTwo & ch
                    If Len(carryTwo) = 3 Then
                        VfxEmit result, carryTwo, wanted
                        carryTwo = ""
                    End If
                End If
                If ch <> previous And Len(prefix) < 3 Then
                    prefix = prefix & ch
                    If Len(prefix) = 3 Then VfxEmit result, prefix, wanted
                End If
                previous = ch
                If ch <> previousWhole And Len(whole) < 3 Then
                    whole = whole & ch
                    If Len(whole) = 3 Then VfxEmit result, whole, wanted
                End If
                previousWhole = ch
                If ch <> previousJoined And Len(joined) < 3 Then
                    joined = joined & ch
                    If Len(joined) = 3 Then VfxEmit result, joined, wanted
                End If
                previousJoined = ch
            End If
        Next j
        If i Mod 8192 = 0 Then VfbPulse VfText("Belgedeki ad anahtarlar\u0131 okunuyor"), i, Len(text)
    Next i
    VfxPrefixCodes = result
End Function

Private Sub VfxEmit(ByRef result As String, ByVal prefix As String, ByVal wanted As Object)
    If wanted Is Nothing Then
        result = result & "|" & prefix & "|"
    ElseIf Not wanted.Exists(prefix) Then
        wanted.Add prefix, True
    End If
End Sub

Private Sub VfxAddHeads(ByVal text As String, ByVal wanted As Object)
    Dim ignored As String
    ignored = VfxPrefixCodes(text, wanted)
End Sub

Private Sub VfxReadRegion(ByVal region As Word.Range, ByVal variables As Object, ByVal protectNames As Boolean, ByVal wanted As Object)
    Dim segment As Word.Range
    For Each segment In VfbPlainSegments(region, variables, protectNames)
        VfxAddHeads segment.Text, wanted
    Next segment
End Sub

Private Function VfxDocumentHeads(ByVal doc As Word.Document, ByVal includeOwnedNames As Boolean) As Object
    Dim wanted As Object, variables As Object, variable As Word.Variable, note As Word.Footnote, done As Long
    Set wanted = CreateObject("Scripting.Dictionary")
    Set variables = CreateObject("Scripting.Dictionary")
    For Each variable In doc.Variables
        variables(variable.Name) = variable.Value
    Next variable
    VfxReadRegion doc.Content, variables, Not includeOwnedNames, wanted
    For Each note In doc.Footnotes
        VfxReadRegion note.Range, variables, Not includeOwnedNames, wanted
        done = done + 1
        If done Mod 16 = 0 Then VfbPulse VfText("Dipnot ad anahtarlar\u0131 okunuyor"), done, doc.Footnotes.Count
    Next note
    Set VfxDocumentHeads = wanted
End Function

Private Sub VfxCatalogCodes(ByVal owner As Collection)
    Dim codes() As String, row As Variant, i As Long
    If Not mVfxOwner Is Nothing Then
        If mVfxOwner Is owner Then Exit Sub
    End If
    ReDim codes(0 To owner.Count)
    ' Sequential enumeration, never N indexed reads of VBA's linked Collection.
    For Each row In owner
        i = i + 1
        codes(i) = VfxPrefixCodes(CStr(row(0))) & VfxPrefixCodes(CStr(row(1))) & VfxPrefixCodes(CStr(row(2))) & VfxPrefixCodes(VfbNaturalOrder(CStr(row(0))))
        If i Mod 64 = 0 Then VfbPulse VfText("Katalog ad anahtarlar\u0131 okunuyor"), i, owner.Count
    Next row
    If Not (owner Is mVfRows) Then Err.Raise vbObjectError + 2461, "Vefeyat", "Catalog changed during prefix preparation."
    mVfxCodes = codes
    Set mVfxOwner = owner
End Sub

Private Function VfxSubset(ByVal wanted As Object, ByVal available As Object) As Boolean
    Dim key As Variant
    If available Is Nothing Then Exit Function
    For Each key In wanted.Keys
        If Not available.Exists(CStr(key)) Then Exit Function
    Next key
    VfxSubset = True
End Function

Private Function VfxRelevant(ByVal codes As String, ByVal wanted As Object) As Boolean
    Dim key As Variant
    For Each key In Split(codes, "|")
        If Len(CStr(key)) > 0 Then
            If wanted.Exists(CStr(key)) Then VfxRelevant = True: Exit Function
        End If
    Next key
End Function

Private Sub VfxPrepareDocument(ByVal doc As Word.Document, ByRef patterns As Object, ByRef heads As Object, _
                               ByVal cache As Object, Optional ByVal includeOwnedNames As Boolean = True)
    Dim owner As Collection, wanted As Object, ids As Object, i As Long, number As Long, detail As String
    If mVfxBusy Then Err.Raise vbObjectError + 2460, "Vefeyat", "Document index preparation is already running."
    Set owner = mVfRows
    mVfxBusy = True
    On Error GoTo Failed
    VfbPulse VfText("\u00d6nce belgedeki ad anahtarlar\u0131 belirleniyor"), 0, 1
    Set wanted = VfxDocumentHeads(doc, includeOwnedNames)
    Set cache(vbNullChar & "VfxWanted") = wanted
    mVfxLastHeads = wanted.Count: mVfxLastReused = False
    If Not mVfxPatternOwner Is Nothing Then
        If mVfxPatternOwner Is owner And VfxSubset(wanted, mVfxWanted) Then
            Set patterns = mVfxPatterns: Set heads = mVfxHeads
            mVfxLastReused = True
            GoTo Ready
        End If
    End If
    Set patterns = CreateObject("Scripting.Dictionary")
    Set heads = CreateObject("Scripting.Dictionary")
    Set ids = CreateObject("Scripting.Dictionary")
    If wanted.Count > 0 Then
        VfxCatalogCodes owner
        For i = 1 To owner.Count
            If VfxRelevant(mVfxCodes(i), wanted) Then ids(CStr(i)) = True
            If i Mod 256 = 0 Then VfbPulse VfText("Belgeyle ilgili katalog kay\u0131tlar\u0131 ayr\u0131l\u0131yor"), i, owner.Count
        Next i
    End If
    mVfxLastRows = ids.Count
    VfbBuildPatterns patterns, heads, cache, ids
    If Not (owner Is mVfRows) Then Err.Raise vbObjectError + 2462, "Vefeyat", "Catalog changed during document index preparation."
    Set mVfxPatternOwner = owner: Set mVfxWanted = wanted
    Set mVfxPatterns = patterns: Set mVfxHeads = heads
Ready:
    If Not (owner Is mVfRows) Then Err.Raise vbObjectError + 2462, "Vefeyat", "Catalog changed during document preparation."
    mVfxBusy = False
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    mVfxBusy = False
    Set mVfxPatternOwner = Nothing
    Err.Raise number, "Vefeyat", detail
End Sub

' Check the already captured scan text, without reading the Word story again.
Private Sub VfxVerifySnapshots(ByVal snapshots As Collection, ByVal cache As Object)
    Dim snapshot As Object, actual As Object, expected As Object
    If Not cache.Exists(vbNullChar & "VfxWanted") Then Exit Sub
    Set expected = cache(vbNullChar & "VfxWanted")
    For Each snapshot In snapshots
        Set actual = CreateObject("Scripting.Dictionary")
        VfxAddHeads CStr(snapshot("Text")), actual
        If Not VfxSubset(actual, expected) Then Err.Raise vbObjectError + 2463, "Vefeyat", "Document changed while preparing names; start a new scan."
    Next snapshot
End Sub

Private Function VfxDiagnostic() As String
    If mVfxWanted Is Nothing Then Exit Function
    VfxDiagnostic = VfText("Son tarama: ") & CStr(mVfxLastHeads) & VfText(" belge anahtar\u0131; ") & _
        CStr(mVfxLastRows) & VfText(" katalog kayd\u0131 ayr\u0131nt\u0131l\u0131 dizinde.")
    If mVfxLastReused Then VfxDiagnostic = VfxDiagnostic & VfText(" Haz\u0131r dizin kullan\u0131ld\u0131.")
End Function

Private Function VfmDictionary() As Object
    Set VfmDictionary = CreateObject("Scripting.Dictionary")
End Function

Private Function VfmState(ByVal doc As Word.Document) As Object
    Dim state As Object, owner As Word.Document
    If mVfmStates Is Nothing Then Set mVfmStates = New Collection
    For Each state In mVfmStates
        Set owner = state("Document")
        If owner Is doc Then
            Set VfmState = state
            Exit Function
        End If
    Next state
End Function

Private Sub VfmInvalidate(ByVal doc As Word.Document)
    Dim state As Object
    Set state = VfmState(doc)
    If Not state Is Nothing Then state("Ready") = False
End Sub

Private Sub VfmForget(ByVal doc As Word.Document)
    Dim i As Long, state As Object, owner As Word.Document
    If mVfmStates Is Nothing Then Exit Sub
    For i = mVfmStates.Count To 1 Step -1
        Set state = mVfmStates(i)
        Set owner = state("Document")
        If owner Is doc Then mVfmStates.Remove i
    Next i
End Sub

Private Sub VfmPrune()
    Dim i As Long, state As Object, owner As Word.Document, doc As Word.Document, found As Boolean
    If mVfmStates Is Nothing Then Exit Sub
    For i = mVfmStates.Count To 1 Step -1
        Set state = mVfmStates(i)
        Set owner = state("Document")
        found = False
        For Each doc In Application.Documents
            If owner Is doc Then found = True
        Next doc
        If Not found Then mVfmStates.Remove i
    Next i
End Sub

Private Function VfmEmptyArray() As Variant
    Dim result(0 To 0) As Variant
    VfmEmptyArray = result
End Function

Private Function VfmNewState(ByVal doc As Word.Document) As Object
    Dim state As Object
    Set state = VfmDictionary()
    Set state("Document") = doc
    Set state("Chunks") = VfmDictionary()
    Set state("Index") = VfmDictionary()
    Set state("Words") = VfmDictionary()
    Set state("Choices") = New Collection
    Set state("Pending") = New Collection
    Set state("Dirty") = New Collection
    state("ObservedAt") = 0!
    Set state("Aliases") = VfmDictionary()
    Set state("Bindings") = VfmDictionary()
    VfmLoadAliases state, doc
    state("Main") = VfmEmptyArray()
    state("Notes") = VfmEmptyArray()
    state("Next") = 1&
    state("Ready") = False
    state("MainLength") = doc.Content.End
    state("NotesCount") = doc.Footnotes.Count
    Set VfmNewState = state
End Function

Private Sub VfmAddKey(ByVal keys As Object, ByVal key As String)
    If Len(key) >= 2 Then keys(key) = True
End Sub

Private Function VfmMakeChunk(ByVal state As Object, ByVal region As Word.Range, ByVal anchor As Word.Range) As Object
    Dim chunk As Object, tokens As Collection, token As Variant, part As Variant, value As String
    Set chunk = VfmDictionary()
    chunk("Id") = CStr(state("Next"))
    state("Next") = CLng(state("Next")) + 1
    Set chunk("Range") = region.Duplicate
    Set chunk("Anchor") = anchor
    chunk("Text") = region.Text
    chunk("Length") = region.End - region.Start
    Set chunk("Keys") = VfmDictionary()
    Set tokens = VfbTokens(CStr(chunk("Text")), state("Words"))
    For Each token In tokens
        VfmAddKey chunk("Keys"), CStr(token(0))
        value = Mid$(CStr(chunk("Text")), CLng(token(1)) + 1, CLng(token(2)))
        If InStr(value, "-") > 0 Or InStr(value, "'") > 0 Or InStr(value, ChrW(&H2019)) > 0 Then
            For Each part In Split(VefeyatNormalize(value), " ")
                VfmAddKey chunk("Keys"), VfbKey(CStr(part))
            Next part
        End If
    Next token
    Set VfmMakeChunk = chunk
End Function

Private Sub VfmIndexChunk(ByVal state As Object, ByVal chunk As Object)
    Dim key As Variant, bucket As Object
    state("Chunks").Add CStr(chunk("Id")), chunk
    For Each key In chunk("Keys").Keys
        If Not state("Index").Exists(CStr(key)) Then
            Set bucket = VfmDictionary()
            state("Index").Add CStr(key), bucket
        End If
        Set bucket = state("Index")(CStr(key))
        bucket(CStr(chunk("Id"))) = True
    Next key
End Sub

Private Sub VfmUnindexChunk(ByVal state As Object, ByVal chunk As Object)
    Dim key As Variant, bucket As Object, id As String
    id = CStr(chunk("Id"))
    For Each key In chunk("Keys").Keys
        If state("Index").Exists(CStr(key)) Then
            Set bucket = state("Index")(CStr(key))
            If bucket.Exists(id) Then bucket.Remove id
            If bucket.Count = 0 Then state("Index").Remove CStr(key)
        End If
    Next key
    If state("Chunks").Exists(id) Then state("Chunks").Remove id
End Sub

Private Function VfmPieces(ByVal state As Object, ByVal area As Word.Range, ByVal anchor As Word.Range) As Collection
    Dim result As New Collection, paragraph As Word.Paragraph, segment As Word.Range, region As Word.Range
    Dim noVariables As Object, first As Long, finish As Long, last As Long, text As String, cut As Long
    Set noVariables = VfmDictionary()
    For Each paragraph In area.Paragraphs
        ' Read field/control geometry from the full paragraph before clipping.
        ' Keep names preceding managed dates: dated and undated mentions both count.
        For Each segment In VfbPlainSegments(paragraph.Range, noVariables, False)
            first = segment.Start: last = segment.End
            If first < area.Start Then first = area.Start
            If last > area.End Then last = area.End
            Do While first < last
                finish = first + VfmBlockSize
                If finish > last Then finish = last
                Set region = segment.Duplicate
                region.SetRange first, finish
                If finish < last Then
                    text = region.Text
                    cut = InStrRev(text, " ")
                    If cut > VfmBlockSize \ 2 Then finish = first + cut
                    region.SetRange first, finish
                End If
                result.Add VfmMakeChunk(state, region, anchor)
                first = finish
                VfbPulse VfText("Belge belle\u011fi haz\u0131rlan\u0131yor"), first - area.Start, area.End - area.Start
            Loop
        Next segment
    Next paragraph
    Set VfmPieces = result
End Function

Private Function VfmObjects(ByVal values As Collection) As Variant
    Dim result() As Variant, value As Object, i As Long
    ReDim result(0 To values.Count)
    For Each value In values
        i = i + 1
        Set result(i) = value
    Next value
    VfmObjects = result
End Function

Private Sub VfmWarm(ByVal doc As Word.Document)
    Dim state As Object, main As Collection, notes As New Collection, pieces As Collection
    Dim note As Word.Footnote, chunk As Object, count As Long
    On Error GoTo Failed
    If mVfmStates Is Nothing Then Set mVfmStates = New Collection
    Set state = VfmNewState(doc)
    Set main = VfmPieces(state, doc.Content, Nothing)
    For Each note In doc.Footnotes
        Set pieces = VfmPieces(state, note.Range, note.Reference.Duplicate)
        For Each chunk In pieces
            notes.Add chunk
        Next chunk
    Next note
    For Each chunk In main
        VfmIndexChunk state, chunk
        count = count + 1
        VfbPulse VfText("Belge belle\u011fi dizini kuruluyor"), count, main.Count + notes.Count
    Next chunk
    For Each chunk In notes
        VfmIndexChunk state, chunk
        count = count + 1
        VfbPulse VfText("Belge belle\u011fi dizini kuruluyor"), count, main.Count + notes.Count
    Next chunk
    state("Main") = VfmObjects(main)
    state("Notes") = VfmObjects(notes)
    state("Ready") = True
    VfmForget doc
    mVfmStates.Add state
    Application.StatusBar = VfText("Belge belle\u011fi haz\u0131r. Sonraki denetimler de\u011fi\u015fen b\u00f6l\u00fcmleri kullan\u0131r.")
    Exit Sub
Failed:
    ' Never replace a complete old index with an incomplete initialization.
    Err.Raise Err.Number, "VefeyatMemory", Err.Description
End Sub

Private Function VfmStoryKey(ByVal region As Word.Range) As String
    If region.StoryType = wdMainTextStory Then VfmStoryKey = "Main"
    If region.StoryType = wdFootnotesStory Then VfmStoryKey = "Notes"
End Function

Private Function VfmFloor(ByRef items As Variant, ByVal position As Long) As Long
    Dim low As Long, high As Long, middle As Long, region As Word.Range
    low = 1: high = UBound(items)
    Do While low <= high
        middle = (low + high) \ 2
        Set region = items(middle)("Range")
        If region.Start <= position Then
            VfmFloor = middle
            low = middle + 1
        Else
            high = middle - 1
        End If
    Loop
End Function

Private Function VfmScope(ByVal doc As Word.Document, ByVal target As Word.Range, ByRef anchor As Word.Range) As Word.Range
    Dim note As Word.Footnote
    If target.StoryType = wdMainTextStory Then
        Set VfmScope = doc.Content.Duplicate
        Exit Function
    End If
    If target.StoryType <> wdFootnotesStory Then Exit Function
    For Each note In doc.Footnotes
        If target.Start >= note.Range.Start And target.Start < note.Range.End Then
            Set anchor = note.Reference.Duplicate
            Set VfmScope = note.Range.Duplicate
            Exit Function
        End If
    Next note
End Function

Private Function VfmSplice(ByRef items As Variant, ByVal first As Long, ByVal last As Long, ByVal additions As Collection) As Variant
    Dim result() As Variant, i As Long, at As Long, item As Object, removed As Long
    removed = last - first + 1
    If removed < 0 Then removed = 0
    ReDim result(0 To UBound(items) - removed + additions.Count)
    For i = 1 To first - 1
        at = at + 1: Set result(at) = items(i)
    Next i
    For Each item In additions
        at = at + 1: Set result(at) = item
    Next item
    For i = last + 1 To UBound(items)
        at = at + 1: Set result(at) = items(i)
    Next i
    VfmSplice = result
End Function

Private Sub VfmRepair(ByVal state As Object, ByVal target As Word.Range)
    Dim items As Variant, key As String, first As Long, last As Long, i As Long
    Dim area As Word.Range, scope As Word.Range, region As Word.Range, anchor As Word.Range
    Dim start As Long, finish As Long, additions As Collection, chunk As Object, changed As Boolean
    key = VfmStoryKey(target)
    If Len(key) = 0 Then Exit Sub
    Set scope = VfmScope(state("Document"), target, anchor)
    If scope Is Nothing Then Exit Sub
    items = state(key)
    start = target.Start - VfaMaxName
    finish = target.End + VfaMaxName
    If start < scope.Start Then start = scope.Start
    If finish > scope.End Then finish = scope.End
    first = VfmFloor(items, start)
    If first = 0 Then first = 1
    Do While first <= UBound(items)
        Set region = items(first)("Range")
        If region.End >= start Then Exit Do
        first = first + 1
    Loop
    last = first - 1
    For i = first To UBound(items)
        Set region = items(i)("Range")
        If region.Start >= finish Then Exit For
        If region.Start < start Then start = region.Start
        If region.End > finish Then finish = region.End
        last = i
    Next i
    If start < scope.Start Then start = scope.Start
    If finish > scope.End Then finish = scope.End
    ' Avoid reading unchanged text blocks on every ordinary caret movement.
    changed = (last < first)
    For i = first To last
        Set chunk = items(i)
        Set region = chunk("Range")
        If region.End - region.Start <> CLng(chunk("Length")) Then changed = True
        If region.Fields.Count > 0 Or region.ContentControls.Count > 0 Then changed = True
        If StrComp(region.Text, CStr(chunk("Text")), vbBinaryCompare) <> 0 Then changed = True
    Next i
    If Not changed Then Exit Sub
    If finish - start > 32768 And mVfbProgress Is Nothing Then
        Set area = scope.Duplicate
        area.SetRange start, finish
        VfmQueue state, area
        Exit Sub
    End If
    Set area = scope.Duplicate
    area.SetRange start, finish
    Set additions = VfmPieces(state, area, anchor)
    ' Publish only after the replacement blocks have been completely read.
    For i = first To last
        VfmUnindexChunk state, items(i)
    Next i
    For Each chunk In additions
        VfmIndexChunk state, chunk
    Next chunk
    state(key) = VfmSplice(items, first, last, additions)
End Sub

Private Sub VfmObserve(ByVal doc As Word.Document, ByVal selected As Word.Range, Optional ByVal force As Boolean = True)
    Dim state As Object, current As Word.Range, previous As Word.Range, area As Word.Range
    Dim stamp As Single, elapsed As Single
    On Error GoTo Uncertain
    If mVfmBusy Or mVfbRunning Then Exit Sub
    Set state = VfmState(doc)
    If state Is Nothing Then Exit Sub
    If Not CBool(state("Ready")) Then Exit Sub
    Set current = selected.Duplicate
    current.Collapse wdCollapseEnd
    VfmDirty state, current
    stamp = VfbUiClock()
    elapsed = stamp - CSng(state("ObservedAt"))
    If elapsed < 0 Then elapsed = elapsed + 86400!
    If Not force And elapsed < 0.2 Then Exit Sub
    mVfmBusy = True
    If state.Exists("Last") Then
        Set previous = state("Last")
        VfmDirty state, previous
    End If
    For Each area In state("Dirty")
        VfmRepair state, area
    Next area
    Set state("Dirty") = New Collection
    Set state("Last") = current.Duplicate
    state("ObservedAt") = stamp
    state("MainLength") = doc.Content.End
    state("NotesCount") = doc.Footnotes.Count
    mVfmBusy = False
    Exit Sub
Uncertain:
    If Not state Is Nothing Then state("Ready") = False
    mVfmBusy = False
End Sub

Private Sub VfmDirty(ByVal state As Object, ByVal current As Word.Range)
    Dim previous As Word.Range, first As Long, last As Long
    ' Coalesce nearby main-text carets. Footnotes are kept separate because
    ' different notes share a story but not the same editing scope.
    If current.StoryType = wdMainTextStory Then
        For Each previous In state("Dirty")
            If previous.StoryType = wdMainTextStory Then
                If previous.Start <= current.End + 200 And current.Start <= previous.End + 200 Then
                    first = previous.Start: last = previous.End
                    If current.Start < first Then first = current.Start
                    If current.End > last Then last = current.End
                    previous.SetRange first, last
                    Exit Sub
                End If
            End If
        Next previous
    End If
    state("Dirty").Add current.Duplicate
End Sub

Private Sub VfmRemember(ByVal doc As Word.Document, ByVal named As Word.Range, ByVal person As String)
    Dim state As Object, item As Object
    Set state = VfmState(doc)
    If state Is Nothing Then Exit Sub
    Set item = VfmDictionary()
    Set item("Range") = named.Duplicate
    item("Text") = named.Text
    item("Person") = person
    state("Choices").Add item
    VfmKnownAlias state, person, CStr(item("Text"))
End Sub

Private Sub VfmAlias(ByVal state As Object, ByVal aliases As Object, ByVal candidates As Object, ByVal value As String)
    Dim key As String, tokens As Collection, token As Variant, part As Variant, bucket As Object
    Dim rare As Object, count As Long, id As Variant
    If value = "---" Or Not VfaValidAlias(value) Then Exit Sub
    key = VfbKey(value)
    If Len(key) < 2 Then Exit Sub
    aliases(key) = True
    Set tokens = VfbTokens(VefeyatNormalize(value), state("Words"))
    count = 2147483647
    If state("Index").Exists(key) Then
        Set rare = state("Index")(key)
        count = rare.Count
    End If
    For Each token In tokens
        key = CStr(token(0))
        If state("Index").Exists(key) Then
            Set bucket = state("Index")(key)
            If bucket.Count < count Then
                count = bucket.Count
                Set rare = bucket
            End If
        End If
    Next token
    If rare Is Nothing Then Exit Sub
    For Each id In rare.Keys
        candidates(CStr(id)) = True
    Next id
End Sub

Private Sub VfmAliases(ByVal state As Object, ByRef row As Variant, ByVal query As String, _
                       ByVal aliases As Object, ByVal candidates As Object)
    Dim i As Long, value As String, words As Variant, shortName As String, part As Variant, token As Variant, known As Object
    VfmAlias state, aliases, candidates, VfbBareAlias(query)
    For i = 0 To 2
        value = CStr(row(i))
        VfmAlias state, aliases, candidates, value
        VfmAlias state, aliases, candidates, VfbWithoutArticles(value)
    Next i
    If state("Aliases").Exists(VfpPersonKey(row)) Then
        Set known = state("Aliases")(VfpPersonKey(row))
        For Each part In known.Keys
            VfmAlias state, aliases, candidates, CStr(part)
        Next part
    End If
    For Each token In VfbTokens(CStr(row(1)), state("Words"))
        value = Mid$(CStr(row(1)), CLng(token(1)) + 1, CLng(token(2)))
        If InStr(value, "-") > 0 And (InStr(value, "'") > 0 Or InStr(value, ChrW(&H2019)) > 0) Then
            If Left$(VfbKey(value), 3) <> "ebu" And Left$(VfbKey(value), 3) <> "abu" Then VfmAlias state, aliases, candidates, value
        End If
    Next token
    shortName = CStr(row(0))
    If InStr(shortName, ",") > 0 Then
        shortName = Left$(shortName, InStr(shortName, ",") - 1)
    Else
        words = Split(VefeyatNormalize(shortName), " ")
        shortName = CStr(words(UBound(words)))
    End If
    If Len(VfbKey(shortName)) >= 3 And Not VfbNoise(VfbKey(shortName)) Then
        VfmAlias state, aliases, candidates, shortName
        For Each part In Array("el-", "er-", "es-", "et-", "ez-", "en-", "ed-")
            VfmAlias state, aliases, candidates, CStr(part) & shortName
        Next part
    End If
End Sub

Private Function VfmPosition(ByVal state As Object, ByVal target As Word.Range) As Variant
    Dim anchor As Word.Range, scope As Word.Range
    If target.StoryType = wdMainTextStory Then
        VfmPosition = Array(target.Start, 0&, 0&)
    Else
        Set scope = VfmScope(state("Document"), target, anchor)
        If anchor Is Nothing Then Err.Raise VfaInvalidData, "VefeyatMemory", "Footnote location changed."
        VfmPosition = Array(anchor.Start, 1&, target.Start)
    End If
End Function

Private Function VfmEarlierPosition(ByRef a As Variant, ByRef b As Variant) As Boolean
    VfmEarlierPosition = VfaOrderBefore(CLng(a(0)), CLng(a(1)), CLng(a(2)), CLng(b(0)), CLng(b(1)), CLng(b(2)))
End Function

Private Function VfmIdentity(ByVal state As Object, ByVal named As Word.Range, ByVal person As String) As Long
    Dim choice As Object, region As Word.Range, field As Word.Field, id As String, key As String, i As Long
    For i = state("Choices").Count To 1 Step -1
        Set choice = state("Choices")(i)
        Set region = choice("Range")
        If named.InStory(region) Then
            If region.Start <= named.Start And region.End >= named.End Then
                If StrComp(region.Text, CStr(choice("Text")), vbBinaryCompare) = 0 Then
                    If CStr(choice("Person")) = person Then VfmIdentity = 1 Else VfmIdentity = -1
                    Exit Function
                End If
            End If
        End If
    Next i
    For Each field In named.Paragraphs(1).Range.Fields
        If field.Code.Start - 1 >= named.End And field.Code.Start - 1 <= named.End + 2 Then
            id = VfaFieldBinding(state("Document"), field)
            If Len(id) > 0 Then
                key = VfaReadVariable(state("Document"), VfaVariableName(id, "Key"))
                If Len(key) > 0 Then
                    If key = person Then VfmIdentity = 1 Else VfmIdentity = -1
                    Exit Function
                End If
            End If
        End If
    Next field
End Function

Private Function VfmFindEarlier(ByVal state As Object, ByRef row As Variant, ByVal query As String, ByVal target As Word.Range, Optional ByVal started As Single = -1, Optional ByVal contextAware As Boolean = False, Optional ByVal exclusions As Object = Nothing) As Object
    Dim aliases As Object, candidates As Object, heads As Object, key As Variant, id As Variant
    Dim chunk As Object, region As Word.Range, named As Word.Range, matches As Collection, match As Variant
    Dim selectedPosition As Variant, position As Variant, bestPosition As Variant, result As Object, identity As Long, anchor As Word.Range
    Dim blockPosition As Variant
    If started < 0 Then started = VfbUiClock()
    selectedPosition = VfmPosition(state, target)
    If CLng(selectedPosition(0)) = 0 And CLng(selectedPosition(1)) = 0 And CLng(selectedPosition(2)) = 0 Then Exit Function
    Set aliases = VfmDictionary(): Set candidates = VfmDictionary(): Set heads = VfmDictionary()
    VfmAliases state, row, query, aliases, candidates
    For Each key In aliases.Keys
        heads(Left$(CStr(key), 3)) = True
    Next key
    For Each id In candidates.Keys
        If VfmLookupExpired(started) Then GoTo TimedOut
        If state("Chunks").Exists(CStr(id)) Then
            Set chunk = state("Chunks")(CStr(id))
            Set region = chunk("Range")
            If region.End > region.Start Then
                If region.StoryType = wdMainTextStory Then
                    blockPosition = Array(region.Start, 0&, 0&)
                Else
                    Set anchor = chunk("Anchor")
                    blockPosition = Array(anchor.Start, 1&, region.Start)
                End If
                ' A later block cannot contain an earlier occurrence. Once a
                ' result exists, blocks starting after it cannot improve it.
                If Not VfmEarlierPosition(blockPosition, selectedPosition) Then GoTo NextChunk
                If Not result Is Nothing Then
                    If Not VfmEarlierPosition(blockPosition, bestPosition) Then GoTo NextChunk
                End If
                Set matches = VfbTextMatches(CStr(chunk("Text")), aliases, heads, state("Words"), False)
                For Each match In matches
                    If VfmLookupExpired(started) Then GoTo TimedOut
                    If contextAware Then
                        If VfiContains(exclusions, CStr(match(2))) Then GoTo NextMatch
                        If VfcExclude(CStr(chunk("Text")), CLng(match(0)), CLng(match(1)), CStr(match(3)), (region.StoryType <> wdFootnotesStory)) Then GoTo NextMatch
                    End If
                    If CLng(match(0)) + CLng(match(1)) > region.End - region.Start Then
                        VfmRepair state, region
                        Set result = VfmDictionary()
                        result("Retry") = True
                        Set VfmFindEarlier = result
                        Exit Function
                    End If
                    If CLng(match(0)) + CLng(match(1)) <= region.End - region.Start Then
                        Set named = region.Duplicate
                        named.SetRange region.Start + CLng(match(0)), region.Start + CLng(match(0)) + CLng(match(1))
                        If StrComp(named.Text, CStr(match(2)), vbBinaryCompare) <> 0 Then
                            VfmRepair state, region
                            Set result = VfmDictionary()
                            result("Retry") = True
                            Set VfmFindEarlier = result
                            Exit Function
                        End If
                        If named.InStory(target) Then
                            If named.Start < target.End And named.End > target.Start Then GoTo NextMatch
                        End If
                        If contextAware Then
                            If VfzReferenceAt(named) Then GoTo NextMatch
                            If VfcNeedsContext(CStr(chunk("Text")), CLng(match(0)), CLng(match(1)), CStr(match(3))) Then
                                If VfcRangeExcluded(named, (region.StoryType <> wdFootnotesStory)) Then GoTo NextMatch
                            End If
                        End If
                        If named.StoryType = wdMainTextStory Then
                            position = Array(named.Start, 0&, 0&)
                        Else
                            Set anchor = chunk("Anchor")
                            position = Array(anchor.Start, 1&, named.Start)
                        End If
                        If VfmEarlierPosition(position, selectedPosition) Then
                            If result Is Nothing Then
                                identity = VfmIdentity(state, named, VfpPersonKey(row))
                                If identity >= 0 Then
                                    Set result = VfmDictionary()
                                    Set result("Range") = named
                                    Set result("Chunk") = chunk
                                    result("Identity") = identity
                                    bestPosition = position
                                End If
                            ElseIf VfmEarlierPosition(position, bestPosition) Then
                                identity = VfmIdentity(state, named, VfpPersonKey(row))
                                If identity >= 0 Then
                                    Set result("Range") = named
                                    Set result("Chunk") = chunk
                                    result("Identity") = identity
                                    bestPosition = position
                                End If
                            End If
                        End If
                    End If
NextMatch:
                    If Not result Is Nothing Then
                        If CLng(bestPosition(0)) = 0 And CLng(bestPosition(1)) = 0 And CLng(bestPosition(2)) = 0 Then
                            Set VfmFindEarlier = result
                            Exit Function
                        End If
                    End If
                Next match
            End If
        End If
NextChunk:
    Next id
    Set VfmFindEarlier = result
    Exit Function
TimedOut:
    Set result = VfmDictionary()
    result("Timeout") = True
    Set VfmFindEarlier = result
End Function

Private Function VfmLookupExpired(ByVal started As Single) As Boolean
    Dim elapsed As Single
    elapsed = VfbUiClock() - started
    If elapsed < 0 Then elapsed = elapsed + 86400!
    VfmLookupExpired = (elapsed >= VfmLookupSeconds)
End Function

Private Function VfmCheckEarlier(ByVal doc As Word.Document, ByRef row As Variant, ByVal aliasText As String, ByVal named As Word.Range, Optional ByRef forceHere As Boolean = False, Optional ByVal contextAware As Boolean = False, Optional ByVal exclusions As Object = Nothing, Optional ByRef prompted As Boolean = False) As Boolean
    Dim state As Object, earlier As Object, region As Word.Range, chunk As Object, retry As Long, message As String, excerpt As String, offset As Long, block As Word.Range
    Dim decision As Long, started As Single, rebuilt As Boolean, fallbackKey As String
    forceHere = False: prompted = False
    Set state = VfmState(doc)
    If state Is Nothing Then
        Vefeyat0341UiWarmMemory doc
        Set state = VfmState(doc)
    End If
    If state Is Nothing Then
        VfaWarn VfText("Belge belle\u011fi haz\u0131rlanamad\u0131; tarih eklenmedi. Belle\u011fi yenile komutunu kullan\u0131n.")
        Exit Function
    End If
    If Not CBool(state("Ready")) Then
        VfaWarn VfText("Belge belle\u011fi g\u00fcncel de\u011fil. Belge belle\u011fini yenile komutunu kullan\u0131n; tarih eklenmedi.")
        Exit Function
    End If
    VfmObserve doc, named
    If state("Pending").Count > 0 Then Vefeyat0341UiWarmMemory doc, True
    If state("Pending").Count > 0 Or Not CBool(state("Ready")) Then
        VfaWarn VfText("Yeni metin hen\u00fcz belge belle\u011fine aktar\u0131lmad\u0131. Belle\u011fi yenile komutunu kullan\u0131n; tarih eklenmedi.")
        Exit Function
    End If
    started = VfbUiClock()
    For retry = 1 To 8
        Set earlier = VfmFindEarlier(state, row, aliasText, named, started, contextAware, exclusions)
        If earlier Is Nothing Then
            ' The memory follows the caret. Pasted or remotely edited text can be
            ' missing from it; Word's own Find decides whether to rebuild once.
            If Not rebuilt Then
                If VfmStaleEarlier(doc, state, row, aliasText, named) Then
                    rebuilt = True
                    fallbackKey = VfmFallbackKey(doc, row, named)
                    If Not VfmRebuild(doc, state) Then Exit Function
                    state("FallbackKey") = fallbackKey
                    started = VfbUiClock()
                    GoTo NextRetry
                End If
            End If
            VfmCheckEarlier = True
            Exit Function
        End If
        If earlier.Exists("Timeout") Then
            VfaWarn VfText("\u0130lk ge\u00e7i\u015f kontrol\u00fc s\u00fcre s\u0131n\u0131r\u0131na ula\u015ft\u0131. Tarih eklenmedi; belgeniz de\u011fi\u015ftirilmedi. Belle\u011fi yenileyip daha a\u00e7\u0131k bir adla yeniden deneyin.")
            Exit Function
        End If
        If earlier.Exists("Retry") Then GoTo NextRetry
        Set chunk = earlier("Chunk")
        Set region = chunk("Range")
        If StrComp(region.Text, CStr(chunk("Text")), vbBinaryCompare) = 0 Then Exit For
        VfmRepair state, region
NextRetry:
    Next retry
    If retry > 8 Then
        state("Ready") = False
        VfaWarn VfText("Belgede toplu bir de\u011fi\u015fiklik alg\u0131land\u0131. Belge belle\u011fini yenileyin; tarih eklenmedi.")
        Exit Function
    End If
    Set region = earlier("Range")
    message = VfText("Bu \u015fahs\u0131n ad\u0131na uyan daha \u00f6nceki bir ge\u00e7i\u015f var: ") & region.Text & vbCrLf & vbCrLf
    If CLng(earlier("Identity")) = 0 Then message = message & VfText("\u00d6nceki ad\u0131n ki\u015fi e\u015fle\u015ftirmesi hen\u00fcz do\u011frulanmad\u0131; ba\u011flam\u0131 kontrol edin.") & vbCrLf
    Set block = chunk("Range")
    offset = region.Start - block.Start - 60
    If offset < 0 Then offset = 0
    excerpt = Replace(Mid$(CStr(chunk("Text")), offset + 1, 180), vbCr, " ")
    message = message & vbCrLf & VfText("Ba\u011flam: ") & excerpt & vbCrLf & vbCrLf
    If VfmDatePresent(doc, region) Then
        message = message & VfText("Bu ge\u00e7i\u015fin yan\u0131nda zaten tarih bulunuyor.")
    Else
        message = message & VfText("Bu ge\u00e7i\u015fin yan\u0131nda tarih olmasa da ilk ge\u00e7i\u015f kural\u0131 ge\u00e7erlidir.")
    End If
    message = message & vbCrLf & vbCrLf & VfText("\u0130sterseniz ilk ge\u00e7i\u015f kural\u0131na istisna yaparak se\u00e7ili ada yine de tarih ekleyebilirsiniz.")
    prompted = True
    decision = Vefeyat0341UiEarlierChoice(message)
    If decision = 1 Then
        forceHere = True
        VfmCheckEarlier = True
    ElseIf decision = 2 Then
        region.Select
    End If
End Function

Private Sub VfmInsertFirst(ByVal doc As Word.Document, ByVal selected As Word.Range, ByVal original As String, _
                           ByVal aliasText As String, ByRef row As Variant, ByRef inserted As Long, ByRef skipped As Long, ByRef changed As Boolean, Optional ByVal forceHere As Boolean = False)
    Dim named As Word.Range, offset As Long, id As String
    Set named = selected.Duplicate
    offset = InStr(1, original, aliasText, vbBinaryCompare) - 1
    If offset < 0 Then Exit Sub
    named.SetRange selected.Start + offset, selected.Start + offset + Len(aliasText)
    If VfmDatePresent(doc, named) Or VfaAlreadyHasDate(named, CStr(row(3))) Then
        mVfaSkipReason = VfaDateBlockReason(named)
        skipped = skipped + 1
        Exit Sub
    End If
    VfaEnsureNamespace doc, changed
    id = VfmBinding(doc, row, aliasText, changed)
    If forceHere Then VfaWriteVariable doc, VfaVariableName(id, "Pinned"), "1"
    ' No Find, DeleteManaged or whole-document ApplyBinding on this fast path.
    VfaInsertManaged doc, named, VfpFormatDate(CStr(row(3)), VfpGetPrefix(doc)), id, changed, True
    inserted = inserted + 1
    named.SetRange selected.Start + offset, selected.Start + offset + Len(aliasText)
    VfmRemember doc, named, VfpPersonKey(row)
End Sub

Private Function VfmDiagnostic() As String
    Dim state As Object
    If Application.Documents.Count = 0 Then Exit Function
    Set state = VfmState(ActiveDocument)
    If state Is Nothing Then
        VfmDiagnostic = VfText("Belge belle\u011fi: hen\u00fcz haz\u0131rlanmad\u0131.")
    ElseIf state("Pending").Count > 0 Then
        VfmDiagnostic = VfText("Belge belle\u011fi: yeni metin i\u00e7in b\u00f6lgesel g\u00fcncelleme bekliyor.")
    ElseIf CBool(state("Ready")) Then
        VfmDiagnostic = VfText("Belge belle\u011fi: haz\u0131r; ") & CStr(state("Chunks").Count) & VfText(" metin b\u00f6l\u00fcm\u00fc RAM'de.")
    Else
        VfmDiagnostic = VfText("Belge belle\u011fi: yenileme gerekiyor.")
    End If
End Function

Private Sub VfmQueue(ByVal state As Object, ByVal area As Word.Range)
    Dim previous As Word.Range
    For Each previous In state("Pending")
        If previous.InStory(area) Then
            If previous.Start <= area.Start And previous.End >= area.End Then Exit Sub
        End If
    Next previous
    state("Pending").Add area.Duplicate
End Sub

Private Sub VfmFlushPending(ByVal state As Object)
    Dim area As Word.Range
    For Each area In state("Pending")
        VfmRepair state, area
    Next area
    Set state("Pending") = New Collection
End Sub

Private Sub VfmKnownAlias(ByVal state As Object, ByVal person As String, ByVal value As String)
    Dim aliases As Object
    If Not VfaValidAlias(value) Then Exit Sub
    If Not state("Aliases").Exists(person) Then
        Set aliases = VfmDictionary()
        state("Aliases").Add person, aliases
    End If
    Set aliases = state("Aliases")(person)
    aliases(value) = True
End Sub

Private Sub VfmLoadAliases(ByVal state As Object, ByVal doc As Word.Document)
    Dim variables As Object, variable As Word.Variable, name As Variant, id As String, count As Long
    Set variables = VfmDictionary()
    For Each variable In doc.Variables
        count = count + 1
        If count Mod 64 = 0 Then VfbPulse VfText("Belge ba\u011flar\u0131 okunuyor"), count, doc.Variables.Count
        If Left$(variable.Name, Len(VfaVariablePrefix)) = VfaVariablePrefix Then variables(variable.Name) = variable.Value
    Next variable
    state("VariableCount") = doc.Variables.Count
    Set state("Bindings") = VfmDictionary()
    For Each name In variables.Keys
        If Right$(CStr(name), 4) = ".Key" Then
            id = Left$(CStr(name), Len(CStr(name)) - 4)
            If variables.Exists(id & ".Alias") Then
                VfmKnownAlias state, CStr(variables(CStr(name))), CStr(variables(id & ".Alias"))
                If Not variables.Exists(id & ".OccurrenceOnly") Then
                    state("Bindings")(VfpKeyPart(VfaAliasKey(CStr(variables(id & ".Alias")))) & CStr(variables(CStr(name)))) = Mid$(id, Len(VfaVariablePrefix) + 1)
                ElseIf CStr(variables(id & ".OccurrenceOnly")) <> "1" Then
                    state("Bindings")(VfpKeyPart(VfaAliasKey(CStr(variables(id & ".Alias")))) & CStr(variables(CStr(name)))) = Mid$(id, Len(VfaVariablePrefix) + 1)
                End If
            End If
        End If
    Next name
End Sub

' Cache the metadata lookup too: the old binding helper enumerates every
' existing binding and rereads the variable collection for each of them.
Private Function VfmBinding(ByVal doc As Word.Document, ByRef row As Variant, ByVal aliasText As String, ByRef changed As Boolean) As String
    Dim state As Object, person As String, key As String, id As String, value As String, number As Long
    Set state = VfmState(doc)
    If state Is Nothing Then
        VfmBinding = VfaGetBinding(doc, row, aliasText, changed)
        Exit Function
    End If
    person = VfpPersonKey(row)
    key = VfpKeyPart(VfaAliasKey(aliasText)) & person
    If doc.Variables.Count <> CLng(state("VariableCount")) Or Not state("Bindings").Exists(key) Then VfmLoadAliases state, doc
    If state("Bindings").Exists(key) Then
        id = CStr(state("Bindings")(key))
        If VfaReadVariable(doc, VfaVariableName(id, "Key")) = person Then
            If VfaAliasKey(VfaReadVariable(doc, VfaVariableName(id, "Alias"))) = VfaAliasKey(aliasText) Then
                VfaWriteVariable doc, VfaVariableName(id, "Raw"), CStr(row(3))
                changed = True
                VfmBinding = id
                Exit Function
            End If
        End If
        state("Bindings").Remove key
    End If
    number = 1
    value = VfaReadVariable(doc, VfaNextVariable)
    If IsNumeric(value) Then
        If CDbl(value) >= 1 And CDbl(value) < 2000000000# Then number = CLng(value)
    End If
    Do
        id = "P" & CStr(number)
        number = number + 1
    Loop While VfaVariableExists(doc, VfaVariableName(id, "Key"))
    changed = True
    VfaWriteVariable doc, VfaNextVariable, CStr(number)
    VfaWriteVariable doc, VfaVariableName(id, "Alias"), aliasText
    VfaWriteVariable doc, VfaVariableName(id, "Raw"), CStr(row(3))
    VfaWriteVariable doc, VfaVariableName(id, "Key"), person
    state("Bindings")(key) = id
    state("VariableCount") = doc.Variables.Count
    VfmKnownAlias state, person, aliasText
    VfmBinding = id
End Function

Private Function VfmDateText(ByVal value As String) As Boolean
    Dim prefix As Variant, text As String, ch As String, at As Long, digits As Long
    text = LCase$(VfaTrimSpaces(value))
    For Each prefix In Array("(o.", VfText("(\u00f6."), VfText("(\u00f6l."), "(v.", "(d.")
        If Left$(text, Len(CStr(prefix))) = CStr(prefix) Then
            VfmDateText = True
            Exit Function
        End If
    Next prefix
    If Left$(text, 1) = "(" Then
        text = VfaTrimSpaces(Mid$(text, 2))
        at = 1
        Do While at <= Len(text)
            ch = Mid$(text, at, 1)
            If ch < "0" Or ch > "9" Then Exit Do
            digits = digits + 1: at = at + 1
        Loop
        ' A numbered comment or page locator is not a biographical date.
        If digits < 3 Or digits > 4 Then Exit Function
        text = VfaTrimSpaces(Mid$(text, at))
        ch = Left$(text, 1)
        VfmDateText = (Len(text) = 0 Or ch = ")" Or ch = "/" Or ch = "-" Or _
            ch = ChrW(&H2013) Or ch = ChrW(&H2014) Or ch = "?" Or ch = "[")
    End If
End Function

Private Function VfmDatePresent(ByVal doc As Word.Document, ByVal named As Word.Range) As Boolean
    Dim field As Word.Field, following As Word.Range
    For Each field In named.Paragraphs(1).Range.Fields
        If field.Code.Start - 1 >= named.End And field.Code.Start - 1 <= named.End + 2 Then
            If Len(VfaFieldBinding(doc, field)) > 0 Then
                VfmDatePresent = True
                Exit Function
            End If
        End If
    Next field
    Set following = named.Duplicate
    following.SetRange named.End, named.End
    Call following.MoveEnd(wdCharacter, 128)
    VfmDatePresent = VfmDateText(following.Text)
End Function

Private Function VfmProtectedPoint(ByVal point As Word.Range) As Boolean
    Dim control As Word.ContentControl
    If point.Information(wdInFieldCode) Or point.Information(wdInFieldResult) Then
        VfmProtectedPoint = True
        Exit Function
    End If
    Set control = point.ParentContentControl
    If Not control Is Nothing Then VfmProtectedPoint = True
End Function

' Paragraph guards must not enumerate all document variables once per field.
' Cache a case-folded snapshot; our writes and count changes invalidate it.
Private Sub VfmMetadataChanged(ByVal doc As Word.Document)
    Dim state As Object
    Set state = VfmState(doc)
    If Not state Is Nothing Then state("MetadataCount") = -1&
End Sub

Private Function VfmMetadata(ByVal doc As Word.Document) As Object
    Dim state As Object, values As Object, variable As Word.Variable
    Set state = VfmState(doc)
    If Not state Is Nothing Then
        If state.Exists("Metadata") Then
            If CLng(state("MetadataCount")) = doc.Variables.Count Then
                Set VfmMetadata = state("Metadata")
                Exit Function
            End If
        End If
    End If
    Set values = VfmDictionary()
    For Each variable In doc.Variables
        values(LCase$(variable.Name)) = variable.Value
    Next variable
    If Not state Is Nothing Then
        Set state("Metadata") = values
        state("MetadataCount") = doc.Variables.Count
    End If
    Set VfmMetadata = values
End Function

Private Function VfmMetaValue(ByVal values As Object, ByVal name As String) As String
    name = LCase$(name)
    If values.Exists(name) Then VfmMetaValue = CStr(values(name))
End Function

' Selection validation must not hide a global exclusion/Find pass behind the
' first-mode cache. Check only the selected paragraph and its owning objects.
Private Function VfmTargetExcluded(ByVal doc As Word.Document, ByVal target As Word.Range, ByVal aliasText As String) As Boolean
    Dim point As Word.Range, paragraph As Word.Range, field As Word.Field, control As Word.ContentControl
    Dim id As String, value As String, region As Word.Range, start As Long, finish As Long
    Dim metadata As Object
    Set point = target.Duplicate
    point.SetRange target.Start, target.Start
    If VfmProtectedPoint(point) Then
        VfmTargetExcluded = True
        Exit Function
    End If
    If target.End > target.Start Then
        point.SetRange target.End - 1, target.End - 1
        If VfmProtectedPoint(point) Then
            VfmTargetExcluded = True
            Exit Function
        End If
    End If
    Set paragraph = target.Paragraphs(1).Range
    For Each control In paragraph.ContentControls
        If target.Start < control.Range.End And target.End > control.Range.Start Then
            VfmTargetExcluded = True
            Exit Function
        End If
    Next control
    For Each field In paragraph.Fields
        start = field.Code.Start - 1: finish = VfdFieldEnd(field)
        If target.Start < finish And target.End > start Then
            VfmTargetExcluded = True
            Exit Function
        End If
        If start < target.Start Or start - target.End > VfaMaxName Then GoTo NextTargetField
        If metadata Is Nothing Then Set metadata = VfmMetadata(doc)
        id = VfaBindingFromFieldCode(field.Code.Text, VfmMetaValue(metadata, VfaNamespaceVariable))
        If Len(id) > 0 Then
            value = VfmMetaValue(metadata, VfaVariableName(id, "Alias"))
            If Len(value) > Len(aliasText) And start - Len(value) >= paragraph.Start Then
                Set region = paragraph.Duplicate
                region.SetRange start - Len(value), start
                If StrComp(region.Text, value, vbTextCompare) = 0 Then
                    If target.Start < region.End And target.End > region.Start Then
                        VfmTargetExcluded = True
                        Exit Function
                    End If
                End If
            End If
        End If
NextTargetField:
    Next field
End Function

Private Function VfwOpenDocument() As Boolean
    Dim doc As Word.Document
    If mVfwDoc Is Nothing Then Exit Function
    For Each doc In Documents
        If doc Is mVfwDoc Then
            VfwOpenDocument = True
            Exit Function
        End If
    Next doc
End Function

Private Sub VfwBumpSession()
    If mVfwSession >= 2000000000 Then mVfwSession = 0
    mVfwSession = mVfwSession + 1
End Sub

Private Sub VfwPrepare()
    Dim heads As Object, cache As Object, snapshots As New Collection, state As Object
    mVfwPrepared = False
    VfwBumpSession
    If Not VfrEditableDocument() Then Exit Sub
    If Not VefeyatEnsureData() Then Exit Sub
    Set mVfwDoc = ActiveDocument
    mVfwPaused = False: mVfwPausing = False: mVfwFaulted = False
    Set mVfwExclusions = VfiRead(mVfwDoc)
    mVfwExclusionStamp = VfiPayload(mVfwExclusions)
    mVfwSkipDated = (VfpGetAuditMode(mVfwDoc) = "audit_undated")
    mVfwFirstOnly = (VfpGetScope(mVfwDoc) = "first")
    Set mVfwCatalog = mVfRows
    Set mVfwPatterns = CreateObject("Scripting.Dictionary")
    Set heads = CreateObject("Scripting.Dictionary")
    Set cache = CreateObject("Scripting.Dictionary")
    VfxPrepareDocument mVfwDoc, mVfwPatterns, heads, cache
    VfzBegin mVfwDoc, mVfwPatterns, heads, cache
    Set mVfwHits = VfbOrderedHits(mVfwDoc, mVfwPatterns, heads, cache, snapshots, True)
    VfxVerifySnapshots snapshots, cache
    If mVfwFirstOnly Then
        ' One preparation, never one full-document scan per review step.
        VfmWarm mVfwDoc
    End If
    VfwLoadVariables
    mVfwIndex = 1
    mVfwAdded = 0
    mVfwSkipped = 0
    Set mVfwSkippedIndices = CreateObject("Scripting.Dictionary")
    Set mVfwAppliedIndices = CreateObject("Scripting.Dictionary")
    Set mVfwChosenPeople = CreateObject("Scripting.Dictionary")
    Set mVfwConfirmed = CreateObject("Scripting.Dictionary")
    Set mVfwCandidateCache = CreateObject("Scripting.Dictionary")
    Set mVfwVisited = CreateObject("Scripting.Dictionary")
    Set mVfwAutomatic = CreateObject("Scripting.Dictionary")
    mVfwFirstVisited = 0: mVfwDatedSkipped = 0: mVfwRepeatedSkipped = 0: mVfwIgnoredSkipped = 0
    mVfwPrepared = True
End Sub

Private Sub VfwLoadVariables()
    Dim variable As Word.Variable
    Set mVfwVariables = CreateObject("Scripting.Dictionary")
    mVfwVariables.CompareMode = vbTextCompare
    For Each variable In mVfwDoc.Variables
        mVfwVariables(variable.Name) = variable.Value
    Next variable
    mVfwVariableCount = mVfwDoc.Variables.Count
End Sub

Private Sub VfwWriteVariable(ByVal name As String, ByVal value As String)
    VfmMetadataChanged mVfwDoc
    If mVfwVariables.Exists(name) Then
        mVfwDoc.Variables.Item(name).Value = value
    Else
        mVfwDoc.Variables.Add Name:=name, Value:=value
    End If
    mVfwVariables(name) = value
    mVfwVariableCount = mVfwDoc.Variables.Count
End Sub

' Never reuse a spelling-wide binding for an explicitly reviewed occurrence.
Private Function VfwBinding(ByRef row As Variant, ByVal aliasText As String) As String
    Dim number As Long, id As String, value As String
    If mVfwVariableCount <> mVfwDoc.Variables.Count Then VfwLoadVariables
    number = 1
    If mVfwVariables.Exists(VfaNextVariable) Then
        value = CStr(mVfwVariables(VfaNextVariable))
        If IsNumeric(value) Then
            If CDbl(value) >= 1 And CDbl(value) < 2000000000# Then number = CLng(value)
        End If
    End If
    Do
        id = "P" & CStr(number)
        number = number + 1
    Loop While mVfwVariables.Exists(VfaVariableName(id, "Key"))
    VfwWriteVariable VfaNextVariable, CStr(number)
    VfwWriteVariable VfaVariableName(id, "Alias"), aliasText
    VfwWriteVariable VfaVariableName(id, "Raw"), CStr(row(3))
    VfwWriteVariable VfaVariableName(id, "OccurrenceOnly"), "1"
    VfwWriteVariable VfaVariableName(id, "Key"), VfpPersonKey(row)
    VfwBinding = id
End Function

Private Sub VfwReserveSpace()
    On Error GoTo NoSpace
    mVfwState = Application.WindowState
    mVfwLeft = Application.Left: mVfwTop = Application.Top
    mVfwWidth = Application.Width: mVfwHeight = Application.Height
    mVfwLayout = False
    If mVfwWidth < 780 Then Exit Sub
    Application.WindowState = wdWindowStateNormal
    Application.Left = mVfwLeft: Application.Top = mVfwTop
    Application.Height = mVfwHeight
    mVfwReserved = mVfwWidth - 350
    Application.Width = mVfwReserved
    mVfwLayout = True
    mVfwReserved = Application.Width
    mVfwAppliedLeft = Application.Left: mVfwAppliedTop = Application.Top: mVfwAppliedHeight = Application.Height
    Exit Sub
NoSpace:
    On Error Resume Next
    If mVfwWidth > 0 Then
        Application.Left = mVfwLeft: Application.Top = mVfwTop
        Application.Width = mVfwWidth: Application.Height = mVfwHeight
        Application.WindowState = mVfwState
    End If
    mVfwLayout = False
End Sub

Private Sub VfwPosition(ByVal panel As VefeyatKisiSec0341)
    Dim left As Single, top As Single
    left = Application.Left + Application.Width - 344
    If mVfwLayout Then left = Application.Left + Application.Width + 6
    top = Application.Top + Application.Height - 500
    If top < Application.Top + 8 Then top = Application.Top + 8
    If top > Application.Top + 80 Then top = Application.Top + 80
    panel.StartUpPosition = 0
    panel.Left = left: panel.Top = top
    If Application.Height > 300 And panel.Height > Application.Height - 16 Then panel.Height = Application.Height - 16
End Sub

Private Sub VfwOpen()
    If Not mVfwPrepared Then Exit Sub
    If Not VfwOpenDocument() Then
        VfwClose
        Exit Sub
    End If
    Set mVfwPanel = New VefeyatKisiSec0341
    mVfwPanel.InitLiveReview
    VfwReserveSpace
    VfwPosition mVfwPanel
    mVfwPanel.Show vbModeless
    VfwShowCurrent
End Sub

Private Function VfwNameAt(ByVal index As Long) As Word.Range
    Dim hit As Object, named As Word.Range
    On Error GoTo Unavailable
    If Not VfwOpenDocument() Then Exit Function
    If index < 1 Or index > mVfwHits.Count Then Exit Function
    Set hit = mVfwHits(index)
    Set named = hit("Range")
    If StrComp(named.Text, CStr(hit("Alias")), vbBinaryCompare) <> 0 Then Exit Function
    If VfzReferenceAt(named) Then Exit Function
    If Not VfaSupportedRange(mVfwDoc, named) Then Exit Function
    Set VfwNameAt = named
Unavailable:
End Function

Private Function VfwCurrentName() As Word.Range
    Set VfwCurrentName = VfwNameAt(mVfwIndex)
End Function

Private Function VfwRowsFor(ByVal hit As Object) As Collection
    Dim key As String, rows As Collection, query As String, tick As Single
    query = VfbBareAlias(CStr(hit("Alias")))
    key = VfpKeyPart(CStr(hit("Key"))) & VfpKeyPart(query)
    If mVfwCandidateCache.Exists(key) Then
        Set VfwRowsFor = mVfwCandidateCache(key)
        Exit Function
    End If
    tick = VfgClock()
    Set rows = VfbRankRows(mVfwPatterns(CStr(hit("Key"))), query)
    VfgEnd "candidates", tick
    If mVfwCandidateCache.Count >= 256 Then Set mVfwCandidateCache = CreateObject("Scripting.Dictionary")
    Set mVfwCandidateCache(key) = rows
    Set VfwRowsFor = rows
End Function

Private Function VfwOwnedPerson(ByVal named As Word.Range) As String
    Dim field As Word.Field, bindingNamespace As String, id As String, key As String
    On Error GoTo Unavailable
    If mVfwVariableCount <> mVfwDoc.Variables.Count Then VfwLoadVariables
    If Not mVfwVariables.Exists(VfaNamespaceVariable) Then Exit Function
    bindingNamespace = mVfwDoc.Variables.Item(VfaNamespaceVariable).Value
    For Each field In named.Paragraphs(1).Range.Fields
        If field.Code.Start - 1 >= named.End And field.Code.Start - 1 <= named.End + 2 Then
            id = VfaBindingFromFieldCode(field.Code.Text, bindingNamespace)
            If Len(id) > 0 Then
                If Not VfaValidDate(VfaTrimSpaces(field.Result.Text)) Then Exit Function
                key = VfaVariableName(id, "Key")
                If mVfwVariables.Exists(key) Then VfwOwnedPerson = mVfwDoc.Variables.Item(key).Value
                Exit Function
            End If
        End If
    Next field
Unavailable:
End Function

Private Function VfwHasDate(ByVal named As Word.Range, ByVal ownedPerson As String) As Boolean
    Dim following As Word.Range, text As String, prefix As Variant, closingAt As Long, first As String, finish As Long
    If Len(ownedPerson) > 0 Then
        VfwHasDate = True
        Exit Function
    End If
    Set following = named.Duplicate
    finish = named.End + 100
    If finish > named.Paragraphs(1).Range.End Then finish = named.Paragraphs(1).Range.End
    following.SetRange named.End, finish
    following.TextRetrievalMode.IncludeFieldCodes = False
    text = LCase$(VfaTrimSpaces(following.Text))
    For Each prefix In Array("(o.", VfText("(\u00f6."), VfText("(\u00f6l."), "(v.")
        If Left$(text, Len(CStr(prefix))) = CStr(prefix) Then
            VfwHasDate = True
            Exit Function
        End If
    Next prefix
    If Left$(text, 1) <> "(" Then Exit Function
    closingAt = InStr(text, ")")
    If closingAt < 3 Then Exit Function
    text = VfaTrimSpaces(Mid$(text, 2, closingAt - 2))
    first = Left$(text, 1)
    If first >= "0" And first <= "9" And InStr(text, "/") > 0 Then VfwHasDate = True
End Function

Private Function VfwHitPosition(ByVal index As Long, ByVal named As Word.Range) As Variant
    Dim hit As Object, anchor As Word.Range, state As Object
    If named.StoryType = wdMainTextStory Then
        VfwHitPosition = Array(named.Start, 0&, 0&)
    Else
        Set hit = mVfwHits(index)
        If hit.Exists("Anchor") Then
            Set anchor = hit("Anchor")
            VfwHitPosition = Array(anchor.Start, 1&, named.Start)
        Else
            Set state = VfmState(mVfwDoc)
            If state Is Nothing Then Err.Raise VfaInvalidData, "Vefeyat", "Footnote review order is unavailable."
            VfwHitPosition = VfmPosition(state, named)
        End If
    End If
End Function

Private Function VfwBefore(ByVal first As Long, ByVal second As Long) As Boolean
    Dim a As Word.Range, b As Word.Range, ap As Variant, bp As Variant
    On Error GoTo Unavailable
    If first = second Then Exit Function
    Set a = VfwNameAt(first): Set b = VfwNameAt(second)
    If a Is Nothing Or b Is Nothing Then Exit Function
    ap = VfwHitPosition(first, a): bp = VfwHitPosition(second, b)
    VfwBefore = VfmEarlierPosition(ap, bp)
Unavailable:
End Function

Private Function VfwProof(ByVal index As Long, ByVal person As String) As Boolean
    Dim named As Word.Range, owned As String
    Set named = VfwNameAt(index)
    If named Is Nothing Then Exit Function
    owned = VfwOwnedPerson(named)
    If Len(owned) > 0 Then
        VfwProof = (owned = person)
        Exit Function
    End If
    ' A cached picker choice is not live proof after its date was removed.
End Function

Private Sub VfwConfirm(ByVal person As String, ByVal index As Long)
    Dim old As Long
    If Len(person) = 0 Then Exit Sub
    If mVfwConfirmed.Exists(person) Then
        old = CLng(mVfwConfirmed(person))
        If VfwProof(old, person) Then
            If Not VfwBefore(index, old) Then Exit Sub
        End If
    End If
    mVfwConfirmed(person) = index
End Sub

Private Function VfwConfirmedBefore(ByVal person As String, ByVal index As Long) As Boolean
    Dim old As Long
    If Not mVfwConfirmed.Exists(person) Then Exit Function
    old = CLng(mVfwConfirmed(person))
    If Not VfwProof(old, person) Then
        mVfwConfirmed.Remove person
        Exit Function
    End If
    VfwConfirmedBefore = VfwBefore(old, index)
End Function

Private Sub VfwMarkAutomatic(ByVal reason As String)
    Dim key As String, old As String
    key = CStr(mVfwIndex)
    If mVfwAppliedIndices.Exists(key) Then Exit Sub
    If mVfwAutomatic.Exists(key) Then old = CStr(mVfwAutomatic(key))
    If old = reason Then Exit Sub
    If old = "dated" Then mVfwDatedSkipped = mVfwDatedSkipped - 1
    If old = "repeat" Then mVfwRepeatedSkipped = mVfwRepeatedSkipped - 1
    If old = "ignored" Then mVfwIgnoredSkipped = mVfwIgnoredSkipped - 1
    mVfwAutomatic(key) = reason
    If reason = "dated" Then mVfwDatedSkipped = mVfwDatedSkipped + 1
    If reason = "repeat" Then mVfwRepeatedSkipped = mVfwRepeatedSkipped + 1
    If reason = "ignored" Then mVfwIgnoredSkipped = mVfwIgnoredSkipped + 1
    If mVfwSkippedIndices.Exists(key) Then mVfwSkippedIndices.Remove key
    mVfwSkipped = mVfwSkippedIndices.Count
End Sub

Private Function VfwCanBack() As Boolean
    If mVfwFirstVisited > 0 Then VfwCanBack = (mVfwIndex > mVfwFirstVisited)
End Function

Private Sub VfwRemoveAutomatic(ByVal index As Long)
    Dim key As String, reason As String
    key = CStr(index)
    If Not mVfwAutomatic.Exists(key) Then Exit Sub
    reason = CStr(mVfwAutomatic(key))
    If reason = "dated" Then mVfwDatedSkipped = mVfwDatedSkipped - 1
    If reason = "repeat" Then mVfwRepeatedSkipped = mVfwRepeatedSkipped - 1
    If reason = "ignored" Then mVfwIgnoredSkipped = mVfwIgnoredSkipped - 1
    mVfwAutomatic.Remove key
End Sub

Private Sub VfwAdvanceAutomatic()
    Dim named As Word.Range, hit As Object, rows As Collection, row As Variant, person As String, reason As String, advanced As Long
    Dim session As Long
    session = mVfwSession
    Do While mVfwIndex <= mVfwHits.Count
        If session <> mVfwSession Then Exit Sub
        Set named = VfwCurrentName()
        If named Is Nothing Then Exit Sub
        reason = ""
        person = ""
        If VfiContains(mVfwExclusions, named.Text) Then reason = "ignored"
        If Len(reason) = 0 Then
            person = VfwOwnedPerson(named)
            If Len(person) > 0 Then VfwConfirm person, mVfwIndex
        End If
        If mVfwSkipDated And Len(reason) = 0 Then
            If VfwHasDate(named, person) Then reason = "dated"
        End If
        If Len(reason) = 0 And mVfwFirstOnly Then
            If Len(person) = 0 And mVfwChosenPeople.Exists(CStr(mVfwIndex)) Then person = CStr(mVfwChosenPeople(CStr(mVfwIndex)))
            If Len(person) = 0 Then
                Set hit = mVfwHits(mVfwIndex)
                Set rows = VfwRowsFor(hit)
                If rows.Count = 1 Then
                    row = rows(1)
                    person = VfpPersonKey(row)
                End If
            End If
            If Len(person) > 0 Then
                If VfwConfirmedBefore(person, mVfwIndex) Then reason = "repeat"
            End If
        End If
        If Len(reason) = 0 Then Exit Sub
        VfwMarkAutomatic reason
        mVfwIndex = mVfwIndex + 1
        advanced = advanced + 1
        If advanced Mod 32 = 0 Then
            mVfwPanel.LiveReviewStatus False, VfText("Do\u011frulanm\u0131\u015f tekrarlar ve se\u00e7ime g\u00f6re tarihli ge\u00e7i\u015fler atlan\u0131yor..."), False
            DoEvents
            If session <> mVfwSession Then Exit Sub
            If mVfwPanel Is Nothing Then Exit Sub
            VfzNextPass
            If Not VfwOpenDocument() Then
                VfwClose
                Exit Sub
            End If
        End If
    Loop
End Sub

Private Function VfwContextRange(ByVal named As Word.Range) As Word.Range
    Dim result As Word.Range, paragraph As Word.Range, first As Long, last As Long
    Set result = named.Duplicate
    Set paragraph = named.Paragraphs(1).Range
    first = named.Start - 90: last = named.End + 150
    If first < paragraph.Start Then first = paragraph.Start
    If last > paragraph.End Then last = paragraph.End
    result.SetRange first, last
    Set VfwContextRange = result
End Function

Private Sub VfwShowCurrent(Optional ByVal browse As Boolean = False)
    Dim hit As Object, named As Word.Range, heading As String, page As Long, choice As Long, i As Long, row As Variant, chosenPerson As String, identityProof As Object
    On Error GoTo Failed
    If mVfwPanel Is Nothing Or mVfwRendering Then Exit Sub
    If mVfwFaulted Then Exit Sub
    Dim session As Long
    session = mVfwSession
    VfzNextPass
    If Not VfwOpenDocument() Then
        VfwClose
        Exit Sub
    End If
    mVfwRendering = True
    Set mVfwContext = Nothing
    mVfwContextText = ""
    If Not browse Then VfwAdvanceAutomatic
    If session <> mVfwSession Then Exit Sub
    If mVfwPanel Is Nothing Then GoTo Finished
    If mVfwIndex > mVfwHits.Count Then
        mVfwDoc.Activate
        Set named = mVfwDoc.Content.Duplicate
        named.Collapse wdCollapseEnd
        named.Select
        Application.ActiveWindow.ScrollIntoView named, True
        mVfwPanel.LiveReviewDone CStr(mVfwAdded) & VfText(" ekleme i\u015flemi yap\u0131ld\u0131. ") & _
            CStr(mVfwSkipped) & VfText(" ge\u00e7i\u015f elle atland\u0131. ") & _
            CStr(mVfwDatedSkipped) & VfText(" tarihli ge\u00e7i\u015f, ") & CStr(mVfwRepeatedSkipped) & _
            VfText(" do\u011frulanm\u0131\u015f tekrar, ") & CStr(mVfwIgnoredSkipped) & _
            VfText(" belge istisnas\u0131 otomatik atland\u0131. \u0130nceleme tamamland\u0131.")
        mVfwPanel.LiveReviewStatus False, "", VfwCanBack()
        GoTo Finished
    End If
    Set hit = mVfwHits(mVfwIndex)
    Set named = VfwCurrentName()
    Set mVfwRows = VfwRowsFor(hit)
    mVfwVisited(CStr(mVfwIndex)) = True
    If mVfwFirstVisited = 0 Or mVfwIndex < mVfwFirstVisited Then mVfwFirstVisited = mVfwIndex
    choice = 1
    If mVfwChosenPeople.Exists(CStr(mVfwIndex)) Then
        chosenPerson = CStr(mVfwChosenPeople(CStr(mVfwIndex)))
    ElseIf Not named Is Nothing Then
        chosenPerson = VfwOwnedPerson(named)
    End If
    If Len(chosenPerson) > 0 Then
        For i = 1 To mVfwRows.Count
            row = mVfwRows(i)
            If VfpPersonKey(row) = chosenPerson Then choice = i
        Next i
    End If
    heading = mVfwDoc.Name & " | " & CStr(mVfwIndex) & " / " & CStr(mVfwHits.Count) & vbCrLf & CStr(hit("Alias"))
    If named Is Nothing Then
        mVfwPanel.ShowLiveReview mVfwRows, heading, VfText("Bu ad silindi veya de\u011fi\u015fti. Tarih eklenmez. Atlay\u0131n; yeni adlar i\u00e7in incelemeyi yeniden ba\u015flat\u0131n.")
        mVfwPanel.LiveReviewStatus False, VfText("De\u011fi\u015fen ge\u00e7i\u015f"), VfwCanBack()
        GoTo Finished
    End If
    mVfwDoc.Activate
    named.Select
    Application.ActiveWindow.ScrollIntoView named, True
    If session <> mVfwSession Then Exit Sub
    On Error Resume Next
    page = named.Information(wdActiveEndAdjustedPageNumber)
    On Error GoTo Failed
    If page > 0 Then heading = heading & VfText(" | Sayfa ") & CStr(page)
    Set mVfwContext = VfwContextRange(named)
    mVfwContextText = mVfwContext.Text
    Set identityProof = VdiBuild(mVfwRows, CStr(hit("Alias")), named)
    Set mVfwRows = identityProof("Rows")
    If Len(chosenPerson) > 0 Then
        For i = 1 To mVfwRows.Count
            row = mVfwRows(i): If VfpPersonKey(row) = chosenPerson Then choice = i
        Next i
    End If
    mVfwPanel.ShowLiveReview mVfwRows, heading, mVfwContextText, choice
    mVfwPanel.SetIdentityEvidence identityProof
Finished:
    If session <> mVfwSession Then Exit Sub
    mVfwRendering = False
    VfwChoiceChanged
    Exit Sub
Failed:
    If session <> mVfwSession Then Exit Sub
    mVfwRendering = False
    Set mVfwContext = Nothing
    If Not mVfwPanel Is Nothing Then mVfwPanel.LiveReviewStatus False, VfText("\u0130nceleme g\u00f6r\u00fcnt\u00fclenemedi; tarih eklenmez. \u0130ncelemeyi yeniden ba\u015flat\u0131n."), VfwCanBack()
    VfrReport Err.Description
End Sub

Private Sub VfwChoiceChanged()
    Dim named As Word.Range, canApply As Boolean, text As String, choice As Long, row As Variant
    If mVfwRendering Or mVfwPanel Is Nothing Then Exit Sub
    If Not mVfwDoc Is Nothing Then
        If Not VfpDateAllowed(mVfwDoc) Then
            mVfwPanel.LiveReviewStatus False, "Review only", VfwCanBack()
            Exit Sub
        End If
    End If
    If mVfwFaulted Then
        mVfwPanel.LiveReviewStatus False, VfText("Hata sonras\u0131 inceleme durduruldu. Bitir ile kapat\u0131n; belgeyi kontrol edip yeni denetim ba\u015flat\u0131n."), False, False
        Exit Sub
    End If
    If mVfwIndex > mVfwHits.Count Then Exit Sub
    Set named = VfwCurrentName()
    text = VfText("Yaln\u0131z bu ge\u00e7i\u015fe eklenir. Yeni yaz\u0131lan adlar i\u00e7in incelemeyi yeniden ba\u015flat\u0131n.")
    If Not named Is Nothing Then
        choice = mVfwPanel.LiveReviewChoice
        canApply = VfaEditable(mVfwDoc) And choice > 0 And choice <= mVfwRows.Count
        If Not VfaEditable(mVfwDoc) Then
            canApply = False
            text = VfText("Belge salt okunur veya korumal\u0131; tarih eklenmez.")
        ElseIf VfiContains(mVfwExclusions, named.Text) Then
            canApply = False
            text = VfText("Bu ifade belge istisnas\u0131nda; toplu ekleme yap\u0131lmaz. \u0130stisnalar\u0131 temizleyip yeni denetim ba\u015flatabilirsiniz.")
        ElseIf VfbAnyDate(named) Then
            canApply = False
            text = VfText("Bu ge\u00e7i\u015fte tarih veya alan var; korunuyor. Sonraki ge\u00e7i\u015fe ilerleyebilirsiniz.")
        ElseIf mVfwFirstOnly And canApply Then
            row = mVfwRows(choice)
            If VfwConfirmedBefore(VfpPersonKey(row), mVfwIndex) Then text = VfText("Se\u00e7ili \u015fah\u0131s daha \u00f6nce do\u011fruland\u0131. \u0130lk ge\u00e7i\u015f ayar\u0131 gere\u011fi onay\u0131n\u0131zla tarih eklenmeden sonraki ge\u00e7i\u015fe ilerlenir.")
        End If
    Else
        text = VfText("Ad de\u011fi\u015fti veya silindi; tarih eklenmez. Atlay\u0131n veya yeniden taray\u0131n.")
    End If
    mVfwPanel.LiveReviewStatus canApply, text, VfwCanBack()
End Sub

Private Sub VfwApply(ByVal choice As Long)
    Dim named As Word.Range, row As Variant, aliasText As String, id As String, forceHere As Boolean
    Dim undoObject As Object, started As Boolean, changed As Boolean, detail As String, first As Long, tick As Single, number As Long, fieldTick As Single
    Dim session As Long, owner As Word.Document, catalog As Collection, prompted As Boolean
    Dim approvedContext As Word.Range, approvedText As String, blockedReason As String
    If mVfwBusy Or mVfwRendering Or mVfwPanel Is Nothing Then Exit Sub
    If mVfwFaulted Then Exit Sub
    session = mVfwSession: Set owner = mVfwDoc: Set catalog = mVfwCatalog
    tick = VfgClock()
    VfzNextPass
    mVfwBusy = True
    On Error GoTo Failed
    If Not VfwOpenDocument() Then
        VfwClose
        GoTo Clean
    End If
    If Not (mVfwCatalog Is mVfRows) Then Err.Raise VfaInvalidData, "Vefeyat", "Catalog changed; restart review."
    If choice < 1 Or choice > mVfwRows.Count Then GoTo Clean
    If Not (ActiveDocument Is mVfwDoc) Then
        VfwShowCurrent
        GoTo Clean
    End If
    Set named = VfwCurrentName()
    If named Is Nothing Then GoTo ChangedContext
    If VfiContains(mVfwExclusions, named.Text) Then GoTo Clean
    If mVfwContext Is Nothing Then GoTo ChangedContext
    Set approvedContext = mVfwContext.Duplicate
    approvedText = mVfwContextText
    If StrComp(approvedContext.Text, approvedText, vbBinaryCompare) <> 0 Then GoTo ChangedContext
    If Not VfaEditable(mVfwDoc) Or Not VfpDateAllowed(mVfwDoc) Then GoTo Clean
    named.Select
    Application.ActiveWindow.ScrollIntoView named, True
    If session <> mVfwSession Then GoTo Clean
    aliasText = named.Text
    first = named.Start
    If VfmTargetExcluded(mVfwDoc, named, aliasText) Then GoTo ChangedContext
    row = mVfwRows(choice)
    If Not VfaValidDate(CStr(row(3))) Then GoTo Clean
    If VfbAnyDate(named) Or VfaAlreadyHasDate(named, CStr(row(3))) Then GoTo DateBlocked
    If mVfwFirstOnly Then
        If VfwConfirmedBefore(VfpPersonKey(row), mVfwIndex) Then
            If mVfwRows.Count > 1 Then
                VdiCheckStore named, row
                VdiStore named, row
            End If
            mVfwChosenPeople(CStr(mVfwIndex)) = VfpPersonKey(row)
            VfmRemember mVfwDoc, named, VfpPersonKey(row)
            VfwMarkAutomatic "repeat"
            mVfwIndex = mVfwIndex + 1
            VfwShowCurrent
            GoTo Clean
        End If
        ' No UI yield while this synchronous local-memory catchup runs.
        mVfbRunning = False
        VfmObserve mVfwDoc, named
        mVfbRunning = True
        If Not VfmCheckEarlier(mVfwDoc, row, aliasText, named, forceHere, True, mVfwExclusions, prompted) Then GoTo Clean
    End If
    ' The warning can navigate or an external edit can change the live context.
    If session <> mVfwSession Then GoTo Clean
    If Not (mVfwDoc Is owner) Then GoTo Clean
    If Not (named.Document Is owner) Then GoTo Clean
    If Not (mVfRows Is catalog) Then Err.Raise VfaInvalidData, "Vefeyat", "Catalog changed during confirmation."
    If Not VfaEditable(owner) Or Not VfpDateAllowed(owner) Then
        VfwChoiceChanged
        GoTo Clean
    End If
    If prompted Then
        If (VfpGetScope(owner) = "first") <> mVfwFirstOnly Then Err.Raise VfaInvalidData, "Vefeyat", "Scope changed during confirmation."
        If (VfpGetAuditMode(owner) = "audit_undated") <> mVfwSkipDated Then Err.Raise VfaInvalidData, "Vefeyat", "Audit mode changed during confirmation."
        If VfiPayload(VfiRead(owner)) <> mVfwExclusionStamp Then Err.Raise VfaInvalidData, "Vefeyat", "Exclusions changed during confirmation."
        VfwLoadVariables
    End If
    If Not (ActiveDocument Is mVfwDoc) Then GoTo ChangedContext
    If StrComp(named.Text, aliasText, vbBinaryCompare) <> 0 Then GoTo ChangedContext
    If mVfwContext Is Nothing Then GoTo ChangedContext
    If StrComp(approvedContext.Text, approvedText, vbBinaryCompare) <> 0 Then GoTo ChangedContext
    If StrComp(mVfwContextText, approvedText, vbBinaryCompare) <> 0 Then GoTo ChangedContext
    If VfmTargetExcluded(mVfwDoc, named, aliasText) Then GoTo ChangedContext
    If VfbAnyDate(named) Then GoTo DateBlocked
    If Not VdiRowCurrent(row) Then GoTo ChangedContext
    If mVfwRows.Count > 1 Then VdiCheckStore named, row
    started = VfaBeginUndo(undoObject)
    VfaEnsureNamespace mVfwDoc, changed
    id = VfwBinding(row, aliasText)
    changed = True
    fieldTick = VfgClock()
    VfaInsertManaged mVfwDoc, named, VfpFormatDate(CStr(row(3)), VfpGetPrefix(mVfwDoc)), id, changed
    VfgEnd "field", fieldTick
    VftLog mVfwDoc, VfText("\u0130ncelemede tarih"), named, aliasText, aliasText & " " & CStr(row(3))
    named.SetRange first, first + Len(aliasText)
    VfmRemember mVfwDoc, named, VfpPersonKey(row)
    If mVfwRows.Count > 1 Then VdiStore named, row
    VfaEndUndo undoObject, started
    mVfwAdded = mVfwAdded + 1
    VfwRemoveAutomatic mVfwIndex
    mVfwAppliedIndices(CStr(mVfwIndex)) = True
    mVfwChosenPeople(CStr(mVfwIndex)) = VfpPersonKey(row)
    VfwConfirm VfpPersonKey(row), mVfwIndex
    If mVfwSkippedIndices.Exists(CStr(mVfwIndex)) Then
        mVfwSkippedIndices.Remove CStr(mVfwIndex)
        mVfwSkipped = mVfwSkippedIndices.Count
    End If
    mVfwIndex = mVfwIndex + 1
    VfwShowCurrent
    GoTo Clean
ChangedContext:
    VfwShowCurrent
    If mVfwPanel Is Nothing Then GoTo Clean
    mVfwPanel.LiveReviewStatus False, VfText("Metin de\u011fi\u015fti. Ba\u011flam\u0131 yeniden kontrol edin; Metinde g\u00f6ster ile yenileyip tekrar onaylay\u0131n."), VfwCanBack()
    GoTo Clean
DateBlocked:
    blockedReason = VfaDateBlockReason(named)
    VfwShowCurrent
    If mVfwPanel Is Nothing Then GoTo Clean
    mVfwPanel.LiveReviewStatus False, blockedReason, VfwCanBack()
    GoTo Clean
Failed:
    number = Err.Number
    detail = Err.Description
    On Error Resume Next
    VfaEndUndo undoObject, started
    If session = mVfwSession Then
        mVfwFaulted = True
        VfwChoiceChanged
    End If
    VfaReportFailure detail, changed
Clean:
    VfaEndUndo undoObject, started
    VfgEnd "apply", tick, number, changed
    If session = mVfwSession Then
        mVfbRunning = Not (mVfwPanel Is Nothing)
        mVfwBusy = False
    End If
End Sub

Private Sub VfwMove(ByVal direction As Long)
    Dim previous As Long
    If mVfwBusy Or mVfwRendering Or mVfwPanel Is Nothing Then Exit Sub
    If mVfwFaulted Then Exit Sub
    If direction > 0 Then
        If mVfwIndex <= mVfwHits.Count Then
            If Not mVfwAppliedIndices.Exists(CStr(mVfwIndex)) Then mVfwSkippedIndices(CStr(mVfwIndex)) = True
            If Not mVfwAppliedIndices.Exists(CStr(mVfwIndex)) Then VfwRemoveAutomatic mVfwIndex
            mVfwSkipped = mVfwSkippedIndices.Count
            mVfwIndex = mVfwIndex + 1
        End If
    ElseIf mVfwIndex > 1 Then
        For previous = mVfwIndex - 1 To 1 Step -1
            If mVfwVisited.Exists(CStr(previous)) Then
                mVfwIndex = previous
                VfwShowCurrent True
                Exit Sub
            End If
        Next previous
        Exit Sub
    End If
    VfwShowCurrent
End Sub

Private Sub VfwRestoreLayout()
    Dim updating As Boolean, captured As Boolean
    On Error GoTo Finished
    If Not mVfwLayout Then Exit Sub
    If Not VfwOpenDocument() Then Exit Sub
    ' Do not reactivate an old document when the user switched windows.
    If Not (ActiveDocument Is mVfwDoc) Then Exit Sub
    If Application.WindowState <> wdWindowStateNormal Then Exit Sub
    If Abs(Application.Width - mVfwReserved) >= 2 Or _
       Abs(Application.Left - mVfwAppliedLeft) >= 2 Or Abs(Application.Top - mVfwAppliedTop) >= 2 Or _
       Abs(Application.Height - mVfwAppliedHeight) >= 2 Then Exit Sub
    updating = Application.ScreenUpdating: captured = True
    Application.ScreenUpdating = False
    ' Only width changed in a normal window. A maximized window needs one setter.
    If mVfwState = wdWindowStateNormal Then
        If Abs(Application.Width - mVfwWidth) >= 2 Then Application.Width = mVfwWidth
    Else
        Application.WindowState = mVfwState
    End If
Finished:
    On Error Resume Next
    If captured Then Application.ScreenUpdating = updating
    On Error GoTo 0
End Sub

Private Sub VfwPause()
    If mVfwBusy Or mVfwRendering Or mVfwPanel Is Nothing Or mVfwFaulted Then Exit Sub
    If mVfwIndex > mVfwHits.Count Then Exit Sub
    mVfwPausing = True
    VfwClose
End Sub

Private Sub VfwResume()
    Dim tick As Single, number As Long, detail As String, currentExclusions As Object, state As Object
    Dim session As Long, owner As Word.Document
    If Not mVfwPaused Or mVfbRunning Or mVfwFaulted Then Exit Sub
    session = mVfwSession: Set owner = mVfwDoc
    tick = VfgClock()
    On Error GoTo Failed
    If Not VfwOpenDocument() Then Err.Raise VfaInvalidData, "Vefeyat", "The paused document was closed."
    If Not (ActiveDocument Is mVfwDoc) Then
        VfaWarn VfText("Duraklat\u0131lan belgeye ge\u00e7ip Denetime devam et komutunu kullan\u0131n.")
        Exit Sub
    End If
    If Not VefeyatEnsureData() Then Exit Sub
    If Not (mVfwCatalog Is mVfRows) Then Err.Raise VfaInvalidData, "Vefeyat", "Catalog changed; start a new review."
    If (VfpGetScope(mVfwDoc) = "first") <> mVfwFirstOnly Then Err.Raise VfaInvalidData, "Vefeyat", "Scope changed; start a new review."
    If (VfpGetAuditMode(mVfwDoc) = "audit_undated") <> mVfwSkipDated Then Err.Raise VfaInvalidData, "Vefeyat", "Audit mode changed; start a new review."
    Set currentExclusions = VfiRead(mVfwDoc)
    If VfiPayload(currentExclusions) <> mVfwExclusionStamp Then Err.Raise VfaInvalidData, "Vefeyat", "Exclusions changed; start a new review."
    If mVfwFirstOnly Then
        Set state = VfmState(mVfwDoc)
        If state Is Nothing Then
            Vefeyat0341UiWarmMemory mVfwDoc
        ElseIf Not CBool(state("Ready")) Then
            Vefeyat0341UiWarmMemory mVfwDoc
        End If
        If session <> mVfwSession Then Exit Sub
        If Not (mVfwDoc Is owner) Then Exit Sub
        If Not mVfwPaused Or mVfwDoc Is Nothing Then Exit Sub
        Set state = VfmState(mVfwDoc)
        If state Is Nothing Then Exit Sub
        If Not CBool(state("Ready")) Then Exit Sub
    End If
    If Not (ActiveDocument Is owner) Then Exit Sub
    If Not (mVfRows Is mVfwCatalog) Then Err.Raise VfaInvalidData, "Vefeyat", "Catalog changed during resume."
    If (VfpGetScope(owner) = "first") <> mVfwFirstOnly Then Err.Raise VfaInvalidData, "Vefeyat", "Scope changed during resume."
    If (VfpGetAuditMode(owner) = "audit_undated") <> mVfwSkipDated Then Err.Raise VfaInvalidData, "Vefeyat", "Audit mode changed during resume."
    If VfiPayload(VfiRead(owner)) <> mVfwExclusionStamp Then Err.Raise VfaInvalidData, "Vefeyat", "Exclusions changed during resume."
    VfwLoadVariables
    ' Edits while paused can change earlier context without changing a name.
    ' Do not replay RAM-only identity approvals for automatic repeat skipping.
    Set mVfwConfirmed = CreateObject("Scripting.Dictionary")
    Set mVfwChosenPeople = CreateObject("Scripting.Dictionary")
    mVfwPaused = False
    mVfbRunning = True: mVflSuspended = True
    VfwOpen
    VfgEnd "resume", tick
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    If session <> mVfwSession Then Exit Sub
    VfwClose
    VfgEnd "resume", tick, number
    VfaWarn VfText("Denetime devam edilemedi. Ayar veya veri de\u011fi\u015fmi\u015f olabilir; T\u00fcm belgeyi denetle ile yeniden ba\u015flat\u0131n. Hata: ") & CStr(number)
End Sub

Private Sub VfwIgnoreCurrent()
    Dim named As Word.Range, literal As String, nextExclusions As Object, payload As String
    Dim undoObject As Object, undoStarted As Boolean, tick As Single, number As Long, changed As Boolean
    Dim session As Long, owner As Word.Document, invalidated As Boolean, current As Object
    If mVfwBusy Or mVfwRendering Or mVfwFaulted Or mVfwPanel Is Nothing Then Exit Sub
    session = mVfwSession: Set owner = mVfwDoc
    tick = VfgClock()
    On Error GoTo Failed
    Set named = VfwCurrentName()
    If named Is Nothing Or mVfwContext Is Nothing Then Exit Sub
    If Not VfaEditable(mVfwDoc) Then Exit Sub
    If StrComp(mVfwContext.Text, mVfwContextText, vbBinaryCompare) <> 0 Then Exit Sub
    literal = named.Text
    mVfwBusy = True
    If VfMessage(VfText("Bu belge i\u00e7in \u015fu yaz\u0131m\u0131 toplu denetim d\u0131\u015f\u0131 b\u0131rak: ") & literal & vbCrLf & _
        VfText("Ayn\u0131 yaz\u0131m\u0131n di\u011fer kullan\u0131mlar\u0131 da sorulmaz. Metin ve tarihler de\u011fi\u015fmez. Tercih belgeyle kaydedilir; \u0130stisnalar\u0131 temizle ile kald\u0131r\u0131labilir. Onayl\u0131yor musunuz?"), _
        vbYesNo Or vbQuestion Or vbDefaultButton2, VfText("Belge istisnas\u0131n\u0131 onaylay\u0131n")) <> vbYes Then GoTo Clean
    If session <> mVfwSession Then GoTo Clean
    If Not (mVfwDoc Is owner) Then GoTo Clean
    If Not (ActiveDocument Is mVfwDoc) Then GoTo Clean
    If StrComp(named.Text, literal, vbBinaryCompare) <> 0 Then GoTo Clean
    If StrComp(mVfwContext.Text, mVfwContextText, vbBinaryCompare) <> 0 Then GoTo Clean
    If Not VfaEditable(mVfwDoc) Then GoTo Clean
    invalidated = True
    Set current = VfiRead(owner)
    If VfiPayload(current) <> mVfwExclusionStamp Then Err.Raise VfiInvalid, "Vefeyat", "Exclusions changed during confirmation."
    invalidated = False
    Set nextExclusions = VfiDecode(VfiPayload(mVfwExclusions))
    nextExclusions(VfiKey(literal)) = True
    payload = VfiPayload(nextExclusions)
    undoStarted = VfaBeginUndo(undoObject)
    changed = True
    VfaWriteVariable mVfwDoc, VfiVariable, payload
    VfaEndUndo undoObject, undoStarted
    Set mVfwExclusions = nextExclusions: mVfwExclusionStamp = payload
    VfgEnd "ignore", tick
    VfwShowCurrent
    GoTo Clean
Failed:
    number = Err.Number
    On Error Resume Next
    VfaEndUndo undoObject, undoStarted
    If session = mVfwSession Then
        If changed Or invalidated Then mVfwFaulted = True
        VfwChoiceChanged
    End If
    VfgEnd "ignore", tick, number, changed
    VfaWarn VfText("Belge istisnas\u0131 kaydedilemedi. Hata: ") & CStr(number)
Clean:
    If session = mVfwSession Then mVfwBusy = False
    On Error GoTo 0
End Sub

Private Sub VfwClose(Optional ByVal panelClosing As Boolean = False)
    Dim doc As Word.Document, panel As VefeyatKisiSec0341, tick As Single
    If mVfwClosing Then Exit Sub
    If mVfwPanel Is Nothing And mVfwDoc Is Nothing And Not mVfwPrepared Then Exit Sub
    mVfwClosing = True
    VfwBumpSession
    On Error Resume Next
    tick = VfgClock()
    mVfbRunning = True
    mVflSuspended = True
    Set doc = mVfwDoc
    mVfwRendering = True
    Set panel = mVfwPanel
    Set mVfwPanel = Nothing
    If Not panelClosing And Not panel Is Nothing Then
        Unload panel
    End If
    VfwRestoreLayout
    ' Mark stale without destroying/rebuilding the thesis index inside QueryClose.
    If Not doc Is Nothing Then VfmInvalidate doc
    If mVfwPausing Then
        mVfwPaused = True: mVfwPausing = False
        Set mVfwContext = Nothing: Set mVfwRows = Nothing
        mVfwContextText = vbNullString
        GoTo Finished
    End If
    Set mVfwDoc = Nothing: Set mVfwHits = Nothing: Set mVfwRows = Nothing
    Set mVfwPatterns = Nothing: Set mVfwContext = Nothing: Set mVfwVariables = Nothing
    Set mVfwCatalog = Nothing
    Set mVfwSkippedIndices = Nothing: Set mVfwAppliedIndices = Nothing
    Set mVfwChosenPeople = Nothing
    VfzClear
    Set mVfwConfirmed = Nothing: Set mVfwCandidateCache = Nothing
    Set mVfwVisited = Nothing: Set mVfwAutomatic = Nothing
    Set mVfwExclusions = Nothing
    mVfwExclusionStamp = vbNullString
    mVfwPaused = False: mVfwPausing = False: mVfwFaulted = False
    mVfwFirstVisited = 0: mVfwFirstOnly = False: mVfwSkipDated = False
    mVfwPrepared = False
Finished:
    mVfwLayout = False: mVfwRendering = False: mVfwBusy = False
    mVfbRunning = False: mVflSuspended = False
    mVfwClosing = False
    VfgEnd "close", tick
    Vefeyat0341UiInvalidate
    On Error GoTo 0
End Sub

Private Function VfgLabel(ByVal key As String) As String
    Select Case key
        Case "insert": VfgLabel = VfText("Se\u00e7ili ada ekleme")
        Case "prepare": VfgLabel = VfText("Denetim haz\u0131rl\u0131\u011f\u0131")
        Case "apply": VfgLabel = VfText("\u0130ncelemede ekleme")
        Case "close": VfgLabel = VfText("\u0130ncelemeyi kapatma")
        Case "resume": VfgLabel = VfText("\u0130ncelemeye devam")
        Case "ignore": VfgLabel = VfText("Belge istisnas\u0131")
        Case "search": VfgLabel = VfText("Arama (veri y\u00fcklendikten sonra)")
        Case "candidates": VfgLabel = VfText("\u0130nceleme aday s\u0131ralamas\u0131")
        Case "field": VfgLabel = VfText("\u0130ncelemede tek tarih alan\u0131 yazma")
    End Select
End Function

Private Function VfgClock() As Single
    VfgClock = Timer
End Function

Private Sub VfgEnd(ByVal key As String, ByVal started As Single, Optional ByVal errorNumber As Long = 0, Optional ByVal partialChange As Boolean = False)
    Dim elapsed As Double, row As Object
    On Error Resume Next
    If Len(VfgLabel(key)) = 0 Then Exit Sub
    elapsed = CDbl(VfgClock()) - CDbl(started)
    If elapsed < 0 Then elapsed = elapsed + 86400#
    elapsed = Fix(elapsed * 1000#)
    If mVfgStats Is Nothing Then Set mVfgStats = CreateObject("Scripting.Dictionary")
    If Not mVfgStats.Exists(key) Then
        Set row = CreateObject("Scripting.Dictionary")
        row("Count") = 0&: row("Max") = 0#: row("Errors") = 0&
        row("Error") = 0&: row("Partial") = False
        Set mVfgStats(key) = row
    End If
    Set row = mVfgStats(key)
    row("Count") = CLng(row("Count")) + 1
    row("Last") = elapsed
    If elapsed > CDbl(row("Max")) Then row("Max") = elapsed
    If errorNumber <> 0 Then
        row("Errors") = CLng(row("Errors")) + 1
        row("Error") = errorNumber
        row("Partial") = partialChange
    End If
End Sub

Private Function VfgReport() As String
    Dim key As Variant, row As Object, result As String
    On Error GoTo Unavailable
    result = VfText("Bu Word oturumunun tan\u0131s\u0131. Metin, ad ve dosya yolu kaydedilmez.") & vbCrLf
    If Not mVfgStats Is Nothing Then
        For Each key In Array("insert", "prepare", "apply", "close", "resume", "ignore", "search", "candidates", "field")
            If mVfgStats.Exists(CStr(key)) Then
                Set row = mVfgStats(CStr(key))
                result = result & vbCrLf & VfgLabel(CStr(key)) & ": " & CStr(row("Count")) & _
                    VfText(" i\u015flem; son ") & CStr(row("Last")) & " ms; en uzun " & CStr(row("Max")) & " ms."
                If CLng(row("Errors")) > 0 Then
                    result = result & VfText(" Hata say\u0131s\u0131: ") & CStr(row("Errors")) & _
                        VfText("; son kod: ") & CStr(row("Error")) & VfText("; k\u0131smi de\u011fi\u015fiklik olas\u0131l\u0131\u011f\u0131: ") & CStr(row("Partial"))
                End If
            End If
        Next key
    End If
    VfgReport = result & vbCrLf & vbCrLf & VfText("Duvar s\u00fcresidir; se\u00e7im penceresinde bekleme dahildir. Tek ba\u015f\u0131na i\u015flem h\u0131z\u0131 \u00f6l\u00e7\u00fct\u00fc de\u011fildir. D\u0131\u015far\u0131ya g\u00f6nderilmez; Word kapan\u0131nca b\u0131rak\u0131l\u0131r.")
    Exit Function
Unavailable:
    VfgReport = VfText("Bu oturumda tam bir tan\u0131 kayd\u0131 bulunamad\u0131. Belge metni okunmad\u0131 veya payla\u015f\u0131lmad\u0131.")
End Function

Private Function VfiKey(ByVal value As String) As String
    If Len(value) > 400 Then Exit Function
    value = VfaNormalizeWhitespace(value)
    If Not VfaValidAlias(value) Then Exit Function
    value = Replace(value, "I", ChrW(&H131))
    value = Replace(value, ChrW(&H130), "i")
    VfiKey = LCase$(value)
End Function

Private Function VfiDecode(ByVal value As String) As Object
    Dim result As Object, parts As Variant, i As Long, key As String
    Dim at As Long, colon As Long, size As Long, digits As String, count As Long
    Set result = CreateObject("Scripting.Dictionary")
    If Len(value) = 0 Then Set VfiDecode = result: Exit Function
    If Len(value) > 12000 Then Err.Raise VfiInvalid, "Vefeyat", "Invalid document exclusions."
    ' New writes use length prefixes, with no line breaks in a Word XML attribute.
    If Left$(value, 5) = "Vfi1:" Then
        at = 6
        Do While at <= Len(value)
            count = count + 1
            If count > 200 Then Err.Raise VfiInvalid, "Vefeyat", "Too many document exclusions."
            colon = InStr(at, value, ":", vbBinaryCompare)
            If colon = 0 Or colon - at < 1 Or colon - at > 3 Then Err.Raise VfiInvalid, "Vefeyat", "Invalid exclusion length."
            digits = Mid$(value, at, colon - at)
            size = 0
            For i = 1 To Len(digits)
                If Mid$(digits, i, 1) < "0" Or Mid$(digits, i, 1) > "9" Then Err.Raise VfiInvalid, "Vefeyat", "Invalid exclusion length."
                size = size * 10 + CLng(Mid$(digits, i, 1))
            Next i
            If size < 1 Or size > 200 Or size > Len(value) - colon Then Err.Raise VfiInvalid, "Vefeyat", "Invalid exclusion length."
            key = VfiKey(Mid$(value, colon + 1, size))
            If Len(key) = 0 Then Err.Raise VfiInvalid, "Vefeyat", "Invalid document exclusion."
            result(key) = True
            at = colon + size + 1
        Loop
        Set VfiDecode = result
        Exit Function
    End If
    parts = Split(value, vbLf)
    If CStr(parts(0)) <> "Vfi1" Then Err.Raise VfiInvalid, "Vefeyat", "Invalid document exclusions."
    If UBound(parts) > 200 Then Err.Raise VfiInvalid, "Vefeyat", "Too many document exclusions."
    For i = 1 To UBound(parts)
        key = VfiKey(CStr(parts(i)))
        If Len(key) = 0 Then Err.Raise VfiInvalid, "Vefeyat", "Invalid document exclusion."
        result(key) = True
    Next i
    Set VfiDecode = result
End Function

Private Function VfiRead(ByVal doc As Word.Document) As Object
    Set VfiRead = VfiDecode(VfaReadVariable(doc, VfiVariable))
End Function

Private Function VfiPayload(ByVal exclusions As Object) As String
    Dim key As Variant, result As String
    If exclusions.Count > 200 Then Err.Raise VfiInvalid, "Vefeyat", "Too many document exclusions."
    result = "Vfi1:"
    For Each key In exclusions.Keys
        If Len(VfiKey(CStr(key))) = 0 Then Err.Raise VfiInvalid, "Vefeyat", "Invalid document exclusion."
        result = result & CStr(Len(CStr(key))) & ":" & CStr(key)
    Next key
    If Len(result) > 12000 Then Err.Raise VfiInvalid, "Vefeyat", "Document exclusion limit reached."
    VfiPayload = result
End Function

Private Function VfiContains(ByVal exclusions As Object, ByVal value As String) As Boolean
    Dim key As String
    If exclusions Is Nothing Then Exit Function
    key = VfiKey(value)
    If Len(key) > 0 Then VfiContains = exclusions.Exists(key)
End Function

Private Function VfiWithout(ByVal exclusions As Object, ByVal key As String) As Object
    Dim result As Object
    Set result = VfiDecode(VfiPayload(exclusions))
    If Not result.Exists(key) Then Err.Raise VfiInvalid, "Vefeyat", "The exclusion no longer exists."
    result.Remove key
    Set VfiWithout = result
End Function

Private Sub VfiEdit(ByVal doc As Word.Document)
    Dim exclusions As Object, current As Object, rows As New Collection, key As Variant, row As Variant
    Dim picker As VefeyatKisiSec0341, choice As Long, stamp As String, undoObject As Object, started As Boolean
    Dim number As Long
    On Error GoTo Failed
    Set exclusions = VfiRead(doc)
    If exclusions.Count = 0 Then VfaWarn VfText("Bu belgede kay\u0131tl\u0131 istisna yok."): Exit Sub
    stamp = VfiPayload(exclusions)
    For Each key In exclusions.Keys
        rows.Add Array(CStr(key), "", "", "", VfText("Listeden \u00e7\u0131kar\u0131rsan\u0131z yeni toplu denetimde bu ifade yeniden sorulabilir. Metin ve tarihler de\u011fi\u015fmez."), "")
    Next key
    Set picker = New VefeyatKisiSec0341
    picker.InitCandidates rows, VfText("Belge istisnalar\u0131")
    picker.SetExclusionMode
    picker.Show vbModal
    choice = picker.ChosenIndex
    Unload picker: Set picker = Nothing
    If choice < 1 Or choice > rows.Count Then Exit Sub
    If Not (ActiveDocument Is doc) Then Exit Sub
    If Not VfaEditable(doc) Then Exit Sub
    Set current = VfiRead(doc)
    If VfiPayload(current) <> stamp Then VfaWarn VfText("Liste de\u011fi\u015fti; yeniden a\u00e7\u0131n."): Exit Sub
    row = rows(choice)
    Set current = VfiWithout(current, CStr(row(0)))
    started = VfaBeginUndo(undoObject)
    VfaWriteVariable doc, VfiVariable, VfiPayload(current)
    VfaEndUndo undoObject, started
    Exit Sub
Failed:
    number = Err.Number
    On Error Resume Next
    If Not picker Is Nothing Then Unload picker
    VfaEndUndo undoObject, started
    VfaWarn VfText("\u0130stisna kald\u0131r\u0131lamad\u0131. Hata: ") & CStr(number)
    On Error GoTo 0
End Sub

Private Sub VfiShow(ByVal doc As Word.Document)
    Dim exclusions As Object, key As Variant, message As String
    Set exclusions = VfiRead(doc)
    message = VfText("Bu belgeye kay\u0131tl\u0131 ifadeler: ") & CStr(exclusions.Count) & vbCrLf
    For Each key In exclusions.Keys
        message = message & vbCrLf & CStr(key)
    Next key
    VfMessage message & vbCrLf & vbCrLf & VfText("Toplu denetim ve \u015fah\u0131s d\u00f6k\u00fcm\u00fcnde uygulan\u0131r. Elle ad se\u00e7ip arama ve metin de\u011fi\u015fmez. Listeyi s\u0131f\u0131rlamak i\u00e7in \u0130stisnalar\u0131 temizle komutunu kullan\u0131n."), vbInformation, VfText("Belge istisnalar\u0131")
End Sub

Private Function VfePersonParts(ByVal key As String) As Variant
    Dim parts(0 To 3) As String, i As Long, at As Long, colon As Long, number As String, size As Long, digit As Long
    If Left$(key, 5) <> "Vfp1:" Or Len(key) > 16432 Then Exit Function
    at = 6
    For i = 0 To 3
        colon = InStr(at, key, ":", vbBinaryCompare)
        If colon = 0 Or colon - at < 1 Or colon - at > 4 Then Exit Function
        number = Mid$(key, at, colon - at): size = 0
        For digit = 1 To Len(number)
            If Mid$(number, digit, 1) < "0" Or Mid$(number, digit, 1) > "9" Then Exit Function
            size = size * 10 + CLng(Mid$(number, digit, 1))
        Next digit
        If size > 4096 Or size > Len(key) - colon Then Exit Function
        parts(i) = Mid$(key, colon + 1, size)
        at = colon + size + 1
    Next i
    If at <> Len(key) + 1 Then Exit Function
    If Len(parts(1)) + Len(parts(2)) + Len(parts(3)) = 0 Then Exit Function
    VfePersonParts = parts
End Function

Private Function VfeDescription(ByVal personKey As String, ByVal raw As String, ByVal aliasText As String) As String
    Dim parts As Variant, url As String
    parts = VfePersonParts(personKey)
    If IsEmpty(parts) Then Exit Function
    If Not VfaValidDate(raw) Or Not VfaValidAlias(aliasText) Then Exit Function
    url = VfdArticleUrl(CStr(parts(0)))
    VfeDescription = VfText("Belgedeki ad: ") & aliasText & vbCrLf & _
        VfText("Kay\u0131tl\u0131 \u015fah\u0131s: ") & CStr(parts(1)) & vbCrLf & _
        VfText("Tam ad: ") & CStr(parts(2)) & vbCrLf & _
        VfText("Belgede kay\u0131tl\u0131 ham tarih: ") & raw & vbCrLf
    If Len(url) > 0 Then
        VfeDescription = VfeDescription & "DIA: " & url & vbCrLf
    Else
        VfeDescription = VfeDescription & VfText("Ge\u00e7erli D\u0130A madde ba\u011f\u0131 bulunamad\u0131.") & vbCrLf
    End If
    VfeDescription = VfeDescription & vbCrLf & VfText("Bu bilgi belgedeki Vefey\u00e2t ba\u011f\u0131ndan okundu. G\u00fcncel CSV'den yeni e\u015fle\u015ftirme veya \u00e7evrimi\u00e7i do\u011frulama yap\u0131lmad\u0131. Belge de\u011fi\u015ftirilmedi.")
End Function

Private Sub VfeShowSelected()
    Dim doc As Word.Document, selected As Word.Range, field As Word.Field, bindingNamespace As String
    Dim id As String, message As String, found As String, current As String, count As Long, first As Long, last As Long
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument: Set selected = Selection.Range.Duplicate
    If selected.End - selected.Start > 512 Then GoTo NoSource
    If selected.StoryType <> wdMainTextStory And selected.StoryType <> wdFootnotesStory Then GoTo NoSource
    bindingNamespace = VfaReadVariable(doc, VfaNamespaceVariable)
    If Len(bindingNamespace) = 0 Then GoTo NoSource
    For Each field In selected.Paragraphs(1).Range.Fields
        count = count + 1
        If count > 64 Then GoTo NoSource
        first = field.Code.Start - 1: last = field.Result.End + 1
        If (selected.Start <= last And selected.End >= first) Or (first >= selected.End And first <= selected.End + 2) Then
            id = VfaBindingFromFieldCode(field.Code.Text, bindingNamespace)
            If Len(id) > 0 Then
                If Len(found) > 0 And found <> id Then GoTo NoSource
                found = id
                current = VfeDescription(VfaReadVariable(doc, VfaVariableName(id, "Key")), _
                    VfaReadVariable(doc, VfaVariableName(id, "Raw")), VfaReadVariable(doc, VfaVariableName(id, "Alias")))
                If Len(current) = 0 Then GoTo NoSource
                message = current
            End If
        End If
    Next field
    If Len(message) > 0 Then
        VfMessage message, vbInformation, VfText("Tarihin kay\u0131tl\u0131 kayna\u011f\u0131")
        Exit Sub
    End If
NoSource:
    VfaWarn VfText("Kaynak ba\u011f\u0131 bulunamad\u0131. Vefey\u00e2t ile eklenmi\u015f tarihi veya hemen \u00f6n\u00fcndeki ad\u0131 se\u00e7in. Elle yaz\u0131lm\u0131\u015f tarihe kaynak atanmaz.")
End Sub

Private Function VfqIdentity(ByRef row As Variant) As String
    Dim url As String, fragment As Long
    url = VfdArticleUrl(CStr(row(5)))
    If Len(url) > 0 Then
        fragment = InStr(url, "#")
        If fragment > 0 Then url = Left$(url, fragment - 1)
        Do While Right$(url, 1) = "/"
            url = Left$(url, Len(url) - 1)
        Loop
        If Len(url) <= Len("https://islamansiklopedisi.org.tr") Then url = ""
    End If
    If Len(url) > 0 Then
        VfqIdentity = "DIA:" & LCase$(url)
    Else
        VfqIdentity = VfpPersonKey(row)
    End If
End Function

' Shared article links are not sufficient to merge two people. Keep a legacy
' full identity for preferences when the current catalog has duplicate links.
Private Function VfqPreferenceIdentity(ByRef row As Variant) As String
    Dim key As String
    key = VfqIdentity(row)
    VfqPreferenceIdentity = VfpPersonKey(row)
    If Left$(key, 4) <> "DIA:" Or mVfRows Is Nothing Then Exit Function
    If Not (mVfqCatalogOwner Is mVfRows) Then
        Set mVfqCatalogIndex = VfqIndex(mVfRows)
        Set mVfqCatalogOwner = mVfRows
    End If
    If mVfqCatalogIndex.Exists(key) Then
        If mVfqCatalogIndex(key).Count = 1 Then VfqPreferenceIdentity = key
    End If
End Function

Private Function VfqIndex(ByVal rows As Collection) As Object
    Dim result As Object, row As Variant, key As String, bucket As Collection
    Set result = CreateObject("Scripting.Dictionary")
    For Each row In rows
        key = VfqIdentity(row)
        If Not result.Exists(key) Then
            Set bucket = New Collection
            Set result(key) = bucket
        End If
        result(key).Add row
    Next row
    Set VfqIndex = result
End Function

Private Function VfqRowStamp(ByRef row As Variant) As String
    Dim i As Long
    For i = 0 To 5
        VfqRowStamp = VfqRowStamp & CStr(Len(CStr(row(i)))) & ":" & CStr(row(i))
    Next i
End Function

Private Function VfqCatalogSame(ByVal first As Collection, ByVal second As Collection) As Boolean
    Dim i As Long, a As Variant, b As Variant
    If first.Count <> second.Count Then Exit Function
    For i = 1 To first.Count
        a = first(i): b = second(i)
        If VfqRowStamp(a) <> VfqRowStamp(b) Then Exit Function
    Next i
    VfqCatalogSame = True
End Function

Private Function VfqCatalogSummary(ByVal previous As Collection, ByVal incoming As Collection) As String
    Dim oldIndex As Object, newIndex As Object, key As Variant, oldRow As Variant, newRow As Variant
    Dim added As Long, removed As Long, changed As Long, ambiguous As Long, dates As Long
    Set oldIndex = VfqIndex(previous): Set newIndex = VfqIndex(incoming)
    For Each key In newIndex.Keys
        If newIndex(key).Count > 1 Then
            ambiguous = ambiguous + 1
        ElseIf Not oldIndex.Exists(key) Then
            added = added + 1
        ElseIf oldIndex(key).Count = 1 Then
            oldRow = oldIndex(key)(1): newRow = newIndex(key)(1)
            If VfqRowStamp(oldRow) <> VfqRowStamp(newRow) Then changed = changed + 1
            If CStr(oldRow(3)) <> CStr(newRow(3)) Then dates = dates + 1
        Else
            ambiguous = ambiguous + 1
        End If
    Next key
    For Each key In oldIndex.Keys
        If Not newIndex.Exists(key) Then removed = removed + 1
    Next key
    VfqCatalogSummary = VfText("Katalog \u00f6nizlemesi") & vbCrLf & _
        VfText("Yeni kimlik: ") & CStr(added) & VfText("; kald\u0131r\u0131lan: ") & CStr(removed) & vbCrLf & _
        VfText("De\u011fi\u015fen: ") & CStr(changed) & VfText("; tarihi de\u011fi\u015fen: ") & CStr(dates) & vbCrLf & _
        VfText("Yinelenen / belirsiz kimlik: ") & CStr(ambiguous) & vbCrLf & vbCrLf & _
        VfText("D\u0130A ba\u011f\u0131 ayn\u0131 olan kay\u0131tlar ad de\u011fi\u015fse de birlikte izlenir. Belirsiz kay\u0131tlar otomatik birle\u015ftirilmez. Belgedeki tarihler ve kimlik ba\u011flar\u0131 bu i\u015flemle de\u011fi\u015fmez.")
End Function

Private Function VfqStrength(ByRef row As Variant, ByVal query As String) As Long
    Dim i As Long, score As Long
    For i = 0 To 2
        score = VefeyatNameMatch(CStr(row(i)), query)
        If score > VfqStrength Then VfqStrength = score
    Next i
End Function

Private Function VfqSingleSafe(ByVal rows As Collection, ByVal query As String) As Boolean
    Dim row As Variant
    If rows.Count <> 1 Then Exit Function
    row = rows(1)
    VfqSingleSafe = (VfqStrength(row, query) >= 3)
End Function

Private Function VfqDateNote(ByVal raw As String) As String
    Dim key As String, part As Variant
    key = VefeyatNormalize(raw)
    For Each part In Split("yaklasik|once|sonra|civari|tahminen|bilinmiyor|ihtilaf|muhtemel", "|")
        If InStr(1, key, CStr(part), vbBinaryCompare) > 0 Then
            VfqDateNote = VfText("Belirsizlik i\u00e7eren kaynak tarihi aynen korunur; kesin tarihe d\u00f6n\u00fc\u015ft\u00fcr\u00fclmez.")
            Exit Function
        End If
    Next part
    If InStr(raw, "?") > 0 Or InStr(raw, ";") > 0 Then VfqDateNote = VfText("Tarih belirsizlik veya alternatif i\u00e7erebilir. Ham kaynak ifadesini kontrol edin.")
    If InStr(raw, "/") > 0 Then VfqDateNote = VfqDateNote & vbCrLf & VfText("Kaynak takvimleri aras\u0131nda yeni bir d\u00f6n\u00fc\u015f\u00fcm hesaplanmad\u0131.")
End Function

Private Function VfqExplain(ByRef row As Variant, ByVal query As String) As String
    Dim score As Long, reason As String
    score = VfqStrength(row, query)
    Select Case score
        Case 4: reason = VfText("Ad alanlar\u0131ndan biriyle normalle\u015ftirilmi\u015f tam e\u015fle\u015fme.")
        Case 3: reason = VfText("Bo\u015fluk ve ad i\u00e7i i\u015faret farklar\u0131 kald\u0131r\u0131ld\u0131\u011f\u0131nda e\u015fle\u015fme.")
        Case 2: reason = VfText("Ad\u0131n bir b\u00f6l\u00fcm\u00fcyle s\u00f6zc\u00fck e\u015fle\u015fmesi; ayn\u0131 \u015fah\u0131s oldu\u011funu kontrol edin.")
        Case Else: reason = VfText("Yak\u0131n yaz\u0131m veya k\u0131smi benzerlik; kimlik do\u011frulamas\u0131 gerekir.")
    End Select
    VfqExplain = VfText("E\u015fle\u015fme gerek\u00e7esi: ") & reason & vbCrLf & _
        VfText("S\u0131ralama kimlik kan\u0131t\u0131 de\u011fildir. Metin ba\u011flam\u0131 ve tan\u0131t\u0131m birlikte incelenmelidir.") & vbCrLf & VfqDateNote(CStr(row(3)))
End Function

Private Function VfuPack(ByVal values As Collection) As String
    Dim value As Variant
    VfuPack = "Vfu1:"
    If values.Count > 400 Then Err.Raise VfuInvalid, "Vefeyat", "Profile entry limit."
    For Each value In values
        If Len(CStr(value)) > 12000 Then Err.Raise VfuInvalid, "Vefeyat", "Profile value limit."
        VfuPack = VfuPack & CStr(Len(CStr(value))) & ":" & CStr(value)
    Next value
    If Len(VfuPack) > 30000 Then Err.Raise VfuInvalid, "Vefeyat", "Profile size limit."
End Function

Private Function VfuUnpack(ByVal payload As String) As Collection
    Dim result As New Collection, at As Long, colon As Long, size As Long, i As Long, digits As String
    If Left$(payload, 5) <> "Vfu1:" Or Len(payload) > 30000 Then Err.Raise VfuInvalid, "Vefeyat", "Invalid profile."
    at = 6
    Do While at <= Len(payload)
        colon = InStr(at, payload, ":", vbBinaryCompare)
        If colon < at + 1 Or colon > at + 5 Then Err.Raise VfuInvalid, "Vefeyat", "Invalid profile length."
        digits = Mid$(payload, at, colon - at): size = 0
        For i = 1 To Len(digits)
            If Mid$(digits, i, 1) < "0" Or Mid$(digits, i, 1) > "9" Then Err.Raise VfuInvalid, "Vefeyat", "Invalid profile length."
            size = size * 10 + CLng(Mid$(digits, i, 1))
        Next i
        If size > 12000 Or size > Len(payload) - colon Then Err.Raise VfuInvalid, "Vefeyat", "Invalid profile length."
        result.Add Mid$(payload, colon + 1, size)
        If result.Count > 400 Then Err.Raise VfuInvalid, "Vefeyat", "Profile entry limit."
        at = colon + size + 1
    Loop
    Set VfuUnpack = result
End Function

Private Function VfuKnownMode(ByVal mode As String) As Boolean
    Select Case mode
        Case "ask", "near", "short", "full", "original": VfuKnownMode = True
    End Select
End Function

Private Function VfuValueLabel(ByVal value As String) As String
    Select Case value
        Case "original": VfuValueLabel = VfText("\u00d6zg\u00fcn bi\u00e7im")
        Case "o": VfuValueLabel = VfText("\u00f6.")
        Case "ol": VfuValueLabel = VfText("\u00f6l.")
        Case "v": VfuValueLabel = "v."
        Case "first": VfuValueLabel = VfText("\u0130lk ge\u00e7ti\u011fi yer")
        Case "none": VfuValueLabel = VfText("Yaln\u0131z denetle; tarih ekleme")
        Case "all": VfuValueLabel = VfText("Her ge\u00e7ti\u011fi yer")
        Case "audit_all": VfuValueLabel = VfText("Tamam\u0131n\u0131 denetle")
        Case "audit_undated": VfuValueLabel = VfText("Tarih eklenmi\u015f olanlar\u0131 atla")
        Case "ask": VfuValueLabel = "Her seferinde sor"
        Case "near": VfuValueLabel = VfText("Se\u00e7ili ada yak\u0131n")
        Case "short": VfuValueLabel = VfText("K\u0131sa ad")
        Case "full": VfuValueLabel = "Tam / uzun ad"
    End Select
End Function

Private Function VfuNameMode(ByVal doc As Word.Document) As String
    VfuNameMode = VfaReadVariable(doc, VfuModeVariable)
    If Not VfuKnownMode(VfuNameMode) Then VfuNameMode = "ask"
End Function

Private Function VfuProposal(ByRef row As Variant, ByVal query As String, ByVal mode As String) As String
    Select Case mode
        Case "near": VfuProposal = VfnSourceName(row, query)
        Case "short": VfuProposal = VfnTextName(CStr(row(0)))
        Case "full": VfuProposal = VfnTextName(CStr(row(1)))
        Case "original": VfuProposal = CStr(row(2))
    End Select
    If VfuProposal = "---" Or Not VfaValidAlias(VfuProposal) Then VfuProposal = ""
End Function

Private Function VfuExportPayload(ByVal doc As Word.Document) As String
    Dim values As New Collection
    values.Add VfpGetPrefix(doc): values.Add VfpGetScope(doc): values.Add VfpGetAuditMode(doc)
    values.Add VfuNameMode(doc): values.Add VfiPayload(VfiRead(doc))
    values.Add VfuNamesPayload(VfuNames(doc))
    VfuExportPayload = VfuPack(values)
End Function

Private Function VfuValidateProfile(ByVal payload As String) As Collection
    Dim values As Collection, excludes As Object
    Set values = VfuUnpack(payload)
    If values.Count = 5 Then
        values.Add "Vfu1:"
    End If
    If values.Count <> 6 Then Err.Raise VfuInvalid, "Vefeyat", "Invalid project profile."
    If Not VfpKnownPrefix(CStr(values(1))) Or Not VfpKnownScope(CStr(values(2))) Then Err.Raise VfuInvalid, "Vefeyat", "Invalid profile policy."
    If CStr(values(3)) <> "audit_all" And CStr(values(3)) <> "audit_undated" Then Err.Raise VfuInvalid, "Vefeyat", "Invalid profile audit policy."
    If Not VfuKnownMode(CStr(values(4))) Then Err.Raise VfuInvalid, "Vefeyat", "Invalid profile name form."
    Set excludes = VfiDecode(CStr(values(5)))
    Set excludes = VfuReadNames(CStr(values(6)))
    Set VfuValidateProfile = values
End Function

Private Sub VfuProject(ByVal importing As Boolean)
    Dim doc As Word.Document, dialog As Object, path As String, payload As String, stamp As String
    Dim values As Collection, stream As Object, undoObject As Object, started As Boolean, number As Long, detail As String
    On Error GoTo Failed
    Set doc = ActiveDocument: stamp = VfuExportPayload(doc)
    Set dialog = Application.FileDialog(3)
    If importing Then
        dialog.Title = VfText("Proje tercihlerini al (.vfp)")
        dialog.Filters.Clear: dialog.Filters.Add "Vefeyat project", "*.vfp"
        dialog.AllowMultiSelect = False
        If dialog.Show <> -1 Then Exit Sub
        path = CStr(dialog.SelectedItems(1))
        VfCheckLocalPath path
        If CreateObject("Scripting.FileSystemObject").GetFile(path).Size > 100000 Then Err.Raise VfuInvalid, "Vefeyat", "Profile file too large."
        payload = VfsReadUtf8(path)
        Set values = VfuValidateProfile(payload)
        If VfMessage(VfText("Proje tercihleri bu belgenin ayarlar\u0131n\u0131 de\u011fi\u015ftirecek.") & vbCrLf & _
            VfuValueLabel(CStr(values(1))) & "; " & VfuValueLabel(CStr(values(2))) & "; " & VfuValueLabel(CStr(values(3))) & vbCrLf & _
            VfText("Ad bi\u00e7imi: ") & VfuValueLabel(CStr(values(4))) & vbCrLf & VfText("\u0130stisnalar ve tercih edilen \u015fah\u0131s yaz\u0131mlar\u0131 da de\u011fi\u015ftirilir. Metin ve tarihler de\u011fi\u015fmez. Uygulans\u0131n m\u0131?"), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") <> vbYes Then Exit Sub
        If Not (ActiveDocument Is doc) Then Exit Sub
        If Not VfaEditable(doc) Or stamp <> VfuExportPayload(doc) Then Exit Sub
        started = VfaBeginUndo(undoObject)
        VfaWriteVariable doc, VfuModeVariable, CStr(values(4))
        VfpSetPrefix doc, CStr(values(1)): VfpSetScope doc, CStr(values(2)): VfpSetAuditMode doc, CStr(values(3))
        VfaWriteVariable doc, VfiVariable, CStr(values(5))
        VfaWriteVariable doc, VfuNamesVariable, CStr(values(6))
        VfaEndUndo undoObject, started
    Else
        path = InputBox(VfText("Tercih dosyas\u0131 i\u00e7in yerel tam yol (.vfp). Belge metni ve katalog eklenmez."), "Vefeyat")
        If Len(path) = 0 Then Exit Sub
        VfCheckLocalPath path
        If LCase$(Right$(path, 4)) <> ".vfp" Then Err.Raise VfuInvalid, "Vefeyat", "Use a .vfp file."
        If CreateObject("Scripting.FileSystemObject").FileExists(path) Then Err.Raise VfuInvalid, "Vefeyat", "Choose a new file; existing file preserved."
        VfsWriteVerifiedUtf8 path, stamp
    End If
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    VfaEndUndo undoObject, started
    Err.Raise number, "Vefeyat", detail
End Sub

Private Function VfuReadNames(ByVal payload As String) As Object
    Dim result As Object, values As Collection, i As Long, key As String, parts As Variant, url As String
    Set result = CreateObject("Scripting.Dictionary")
    If Len(payload) = 0 Then Set VfuReadNames = result: Exit Function
    If Len(payload) > 12000 Then Err.Raise VfuInvalid, "Vefeyat", "Preferred-name limit."
    Set values = VfuUnpack(payload)
    If values.Count Mod 2 <> 0 Or values.Count > 200 Then Err.Raise VfuInvalid, "Vefeyat", "Invalid preferred names."
    For i = 1 To values.Count Step 2
        key = CStr(values(i))
        If Left$(key, 4) = "DIA:" Then
            url = VfdArticleUrl(Mid$(key, 5))
            If Len(url) = 0 Then Err.Raise VfuInvalid, "Vefeyat", "Invalid source identity."
        Else
            parts = VfePersonParts(key)
            If IsEmpty(parts) Then Err.Raise VfuInvalid, "Vefeyat", "Invalid source identity."
        End If
        If Not VfaValidAlias(CStr(values(i + 1))) Or result.Exists(key) Then Err.Raise VfuInvalid, "Vefeyat", "Invalid preferred name."
        result(key) = CStr(values(i + 1))
    Next i
    Set VfuReadNames = result
End Function

Private Function VfuNames(ByVal doc As Word.Document) As Object
    Set VfuNames = VfuReadNames(VfaReadVariable(doc, VfuNamesVariable))
End Function

Private Function VfuNamesPayload(ByVal names As Object) As String
    Dim values As New Collection, key As Variant
    If names.Count > 100 Then Err.Raise VfuInvalid, "Vefeyat", "Preferred-name entry limit (100)."
    For Each key In names.Keys
        values.Add CStr(key)
        values.Add CStr(names(key))
    Next key
    VfuNamesPayload = VfuPack(values)
    If Len(VfuNamesPayload) > 12000 Then Err.Raise VfuInvalid, "Vefeyat", "Preferred-name limit."
End Function

Private Function VfuPreferred(ByVal doc As Word.Document, ByRef row As Variant) As String
    Dim names As Object, key As String
    Set names = VfuNames(doc): key = VfqPreferenceIdentity(row)
    If names.Exists(key) Then VfuPreferred = CStr(names(key))
End Function

Private Sub VfuRememberName()
    Dim doc As Word.Document, named As Word.Range, original As String, aliasText As String, reason As String
    Dim first As Long, last As Long, story As Long, rows As Collection, choice As Long, row As Variant
    Dim names As Object, stamp As String, proposal As String
    If Not VfaCapture(doc, named, original, aliasText, first, last, story) Then Exit Sub
    proposal = VfNameQuery(aliasText)
    If Not VfaValidAlias(proposal) Then Exit Sub
    Set names = VfuNames(doc): stamp = VfuNamesPayload(names)
    If Not VefeyatEnsureData() Then Exit Sub
    Set rows = VfaConfirmationCandidates(VefeyatFind(aliasText, False), aliasText)
    If rows.Count = 0 Then Exit Sub
    choice = VfePickPerson(rows, aliasText)
    If choice < 1 Or choice > rows.Count Then Exit Sub
    row = rows(choice)
    If VfMessage(VfText("Bu projede tercih edilen yaz\u0131m: ") & proposal & vbCrLf & _
        VfText("\u015eah\u0131s: ") & CStr(row(0)) & vbCrLf & VfText("Bu tercih yaz\u0131m denetiminde \u00f6nerilecek; metin kendili\u011finden de\u011fi\u015fmez. Kaydedilsin mi?"), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") <> vbYes Then Exit Sub
    If Not VfaTargetCurrent(doc, named, original, first, last, story, reason) Then Exit Sub
    If stamp <> VfuNamesPayload(VfuNames(doc)) Then Err.Raise VfuInvalid, "Vefeyat", "Preferred names changed."
    names(VfqPreferenceIdentity(row)) = proposal
    VfaWriteVariable doc, VfuNamesVariable, VfuNamesPayload(names)
End Sub

Private Function VfuReasonLabel(ByVal reason As String) As String
    Select Case reason
        Case "not_person": VfuReasonLabel = VfText("Bu kullan\u0131m \u015fah\u0131s de\u011fil")
        Case "work_title": VfuReasonLabel = VfText("Eser ad\u0131 / at\u0131f")
        Case "other_person": VfuReasonLabel = VfText("Ba\u015fka \u015fah\u0131s; yaln\u0131z bu ge\u00e7i\u015fi atla")
    End Select
End Function

Private Sub VfuFeedback(ByVal reason As String)
    Dim doc As Word.Document, named As Word.Range, literal As String, exclusions As Object
    Dim stamp As String, original As String, why As String, first As Long, last As Long, story As Long
    If Len(VfuReasonLabel(reason)) = 0 Then Exit Sub
    If Not VfaCapture(doc, named, original, literal, first, last, story) Then Exit Sub
    If reason = "other_person" Then
        VftLog doc, VfuReasonLabel(reason), named, literal, literal
        VfMessage VfText("Bu ge\u00e7i\u015fe tarih eklenmedi. Ba\u015fka \u015fah\u0131s olabilece\u011fi i\u00e7in bu yaz\u0131m genel istisnaya al\u0131nmad\u0131. Toplu incelemede Atla / Sonraki kullan\u0131n."), vbInformation, "Vefeyat"
        Exit Sub
    End If
    Set exclusions = VfiRead(doc): stamp = VfiPayload(exclusions)
    If VfMessage(VfuReasonLabel(reason) & ": " & literal & vbCrLf & VfText("Bu yaz\u0131m\u0131n b\u00fct\u00fcn kullan\u0131mlar\u0131 bu belgenin toplu denetiminden \u00e7\u0131kar\u0131lacak. Ba\u015fka belgeler etkilenmez. Onayl\u0131yor musunuz?"), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") <> vbYes Then Exit Sub
    If Not VfaTargetCurrent(doc, named, original, first, last, story, why) Then Exit Sub
    If stamp <> VfiPayload(VfiRead(doc)) Then Err.Raise VfuInvalid, "Vefeyat", "Exclusions changed."
    exclusions(VfiKey(literal)) = True
    VfaWriteVariable doc, VfiVariable, VfiPayload(exclusions)
    VftLog doc, VfuReasonLabel(reason), named, literal, VfText("Belge istisnas\u0131na al\u0131nd\u0131")
End Sub

Private Function VdiContains(ByVal text As String, ByVal phrase As String) As Boolean
    Dim key As String
    key = VefeyatNormalize(phrase)
    If Len(key) < 3 Then Exit Function
    VdiContains = InStr(1, " " & VefeyatNormalize(text) & " ", " " & key & " ", vbBinaryCompare) > 0
End Function

Private Sub VdiFact(ByVal facts As Object, ByVal kind As String, ByVal phrase As String, ByVal quote As String)
    Dim key As String
    phrase = VfaTrimSpaces(phrase)
    If Len(phrase) < 3 Or Len(phrase) > 180 Or Not VfaSingleLine(phrase) Then Exit Sub
    key = kind & ":" & VefeyatNormalize(phrase)
    If Not facts.Exists(key) Then facts.Add key, Array(kind, phrase, Left$(quote, 600))
End Sub

Private Function VdiFacts(ByRef row As Variant, ByVal query As String) As Object
    Dim facts As Object, text As String, key As String, kind As String, phrase As String
    Dim i As Long, j As Long, opening As String, closing As String, first As Long, last As Long
    Dim item As Variant, part As Variant, at As Long, marker As String, near As String
    Set facts = CreateObject("Scripting.Dictionary")
    For i = 0 To 2
        phrase = CStr(row(i))
        If i < 2 Then phrase = VfnTextName(phrase)
        If Len(VefeyatNormalize(phrase)) > Len(VefeyatNormalize(VfNameQuery(query))) + 3 Then
            If InStr(VefeyatNormalize(phrase), " ") > 0 Then VdiFact facts, "name", phrase, CStr(row(i))
        End If
    Next i
    text = CStr(row(4))
    ' Explicit labels are literal catalog statements, never invented relations.
    For Each item In Array("eser:", "eseri:", "kitabi:", "donem:", "hocasi:", "ogretmeni:", "talebesi:", "ogrencisi:")
        key = VefeyatNormalize(CStr(item))
        For Each part In Split(Replace(text, vbCr, ";"), ";")
            at = InStr(CStr(part), ":")
            If at > 1 Then
                If VefeyatNormalize(Left$(CStr(part), at - 1)) = key Then
                    kind = "relation"
                    If key = "eser" Or key = "eseri" Or key = "kitabi" Then kind = "work"
                    If key = "donem" Then kind = "period"
                    VdiFact facts, kind, Mid$(CStr(part), at + 1), CStr(part)
                End If
            End If
        Next part
    Next item
    ' Quoted source phrases are shown with the original source sentence.
    For j = 1 To 3
        opening = Chr$(34): closing = opening
        If j = 2 Then opening = ChrW(&H201C): closing = ChrW(&H201D)
        If j = 3 Then opening = ChrW(&HAB): closing = ChrW(&HBB)
        first = InStr(1, text, opening, vbBinaryCompare)
        Do While first > 0
            last = InStr(first + 1, text, closing, vbBinaryCompare)
            If last = 0 Then Exit Do
            phrase = Mid$(text, first + 1, last - first - 1)
            at = first - 60: If at < 1 Then at = 1
            near = VefeyatNormalize(Mid$(text, at, last - at + 80))
            kind = "quote"
            If InStr(near, "eser") > 0 Or InStr(near, "kitap") > 0 Then kind = "work"
            If InStr(near, "hocasi") > 0 Or InStr(near, "talebesi") > 0 Or InStr(near, "ogrencisi") > 0 Then
                If kind = "work" Then kind = "quote" Else kind = "relation"
            End If
            VdiFact facts, kind, phrase, text
            first = InStr(last + 1, text, opening, vbBinaryCompare)
        Loop
    Next j
    ' A nisba in the name is not proof of a school: use the description only.
    For Each item In Array("Hanefi", VfText("\u015eafii"), "Maliki", "Hanbeli", VfText("E\u015f'ari"), "Maturidi", VfText("Mu'tezili"))
        key = Replace(VefeyatNormalize(CStr(item)), " ", "")
        near = Replace(VefeyatNormalize(text), " ", "")
        If InStr(1, near, key, vbBinaryCompare) > 0 Then VdiFact facts, "school", CStr(item), text
    Next item
    For Each item In Array("hekim", "tabip", "filozof", "fakih", VfText("kel\u00e2mc\u0131s\u0131"), VfText("m\u00fcfessir"), VfText("muhaddis"), VfText("\u015fair"), VfText("tarih\u00e7i"), VfText("mutasavv\u0131f"))
        If VdiContains(text, CStr(item)) Then VdiFact facts, "profession", CStr(item), text
    Next item
    Set VdiFacts = facts
End Function

Private Function VdiNegated(ByVal text As String, ByVal phrase As String) As Boolean
    VdiNegated = VdiContains(text, phrase & " degil") Or VdiContains(text, phrase & " olmayan")
End Function

Private Function VdiEvaluate(ByRef row As Variant, ByVal query As String, ByVal sentence As String, ByVal nearby As String) As Object
    Dim result As Object, facts As Object, evidence As New Collection, fact As Variant, key As Variant
    Dim label As String, scope As String, score As Long, points As Long, conflict As Boolean
    Set result = CreateObject("Scripting.Dictionary")
    Set facts = VdiFacts(row, query)
    For Each key In facts.Keys
        fact = facts(key): scope = "": points = 0
        If VdiContains(sentence, CStr(fact(1))) Then
            scope = VfText("Ayn\u0131 c\u00fcmle")
            Select Case CStr(fact(0))
                Case "name": points = 80
                Case "work": points = 50
                Case "relation": points = 30
                Case Else: points = 8
            End Select
        ElseIf VdiContains(nearby, CStr(fact(1))) Then
            scope = VfText("Yak\u0131n paragraf")
            Select Case CStr(fact(0))
                Case "name": points = 25
                Case "work": points = 18
                Case "relation": points = 10
                Case Else: points = 2
            End Select
        End If
        If points > 0 Then
            conflict = VdiNegated(sentence, CStr(fact(1)))
            If scope = VfText("Yak\u0131n paragraf") Then conflict = VdiNegated(nearby, CStr(fact(1)))
            If conflict Then points = 0
            Select Case CStr(fact(0))
                Case "name": label = VfText("Kay\u0131tl\u0131 a\u00e7\u0131k ad")
                Case "work": label = VfText("Kaynakta ge\u00e7en eser ifadesi")
                Case "relation": label = VfText("Kaynakta ge\u00e7en hoca-talebe ifadesi")
                Case "profession": label = VfText("Kaynakta kay\u0131tl\u0131 u\u011fra\u015f alan\u0131 (zay\u0131f)")
                Case "period": label = VfText("Kaynakta kay\u0131tl\u0131 d\u00f6nem (zay\u0131f)")
                Case "school": label = VfText("Mezhep/ekol ipucu (zay\u0131f)")
                Case Else: label = VfText("Kaynak al\u0131nt\u0131s\u0131 (zay\u0131f)")
            End Select
            evidence.Add Array(CStr(key), label, CStr(fact(1)), scope, CStr(fact(2)), points, conflict)
            score = score + points
        End If
    Next key
    ' Only an exact explicit death expression is a temporal hint. Do not infer
    ' teachers, lifetimes, calendar conversions or historical contradictions.
    If Len(CStr(row(3))) > 0 Then
        If InStr(1, sentence, CStr(row(3)), vbTextCompare) > 0 Then
            evidence.Add Array("period", VfText("Kay\u0131tl\u0131 tarih ifadesi (zay\u0131f)"), CStr(row(3)), VfText("Ayn\u0131 c\u00fcmle"), CStr(row(3)), 5, False)
            score = score + 5
        End If
    End If
    result("Score") = score: Set result("Evidence") = evidence
    Set VdiEvaluate = result
End Function

Private Function VdiWindow(ByVal named As Word.Range, ByVal paragraph As Boolean) As Word.Range
    Dim result As Word.Range, first As Long, last As Long
    Set result = named.Duplicate
    If paragraph Then
        Call result.MoveStart(wdParagraph, -2)
        Call result.MoveEnd(wdParagraph, 2)
    Else
        Set result = named.Sentences(1).Duplicate
    End If
    first = result.Start: last = result.End
    If first < named.Start - 700 Then first = named.Start - 700
    If last > named.End + 700 Then last = named.End + 700
    result.SetRange first, last
    result.TextRetrievalMode.IncludeFieldCodes = False
    Set VdiWindow = result
End Function

Private Function VdiNamed(ByVal target As Word.Range, ByVal query As String) As Word.Range
    Dim result As Word.Range, at As Long, text As String
    If target Is Nothing Then Exit Function
    text = target.Text
    at = InStr(1, text, query, vbBinaryCompare)
    If at < 1 Or Len(query) = 0 Then Exit Function
    Set result = target.Duplicate
    result.SetRange target.Start + at - 1, target.Start + at - 1 + Len(query)
    Set VdiNamed = result
End Function

Private Function VdiSignature(ByVal named As Word.Range) As String
    Dim a As String, b As String, closing As Long, field As Word.Field, point As Long
    point = named.End
    Set field = VfnAdjacentField(named.Document, named)
    If Not field Is Nothing Then point = field.Result.End + 1
    ' Index marks (XE) are hidden and are skipped, so adding or removing a
    ' dizin mark never invalidates an identity confirmation. Without marks the
    ' text is exactly the pre-0.3.40 signature text.
    a = VdiSpanText(named, named.Start, 100, True)
    b = VdiSpanText(named, point, 160, False)
    b = VfaTrimSpaces(b)
    If VfmDateText(b) Then
        closing = InStr(1, b, ")", vbBinaryCompare)
        If closing > 0 Then b = Mid$(b, closing + 1)
    End If
    VdiSignature = VfpKeyPart(VfaNormalizeWhitespace(a)) & VfpKeyPart(Left$(VfaNormalizeWhitespace(b), 80))
End Function

Private Function VdiRecordName(ByVal number As Long) As String
    VdiRecordName = VdiPrefix & CStr(number)
End Function

Private Function VdiBookmark(ByVal number As Long) As String
    VdiBookmark = "_Vdc_" & CStr(number)
End Function

Private Function VdiRead(ByVal doc As Word.Document) As Object
    Dim result As Object, variable As Word.Variable, name As String, suffix As String, number As Long, values As Collection
    Set result = CreateObject("Scripting.Dictionary")
    Call VdiRevisionValue(doc)
    For Each variable In doc.Variables
        name = variable.Name
        If StrComp(Left$(name, Len(VdiPrefix)), VdiPrefix, vbTextCompare) = 0 Then
            suffix = Mid$(name, Len(VdiPrefix) + 1)
            If Len(suffix) < 1 Or Len(suffix) > 3 Or Not IsNumeric(suffix) Then Err.Raise VdiInvalid, "Vefeyat", "Invalid contextual identity metadata."
            number = CLng(suffix)
            If number < 1 Or number > VdiLimit Or CStr(number) <> suffix Then Err.Raise VdiInvalid, "Vefeyat", "Invalid contextual identity slot."
            Set values = VfuUnpack(VdiDecode(variable.Value))
            If values.Count <> 6 Then Err.Raise VdiInvalid, "Vefeyat", "Invalid contextual identity record."
            If CStr(values(1)) <> VdiBookmark(number) Or Not VfaValidAlias(CStr(values(2))) Then Err.Raise VdiInvalid, "Vefeyat", "Invalid contextual identity anchor."
            If Len(CStr(values(3))) = 0 Or Len(CStr(values(3))) > 4000 Then Err.Raise VdiInvalid, "Vefeyat", "Invalid contextual identity key."
            If result.Exists(CStr(number)) Then Err.Raise VdiInvalid, "Vefeyat", "Conflicting contextual identity slots."
            Set result(CStr(number)) = values
        End If
    Next variable
    Set VdiRead = result
End Function

Private Function VdiLiveRecord(ByVal doc As Word.Document, ByVal values As Collection) As Word.Range
    Dim named As Word.Range, field As Word.Field, id As String
    If Not doc.Bookmarks.Exists(CStr(values(1))) Then Exit Function
    Set named = doc.Bookmarks(CStr(values(1))).Range
    If StrComp(named.Text, CStr(values(2)), vbBinaryCompare) <> 0 Then Exit Function
    If VdiSignature(named) <> CStr(values(5)) Then Exit Function
    If CStr(named.StoryType) <> CStr(values(6)) Then Exit Function
    Set field = VfnAdjacentField(doc, named)
    If Not field Is Nothing Then
        id = VfaFieldBinding(doc, field)
        If VfaReadVariable(doc, VfaVariableName(id, "Key")) <> CStr(values(3)) Then Exit Function
    End If
    Set VdiLiveRecord = named.Duplicate
End Function

Private Function VdiBuild(ByVal people As Collection, ByVal query As String, ByVal target As Word.Range) As Object
    Dim proof As Object, evaluations As Object, records As Object, named As Word.Range, window As Word.Range
    Dim sentence As String, nearby As String, row As Variant, item As Object, i As Long, j As Long, remembered As Long
    Dim order() As Long, scores() As Long, indices() As Long, ranked As New Collection, counts As Object
    Dim key As Variant, record As Collection, live As Word.Range, evidence As Variant, bucket As Object, person As String, swap As Long
    Set proof = CreateObject("Scripting.Dictionary"): Set evaluations = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    Set named = VdiNamed(target, query)
    proof("Query") = query: proof("Remembered") = 0&: proof("Prior") = ""
    Set proof("Catalog") = mVfRows
    If Not named Is Nothing Then
        Set proof("Named") = named: Set proof("Doc") = named.Document
        Set window = VdiWindow(named, False): sentence = window.Text
        Set window = VdiWindow(named, True): nearby = window.Text
        Set proof("Window") = window: proof("Snapshot") = nearby: proof("Literal") = named.Text
        proof("Settings") = VfpGetScope(named.Document) & vbNullChar & VfpGetPrefix(named.Document)
        proof("Revision") = VdiRevisionValue(named.Document)
        Set records = VdiRead(named.Document)
    Else
        Set records = CreateObject("Scripting.Dictionary")
    End If
    proof("Sentence") = sentence: proof("Nearby") = nearby
    ReDim order(0 To people.Count): ReDim scores(0 To people.Count): ReDim indices(0 To people.Count)
    For i = 1 To people.Count
        row = people(i): person = VfpPersonKey(row)
        If i <= 128 Then
            Set item = VdiEvaluate(row, query, sentence, nearby)
        Else
            Set item = CreateObject("Scripting.Dictionary")
            Set item("Evidence") = New Collection: item("Score") = 0&: item("Limited") = True
        End If
        Set evaluations(person) = item
        For Each evidence In item("Evidence")
            key = CStr(evidence(0))
            If Not counts.Exists(CStr(key)) Then counts(CStr(key)) = 0&
            counts(CStr(key)) = CLng(counts(CStr(key))) + 1
        Next evidence
        order(i) = i: scores(i) = CLng(item("Score"))
        ' Reuse an exact, live, source-unchanged occurrence confirmation only.
        ' Other occurrences are contextual proposals and always need approval.
        For Each key In records.Keys
            Set record = records(key)
            If CStr(record(3)) = person And CStr(record(4)) = VfqRowStamp(row) Then
                If Not named Is Nothing Then
                    If VfaAliasKey(CStr(record(2))) = VfaAliasKey(query) Then
                        item("Prior") = True
                        Set live = Nothing
                        If named.Document.Bookmarks.Exists(CStr(record(1))) Then
                            Set live = named.Document.Bookmarks(CStr(record(1))).Range
                            If live.IsEqual(named) Then Set live = VdiLiveRecord(named.Document, record) Else Set live = Nothing
                        End If
                        If Not live Is Nothing Then
                            If live.IsEqual(named) Then
                                If CLng(proof("Remembered")) = 0 Then
                                    proof("Remembered") = i
                                ElseIf CLng(proof("Remembered")) <> i Then
                                    proof("Remembered") = -1&
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next key
    Next i
    ' Shared cues have no discriminatory rank. Keep original lookup order on ties.
    For i = 1 To people.Count
        row = people(i): Set item = evaluations(VfpPersonKey(row))
        For Each evidence In item("Evidence")
            If CLng(counts(CStr(evidence(0)))) > 1 Then scores(i) = scores(i) - CLng(evidence(5))
        Next evidence
        If scores(i) > 0 And item.Exists("Prior") Then scores(i) = scores(i) + 1
    Next i
    remembered = CLng(proof("Remembered"))
    If remembered > 0 Then
        For i = 1 To people.Count
            If i <> remembered And scores(i) >= 25 And scores(i) > scores(remembered) Then
                proof("Remembered") = 0&: proof("MemoryConflict") = True
                Exit For
            End If
        Next i
        If CLng(proof("Remembered")) > 0 Then scores(remembered) = scores(remembered) + 100000
    End If
    For i = 2 To people.Count
        j = i
        Do While j > 1
            If scores(order(j - 1)) >= scores(order(j)) Then Exit Do
            swap = order(j - 1): order(j - 1) = order(j): order(j) = swap: j = j - 1
        Loop
    Next i
    For i = 1 To people.Count
        row = people(order(i)): ranked.Add row: indices(i) = order(i)
    Next i
    Set proof("Rows") = ranked: Set proof("Evidence") = evaluations: Set proof("Counts") = counts
    proof("Indices") = indices
    Set VdiBuild = proof
End Function

Private Function VdiDescribe(ByVal proof As Object, ByRef row As Variant) As String
    Dim person As String, item As Object, evidence As Variant, text As String, counts As Object
    If proof Is Nothing Then Exit Function
    person = VfpPersonKey(row)
    If Not proof("Evidence").Exists(person) Then Exit Function
    Set item = proof("Evidence")(person): Set counts = proof("Counts")
    text = VfText("Ba\u011flamsal kimlik incelemesi") & vbCrLf & _
        VfText("Belgedeki c\u00fcmle: ") & Left$(CStr(proof("Sentence")), 600) & vbCrLf & _
        VfText("Yak\u0131n paragraflar: ") & Left$(CStr(proof("Nearby")), 900) & vbCrLf
    If item.Exists("Limited") Then text = text & VfText("\u00c7ok aday var; ilk 128 aday\u0131n ba\u011flam\u0131 incelendi. Bu aday elenmedi, ancak ba\u011flam\u0131 ayr\u0131ca incelenmeli. Daha a\u00e7\u0131k bir adla arayabilirsiniz.") & vbCrLf
    If proof.Exists("MemoryConflict") Then text = text & VfText("\u00d6nceki onayla yeni ba\u011flam ipu\u00e7lar\u0131 farkl\u0131 adaylar\u0131 destekliyor; yeniden onay gerekir.") & vbCrLf
    If item("Evidence").Count = 0 Then text = text & VfText("Ay\u0131r\u0131c\u0131 kaynak kan\u0131t\u0131 bulunamad\u0131. Benzer ad ayn\u0131 \u015fah\u0131s demek de\u011fildir.") & vbCrLf
    For Each evidence In item("Evidence")
        text = text & CStr(evidence(1)) & ": " & CStr(evidence(2)) & " | " & CStr(evidence(3))
        If CBool(evidence(6)) Then text = text & VfText(" | olumsuz ifade; destek say\u0131lmad\u0131")
        If CLng(counts(CStr(evidence(0)))) > 1 Then text = text & VfText(" | birden fazla adayda var; ay\u0131r\u0131c\u0131 de\u011fil")
        text = text & vbCrLf & VfText("Kaynak al\u0131nt\u0131s\u0131: ") & CStr(evidence(4)) & vbCrLf
    Next evidence
    If item.Exists("Prior") Then text = text & VfText("Bu belgede bu yaz\u0131m i\u00e7in \u00f6nceki onay var; yeni ge\u00e7i\u015fe otomatik ta\u015f\u0131nmaz.") & vbCrLf
    text = text & VfText("Yeni belirsiz ge\u00e7i\u015f ayr\u0131ca onaylan\u0131r. Bu i\u015flem kimli\u011fi t\u00fcm ayn\u0131 yaz\u0131mlara topluca atamaz.") & vbCrLf
    text = text & VfText("Kaynak: ") & CStr(row(5)) & vbCrLf & _
        VfText("Puan/s\u0131ra kesin kimlik kan\u0131t\u0131 de\u011fildir. Eksik eser, ili\u015fki veya d\u00f6nem bilgisi \u00fcretilmedi.")
    VdiDescribe = text
End Function

Private Function VdiCurrent(ByVal proof As Object) As Boolean
    Dim named As Word.Range, window As Word.Range, doc As Word.Document
    mVdiReason = VfText("Se\u00e7ili ad metinde bulunamad\u0131. Ad\u0131n tamam\u0131n\u0131 se\u00e7ip yeniden deneyin.")
    If proof Is Nothing Then Exit Function
    If Not proof.Exists("Named") Then Exit Function
    Set doc = proof("Doc"): Set named = proof("Named"): Set window = proof("Window")
    mVdiReason = VfText("Belge onay s\u0131ras\u0131nda etkin de\u011fildi. Belgeye d\u00f6n\u00fcp yeniden deneyin.")
    If Not VftOpenDoc(doc) Or Not (ActiveDocument Is doc) Then Exit Function
    mVdiReason = VfText("Belge salt okunur veya korumal\u0131.")
    If doc.ReadOnly Or doc.ProtectionType <> wdNoProtection Then Exit Function
    mVdiReason = VfText("CSV katalo\u011fu onay s\u0131ras\u0131nda de\u011fi\u015fti. Ad\u0131 yeniden se\u00e7in.")
    If Not (proof("Catalog") Is mVfRows) Then Exit Function
    mVdiReason = VfText("Ad\u0131n \u00e7evresindeki metin onay s\u0131ras\u0131nda de\u011fi\u015fti. Ad\u0131 yeniden se\u00e7ip tekrar deneyin.")
    If StrComp(named.Text, CStr(proof("Literal")), vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(window.Text, CStr(proof("Snapshot")), vbBinaryCompare) <> 0 Then Exit Function
    mVdiReason = VfText("Tarih ekleme veya k\u0131saltma ayar\u0131 onay s\u0131ras\u0131nda de\u011fi\u015fti.")
    If VfpGetScope(doc) & vbNullChar & VfpGetPrefix(doc) <> CStr(proof("Settings")) Then Exit Function
    mVdiReason = VfText("Bu belgede ba\u015fka bir kimlik onay\u0131 ayn\u0131 anda kaydedildi. Ad\u0131 yeniden se\u00e7in.")
    If VdiRevisionValue(doc) <> CStr(proof("Revision")) Then Exit Function
    mVdiReason = ""
    VdiCurrent = True
End Function

Private Function VdiRowCurrent(ByRef row As Variant) As Boolean
    Dim candidate As Variant, key As String, stamp As String
    If mVfRows Is Nothing Then Exit Function
    key = VfpPersonKey(row): stamp = VfqRowStamp(row)
    For Each candidate In mVfRows
        If VfpPersonKey(candidate) = key Then
            If VfqRowStamp(candidate) = stamp Then VdiRowCurrent = True: Exit Function
        End If
    Next candidate
End Function

Private Sub VdiStore(ByVal named As Word.Range, ByRef row As Variant)
    Dim doc As Word.Document, records As Object, key As Variant, values As Collection, live As Word.Range
    Dim slot As Long, i As Long, payload As String, revision As Long
    Set doc = named.Document
    If Not VfaEditable(doc) Or Not VdiRowCurrent(row) Then Err.Raise VdiInvalid, "Vefeyat", "Identity confirmation changed; nothing was stored."
    Set records = VdiRead(doc)
    For Each key In records.Keys
        Set values = records(key)
        If doc.Bookmarks.Exists(CStr(values(1))) Then
            Set live = doc.Bookmarks(CStr(values(1))).Range
            If live.IsEqual(named) Then slot = CLng(key): Exit For
        End If
    Next key
    If slot = 0 Then
        For i = 1 To VdiLimit
            If Not records.Exists(CStr(i)) Then slot = i: Exit For
        Next i
    End If
    If slot = 0 Then Err.Raise VdiInvalid, "Vefeyat", "Document identity confirmation limit reached. Remove an unused confirmation."
    Set values = New Collection
    values.Add VdiBookmark(slot): values.Add named.Text: values.Add VfpPersonKey(row)
    values.Add VfqRowStamp(row): values.Add VdiSignature(named): values.Add CStr(named.StoryType)
    payload = VdiEncode(VfuPack(values))
    doc.Bookmarks.Add Name:=VdiBookmark(slot), Range:=named.Duplicate
    VfaWriteVariable doc, VdiRecordName(slot), payload
    revision = Val(VdiRevisionValue(doc))
    If revision < 0 Or revision >= 2000000000 Then revision = 0
    VfaWriteVariable doc, VdiRevisionName, CStr(revision + 1)
End Sub

Private Function VdiNewBinding(ByVal doc As Word.Document, ByRef row As Variant, ByVal aliasText As String, ByRef changed As Boolean) As String
    Dim number As Long, value As String, id As String
    number = 1: value = VfaReadVariable(doc, VfaNextVariable)
    If IsNumeric(value) Then
        If CDbl(value) >= 1 And CDbl(value) < 2000000000# Then number = CLng(value)
    End If
    Do
        id = "P" & CStr(number): number = number + 1
    Loop While VfaVariableExists(doc, VfaVariableName(id, "Key"))
    changed = True
    VfaWriteVariable doc, VfaNextVariable, CStr(number)
    VfaWriteVariable doc, VfaVariableName(id, "Alias"), aliasText
    VfaWriteVariable doc, VfaVariableName(id, "Raw"), CStr(row(3))
    VfaWriteVariable doc, VfaVariableName(id, "OccurrenceOnly"), "1"
    VfaWriteVariable doc, VfaVariableName(id, "Key"), VfpPersonKey(row)
    VdiNewBinding = id
End Function

Private Sub VdiInsertOne(ByVal doc As Word.Document, ByVal target As Word.Range, ByVal aliasText As String, ByRef row As Variant, ByRef inserted As Long, ByRef skipped As Long, ByRef changed As Boolean)
    Dim named As Word.Range, id As String, first As Long
    Set named = VdiNamed(target, aliasText): If named Is Nothing Then Exit Sub
    If Not VfaEditable(doc) Or Not VfpDateAllowed(doc) Then Exit Sub
    If Not VdiRowCurrent(row) Then Err.Raise VdiInvalid, "Vefeyat", "Catalog changed; date was not inserted."
    If VfmDatePresent(doc, named) Or VfbAnyDate(named) Then mVfaSkipReason = VfaDateBlockReason(named): skipped = skipped + 1: Exit Sub
    first = named.Start
    VdiCheckStore named, row
    VfaEnsureNamespace doc, changed
    id = VdiNewBinding(doc, row, aliasText, changed)
    VfaInsertManaged doc, named, VfpFormatDate(CStr(row(3)), VfpGetPrefix(doc)), id, changed, True
    inserted = inserted + 1
    named.SetRange first, first + Len(aliasText)
    VdiStore named, row
    VfmRemember doc, named, VfpPersonKey(row)
End Sub

Private Sub VdiCheckStore(ByVal named As Word.Range, ByRef row As Variant)
    Dim records As Object, key As Variant, values As Collection, live As Word.Range, packed As New Collection
    If Not VdiRowCurrent(row) Then Err.Raise VdiInvalid, "Vefeyat", "Catalog changed; reconfirm the person."
    If Not VfaEditable(named.Document) Then Err.Raise VdiInvalid, "Vefeyat", "Document is not editable."
    packed.Add VdiBookmark(1): packed.Add named.Text: packed.Add VfpPersonKey(row)
    packed.Add VfqRowStamp(row): packed.Add VdiSignature(named): packed.Add CStr(named.StoryType)
    Call VdiEncode(VfuPack(packed))
    Set records = VdiRead(named.Document)
    If records.Count >= VdiLimit Then
        For Each key In records.Keys
            Set values = records(key)
            If named.Document.Bookmarks.Exists(CStr(values(1))) Then
                Set live = named.Document.Bookmarks(CStr(values(1))).Range
                If live.IsEqual(named) Then Exit Sub
            End If
        Next key
        Err.Raise VdiInvalid, "Vefeyat", "Document identity confirmation limit reached; remove an unused confirmation first."
    End If
End Sub

Private Sub VdiCorrect(ByVal named As Word.Range, ByRef row As Variant)
    Dim doc As Word.Document, proof As Object, people As New Collection, field As Word.Field
    Dim previousCode As String, previousResult As String, previousKey As String, id As String, preview As String, formatted As String
    Dim started As Boolean, undoObject As Object, changed As Boolean, number As Long, detail As String
    On Error GoTo Failed
    Set doc = named.Document
    If Not VfaEditable(doc) Or Not VfpDateAllowed(doc) Then Exit Sub
    people.Add row: Set proof = VdiBuild(people, named.Text, named)
    VdiCheckStore named, row
    Set field = VfnAdjacentField(doc, named)
    preview = VfText("Bu ge\u00e7i\u015fin kimlik onay\u0131 de\u011fi\u015ftirilsin mi?") & vbCrLf & _
        named.Text & " -> " & CStr(row(0)) & vbCrLf & CStr(row(4)) & vbCrLf & CStr(row(5))
    If Not field Is Nothing Then
        If Not VfaValidDate(CStr(row(3))) Then Err.Raise VdiInvalid, "Vefeyat", "The selected catalog record has no usable date."
        previousCode = field.Code.Text: previousResult = field.Result.Text
        previousKey = VfaReadVariable(doc, VfaVariableName(VfaFieldBinding(doc, field), "Key"))
        formatted = " " & VfpFormatDate(CStr(row(3)), VfpGetPrefix(doc))
        preview = preview & vbCrLf & VfText("Yaln\u0131z bu ge\u00e7i\u015fteki y\u00f6netilen tarih: ") & previousResult & " -> " & formatted
    Else
        preview = preview & vbCrLf & VfText("Ad ve elle yaz\u0131lm\u0131\u015f tarihler korunur; yaln\u0131z kimlik onay\u0131 kaydedilir.")
    End If
    If VfMessage(preview, vbYesNo Or vbQuestion Or vbDefaultButton2, VfText("Kimlik e\u015fle\u015fmesini d\u00fczelt")) <> vbYes Then Exit Sub
    If Not VdiCurrent(proof) Or Not VfaEditable(doc) Or Not VdiRowCurrent(row) Then Exit Sub
    If Not field Is Nothing Then
        If field.Code.Text <> previousCode Or field.Result.Text <> previousResult Then Exit Sub
        If VfaReadVariable(doc, VfaVariableName(VfaFieldBinding(doc, field), "Key")) <> previousKey Then Exit Sub
    End If
    started = VfaBeginUndo(undoObject)
    If Not field Is Nothing Then
        VfaEnsureNamespace doc, changed
        id = VdiNewBinding(doc, row, named.Text, changed)
        changed = True
        field.Locked = False: field.Result.Text = formatted
        field.Code.Text = " ADDIN " & VfaFieldTag(doc, id) & " "
        field.ShowCodes = False: field.Locked = True
        If field.Result.Text <> formatted Then Err.Raise VdiInvalid, "Vefeyat", "Word did not retain the corrected date."
    End If
    VdiStore named, row
    VfmRemember doc, named, VfpPersonKey(row)
    VfaEndUndo undoObject, started
    VfmInvalidate doc
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not field Is Nothing Then field.Locked = True
    VfaEndUndo undoObject, started
    On Error GoTo 0
    VfaReportFailure detail, changed
End Sub

Private Sub VdiManage()
    Dim doc As Word.Document, records As Object, key As Variant, values As Collection, view As New Collection, slots As New Collection
    Dim picker As VefeyatKisiSec0341, chosen As Long, named As Word.Range, people As Collection, row As Variant, found As Variant
    Dim title As String, revision As String, number As Long, detail As String, count As Long, lookup As Object
    On Error GoTo Failed
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument: Set records = VdiRead(doc)
    revision = VdiRevisionValue(doc)
    If records.Count = 0 Then VfaWarn VfText("Bu belgede ba\u011flamsal kimlik onay\u0131 yok."): Exit Sub
    If Not VefeyatEnsureData() Then Exit Sub
    Set lookup = CreateObject("Scripting.Dictionary")
    For Each found In mVfRows
        lookup(VfpPersonKey(found)) = found
    Next found
    For Each key In records.Keys
        Set values = records(key): title = VfText("Kaynak de\u011fi\u015fmi\u015f veya ge\u00e7i\u015f silinmi\u015f")
        If lookup.Exists(CStr(values(3))) Then
            found = lookup(CStr(values(3))): title = CStr(found(0))
            If VfqRowStamp(found) <> CStr(values(4)) Then title = title & VfText(" [kaynak de\u011fi\u015fti]")
        End If
        If Not doc.Bookmarks.Exists(CStr(values(1))) Then title = title & VfText(" [ge\u00e7i\u015f silindi]")
        view.Add Array(CStr(values(2)), title, "", "", _
            VfText("Belgeye \u00f6zg\u00fc onay. Se\u00e7ili kayd\u0131n ge\u00e7i\u015fine gidip kimli\u011fi yeniden se\u00e7ebilirsiniz. Kaynak bilgisi de\u011fi\u015fti\u011finde eski onay otomatik kullan\u0131lmaz. Onay\u0131 kald\u0131rmak i\u00e7in metindeki ad\u0131 se\u00e7ip Onay\u0131 kald\u0131r komutunu kullan\u0131n."), "")
        slots.Add CStr(key)
    Next key
    Set picker = New VefeyatKisiSec0341
    picker.InitCandidates view, VfText("Bu belgenin kimlik e\u015fle\u015ftirmeleri")
    picker.SetIdentityManageCaption
    picker.Show vbModal: chosen = picker.ChosenIndex
    Unload picker: Set picker = Nothing
    If chosen < 1 Or chosen > slots.Count Then Exit Sub
    If Not VftOpenDoc(doc) Or Not (ActiveDocument Is doc) Then Exit Sub
    If VdiRevisionValue(doc) <> revision Then Exit Sub
    Set values = records(CStr(slots(chosen)))
    If Not doc.Bookmarks.Exists(CStr(values(1))) Then
        If VfMessage(VfText("Ge\u00e7i\u015f silinmi\u015f. Yaln\u0131z bu kullan\u0131lmayan onay kayd\u0131 kald\u0131r\u0131ls\u0131n m\u0131?"), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") = vbYes Then
            VdiDeleteRecord doc, CLng(slots(chosen)), revision
        End If
        Exit Sub
    End If
    Set named = doc.Bookmarks(CStr(values(1))).Range.Duplicate
    If StrComp(named.Text, CStr(values(2)), vbBinaryCompare) <> 0 Then VfaWarn VfText("Ad de\u011fi\u015fmi\u015f; metindeki ad\u0131 yeniden se\u00e7in."): Exit Sub
    named.Select
    Application.ActiveWindow.ScrollIntoView named, True
    Set people = VfaConfirmationCandidates(VefeyatFind(named.Text, False), named.Text)
    If people.Count = 0 Then Exit Sub
    chosen = VfePickPerson(people, named.Text, False, True)
    If chosen < 1 Or chosen > people.Count Then Exit Sub
    row = people(chosen)
    VdiCorrect named, row
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not picker Is Nothing Then Unload picker
    On Error GoTo 0
    Err.Raise number, "Vefeyat", detail
End Sub

Private Sub VdiForgetSelected()
    Dim doc As Word.Document, target As Word.Range, named As Word.Range, original As String, aliasText As String, reason As String
    Dim first As Long, last As Long, story As Long, records As Object, key As Variant, values As Collection, live As Word.Range
    Dim revision As String, selectedKey As String, started As Boolean, undoObject As Object
    If Not VfaCapture(doc, target, original, aliasText, first, last, story) Then Exit Sub
    Set named = VdiNamed(target, aliasText): Set records = VdiRead(doc)
    revision = VdiRevisionValue(doc)
    For Each key In records.Keys
        Set values = records(key)
        If doc.Bookmarks.Exists(CStr(values(1))) Then
            Set live = doc.Bookmarks(CStr(values(1))).Range
            If live.IsEqual(named) Then selectedKey = CStr(key): Exit For
        End If
    Next key
    If Len(selectedKey) = 0 Then VfaWarn VfText("Se\u00e7ili ge\u00e7i\u015fte kay\u0131tl\u0131 kimlik onay\u0131 yok."): Exit Sub
    If VfMessage(VfText("Yaln\u0131z se\u00e7ili ge\u00e7i\u015fin kimlik onay\u0131 kald\u0131r\u0131ls\u0131n m\u0131? Ad ve tarih silinmez."), _
        vbYesNo Or vbQuestion Or vbDefaultButton2, VfText("Kimlik onay\u0131n\u0131 kald\u0131r")) <> vbYes Then Exit Sub
    If Not VfaTargetCurrent(doc, target, original, first, last, story, reason) Then Exit Sub
    If VdiRevisionValue(doc) <> revision Then Exit Sub
    started = VfaBeginUndo(undoObject)
    On Error GoTo Failed
    VfaDeleteVariable doc, VdiRecordName(CLng(selectedKey))
    If doc.Bookmarks.Exists(VdiBookmark(CLng(selectedKey))) Then doc.Bookmarks(VdiBookmark(CLng(selectedKey))).Delete
    VfaWriteVariable doc, VdiRevisionName, CStr((Val(revision) Mod 2000000000) + 1)
    VfaEndUndo undoObject, started
    Exit Sub
Failed:
    VfaEndUndo undoObject, started
    Err.Raise VdiInvalid, "Vefeyat", "The identity confirmation could not be removed."
End Sub

Private Sub VdiDeleteRecord(ByVal doc As Word.Document, ByVal slot As Long, ByVal expectedRevision As String)
    Dim started As Boolean, undoObject As Object
    If slot < 1 Or slot > VdiLimit Then Exit Sub
    If Not VftOpenDoc(doc) Or Not (ActiveDocument Is doc) Then Exit Sub
    If Not VfaEditable(doc) Then Exit Sub
    If VdiRevisionValue(doc) <> expectedRevision Then Exit Sub
    started = VfaBeginUndo(undoObject)
    On Error GoTo Failed
    VfaDeleteVariable doc, VdiRecordName(slot)
    If doc.Bookmarks.Exists(VdiBookmark(slot)) Then doc.Bookmarks(VdiBookmark(slot)).Delete
    VfaWriteVariable doc, VdiRevisionName, CStr((Val(expectedRevision) Mod 2000000000) + 1)
    VfaEndUndo undoObject, started
    Exit Sub
Failed:
    VfaEndUndo undoObject, started
    Err.Raise VdiInvalid, "Vefeyat", "The selected identity record could not be removed."
End Sub

Private Function VdiRevisionValue(ByVal doc As Word.Document) As String
    Dim value As String, i As Long, ch As String
    value = VfaReadVariable(doc, VdiRevisionName)
    If Len(value) > 10 Then Err.Raise VdiInvalid, "Vefeyat", "Invalid identity revision."
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch < "0" Or ch > "9" Then Err.Raise VdiInvalid, "Vefeyat", "Invalid identity revision."
    Next i
    If Val(value) > 2000000000 Then Err.Raise VdiInvalid, "Vefeyat", "Identity revision limit."
    VdiRevisionValue = value
End Function

' XML-attribute-safe outer encoding. Decode exactly once; source percent text
' and multiline catalog descriptions must round-trip without normalization.
Private Function VdiEncode(ByVal value As String) As String
    Dim i As Long, ch As String, code As Long, text As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1): code = AscW(ch) And &HFFFF&
        Select Case code
            Case 37: text = text & "%25"
            Case 9: text = text & "%09"
            Case 10: text = text & "%0A"
            Case 13: text = text & "%0D"
            Case 133: text = text & "%85"
            Case 8232: text = text & "%LS"
            Case 8233: text = text & "%PS"
            Case Else
                If code < 32 Or code = 127 Then
                    ' Word story text carries control marks next to names (footnote
                    ' reference 2, picture 1, cell end 7, line break 11, page break 12,
                    ' field marks 19-21, hyphens 30/31). Keep them losslessly.
                    text = text & "%" & Right$("0" & Hex$(code), 2)
                Else
                    text = text & ch
                End If
        End Select
    Next i
    If Len(text) > 30000 Then Err.Raise VdiInvalid, "Vefeyat", "Identity metadata record too large."
    VdiEncode = "Vdc1:" & text
End Function

' Decodes %00-%1F and %7F written by VdiEncode since 0.3.38. Any other
' escape is still rejected as corrupt metadata.
Private Function VdiControlEscape(ByVal token As String) As String
    Dim code As Long, i As Long, digit As Long
    If Len(token) <> 3 Or Left$(token, 1) <> "%" Then Err.Raise VdiInvalid, "Vefeyat", "Invalid identity escape."
    For i = 2 To 3
        digit = InStr(1, "0123456789ABCDEF", Mid$(token, i, 1), vbBinaryCompare) - 1
        If digit < 0 Then Err.Raise VdiInvalid, "Vefeyat", "Invalid identity escape."
        code = code * 16 + digit
    Next i
    If code = 9 Or code = 10 Or code = 13 Then Err.Raise VdiInvalid, "Vefeyat", "Invalid identity escape."
    If code >= 32 And code <> 127 Then Err.Raise VdiInvalid, "Vefeyat", "Invalid identity escape."
    VdiControlEscape = ChrW(code)
End Function

Private Function VdiDecode(ByVal value As String) As String
    Dim i As Long, token As String, text As String
    If Left$(value, 5) = "Vfu1:" Then
        If InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then Err.Raise VdiInvalid, "Vefeyat", "Unsafe old identity metadata."
        VdiDecode = value: Exit Function
    End If
    If Left$(value, 5) <> "Vdc1:" Or Len(value) > 30005 Then Err.Raise VdiInvalid, "Vefeyat", "Invalid identity encoding."
    i = 6
    Do While i <= Len(value)
        If Mid$(value, i, 1) <> "%" Then
            text = text & Mid$(value, i, 1): i = i + 1
        Else
            token = Mid$(value, i, 3)
            Select Case token
                Case "%25": text = text & "%"
                Case "%09": text = text & vbTab
                Case "%0A": text = text & vbLf
                Case "%0D": text = text & vbCr
                Case "%85": text = text & ChrW(133)
                Case "%LS": text = text & ChrW(8232)
                Case "%PS": text = text & ChrW(8233)
                Case Else
                    text = text & VdiControlEscape(token)
            End Select
            i = i + 3
        End If
    Loop
    VdiDecode = text
End Function

Private Function VrgEnabled(ByVal doc As Word.Document) As Boolean
    On Error GoTo Disabled
    VrgEnabled = (doc.Variables.Item(VrgMode).Value = "1")
Disabled:
    Err.Clear
End Function

Private Function VrgApi(ByVal doc As Word.Document) As Boolean
    Dim bookmarks As Object, variables As Object, format As Long, count As Long
    On Error GoTo Unavailable
    Set bookmarks = CallByName(doc, "Bookmarks", VbGet)
    Set variables = CallByName(doc, "Variables", VbGet)
    count = bookmarks.Count: count = variables.Count
    format = CLng(CallByName(doc, "SaveFormat", VbGet))
    Select Case format
        Case 0, 1, 12 To 16: VrgApi = True
    End Select
Unavailable:
    Err.Clear
End Function

Private Function VrgName(ByVal slot As Long) As String
    VrgName = VrgPrefix & CStr(slot)
End Function

Private Function VrgBookmark(ByVal slot As Long) As String
    VrgBookmark = "_Vrg_" & CStr(slot)
End Function

Private Function VrgRead(ByVal doc As Word.Document) As Object
    Dim result As Object, variable As Word.Variable, suffix As String, slot As Long, values As Collection, row As Variant, id As String, parts As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each variable In doc.Variables
        If StrComp(Left$(variable.Name, Len(VrgPrefix)), VrgPrefix, vbTextCompare) = 0 Then
            suffix = Mid$(variable.Name, Len(VrgPrefix) + 1)
            If Len(suffix) < 1 Or Len(suffix) > 4 Or Not IsNumeric(suffix) Then Err.Raise VrgInvalid, "Vefeyat", "Invalid registry slot."
            slot = CLng(suffix)
            If slot < 1 Or slot > VrgLimit Or CStr(slot) <> suffix Or result.Exists(suffix) Then Err.Raise VrgInvalid, "Vefeyat", "Conflicting registry slot."
            Set values = VfuUnpack(VdiDecode(variable.Value))
            If values.Count <> 17 Then Err.Raise VrgInvalid, "Vefeyat", "Invalid registry record version."
            If CStr(values(1)) <> "Vrg1" Or CStr(values(2)) <> VrgBookmark(slot) Then Err.Raise VrgInvalid, "Vefeyat", "Invalid registry anchor."
            If Not VfaValidAlias(CStr(values(4))) Then Err.Raise VrgInvalid, "Vefeyat", "Invalid recorded spelling."
            row = Array(CStr(values(10)), CStr(values(11)), CStr(values(12)), CStr(values(13)), CStr(values(14)), CStr(values(15)))
            id = VfqIdentity(row)
            If Left$(id, 4) <> "DIA:" Or id <> CStr(values(3)) Then Err.Raise VrgInvalid, "Vefeyat", "Invalid durable source identity."
            parts = VfePersonParts(CStr(values(5)))
            If IsEmpty(parts) Then Err.Raise VrgInvalid, "Vefeyat", "Invalid legacy identity snapshot."
            row = Array(CStr(parts(1)), CStr(parts(2)), CStr(parts(3)), "", "", CStr(parts(0)))
            If VfqIdentity(row) <> id Then Err.Raise VrgInvalid, "Vefeyat", "Registry identity conflicts with the recorded source."
            If CStr(values(16)) <> "1" And CStr(values(16)) <> "2" Then Err.Raise VrgInvalid, "Vefeyat", "Invalid registry story."
            If CStr(values(17)) = "1" And Not VfaValidDate(CStr(values(8))) Then Err.Raise VrgInvalid, "Vefeyat", "Invalid registered date."
            If CStr(values(17)) <> "0" And CStr(values(17)) <> "1" Then Err.Raise VrgInvalid, "Vefeyat", "Invalid date ownership state."
            Set result(suffix) = values
        End If
    Next variable
    Set VrgRead = result
End Function

Private Function VrgRow(ByVal values As Collection) As Variant
    VrgRow = Array(CStr(values(10)), CStr(values(11)), CStr(values(12)), CStr(values(13)), CStr(values(14)), CStr(values(15)))
End Function

Private Function VrgIndex() As Object
    Dim row As Variant
    If mVfRows Is Nothing Then Set VrgIndex = CreateObject("Scripting.Dictionary"): Exit Function
    If Not (mVrgCatalogOwner Is mVfRows) Or mVrgCatalogCount <> mVfRows.Count Then
        Set mVrgCatalogIndex = VfqIndex(mVfRows)
        Set mVrgLegacyIndex = CreateObject("Scripting.Dictionary")
        For Each row In mVfRows: mVrgLegacyIndex(VfpPersonKey(row)) = row: Next row
        Set mVrgCatalogOwner = mVfRows: mVrgCatalogCount = mVfRows.Count
    End If
    Set VrgIndex = mVrgCatalogIndex
End Function

Private Function VrgSource(ByRef row As Variant) As String
    Dim index As Object, id As String
    If mVfRows Is Nothing Then Exit Function
    id = VfqIdentity(row)
    If Left$(id, 4) <> "DIA:" Then Exit Function
    Set index = VrgIndex()
    If Not index.Exists(id) Then Exit Function
    If index(id).Count <> 1 Then Exit Function
    If VfqRowStamp(index(id)(1)) <> VfqRowStamp(row) Then Exit Function
    VrgSource = id
End Function

Private Function VrgLive(ByVal doc As Word.Document, ByVal values As Collection) As Word.Range
    Dim named As Word.Range
    If Not doc.Bookmarks.Exists(CStr(values(2))) Then Exit Function
    Set named = doc.Bookmarks(CStr(values(2))).Range.Duplicate
    If named.StoryType <> CLng(values(16)) Then Exit Function
    If named.StoryType <> wdMainTextStory And named.StoryType <> wdFootnotesStory Then Exit Function
    Set VrgLive = named
End Function

Private Function VrgNameSlot(ByVal doc As Word.Document, ByVal named As Word.Range) As Long
    Dim records As Object, key As Variant, live As Word.Range
    On Error GoTo Unavailable
    If Not VrgEnabled(doc) Or mVrgPaused Then Exit Function
    Set records = VrgRead(doc)
    For Each key In records.Keys
        Set live = VrgLive(doc, records(key))
        If Not live Is Nothing Then
            If live.IsEqual(named) Then VrgNameSlot = CLng(key): Exit Function
        End If
    Next key
Unavailable:
End Function

' Record layout: version, bookmark, durable ID, last name, legacy key,
' date code/result/raw/binding, six source columns, story, managed-date flag.
Private Sub VrgStore(ByVal doc As Word.Document, ByVal named As Word.Range, ByRef row As Variant, _
                     ByVal kind As String, Optional ByVal knownSlot As Long = 0)
    Dim records As Object, values As New Collection, old As Collection, id As String, slot As Long, i As Long, key As Variant
    Dim live As Word.Range, field As Word.Field, code As String, result As String, raw As String, binding As String, owned As String, payload As String
    If Not VrgEnabled(doc) Or Not VrgApi(doc) Then Exit Sub
    If Not VfaEditable(doc) Or doc.TrackRevisions Or doc.Revisions.Count > 0 Then Exit Sub
    If Not VfaValidAlias(named.Text) Or Not VfaSupportedRange(doc, named) Then Err.Raise VrgInvalid, "Vefeyat", "Unsupported registry range."
    id = VrgSource(row)
    If Len(id) = 0 Then Err.Raise VrgInvalid, "Vefeyat", "A unique, current DIA source identity is required."
    Set records = VrgRead(doc)
    If knownSlot > 0 Then
        If Not records.Exists(CStr(knownSlot)) Then Err.Raise VrgInvalid, "Vefeyat", "Registry record changed."
        slot = knownSlot
    Else
        For Each key In records.Keys
            Set live = VrgLive(doc, records(key))
            If Not live Is Nothing Then
                If live.IsEqual(named) Then slot = CLng(key): Exit For
            End If
        Next key
    End If
    If slot > 0 Then
        Set old = records(CStr(slot))
        If CStr(old(3)) <> id Then Err.Raise VrgInvalid, "Vefeyat", "Detach the previous identity before relinking."
    Else
        For i = 1 To VrgLimit
            If Not records.Exists(CStr(i)) And Not doc.Bookmarks.Exists(VrgBookmark(i)) Then slot = i: Exit For
        Next i
        If slot = 0 Then Err.Raise VrgInvalid, "Vefeyat", "Registry capacity reached; existing text is preserved."
    End If
    owned = "0"
    Set field = VfnAdjacentField(doc, named)
    If Not field Is Nothing Then
        binding = VfaFieldBinding(doc, field)
        If VfaReadVariable(doc, VfaVariableName(binding, "Key")) = VfpPersonKey(row) Then
            code = field.Code.Text: result = field.Result.Text
            raw = VfaReadVariable(doc, VfaVariableName(binding, "Raw"))
            If VfaValidDate(raw) Then
                If VfpFormatDate(result, "o") = VfpFormatDate(" " & raw, "o") Then owned = "1"
            End If
        End If
    End If
    ' An approved name edit cannot silently rebase an unrelated manual date edit.
    If Not old Is Nothing And kind = "name" Then
        result = CStr(old(7)): raw = CStr(old(8)): owned = CStr(old(17))
        If Len(code) = 0 Then code = CStr(old(6)): binding = CStr(old(9))
    End If
    values.Add "Vrg1": values.Add VrgBookmark(slot): values.Add id: values.Add named.Text
    values.Add VfpPersonKey(row): values.Add code: values.Add result: values.Add raw: values.Add binding
    For i = 0 To 5: values.Add CStr(row(i)): Next i
    values.Add CStr(named.StoryType): values.Add owned
    payload = VdiEncode(VfuPack(values))
    If Not old Is Nothing Then
        If VfaReadVariable(doc, VrgName(slot)) = payload And doc.Bookmarks.Exists(VrgBookmark(slot)) Then
            Set live = doc.Bookmarks(VrgBookmark(slot)).Range
            If live.IsEqual(named) Then Exit Sub
        End If
    End If
    doc.Bookmarks.Add Name:=VrgBookmark(slot), Range:=named.Duplicate
    VfaWriteVariable doc, VrgName(slot), payload
End Sub

' Best-effort optional hooks. A registry failure never deletes or retries a date.
Private Sub VrgCapture(ByVal doc As Word.Document, ByVal named As Word.Range, ByRef row As Variant, _
                       ByVal kind As String, Optional ByVal knownSlot As Long = 0)
    Dim detail As String
    On Error GoTo Skipped
    If mVrgPaused Or Not VrgEnabled(doc) Then Exit Sub
    VrgStore doc, named, row, kind, knownSlot
    Exit Sub
Skipped:
    detail = Err.Description
    Application.StatusBar = VfText("Metin i\u015flemi tamamland\u0131; iste\u011fe ba\u011fl\u0131 \u015fah\u0131s kayd\u0131 olu\u015fturulamad\u0131: ") & Left$(detail, 160)
    Err.Clear
End Sub

Private Sub VrgCaptureBinding(ByVal doc As Word.Document, ByVal occurrence As Word.Range, ByVal binding As String)
    Dim row As Variant, named As Word.Range, aliasText As String, person As String
    On Error GoTo Unavailable
    If mVrgPaused Or Not VrgEnabled(doc) Or mVfRows Is Nothing Then Exit Sub
    aliasText = VfaReadVariable(doc, VfaVariableName(binding, "Alias"))
    person = VfaReadVariable(doc, VfaVariableName(binding, "Key"))
    Set named = occurrence.Duplicate
    named.SetRange occurrence.Start, occurrence.Start + Len(aliasText)
    If named.Text <> aliasText Then Exit Sub
    Call VrgIndex
    If mVrgLegacyIndex.Exists(person) Then
        row = mVrgLegacyIndex(person)
        VrgCapture doc, named, row, "date"
    End If
Unavailable:
    Err.Clear
End Sub

Private Function VrgDateIntact(ByVal doc As Word.Document, ByVal named As Word.Range, ByVal values As Collection) As Boolean
    Dim field As Word.Field, id As String
    If CStr(values(17)) <> "1" Then Exit Function
    Set field = VfnAdjacentField(doc, named): If field Is Nothing Then Exit Function
    id = VfaFieldBinding(doc, field)
    If id <> CStr(values(9)) Or field.Code.Text <> CStr(values(6)) Or field.Result.Text <> CStr(values(7)) Then Exit Function
    If VfaReadVariable(doc, VfaVariableName(id, "Key")) <> CStr(values(5)) Then Exit Function
    If VfaReadVariable(doc, VfaVariableName(id, "Raw")) <> CStr(values(8)) Then Exit Function
    If Not VbpUniform(field.Result) Then Exit Function
    VrgDateIntact = True
End Function

Private Function VrgStamp(ByVal records As Object) As String
    Dim key As Variant, values As Collection
    For Each key In records.Keys
        Set values = records(key)
        VrgStamp = VrgStamp & VfpKeyPart(CStr(key)) & VfpKeyPart(VdiEncode(VfuPack(values)))
    Next key
End Function

Private Function VrgItems(ByVal doc As Word.Document, ByVal onlyId As String) As Collection
    Dim result As New Collection, records As Object, index As Object, key As Variant, values As Collection
    Dim item As Object, named As Word.Range, row As Variant, oldRow As Variant, proposal As String, mode As String, field As Word.Field, binding As String
    Set records = VrgRead(doc): Set index = CreateObject("Scripting.Dictionary")
    If Not mVfRows Is Nothing Then Set index = VrgIndex()
    For Each key In records.Keys
        Set values = records(key)
        If Len(onlyId) = 0 Or CStr(values(3)) = onlyId Then
            Set item = CreateObject("Scripting.Dictionary")
            item("Id") = CStr(key): Set item("Saved") = values
            item("Name") = False: item("Date") = False: item("Status") = ""
            item("Order") = 2147483646: item("Story") = 0&: item("Offset") = 0&
            Set named = VrgLive(doc, values)
            If named Is Nothing Then
                item("Status") = VfText("Yer imi silinmi\u015f veya alan desteklenmiyor; otomatik yeniden e\u015fle\u015ftirilmez.")
            Else
                Set item("Range") = named: item("Offset") = named.Start: item("Order") = named.Start: item("Story") = named.StoryType
                item("Seal") = VbpSeal(named): item("Literal") = named.Text
                If named.Text <> CStr(values(4)) Then
                    item("Status") = VfText("Ad elle de\u011fi\u015ftirilmi\u015f; ad ve tarih g\u00fcncellemesi engellendi.")
                ElseIf Not index.Exists(CStr(values(3))) Then
                    item("Status") = VfText("Bu bilgisayarda g\u00fcncel kaynak yok; belgeye kay\u0131tl\u0131 kimlik korunur.")
                ElseIf index(CStr(values(3))).Count <> 1 Then
                    item("Status") = VfText("Kaynak kimli\u011fi birden fazla \u015fahsa ait; birle\u015ftirilmez.")
                Else
                    row = index(CStr(values(3)))(1): oldRow = VrgRow(values)
                    item("Row") = row: item("RowStamp") = VfqRowStamp(row)
                    proposal = VfuPreferred(doc, row)
                    mode = VfuNameMode(doc): If mode = "ask" Then mode = "near"
                    If Len(proposal) = 0 Then proposal = VfuProposal(row, VfNameQuery(named.Text), mode)
                    If Len(proposal) > 0 Then proposal = proposal & Mid$(named.Text, Len(VfNameQuery(named.Text)) + 1)
                    item("Proposal") = proposal
                    If VfaValidAlias(proposal) And proposal <> named.Text Then
                        If VbpUniform(named) Then item("Name") = True Else item("Status") = VfText("Karma bi\u00e7imli ad; elle incelenmeli.")
                    End If
                    If CStr(values(8)) <> CStr(row(3)) And VfaValidDate(CStr(row(3))) Then
                        If VrgDateIntact(doc, named, values) Then
                            item("Date") = True
                        Else
                            item("Status") = CStr(item("Status")) & VfText(" Tarih elle de\u011fi\u015ftirilmi\u015f veya yal\u0131n metin; korunur.")
                        End If
                    End If
                    Set field = VfnAdjacentField(doc, named)
                    If Not field Is Nothing Then
                        binding = VfaFieldBinding(doc, field)
                        If field.Code.Text <> CStr(values(6)) Or VfaReadVariable(doc, VfaVariableName(binding, "Key")) <> CStr(values(5)) Then
                            item("Name") = False: item("Date") = False
                            item("Status") = VfText("Tarih ba\u011f\u0131 elle de\u011fi\u015ftirilmi\u015f; ad ve tarih korunur.")
                        End If
                    End If
                    If VfqRowStamp(oldRow) <> VfqRowStamp(row) Then item("Status") = CStr(item("Status")) & VfText(" Kaynak kayd\u0131 de\u011fi\u015fmi\u015f; \u00f6nerileri inceleyin.")
                End If
            End If
            result.Add item
        End If
    Next key
    Set VrgItems = result
End Function

Private Function VrgChanges(ByVal items As Collection) As Collection
    Dim result As New Collection, item As Object, change As Object, kind As Variant, emitted As Boolean
    For Each item In items
        emitted = False
        For Each kind In Array("Name", "Date")
            If CBool(item(CStr(kind))) Then
                Set change = CreateObject("Scripting.Dictionary")
                Set change("Item") = item: change("Kind") = LCase$(CStr(kind))
                change("Key") = CStr(item("Id")) & ":" & CStr(change("Kind"))
                result.Add change: emitted = True
            End If
        Next kind
        If Not emitted Or Len(CStr(item("Status"))) > 0 Then
            Set change = CreateObject("Scripting.Dictionary")
            Set change("Item") = item: change("Kind") = "info": change("Key") = CStr(item("Id")) & ":info"
            result.Add change
        End If
    Next item
    Set VrgChanges = result
End Function

Private Function VrgRows(ByVal changes As Collection, ByVal doc As Word.Document) As Collection
    Dim result As New Collection, change As Object, item As Object, values As Collection, row As Variant, name As String, dateText As String, detail As String
    For Each change In changes
        Set item = change("Item"): Set values = item("Saved")
        name = CStr(values(4)): dateText = ""
        If CStr(change("Kind")) = "name" Then name = CStr(item("Proposal"))
        If CStr(change("Kind")) = "date" Then row = item("Row"): dateText = VfpFormatDate(CStr(row(3)), VfpGetPrefix(doc))
        detail = VfText("Kal\u0131c\u0131 kaynak kimli\u011fi: ") & CStr(values(3)) & vbCrLf & _
            VfText("Son onayl\u0131 yaz\u0131m: ") & CStr(values(4)) & vbCrLf & VfText("Kay\u0131tl\u0131 tarih: ") & CStr(values(7)) & vbCrLf & _
            VfText("\u00d6neri: ") & name & " " & dateText & vbCrLf & CStr(item("Status")) & vbCrLf & _
            CStr(values(14)) & vbCrLf & CStr(values(15))
        If item.Exists("Row") Then
            row = item("Row")
            detail = detail & vbCrLf & VfText("G\u00fcncel kaynak: ") & CStr(row(0)) & vbCrLf & CStr(row(1)) & vbCrLf & CStr(row(2)) & vbCrLf & CStr(row(3)) & vbCrLf & CStr(row(4)) & vbCrLf & CStr(row(5))
        End If
        If item.Exists("Range") Then detail = detail & vbCrLf & VfText("\u015eimdiki ad: ") & CStr(item("Literal")) & vbCrLf & VfwContextRange(item("Range")).Text
        result.Add Array(CStr(values(4)), name, CStr(change("Kind")), dateText, detail, "")
    Next change
    Set VrgRows = result
End Function

Private Sub VrgApply(ByVal doc As Word.Document, ByVal items As Collection, ByVal changes As Collection, _
                     ByVal selected As Object, ByVal expectedRecords As String)
    Dim chosen As Object, change As Object, item As Object, values As Collection, named As Word.Range, field As Word.Field
    Dim row As Variant, i As Long, id As String, first As Long, changed As Boolean, started As Boolean, undoObject As Object, detail As String
    Dim ordered As Variant, count As Long, key As String, raw As String, previous As String, literal As String, appliedName As String
    Dim previousScreen As Boolean, paused As Boolean
    If Not VrgEnabled(doc) Or Not VrgApi(doc) Or Not VfaEditable(doc) Or Not VfpDateAllowed(doc) Then Exit Sub
    If Not VftOpenDoc(doc) Or Not (ActiveDocument Is doc) Then Exit Sub
    If doc.TrackRevisions Or doc.Revisions.Count > 0 Then Err.Raise VrgInvalid, "Vefeyat", "Tracked documents are preview-only."
    If VrgStamp(VrgRead(doc)) <> expectedRecords Then Err.Raise VrgInvalid, "Vefeyat", "Registry changed; prepare a new preview."
    Set chosen = CreateObject("Scripting.Dictionary")
    For Each change In changes
        key = CStr(change("Key"))
        If CStr(change("Kind")) <> "info" And selected.Exists(key) Then
            If CBool(selected(key)) Then chosen(key) = True
        End If
    Next change
    If chosen.Count = 0 Then Exit Sub
    ReDim ordered(0 To items.Count)
    For Each item In items
        key = CStr(item("Id"))
        If chosen.Exists(key & ":name") Or chosen.Exists(key & ":date") Then
            If Not item.Exists("Range") Or Not item.Exists("Row") Then Err.Raise VrgInvalid, "Vefeyat", "No live registered source."
            Set named = item("Range"): Set values = item("Saved"): row = item("Row")
            If named.Text <> CStr(values(4)) Or named.Start <> CLng(item("Offset")) Or VbpSeal(named) <> CStr(item("Seal")) Then Err.Raise VrgInvalid, "Vefeyat", "Manual edit or changed context; text was preserved."
            If VrgSource(row) <> CStr(values(3)) Or VfqRowStamp(row) <> CStr(item("RowStamp")) Then Err.Raise VrgInvalid, "Vefeyat", "Catalog changed; reconfirm."
            If VfmTargetExcluded(doc, named, named.Text) Then Err.Raise VrgInvalid, "Vefeyat", "Protected target."
            If chosen.Exists(key & ":name") Then
                Set field = VfnAdjacentField(doc, named)
                If Not field Is Nothing Then
                    id = VfaFieldBinding(doc, field)
                    If field.Code.Text <> CStr(values(6)) Or VfaReadVariable(doc, VfaVariableName(id, "Key")) <> CStr(values(5)) Then Err.Raise VrgInvalid, "Vefeyat", "Manual identity link is protected."
                End If
                If Not CBool(item("Name")) Or Not VbpUniform(named) Then Err.Raise VrgInvalid, "Vefeyat", "Name update is not eligible."
            End If
            If chosen.Exists(key & ":date") Then
                If Not CBool(item("Date")) Or Not VrgDateIntact(doc, named, values) Then Err.Raise VrgInvalid, "Vefeyat", "Manual date edit is protected."
            End If
            count = count + 1: Set ordered(count) = item
        End If
    Next item
    If count > 1 Then VfjSort ordered, 1, count
    started = VfaBeginUndo(undoObject)
    If Not started Then Err.Raise VrgInvalid, "Vefeyat", "A separate Undo record is required."
    On Error GoTo Failed
    previousScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    paused = mVrgPaused: mVrgPaused = True
    For i = count To 1 Step -1
        Set item = ordered(i): key = CStr(item("Id")): Set values = item("Saved")
        Set named = item("Range"): row = item("Row"): first = named.Start: literal = named.Text
        Set field = VfnAdjacentField(doc, named)
        previous = "": raw = ""
        If Not field Is Nothing Then previous = field.Result.Text: raw = VfaReadVariable(doc, VfaVariableName(VfaFieldBinding(doc, field), "Raw"))
        If chosen.Exists(key & ":name") Then
            changed = True: named.Text = CStr(item("Proposal"))
            named.SetRange first, first + Len(CStr(item("Proposal")))
        End If
        appliedName = named.Text
        If Not field Is Nothing Then
            VfaEnsureNamespace doc, changed
            id = VdiNewBinding(doc, row, named.Text, changed)
            If chosen.Exists(key & ":date") Then
                changed = True: field.Locked = False
                field.Result.Text = " " & VfpFormatDate(CStr(row(3)), VfpGetPrefix(doc))
            Else
                VfaWriteVariable doc, VfaVariableName(id, "Raw"), raw
            End If
            field.Locked = False: field.Code.Text = " ADDIN " & VfaFieldTag(doc, id) & " ": field.ShowCodes = False: field.Locked = True
            If chosen.Exists(key & ":date") Then
                If field.Result.Text <> " " & VfpFormatDate(CStr(row(3)), VfpGetPrefix(doc)) Then Err.Raise VrgInvalid, "Vefeyat", "Updated date result was not retained; stop and Undo."
            End If
            If Not chosen.Exists(key & ":date") And field.Result.Text <> previous Then Err.Raise VrgInvalid, "Vefeyat", "An existing date changed; stop and Undo."
        End If
        named.SetRange first, first + Len(appliedName)
        If chosen.Exists(key & ":date") Then
            VrgStore doc, named, row, "date", CLng(key)
        Else
            VrgStore doc, named, row, "name", CLng(key)
        End If
        VftLog doc, VfText("Onayl\u0131 kay\u0131t g\u00fcncellemesi"), named, literal, named.Text
    Next i
    Application.ScreenUpdating = previousScreen: mVrgPaused = paused
    VfaEndUndo undoObject, started
    VfmInvalidate doc
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not field Is Nothing Then field.Locked = True
    Application.ScreenUpdating = previousScreen: mVrgPaused = paused
    VfaEndUndo undoObject, started
    On Error GoTo 0
    VfaReportFailure detail, changed
End Sub

Private Sub VrgRegisterSelected()
    Dim doc As Word.Document, target As Word.Range, named As Word.Range, original As String, aliasText As String, reason As String
    Dim first As Long, last As Long, story As Long, rows As Collection, choice As Long, row As Variant, stamp As String
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    If Not VfaCapture(doc, target, original, aliasText, first, last, story) Then Exit Sub
    If Not VrgEnabled(doc) Or Not VrgApi(doc) Then VfaWarn VfText("\u00d6nce iste\u011fe ba\u011fl\u0131 kay\u0131t y\u00f6netimini a\u00e7\u0131n. Desteklenmeyen dosyada yal\u0131n metin i\u015flemleri korunur."): Exit Sub
    If Not VfaEditable(doc) Or doc.TrackRevisions Or doc.Revisions.Count > 0 Then VfaWarn VfText("Salt okunur veya izlenen belge: kay\u0131tlar yaln\u0131zca incelenebilir."): Exit Sub
    If Not VefeyatEnsureData() Then Exit Sub
    stamp = VrgStamp(VrgRead(doc))
    Set rows = VfaConfirmationCandidates(VefeyatFind(aliasText, False), aliasText)
    choice = VfePickPerson(rows, aliasText, False, True)
    If choice < 1 Or choice > rows.Count Then Exit Sub
    row = rows(choice)
    If Len(VrgSource(row)) = 0 Then VfaWarn VfText("Kaynakta tekil kal\u0131c\u0131 kimlik yok; belirsiz kay\u0131t olu\u015fturulmad\u0131."): Exit Sub
    If VfMessage(VfText("Se\u00e7ili ge\u00e7i\u015f bu kaynak kimli\u011fine ba\u011flans\u0131n m\u0131? Ad ve tarih de\u011fi\u015ftirilmez.") & vbCrLf & VrgSource(row) & vbCrLf & CStr(row(0)), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") <> vbYes Then Exit Sub
    If Not VfaTargetCurrent(doc, target, original, first, last, story, reason) Then Exit Sub
    If VrgStamp(VrgRead(doc)) <> stamp Then Exit Sub
    Set named = VdiNamed(target, aliasText)
    VrgStore doc, named, row, "register"
End Sub

Private Sub VrgTryCatalog()
    Dim path As String
    On Error GoTo Offline
    If Not mVfRows Is Nothing Then Exit Sub
    path = VefeyatDataPath()
    If CreateObject("Scripting.FileSystemObject").FileExists(path) Then VfLoadCache path
Offline:
    Err.Clear
End Sub

Private Sub VrgManage(ByVal onlySelectedPerson As Boolean)
    Dim doc As Word.Document, items As Collection, changes As Collection, selected As Object, picker As VefeyatKisiSec0341
    Dim records As Object, key As Variant, live As Word.Range, id As String, values As Collection, item As Object, change As Object
    Dim i As Long, action As String, choice As Long, stamp As String, settings As String, detail As String, catalog As Collection
    On Error GoTo Failed
    If mVfbRunning Or mVfrInsertBusy Or Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument: Set records = VrgRead(doc)
    If onlySelectedPerson Then
        For Each key In records.Keys
            Set live = VrgLive(doc, records(key))
            If Not live Is Nothing Then
                If live.InRange(Selection.Range) Or Selection.Range.InRange(live) Then Set values = records(key): id = CStr(values(3)): Exit For
            End If
        Next key
        If Len(id) = 0 Then VfaWarn VfText("Se\u00e7ili ge\u00e7i\u015f hen\u00fcz kayda ba\u011flanmam\u0131\u015f."): Exit Sub
    End If
    VrgTryCatalog
    Set catalog = mVfRows
    stamp = VrgStamp(records): settings = VbpSettings(doc) & VfpKeyPart(CStr(VrgEnabled(doc)))
    Set items = VrgItems(doc, id): Set changes = VrgChanges(items)
    mVfrInsertBusy = True
    Set picker = New VefeyatKisiSec0341
    picker.InitCandidates VrgRows(changes, doc), VfText("Belgeye kay\u0131tl\u0131 \u015fah\u0131slar")
    picker.SetRegistryMode VfText("Ad/tarih g\u00fcncellemeleri ayr\u0131 se\u00e7ilir. Elle de\u011fi\u015ftirilen veya silinen ge\u00e7i\u015fler korunur. Kay\u0131tlar belgeyle ta\u015f\u0131n\u0131r."), _
        VrgEnabled(doc) And VfpDateAllowed(doc) And VfaEditable(doc) And Not doc.TrackRevisions And doc.Revisions.Count = 0
    picker.Show vbModal
    action = picker.PlanAction: choice = picker.ChosenIndex
    Set selected = CreateObject("Scripting.Dictionary")
    For i = 1 To changes.Count
        Set change = changes(i): selected(CStr(change("Key"))) = picker.PlanIncludes(i)
    Next i
    Unload picker: Set picker = Nothing
    If Not VftOpenDoc(doc) Then GoTo Clean
    If Not (ActiveDocument Is doc) Or Not (catalog Is mVfRows) Then GoTo Clean
    If VbpSettings(doc) & VfpKeyPart(CStr(VrgEnabled(doc))) <> settings Then GoTo Clean
    If action = "choose" And choice > 0 And choice <= changes.Count Then
        Set change = changes(choice): Set item = change("Item")
        If item.Exists("Range") Then
            Set live = item("Range"): live.Select
            Application.ActiveWindow.ScrollIntoView live, True
        Else
            If VfMessage(VfText("Ge\u00e7i\u015f bulunam\u0131yor. Yaln\u0131z kullan\u0131lmayan kay\u0131t kald\u0131r\u0131ls\u0131n m\u0131?"), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") = vbYes Then
                If VftOpenDoc(doc) And (ActiveDocument Is doc) And VfaEditable(doc) Then
                    If VrgStamp(VrgRead(doc)) = stamp Then VfaDeleteVariable doc, VrgName(CLng(item("Id")))
                End If
            End If
        End If
    ElseIf action = "apply" Then
        VrgApply doc, items, changes, selected, stamp
    End If
    GoTo Clean
Failed:
    detail = Err.Description
    VfrReport detail
Clean:
    On Error Resume Next
    If Not picker Is Nothing Then Unload picker
    mVfrInsertBusy = False
End Sub

Private Sub VrgForgetSelected()
    Dim doc As Word.Document, records As Object, key As Variant, named As Word.Range, values As Collection, slot As Long, stamp As String
    If Documents.Count = 0 Or mVfbRunning Or mVfrInsertBusy Then Exit Sub
    Set doc = ActiveDocument: Set records = VrgRead(doc): stamp = VrgStamp(records)
    For Each key In records.Keys
        Set named = VrgLive(doc, records(key))
        If Not named Is Nothing Then
            If Selection.Range.InRange(named) Or named.InRange(Selection.Range) Then slot = CLng(key): Exit For
        End If
    Next key
    If slot = 0 Then Exit Sub
    If VfMessage(VfText("Yaln\u0131z bu ge\u00e7i\u015fin dinamik kayd\u0131 ayr\u0131ls\u0131n m\u0131? Metin ve tarih korunur."), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") <> vbYes Then Exit Sub
    If Not VftOpenDoc(doc) Or Not (ActiveDocument Is doc) Or Not VfaEditable(doc) Then Exit Sub
    If VrgStamp(VrgRead(doc)) <> stamp Then Exit Sub
    VrgDetach doc, slot
End Sub

Private Sub VrgDetach(ByVal doc As Word.Document, ByVal slot As Long)
    If Not VfaEditable(doc) Then Exit Sub
    If slot < 1 Or slot > VrgLimit Then Exit Sub
    VfaDeleteVariable doc, VrgName(slot)
    If doc.Bookmarks.Exists(VrgBookmark(slot)) Then doc.Bookmarks(VrgBookmark(slot)).Delete
End Sub

Private Function VbpSettings(ByVal doc As Word.Document) As String
    VbpSettings = VfuExportPayload(doc) & VfpKeyPart(VdiRevisionValue(doc)) & VfpKeyPart(VfaReadVariable(doc, VfaNamespaceVariable)) & _
        VfpKeyPart(CStr(doc.ReadOnly)) & VfpKeyPart(CStr(doc.ProtectionType)) & VfpKeyPart(CStr(doc.TrackRevisions)) & VfpKeyPart(CStr(doc.Revisions.Count))
End Function

Private Function VbpSeal(ByVal named As Word.Range) As String
    Dim window As Word.Range, field As Word.Field
    Set window = VdiWindow(named, True)
    VbpSeal = VfpKeyPart(window.Text)
    For Each field In window.Fields
        VbpSeal = VbpSeal & VfpKeyPart(field.Code.Text) & VfpKeyPart(field.Result.Text)
    Next field
End Function

Private Function VbpUniform(ByVal named As Word.Range) As Boolean
    Dim character As Word.Range, styleName As String
    On Error GoTo Mixed
    With named.Font
        If .Name = "" Or .Size = wdUndefined Or .Bold = wdUndefined Or .Italic = wdUndefined Then Exit Function
        If .Underline = wdUndefined Or .Color = wdUndefined Or .Hidden = wdUndefined Then Exit Function
        If .StrikeThrough = wdUndefined Or .Superscript = wdUndefined Or .Subscript = wdUndefined Then Exit Function
        If .SmallCaps = wdUndefined Or .AllCaps = wdUndefined Or .DoubleStrikeThrough = wdUndefined Then Exit Function
        If .Position = wdUndefined Or .Scaling = wdUndefined Or .Spacing = wdUndefined Or .Kerning = wdUndefined Then Exit Function
    End With
    styleName = CStr(named.Characters(1).Style)
    For Each character In named.Characters
        If CStr(character.Style) <> styleName Then Exit Function
    Next character
    VbpUniform = True
Mixed:
End Function

Private Function VbpRecord(ByVal hit As Object, ByVal rows As Collection) As Object
    Dim item As Object, named As Word.Range, anchor As Word.Range, row As Variant, field As Word.Field, id As String, person As String
    Set item = CreateObject("Scripting.Dictionary"): Set named = hit("Range")
    Set item("Range") = named.Duplicate: Set item("Rows") = rows
    item("Alias") = CStr(hit("Alias")): item("Resolved") = False
    item("Proposal") = "": item("Confirmed") = False
    item("Dated") = VfmDatePresent(named.Document, named) Or VfbAnyDate(named)
    item("Seal") = VbpSeal(named): item("Uniform") = VbpUniform(named)
    item("Story") = 0&: item("Order") = named.Start: item("Offset") = named.Start
    item("Location") = VfText("Ana metin / ba\u015fl\u0131k / tablo")
    If named.StoryType = wdFootnotesStory Then
        item("Story") = 1&: item("Location") = VfText("Dipnot a\u00e7\u0131klamas\u0131")
        Set anchor = hit("Anchor"): item("Order") = anchor.Start: Set item("Anchor") = anchor.Duplicate
    End If
    If VfqSingleSafe(rows, CStr(item("Alias"))) Then
        row = rows(1): VbpResolve item, row, False
    End If
    If Not CBool(item("Resolved")) Then
        Set field = VfnAdjacentField(named.Document, named)
        If Not field Is Nothing Then
            id = VfaFieldBinding(named.Document, field)
            person = VfaReadVariable(named.Document, VfaVariableName(id, "Key"))
            For Each row In rows
                If VfpPersonKey(row) = person Then
                    VbpResolve item, row, False
                    Exit For
                End If
            Next row
        End If
    End If
    Set VbpRecord = item
End Function

Private Sub VbpResolve(ByVal item As Object, ByRef row As Variant, ByVal confirmed As Boolean)
    Dim proposal As String, query As String, mode As String
    item("Row") = row: item("RowStamp") = VfqRowStamp(row): item("Resolved") = True: item("Confirmed") = confirmed
    query = VfNameQuery(CStr(item("Alias")))
    proposal = VfuPreferred(item("Range").Document, row)
    mode = VfuNameMode(item("Range").Document): If mode = "ask" Then mode = "near"
    If Len(proposal) = 0 Then proposal = VfuProposal(row, query, mode)
    If Len(proposal) > 0 Then proposal = proposal & Mid$(CStr(item("Alias")), Len(query) + 1)
    If Not VfaValidAlias(proposal) Then proposal = ""
    item("Proposal") = proposal
End Sub

Private Function VbpPrepare(ByVal doc As Word.Document, ByVal snapshots As Collection) As Collection
    Dim patterns As Object, heads As Object, cache As Object, exclusions As Object, ranked As Object, seen As Object
    Dim hits As Collection, hit As Object, snapshot As Object, match As Variant, named As Word.Range, source As Word.Range
    Dim rows As Collection, all As New Collection, result As New Collection, item As Object, key As String, items As Variant, i As Long
    Set patterns = CreateObject("Scripting.Dictionary"): Set heads = CreateObject("Scripting.Dictionary")
    Set cache = CreateObject("Scripting.Dictionary"): Set ranked = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary"): Set exclusions = VfiRead(doc)
    VfxPrepareDocument doc, patterns, heads, cache
    VfzBegin doc, patterns, heads, cache
    Set hits = VfbOrderedHits(doc, patterns, heads, cache, snapshots, True)
    VfxVerifySnapshots snapshots, cache
    For Each hit In hits
        If Not VfiContains(exclusions, CStr(hit("Alias"))) Then
            key = CStr(hit("Key")) & vbNullChar & VfNameQuery(CStr(hit("Alias")))
            If Not ranked.Exists(key) Then Set ranked(key) = VfbRankRows(patterns(CStr(hit("Key"))), VfNameQuery(CStr(hit("Alias"))))
            Set rows = ranked(key): all.Add VbpRecord(hit, rows)
            Set named = hit("Range"): seen(VfjRangeKey(named)) = True
        End If
    Next hit
    For Each snapshot In snapshots
        Set source = snapshot("Range")
        For Each match In VfjUnknownMatches(CStr(snapshot("Text")), patterns, cache)
            Set named = source.Duplicate
            named.SetRange source.Start + CLng(match(0)), source.Start + CLng(match(0)) + CLng(match(1))
            If Not seen.Exists(VfjRangeKey(named)) And Not VfiContains(exclusions, CStr(match(2))) Then
                If Not VfzReferenceAt(named) Then
                    Set hit = CreateObject("Scripting.Dictionary")
                    Set hit("Range") = named: hit("Alias") = CStr(match(2))
                    If snapshot.Exists("Anchor") Then Set hit("Anchor") = snapshot("Anchor")
                    Set rows = New Collection: all.Add VbpRecord(hit, rows)
                    seen(VfjRangeKey(named)) = True
                End If
            End If
        Next match
    Next snapshot
    ReDim items(0 To all.Count)
    For i = 1 To all.Count: Set items(i) = all(i): Next i
    If all.Count > 1 Then VfjSort items, 1, all.Count
    For i = 1 To all.Count
        Set item = items(i): item("Id") = CStr(i): result.Add item
    Next i
    Set VbpPrepare = result
End Function

Private Sub VbpOperation(ByVal plan As Collection, ByVal item As Object, ByVal kind As String, ByVal detail As String)
    Dim change As Object
    Set change = CreateObject("Scripting.Dictionary")
    Set change("Record") = item: change("Kind") = kind
    change("Key") = CStr(item("Id")) & ":" & kind: change("Detail") = detail
    plan.Add change
End Sub

' First means first identified use, not first selected or first undated use.
Private Function VbpDecide(ByVal records As Collection, ByVal scope As String) As Collection
    Dim result As New Collection, seen As Object, possible As Object, item As Object, row As Variant, candidate As Variant
    Dim person As String, repeated As Boolean, uncertain As Boolean, message As String
    Set seen = CreateObject("Scripting.Dictionary"): Set possible = CreateObject("Scripting.Dictionary")
    For Each item In records
        If Not CBool(item("Resolved")) Then
            For Each candidate In item("Rows"): possible(VfpPersonKey(candidate)) = True: Next candidate
            VbpOperation result, item, "info", VfText("Kimlik do\u011frulanmal\u0131; tahminle de\u011fi\u015fiklik yap\u0131lmaz.")
        Else
            row = item("Row"): person = VfpPersonKey(row)
            repeated = seen.Exists(person): uncertain = possible.Exists(person)
            seen(person) = True
            message = ""
            If scope = "none" Then
                message = VfText("Yaln\u0131z denetim; ad ve tarih de\u011fi\u015ftirilmez.")
            Else
                If Len(CStr(item("Proposal"))) > 0 And CStr(item("Proposal")) <> CStr(item("Alias")) Then
                    If CBool(item("Uniform")) Then
                        VbpOperation result, item, "name", VfText("Yaln\u0131z bu ge\u00e7i\u015fin yaz\u0131m\u0131; mevcut tarih korunur.")
                    Else
                        message = VfText("Karma karakter bi\u00e7imi: ad\u0131 elle inceleyin; topluca de\u011fi\u015ftirilmez.")
                    End If
                End If
                If CBool(item("Dated")) Then
                    message = message & VfText(" Mevcut tarih veya korunan alan var; aynen korunur.")
                ElseIf scope = "first" And repeated Then
                    message = message & VfText(" Ayn\u0131 kimli\u011fin sonraki kullan\u0131m\u0131; tarih eklenmez.")
                ElseIf scope = "first" And uncertain Then
                    message = message & VfText(" Daha \u00f6nceki belirsiz kullan\u0131m do\u011frulanmadan ilk kullan\u0131m kesinle\u015ftirilemez.")
                ElseIf Not VfaValidDate(CStr(row(3))) Then
                    message = message & VfText(" Kaynakta kullan\u0131labilir tarih yok.")
                Else
                    VbpOperation result, item, "date", VfText("Se\u00e7ilen ge\u00e7i\u015fe kay\u0131tl\u0131 kaynak tarihi eklenecek.")
                End If
            End If
            If Len(message) > 0 Then VbpOperation result, item, "info", message
        End If
    Next item
    Set VbpDecide = result
End Function

Private Function VbpRows(ByVal plan As Collection, ByVal doc As Word.Document) As Collection
    Dim result As New Collection, change As Object, item As Object, row As Variant, proposed As String, dateText As String, detail As String
    For Each change In plan
        Set item = change("Record"): proposed = CStr(item("Alias")): dateText = ""
        If CStr(change("Kind")) = "name" Then proposed = CStr(item("Proposal"))
        If CStr(change("Kind")) = "date" Then row = item("Row"): dateText = VfpFormatDate(CStr(row(3)), VfpGetPrefix(doc))
        detail = CStr(item("Location")) & " | " & CStr(item("Offset")) & vbCrLf & _
            VfText("\u00d6nce: ") & CStr(item("Alias")) & vbCrLf & VfText("Sonra: ") & proposed & " " & dateText & vbCrLf & CStr(change("Detail"))
        If CBool(item("Resolved")) Then
            row = item("Row"): detail = detail & vbCrLf & VfText("Kimlik: ") & CStr(row(0)) & vbCrLf & CStr(row(4)) & vbCrLf & CStr(row(5))
        End If
        detail = detail & vbCrLf & vbCrLf & VfwContextRange(item("Range")).Text
        result.Add Array(CStr(item("Alias")), proposed, CStr(change("Kind")), dateText, detail, "")
    Next change
    Set VbpRows = result
End Function

Private Function VbpCurrent(ByVal doc As Word.Document, ByVal catalog As Collection, ByVal records As Collection, _
                            ByVal snapshots As Collection, ByVal settings As String, ByVal extent As Long, ByVal noteCount As Long) As Boolean
    Dim item As Object, snapshot As Object, named As Word.Range, row As Variant
    If Not VftOpenDoc(doc) Then Exit Function
    If Not (ActiveDocument Is doc) Or Not (mVfRows Is catalog) Then Exit Function
    If doc.Content.End <> extent Or doc.Footnotes.Count <> noteCount Then Exit Function
    If VbpSettings(doc) <> settings Then Exit Function
    For Each snapshot In snapshots
        Set named = snapshot("Range")
        If StrComp(named.Text, CStr(snapshot("Text")), vbBinaryCompare) <> 0 Then Exit Function
    Next snapshot
    For Each item In records
        Set named = item("Range")
        If StrComp(named.Text, CStr(item("Alias")), vbBinaryCompare) <> 0 Then Exit Function
        If named.Start <> CLng(item("Offset")) Then Exit Function
        If item.Exists("Anchor") Then
            If item("Anchor").Start <> CLng(item("Order")) Then Exit Function
        End If
        If VbpSeal(named) <> CStr(item("Seal")) Then Exit Function
        If CBool(item("Resolved")) Then
            row = item("Row")
            If VfqRowStamp(row) <> CStr(item("RowStamp")) Or Not VdiRowCurrent(row) Then Exit Function
        End If
    Next item
    VbpCurrent = True
End Function

Private Sub VbpCheckConfirmations(ByVal doc As Word.Document, ByVal records As Collection, ByVal changes As Object)
    Dim saved As Object, item As Object, key As Variant, values As Collection, named As Word.Range, row As Variant, extra As Long, found As Boolean
    Set saved = VdiRead(doc)
    For Each item In records
        If CBool(item("Confirmed")) And item("Rows").Count > 1 Then
            If changes.Exists(CStr(item("Id")) & ":name") Or changes.Exists(CStr(item("Id")) & ":date") Then
                Set named = item("Range"): row = item("Row")
                VdiCheckStore named, row
                found = False
                For Each key In saved.Keys
                    Set values = saved(key)
                    If doc.Bookmarks.Exists(CStr(values(1))) Then
                        If doc.Bookmarks(CStr(values(1))).Range.IsEqual(named) Then
                            item("ContextSlot") = CLng(key): found = True: Exit For
                        End If
                    End If
                Next key
                If Not found Then extra = extra + 1
            End If
        End If
    Next item
    If saved.Count + extra > VdiLimit Then Err.Raise VbpInvalid, "Vefeyat", VfText("Bu grupla belge kimlik onay\u0131 s\u0131n\u0131r\u0131 a\u015f\u0131lacak; hi\u00e7bir de\u011fi\u015fiklik yap\u0131lmad\u0131.")
End Sub

Private Sub VbpApply(ByVal doc As Word.Document, ByVal records As Collection, ByVal plan As Collection, ByVal selected As Object)
    Dim changes As Object, eligible As Object, change As Object, item As Object, named As Word.Range, row As Variant, field As Word.Field
    Dim i As Long, id As String, literal As String, previous As String, raw As String, first As Long, changed As Boolean, appliedName As String
    Dim undoObject As Object, started As Boolean, detail As String, key As String, previousScreen As Boolean, screenCaptured As Boolean, registryPaused As Boolean, registrySlot As Long
    Set changes = CreateObject("Scripting.Dictionary")
    If Not VfpDateAllowed(doc) Or Not VfaEditable(doc) Then Exit Sub
    If doc.TrackRevisions Or doc.Revisions.Count > 0 Then Err.Raise VbpInvalid, "Vefeyat", VfText("De\u011fi\u015fiklik \u0130zleme a\u00e7\u0131k veya belgede izlenen de\u011fi\u015fiklikler var; toplu uygulama kapal\u0131d\u0131r. \u00d6nizleme ile denetleyebilirsiniz.")
    Set eligible = CreateObject("Scripting.Dictionary")
    For Each change In VbpDecide(records, VfpGetScope(doc))
        If CStr(change("Kind")) <> "info" Then eligible(CStr(change("Key"))) = True
    Next change
    For Each change In plan
        key = CStr(change("Key"))
        If CStr(change("Kind")) <> "info" And selected.Exists(key) Then
            If CBool(selected(key)) Then
                If Not eligible.Exists(key) Then Err.Raise VbpInvalid, "Vefeyat", VfText("Tarih ekleme tercihi veya uygulanabilir \u00f6neriler de\u011fi\u015fti; yeniden \u00f6nizleyin.")
                changes(key) = True
            End If
        End If
    Next change
    If changes.Count = 0 Then Exit Sub
    ' All selected targets must pass before the first write or Undo record.
    For Each item In records
        key = CStr(item("Id"))
        If changes.Exists(key & ":name") Or changes.Exists(key & ":date") Then
            Set named = item("Range"): row = item("Row")
            If named.Text <> CStr(item("Alias")) Or Not VdiRowCurrent(row) Then Err.Raise VbpInvalid, "Vefeyat", VfText("\u00d6nizleme g\u00fcncelli\u011fini yitirdi; yeniden denetleyin.")
            If named.Start <> CLng(item("Offset")) Or VfqRowStamp(row) <> CStr(item("RowStamp")) Then Err.Raise VbpInvalid, "Vefeyat", VfText("\u00d6nizlemedeki konum veya kaynak de\u011fi\u015fti; yeniden denetleyin.")
            If VbpSeal(named) <> CStr(item("Seal")) Then Err.Raise VbpInvalid, "Vefeyat", VfText("\u00d6nizlemeden sonra ba\u011flam de\u011fi\u015fti; yeniden denetleyin.")
            If VfmTargetExcluded(doc, named, named.Text) Then Err.Raise VbpInvalid, "Vefeyat", VfText("Bu ge\u00e7i\u015f korunan bir alanda; de\u011fi\u015ftirilmedi.")
            If changes.Exists(key & ":date") Then
                If VfmDatePresent(doc, named) Or VfbAnyDate(named) Then Err.Raise VbpInvalid, "Vefeyat", VfText("Mevcut tarih korunmal\u0131; ikinci tarih eklenmedi.")
            End If
            If changes.Exists(key & ":name") Then
                If Not VbpUniform(named) Then Err.Raise VbpInvalid, "Vefeyat", VfText("Karma bi\u00e7imli ad\u0131 elle inceleyin; topluca de\u011fi\u015ftirilmedi.")
                Set field = VfnAdjacentField(doc, named)
                If Not field Is Nothing Then
                    id = VfaFieldBinding(doc, field)
                    If VfaReadVariable(doc, VfaVariableName(id, "Key")) <> VfpPersonKey(row) Then Err.Raise VbpInvalid, "Vefeyat", VfText("Kay\u0131tl\u0131 kimlik se\u00e7ilen \u015fah\u0131sla \u00e7eli\u015fiyor. \u00d6nce kimli\u011fi do\u011frulay\u0131n.")
                End If
            End If
        End If
    Next item
    VbpCheckConfirmations doc, records, changes
    started = VfaBeginUndo(undoObject)
    If Not started Then Err.Raise VbpInvalid, "Vefeyat", VfText("Ayr\u0131 bir Geri Al kayd\u0131 ba\u015flat\u0131lamad\u0131; belge de\u011fi\u015ftirilmedi.")
    On Error GoTo Failed
    previousScreen = Application.ScreenUpdating: screenCaptured = True
    Application.ScreenUpdating = False
    registryPaused = mVrgPaused: mVrgPaused = True
    For i = records.Count To 1 Step -1
        Set item = records(i): key = CStr(item("Id"))
        If changes.Exists(key & ":name") Or changes.Exists(key & ":date") Then
            Set named = item("Range"): row = item("Row"): literal = named.Text: first = named.Start
            mVrgPaused = registryPaused: registrySlot = VrgNameSlot(doc, named): mVrgPaused = True
            If changes.Exists(key & ":name") Then
                Set field = VfnAdjacentField(doc, named)
                If Not field Is Nothing Then previous = field.Result.Text: raw = VfaReadVariable(doc, VfaVariableName(VfaFieldBinding(doc, field), "Raw"))
                changed = True: named.Text = CStr(item("Proposal"))
                named.SetRange first, first + Len(CStr(item("Proposal")))
                If Not field Is Nothing Then
                    id = VdiNewBinding(doc, row, named.Text, changed)
                    VfaWriteVariable doc, VfaVariableName(id, "Raw"), raw
                    field.Locked = False: field.Code.Text = " ADDIN " & VfaFieldTag(doc, id) & " ": field.Locked = True
                    If field.Result.Text <> previous Then Err.Raise VbpInvalid, "Vefeyat", VfText("Word mevcut tarihi de\u011fi\u015ftirdi. \u0130\u015flem durduruldu; Geri Al ile kontrol edin.")
                End If
            End If
            appliedName = named.Text
            If changes.Exists(key & ":date") Then
                VfaEnsureNamespace doc, changed
                id = VdiNewBinding(doc, row, named.Text, changed)
                VfaInsertManaged doc, named, VfpFormatDate(CStr(row(3)), VfpGetPrefix(doc)), id, changed
            End If
            named.SetRange first, first + Len(appliedName)
            If CBool(item("Confirmed")) And item("Rows").Count > 1 Then
                If item.Exists("ContextSlot") Then doc.Bookmarks.Add Name:=VdiBookmark(CLng(item("ContextSlot"))), Range:=named.Duplicate
                VdiStore named, row
            End If
            mVrgPaused = registryPaused
            If changes.Exists(key & ":date") Then
                VrgCapture doc, named, row, "date", registrySlot
            Else
                VrgCapture doc, named, row, "name", registrySlot
            End If
            mVrgPaused = True
            VftLog doc, VfText("Onayl\u0131 toplu i\u015flem"), named, literal, appliedName
        End If
    Next i
    Application.ScreenUpdating = previousScreen: mVrgPaused = registryPaused
    VfaEndUndo undoObject, started
    VfmInvalidate doc
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not field Is Nothing Then field.Locked = True
    If screenCaptured Then Application.ScreenUpdating = previousScreen
    mVrgPaused = registryPaused
    VfaEndUndo undoObject, started
    On Error GoTo 0
    VfaReportFailure detail, changed
End Sub

Private Function VbpDispatchPrepare() As Boolean
    If mVbpDoc Is Nothing Then Exit Function
    VbpDispatchPrepare = True
    Set mVbpPrepared = VbpPrepare(mVbpDoc, mVbpSnapshots)
    mVbpReady = True
End Function

Private Function VbpPrepareInteractive(ByVal doc As Word.Document, ByVal snapshots As Collection) As Collection
    On Error GoTo Failed
    Set mVbpDoc = doc: Set mVbpSnapshots = snapshots
    Set mVbpPrepared = Nothing: mVbpReady = False
    mVfbRunning = True
    VfbResetProgress
    Set mVfbProgress = New VefeyatKisiSec0341
    mVfbProgress.InitProgress
    mVfbProgress.Show vbModal
    Unload mVfbProgress: Set mVfbProgress = Nothing
    mVfbRunning = False
    If Not mVbpReady Then Err.Raise VfbCancelError, "Vefeyat", VfText("\u00d6nizleme haz\u0131rl\u0131\u011f\u0131 iptal edildi veya tamamlanamad\u0131.")
    Set VbpPrepareInteractive = mVbpPrepared
    Set mVbpDoc = Nothing: Set mVbpSnapshots = Nothing: Set mVbpPrepared = Nothing
    Exit Function
Failed:
    On Error Resume Next
    If Not mVfbProgress Is Nothing Then Unload mVfbProgress
    Set mVfbProgress = Nothing: Set mVbpDoc = Nothing: Set mVbpSnapshots = Nothing: Set mVbpPrepared = Nothing
    mVfbRunning = False
    On Error GoTo 0
    Err.Raise VfbCancelError, "Vefeyat", VfText("\u00d6nizleme haz\u0131rl\u0131\u011f\u0131 iptal edildi veya tamamlanamad\u0131.")
End Function

Private Sub VbpRun()
    Dim doc As Word.Document, catalog As Collection, records As Collection, snapshots As New Collection, plan As Collection
    Dim picker As VefeyatKisiSec0341, selected As Object, change As Object, item As Object, indices As Collection, index As Variant
    Dim settings As String, coverage As String, extent As Long, noteCount As Long, action As String, choice As Long, row As Variant
    Dim i As Long, detail As String, accepted As Boolean
    On Error GoTo Failed
    If mVfbRunning Or mVfrInsertBusy Or Documents.Count = 0 Then Exit Sub
    If Not VefeyatEnsureData() Then Exit Sub
    mVfrInsertBusy = True: mVflSuspended = True
    Set doc = ActiveDocument: Set catalog = mVfRows
    extent = doc.Content.End: noteCount = doc.Footnotes.Count: settings = VbpSettings(doc)
    coverage = VbsReport(doc)
    If doc.TrackRevisions Or doc.Revisions.Count > 0 Then coverage = coverage & vbCrLf & VfText("\u0130zlenen de\u011fi\u015fiklikler: yaln\u0131z \u00f6nizleme; i\u015faretli veya silinmi\u015f kullan\u0131mlar da listelenebilir.")
    Set records = VbpPrepareInteractive(doc, snapshots)
    Set selected = CreateObject("Scripting.Dictionary")
    Do
        If Not VbpCurrent(doc, catalog, records, snapshots, settings, extent, noteCount) Then Err.Raise VbpInvalid, "Vefeyat", VfText("Belge veya tercihler de\u011fi\u015fti. Yeni \u00f6nizleme haz\u0131rlay\u0131n.")
        Set plan = VbpDecide(records, VfpGetScope(doc))
        Set picker = New VefeyatKisiSec0341
        picker.InitCandidates VbpRows(plan, doc), VfText("Toplu d\u00fczeltme \u00f6nizlemesi")
        picker.SetPlanMode coverage, VfpDateAllowed(doc) And VfaEditable(doc) And Not doc.TrackRevisions And doc.Revisions.Count = 0
        For i = 1 To plan.Count
            Set change = plan(i)
            If selected.Exists(CStr(change("Key"))) Then picker.PlanSetSelection i, CBool(selected(CStr(change("Key"))))
        Next i
        picker.Show vbModal
        action = picker.PlanAction: choice = picker.ChosenIndex
        For i = 1 To plan.Count
            Set change = plan(i): selected(CStr(change("Key"))) = picker.PlanIncludes(i)
        Next i
        Unload picker: Set picker = Nothing
        If action = "choose" And choice > 0 And choice <= plan.Count Then
            Set change = plan(choice): Set item = change("Record")
            item("Range").Select
            If item("Rows").Count > 0 Then
                choice = VfePickPerson(item("Rows"), CStr(item("Alias")), True, True)
                If choice > 0 And choice <= item("Rows").Count Then
                    row = item("Rows")(choice): VbpResolve item, row, True
                    detail = VfnChooseName(row, VfNameQuery(CStr(item("Alias"))))
                    If Len(detail) > 0 Then item("Proposal") = detail & Mid$(CStr(item("Alias")), Len(VfNameQuery(CStr(item("Alias")))) + 1)
                End If
            Else
                VfaWarn VfText("Bu ad katalogda yok; kimlik veya tarih uydurulmaz. Mevcut katalog aramas\u0131yla ayr\u0131ca inceleyin.")
            End If
        Else
            accepted = (action = "apply")
            Exit Do
        End If
    Loop
    If accepted Then
        If Not VbpCurrent(doc, catalog, records, snapshots, settings, extent, noteCount) Then Err.Raise VbpInvalid, "Vefeyat", VfText("\u00d6nizleme g\u00fcncelli\u011fini yitirdi. Toplu de\u011fi\u015fiklik yap\u0131lmad\u0131.")
        VbpApply doc, records, plan, selected
    End If
    GoTo Clean
Failed:
    detail = Err.Description
    If Err.Number = VfbCancelError Then
        Application.StatusBar = VfText("\u00d6nizleme haz\u0131rl\u0131\u011f\u0131 durduruldu; belge de\u011fi\u015ftirilmedi.")
    Else
        VfrReport detail
    End If
Clean:
    On Error Resume Next
    If Not picker Is Nothing Then Unload picker
    mVfrInsertBusy = False: mVflSuspended = False
End Sub

Private Function VbsReport(ByVal doc As Word.Document) As String
    Dim region As Word.Range, unsupported As Object, kind As Variant, label As String
    Set unsupported = CreateObject("Scripting.Dictionary")
    VbsReport = VfText("Kapsam: ana metin (ba\u015fl\u0131klar ve tablolar dahil) ve dipnot a\u00e7\u0131klamalar\u0131. At\u0131flar/alanlar korunur.")
    On Error GoTo UnknownCoverage
    For Each region In doc.StoryRanges
        Select Case region.StoryType
            Case wdMainTextStory, wdFootnotesStory, 12 To 17
            Case Else
                Select Case region.StoryType
                    Case 3: label = VfText("sonnotlar")
                    Case 4: label = VfText("yorumlar")
                    Case 5: label = VfText("metin kutular\u0131 / ba\u011fl\u0131 \u00e7er\u00e7eveler")
                    Case 6 To 11: label = VfText("\u00fcstbilgi / altbilgi")
                    Case Else: label = "Word story " & CStr(region.StoryType)
                End Select
                unsupported(label) = True
        End Select
    Next region
    For Each kind In unsupported.Keys
        VbsReport = VbsReport & vbCrLf & VfText("Bu belgede tarama d\u0131\u015f\u0131: ") & CStr(kind)
    Next kind
    VbsReport = VbsReport & vbCrLf & VfText("Denetim yaln\u0131z belirtilen alanlar\u0131 kapsar; alg\u0131lanmayan \u015fah\u0131slar olabilir. Metin kutular\u0131, \u00fcst/altbilgi ve sonnot i\u00e7erikleri taranmaz ve de\u011fi\u015ftirilmez.")
    Exit Function
UnknownCoverage:
    VbsReport = VbsReport & vbCrLf & VfText("Di\u011fer Word alanlar\u0131n\u0131n varl\u0131\u011f\u0131 okunamad\u0131; kapsam eksik/bilinmiyor. Tam denetim olarak kabul etmeyin.")
End Function

Private Sub VftLog(ByVal doc As Word.Document, ByVal kind As String, ByVal target As Word.Range, ByVal before As String, ByVal after As String)
    Dim item As Object
    On Error Resume Next
    If mVftHistory Is Nothing Then Set mVftHistory = New Collection
    Set item = CreateObject("Scripting.Dictionary")
    Set item("Doc") = doc
    If Not target Is Nothing Then Set item("Range") = target.Duplicate
    item("Label") = Format$(Now, "hh:nn:ss") & " - " & kind
    item("Detail") = VfText("\u00d6nce: ") & Left$(before, 500) & vbCrLf & VfText("Sonra: ") & Left$(after, 500) & vbCrLf & vbCrLf & _
        VfText("Bu Word oturumunda tutulur; dosyaya veya tan\u0131 raporuna yaz\u0131lmaz. Toplu i\u015flem i\u00e7in ge\u00e7i\u015f kayd\u0131d\u0131r. Geri alma Word'\u00fcn Geri Al komutuyla yap\u0131l\u0131r; araya giren i\u015flemleri kontrol edin.")
    item("Repair") = "none"
    mVftHistory.Add item
    If mVftHistory.Count > 100 Then mVftHistory.Remove 1
End Sub

Private Function VftOpenDoc(ByVal doc As Word.Document) As Boolean
    Dim current As Word.Document
    For Each current In Documents
        If current Is doc Then VftOpenDoc = True: Exit Function
    Next current
End Function

Private Sub VftClose(Optional ByVal panelClosing As Boolean = False)
    Dim panel As VefeyatKisiSec0341
    If mVftClosing Then Exit Sub
    mVftClosing = True
    On Error GoTo Clean
    mVftEpoch = (mVftEpoch Mod 2000000000) + 1
    Set panel = mVftPanel: Set mVftPanel = Nothing
    Set mVftDoc = Nothing: Set mVftItems = Nothing
    mVftBusy = False
    If Not panel Is Nothing And Not panelClosing Then Unload panel
Clean:
    mVftClosing = False
End Sub

Private Function VftRows() As Collection
    Dim rows As New Collection, item As Object
    For Each item In mVftItems
        rows.Add Array(CStr(item("Label")), "", CStr(item("Repair")), IIf(item.Exists("Field"), "unlink", ""), CStr(item("Detail")), "")
    Next item
    Set VftRows = rows
End Function

Private Sub VftDisplay(ByVal heading As String)
    Dim canChange As Boolean
    If mVftPanel Is Nothing Then Exit Sub
    canChange = (mVftMode <> "history")
    If Not mVftDoc Is Nothing Then canChange = canChange And Not mVftDoc.ReadOnly And mVftDoc.ProtectionType = wdNoProtection
    mVftPanel.InitWorkbench VftRows(), heading, canChange
End Sub

Private Function VftValue(ByVal values As Object, ByVal key As String) As String
    If values.Exists(key) Then VftValue = CStr(values(key))
End Function

Private Function VftMeta(ByVal doc As Word.Document, ByVal id As String) As String
    VftMeta = VfaReadVariable(doc, VfaVariableName(id, "Key")) & vbNullChar & _
        VfaReadVariable(doc, VfaVariableName(id, "Raw")) & vbNullChar & _
        VfaReadVariable(doc, VfaVariableName(id, "Alias")) & vbNullChar & VfpGetPrefix(doc) & vbNullChar & VfaReadVariable(doc, VfaNamespaceVariable)
End Function

Private Function VftProblem(ByVal raw As String, ByVal aliasText As String, ByVal personKey As String, ByVal result As String, ByVal prefix As String) As String
    Dim parts As Variant
    parts = VfePersonParts(personKey)
    If IsEmpty(parts) Or Not VfaValidDate(raw) Or Not VfaValidAlias(aliasText) Then
        VftProblem = "metadata"
    ElseIf Len(VfaTrimSpaces(result)) = 0 Then
        VftProblem = "empty"
    ElseIf StrComp(result, " " & VfpFormatDate(raw, prefix), vbBinaryCompare) <> 0 Then
        VftProblem = "changed"
    End If
End Function

Private Sub VftHealth(ByVal compareCatalog As Boolean)
    Dim doc As Word.Document, region As Word.Range, field As Word.Field, id As String, code As String
    Dim raw As String, aliasText As String, key As String, problem As String, item As Object, parts As Variant
    Dim seenDates As Object, seenAliases As Object, byCatalog As Object, matchRows As Collection, row As Variant
    Dim epoch As Long, count As Long, stable As String, label As String, repair As String, expected As String
    Dim variables As Object, variable As Word.Variable, bindingNamespace As String, prefix As String
    Dim variableCount As Long, variableName As String
    If Not mVftPanel Is Nothing Then VftClose
    Set doc = ActiveDocument
    If compareCatalog Then
        If Not VefeyatEnsureData() Then Exit Sub
        Set byCatalog = VfqIndex(mVfRows)
    End If
    Set mVftDoc = doc: Set mVftItems = New Collection
    mVftMode = "health": epoch = mVftEpoch: mVftBusy = True
    Set mVftPanel = New VefeyatKisiSec0341
    VftDisplay VfText("Belge sa\u011fl\u0131\u011f\u0131 haz\u0131rlan\u0131yor; metin de\u011fi\u015ftirilmiyor. Kapat ile durdurabilirsiniz.")
    mVftPanel.Show vbModeless
    If epoch <> mVftEpoch Then Exit Sub
    Set seenDates = CreateObject("Scripting.Dictionary"): Set seenAliases = CreateObject("Scripting.Dictionary")
    Set variables = CreateObject("Scripting.Dictionary"): variables.CompareMode = vbTextCompare
    For Each variable In doc.Variables
        variableName = variable.Name
        If StrComp(Left$(variableName, Len(VfaVariablePrefix)), VfaVariablePrefix, vbTextCompare) = 0 Or _
            StrComp(variableName, VfaNamespaceVariable, vbTextCompare) = 0 Or StrComp(variableName, VfpPrefixVariable, vbTextCompare) = 0 Then
            variables(variableName) = variable.Value
        End If
        variableCount = variableCount + 1
        If variableCount Mod 128 = 0 Then
            DoEvents
            If epoch <> mVftEpoch Then Exit Sub
            If Not VftOpenDoc(doc) Then VftClose: Exit Sub
        End If
    Next variable
    bindingNamespace = VftValue(variables, VfaNamespaceVariable)
    prefix = VfpNormalizeKey(VftValue(variables, VfpPrefixVariable))
    If Not VfpKnownPrefix(prefix) Then prefix = "original"
    For Each region In VfaRegions(doc)
        For Each field In region.Fields
            count = count + 1
            If count Mod 64 = 0 Then
                Application.StatusBar = VfText("Belge sa\u011fl\u0131\u011f\u0131: alan ") & CStr(count)
                DoEvents
                If epoch <> mVftEpoch Then Exit Sub
                If Not VftOpenDoc(doc) Then VftClose: Exit Sub
            End If
            If field.Type <> wdFieldAddin Then GoTo NextField
            If field.Code.End - field.Code.Start > 512 Then GoTo NextField
            code = field.Code.Text: id = VfaBindingFromFieldCode(code, bindingNamespace)
            If InStr(1, LCase$(code), "addin vefeyat.date.", vbBinaryCompare) > 0 And _
                (Len(id) = 0 Or field.Result.End - field.Result.Start > 2000) Then
                Set item = CreateObject("Scripting.Dictionary")
                item("Label") = VfText("Bozuk veya a\u015f\u0131r\u0131 uzun tarih alan\u0131")
                item("Detail") = VfText("Alan\u0131n ad alan\u0131 / etiketi ge\u00e7ersiz veya sonu\u00e7 boyutu ola\u011fand\u0131\u015f\u0131. Bu panelden onar\u0131lmaz veya ba\u011f\u0131 kald\u0131r\u0131lmaz. Metinde g\u00f6ster ile Word'de inceleyin.")
                item("Repair") = "none": Set item("Range") = field.Result.Duplicate
                mVftItems.Add item
                If mVftItems.Count >= 1000 Then GoTo Complete
                GoTo NextField
            End If
            If Len(id) > 0 Then
                raw = VftValue(variables, VfaVariableName(id, "Raw"))
                aliasText = VftValue(variables, VfaVariableName(id, "Alias"))
                key = VftValue(variables, VfaVariableName(id, "Key"))
                problem = VftProblem(raw, aliasText, key, field.Result.Text, prefix)
                repair = "none": expected = "": stable = key
                parts = VfePersonParts(key)
                If Not IsEmpty(parts) Then
                    row = Array(CStr(parts(1)), CStr(parts(2)), CStr(parts(3)), raw, "", CStr(parts(0)))
                    stable = VfqIdentity(row)
                End If
                If Len(problem) = 0 Then
                    If seenDates.Exists(stable) Then
                        If CStr(seenDates(stable)) <> raw Then problem = "conflict"
                    End If
                    seenDates(stable) = raw
                    If seenAliases.Exists(VfiKey(aliasText)) Then
                        If CStr(seenAliases(VfiKey(aliasText))) <> stable Then problem = "ambiguous"
                    End If
                    seenAliases(VfiKey(aliasText)) = stable
                End If
                If problem = "empty" Or problem = "changed" Then
                    repair = "date": expected = " " & VfpFormatDate(raw, prefix)
                End If
                If Len(problem) = 0 And compareCatalog And Not IsEmpty(parts) Then
                    If byCatalog.Exists(stable) Then
                        Set matchRows = byCatalog(stable)
                        If matchRows.Count = 1 Then
                            row = matchRows(1)
                            If CStr(row(3)) <> raw Then
                                problem = "catalog_date"
                            ElseIf VfpPersonKey(row) <> key Then
                                problem = "catalog_name": repair = "key": expected = VfpPersonKey(row)
                            End If
                        Else
                            problem = "catalog_ambiguous"
                        End If
                    Else
                        problem = "catalog_missing"
                    End If
                End If
                If Len(problem) > 0 Then
                    Set item = CreateObject("Scripting.Dictionary")
                    item("Label") = VftProblemLabel(problem) & " - " & aliasText
                    item("Detail") = VfText("Durum: ") & VftProblemLabel(problem) & vbCrLf & _
                        VfText("Belgedeki tarih: ") & field.Result.Text & vbCrLf & VfText("Kay\u0131tl\u0131 ham tarih: ") & raw & vbCrLf & _
                        VfText("\u00d6neri: ") & IIf(repair = "date", expected, VftProblemLabel(problem)) & vbCrLf & VfqDateNote(raw) & vbCrLf & _
                        VfText("Onar yaln\u0131z kay\u0131tl\u0131 ham tarihi geri koyar veya ayn\u0131 D\u0130A kimli\u011finin ad ba\u011f\u0131n\u0131 yeniler. Katalog tarihi kendili\u011finden yaz\u0131lmaz. Ba\u011f\u0131 kald\u0131r, g\u00f6r\u00fcnen metni koruyarak tarihi d\u00fcz metne \u00e7evirir.")
                    Set item("Range") = field.Result.Duplicate: Set item("Field") = field
                    item("Code") = code: item("Result") = field.Result.Text: item("Id") = id
                    item("Meta") = key & vbNullChar & raw & vbNullChar & aliasText & vbNullChar & prefix & vbNullChar & bindingNamespace
                    item("Repair") = repair: item("Expected") = expected
                    mVftItems.Add item
                    If mVftItems.Count >= 1000 Then GoTo Complete
                End If
            End If
NextField:
        Next field
    Next region
Complete:
    If epoch <> mVftEpoch Then Exit Sub
    mVftBusy = False
    Application.StatusBar = vbNullString
    VftDisplay CStr(mVftItems.Count) & VfText(" bulgu. En \u00e7ok 1000 bulgu listelenir. Salt okunur inceleme; onay\u0131n\u0131z olmadan de\u011fi\u015fiklik yap\u0131lmad\u0131.")
End Sub

Private Function VftProblemLabel(ByVal key As String) As String
    Select Case key
        Case "metadata": VftProblemLabel = VfText("Eksik veya bozuk \u015fah\u0131s ba\u011f\u0131; otomatik onar\u0131lmaz.")
        Case "empty": VftProblemLabel = VfText("Tarih alan\u0131 bo\u015f.")
        Case "changed": VftProblemLabel = VfText("G\u00f6r\u00fcnen tarih kay\u0131tl\u0131 tarihten farkl\u0131.")
        Case "conflict": VftProblemLabel = VfText("Ayn\u0131 kaynak kimli\u011fi i\u00e7in farkl\u0131 ham tarihler var.")
        Case "ambiguous": VftProblemLabel = VfText("Ayn\u0131 yaz\u0131m birden fazla \u015fahsa ba\u011fl\u0131; ba\u011flam\u0131 kontrol edin.")
        Case "catalog_name": VftProblemLabel = VfText("Kaynak kimli\u011fi ayn\u0131; katalog adlar\u0131 de\u011fi\u015fmi\u015f. Tarih korunarak ba\u011f yenilenebilir.")
        Case "catalog_date": VftProblemLabel = VfText("Katalog tarihi farkl\u0131. Kaynak incelemesi gerekir; otomatik aktar\u0131lmaz.")
        Case "catalog_missing": VftProblemLabel = VfText("Kaynak kimli\u011fi yeni katalogda bulunamad\u0131; eski kay\u0131t korundu.")
        Case "catalog_ambiguous": VftProblemLabel = VfText("Yeni katalogda ayn\u0131 kimlikle birden fazla kay\u0131t var; birle\u015ftirilmedi.")
    End Select
End Function

Private Sub VftHistoryShow()
    Dim item As Object
    VftClose
    Set mVftDoc = ActiveDocument: Set mVftItems = New Collection: mVftMode = "history"
    If Not mVftHistory Is Nothing Then
        For Each item In mVftHistory
            If item("Doc") Is mVftDoc Then mVftItems.Add item
        Next item
    End If
    Set mVftPanel = New VefeyatKisiSec0341
    VftDisplay VfText("Bu belgenin son i\u015flemleri. En \u00e7ok 100 ge\u00e7i\u015f; yaln\u0131z bu Word oturumunda.")
    mVftPanel.Show vbModeless
End Sub

Private Sub VftLocate(ByVal index As Long)
    Dim item As Object
    If mVftBusy Or mVftItems Is Nothing Then Exit Sub
    If index < 1 Or index > mVftItems.Count Then Exit Sub
    If Not VftOpenDoc(mVftDoc) Then VftClose: Exit Sub
    Set item = mVftItems(index)
    If Not item.Exists("Range") Then Exit Sub
    mVftDoc.Activate
    item("Range").Select
    Application.ActiveWindow.ScrollIntoView item("Range"), True
End Sub

Private Sub VftApply(ByVal index As Long, ByVal unlink As Boolean)
    Dim item As Object, field As Word.Field, doc As Word.Document, epoch As Long, detail As String
    Dim undoObject As Object, started As Boolean, changed As Boolean, logText As String
    If Not mVftDoc Is Nothing Then
        If Not VfpDateAllowed(mVftDoc) Then Exit Sub
    End If
    If mVftBusy Or mVftItems Is Nothing Or mVftMode = "history" Then Exit Sub
    If index < 1 Or index > mVftItems.Count Then Exit Sub
    Set doc = mVftDoc: epoch = mVftEpoch
    If Not VftOpenDoc(doc) Or Not (ActiveDocument Is doc) Then Exit Sub
    If Not VfaEditable(doc) Then Exit Sub
    Set item = mVftItems(index)
    If Not item.Exists("Field") Then Exit Sub
    Set field = item("Field")
    If CStr(item("Repair")) = "none" And Not unlink Then Exit Sub
    detail = CStr(item("Detail"))
    If unlink Then detail = detail & vbCrLf & VfText("Bu tarihin g\u00fcncelleme ve kaynak ba\u011f\u0131 kald\u0131r\u0131lacak. G\u00f6r\u00fcnen metin kalacak.")
    mVftBusy = True
    On Error GoTo Failed
    If VfMessage(detail & vbCrLf & VfText("Yaln\u0131z bu bulguya uygulans\u0131n m\u0131?"), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") <> vbYes Then GoTo Clean
    If epoch <> mVftEpoch Then GoTo Clean
    If Not VftOpenDoc(doc) Or Not (ActiveDocument Is doc) Then GoTo Clean
    If Not VfaEditable(doc) Or Not VfpDateAllowed(doc) Then GoTo Clean
    If Not (field.Result.Document Is doc) Then GoTo Stale
    If field.Code.Text <> CStr(item("Code")) Or field.Result.Text <> CStr(item("Result")) Then GoTo Stale
    If VftMeta(doc, CStr(item("Id"))) <> CStr(item("Meta")) Then GoTo Stale
    started = VfaBeginUndo(undoObject)
    changed = True
    If unlink Then
        field.Locked = False
        field.Unlink
    ElseIf CStr(item("Repair")) = "date" Then
        field.Locked = False
        field.Result.Text = CStr(item("Expected"))
        field.ShowCodes = False: field.Locked = True
    ElseIf CStr(item("Repair")) = "key" Then
        VfaWriteVariable doc, VfaVariableName(CStr(item("Id")), "Key"), CStr(item("Expected"))
    End If
    VfaEndUndo undoObject, started
    If unlink Then
        logText = VfText("Alan ba\u011f\u0131 kald\u0131r\u0131ld\u0131; g\u00f6r\u00fcnen metin korundu.")
    ElseIf CStr(item("Repair")) = "key" Then
        logText = VfText("\u015eah\u0131s ba\u011f\u0131 g\u00fcncellendi; tarih korundu.")
    Else
        logText = CStr(item("Expected"))
    End If
    VftLog doc, VfText("Belge sa\u011fl\u0131\u011f\u0131"), item("Range"), CStr(item("Result")), logText
    item("Repair") = "none": item("Detail") = VfText("\u0130\u015flem uyguland\u0131. G\u00fcncel bulgular i\u00e7in Belge sa\u011fl\u0131\u011f\u0131'n\u0131 yeniden \u00e7al\u0131\u015ft\u0131r\u0131n.")
    item("Label") = "OK - " & CStr(item("Label"))
    VftDisplay VfText("\u0130\u015flem tamamland\u0131. Yeni denetim kalan durumu g\u00f6sterir.")
    GoTo Clean
Stale:
    VfaWarn VfText("Bulgu g\u00f6r\u00fcnt\u00fclendikten sonra de\u011fi\u015fti. Eski \u00f6neri uygulanmad\u0131; yeniden denetleyin.")
    GoTo Clean
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not field Is Nothing Then field.Locked = True
    VfaEndUndo undoObject, started
    VfaReportFailure detail, changed
Clean:
    If epoch = mVftEpoch Then mVftBusy = False
End Sub

Private Function VfjRole(ByVal key As String) As Boolean
    Select Case key
        Case "dr", "prof", "doc", "yazar", "muellif", "sair", "filozof", "dusunur", "tarihci"
            VfjRole = True
    End Select
End Function

Private Function VfjOrganization(ByVal key As String) As Boolean
    Dim word As Variant
    For Each word In Split("universite|fakulte|enstitu|yayin|matbaa|kutuphane|devleti|bakanlik|belediye|vakfi|dernegi|konseyi|ansiklopedisi|sempozyum|kongresi", "|")
        If InStr(1, key, CStr(word), vbBinaryCompare) > 0 Then VfjOrganization = True: Exit Function
    Next word
End Function

Private Function VfjCased(ByVal text As String) As Boolean
    VfjCased = (UCase$(text) <> LCase$(text))
End Function

Private Function VfjUnknownMatches(ByVal text As String, ByVal patterns As Object, ByVal cache As Object) As Collection
    Dim result As New Collection, tokens As Variant, source As Collection, token As Variant, part As Variant
    Dim i As Long, j As Long, last As Long, first As Long, finish As Long, strong As Long
    Dim raw As String, key As String, role As Boolean, reject As Boolean, gap As String
    Set source = VfbTokens(text, cache)
    tokens = VfbTokenArray(source)
    i = 1
    Do While i <= source.Count
        token = tokens(i): raw = Mid$(text, CLng(token(1)) + 1, CLng(token(2)))
        key = CStr(token(0))
        If VfjRole(key) Or Not VfjCased(raw) Or Not VfbInitialUpper(raw) Or VfbCommonWord(key) Then GoTo NextWord
        last = VfbNameSpanEnd(text, tokens, i, source.Count, patterns)
        strong = 0: reject = False: role = False
        If i > 1 Then
            part = tokens(i - 1)
            gap = Mid$(text, CLng(part(1)) + CLng(part(2)) + 1, CLng(token(1)) - CLng(part(1)) - CLng(part(2)))
            role = VfjRole(CStr(part(0))) And VfjRoleGap(gap, CStr(part(0)))
        End If
        For j = i To last
            part = tokens(j): raw = Mid$(text, CLng(part(1)) + 1, CLng(part(2)))
            If VfjOrganization(CStr(part(0))) Or VfbCommonWord(CStr(part(0))) Then reject = True
            If Not VfbNameLink(CStr(part(0))) And Not VfbNoise(CStr(part(0))) Then
                If VfjCased(raw) And VfbInitialUpper(raw) Then strong = strong + 1 Else reject = True
            End If
        Next j
        first = CLng(token(1)): part = tokens(last): finish = CLng(part(1)) + CLng(part(2))
        raw = Mid$(text, first + 1, finish - first)
        If Not reject And role And strong >= 1 And VfaValidAlias(raw) Then
            If Not VfcExclude(text, first, finish - first, "", True) Then result.Add Array(first, finish - first, raw, "")
        End If
        i = last + 1
        GoTo ContinueScan
NextWord:
        i = i + 1
ContinueScan:
        If i Mod 64 = 0 Then VfbPulse VfText("Katalog d\u0131\u015f\u0131 ad adaylar\u0131 inceleniyor"), i, source.Count
    Loop
    Set VfjUnknownMatches = result
End Function

Private Function VfjRoleGap(ByVal gap As String, ByVal role As String) As Boolean
    Dim i As Long, ch As String
    If Len(gap) > 8 Then Exit Function
    For i = 1 To Len(gap)
        ch = Mid$(gap, i, 1)
        If ch <> " " And ch <> vbTab And ch <> ":" Then
            If ch <> "." Or (role <> "dr" And role <> "prof" And role <> "doc") Then Exit Function
        End If
    Next i
    VfjRoleGap = True
End Function

' Return literal adjacent text only. Never substitute the catalog death here.
Private Function VfjPlainDate(ByVal following As String, ByRef blocked As Boolean) As String
    Dim text As String, closing As Long, lead As Variant, marked As Boolean, inside As String
    text = VfaTrimSpaces(following)
    If Left$(text, 1) <> "(" Then Exit Function
    For Each lead In Array("(o.", VfText("(\u00f6."), VfText("(\u00f6l."), "(v.", "(d.")
        If Left$(LCase$(text), Len(CStr(lead))) = CStr(lead) Then marked = True
    Next lead
    closing = InStr(1, text, ")", vbBinaryCompare)
    If closing < 2 Or closing > VfaMaxDate Then
        blocked = marked
        Exit Function
    End If
    inside = Mid$(text, 2, closing - 2)
    If Not marked Then
        inside = VfaTrimSpaces(inside)
        If Len(inside) > 0 Then marked = VfcDigit(Left$(inside, 1)) And InStr(inside, "/") > 0
    End If
    If marked Then blocked = True: VfjPlainDate = Left$(text, closing)
End Function

Private Function VfjSpellingMatches(ByRef row As Variant, ByVal aliasText As String) As Boolean
    Dim i As Long, query As String, source As String, at As Long, before As Boolean, after As Boolean
    query = VfaNormalizeWhitespace(VfNameQuery(aliasText))
    If Len(query) = 0 Then Exit Function
    For i = 0 To 3
        If i = 3 Then source = VfbNaturalOrder(CStr(row(0))) Else source = CStr(row(i))
        source = VfaNormalizeWhitespace(source)
        at = InStr(1, source, query, vbBinaryCompare)
        Do While at > 0
            before = (at = 1): after = (at + Len(query) > Len(source))
            If Not before Then before = VfnNameBoundary(Mid$(source, at - 1, 1)) Or Mid$(source, at - 1, 1) = ","
            If Not after Then after = VfnNameBoundary(Mid$(source, at + Len(query), 1)) Or Mid$(source, at + Len(query), 1) = ","
            If before And after Then VfjSpellingMatches = True: Exit Function
            at = InStr(at + 1, source, query, vbBinaryCompare)
        Loop
    Next i
End Function

Private Function VfjOrderBefore(ByVal a As Object, ByVal b As Object) As Boolean
    If CLng(a("Order")) <> CLng(b("Order")) Then
        VfjOrderBefore = (CLng(a("Order")) < CLng(b("Order")))
    ElseIf CLng(a("Story")) <> CLng(b("Story")) Then
        VfjOrderBefore = (CLng(a("Story")) < CLng(b("Story")))
    Else
        VfjOrderBefore = (CLng(a("Offset")) < CLng(b("Offset")))
    End If
End Function

Private Sub VfjSort(ByRef items As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Object, swap As Object
    i = low: j = high: Set pivot = items((low + high) \ 2)
    Do While i <= j
        Do While VfjOrderBefore(items(i), pivot): i = i + 1: Loop
        Do While VfjOrderBefore(pivot, items(j)): j = j - 1: Loop
        If i <= j Then
            Set swap = items(i): Set items(i) = items(j): Set items(j) = swap
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then VfjSort items, low, j
    If i < high Then VfjSort items, i, high
End Sub

Private Sub VfjClose(Optional ByVal panelClosing As Boolean = False)
    Dim panel As VefeyatKisiSec0341, ownedRun As Boolean
    If mVfjClosing Then Exit Sub
    mVfjClosing = True: ownedRun = mVfjBusy
    mVfjEpoch = (mVfjEpoch Mod 2000000000) + 1
    On Error GoTo Clean
    Set panel = mVfjPanel: Set mVfjPanel = Nothing
    Set mVfjDoc = Nothing: Set mVfjCatalog = Nothing
    Set mVfjCache = Nothing: Set mVfjFirst = Nothing
    mVfjItems = Array(): mVfjCount = 0: mVfjReady = False
    mVfjBusy = False: mVfjPreparing = False: mVfjFaulted = False
    If ownedRun Then mVfbRunning = False: mVflSuspended = False
    If Not panel Is Nothing And Not panelClosing Then Unload panel
Clean:
    mVfjClosing = False
End Sub

Private Function VfjRangeKey(ByVal named As Word.Range) As String
    VfjRangeKey = CStr(named.StoryType) & ":" & CStr(named.Start) & ":" & CStr(named.End)
End Function

Private Function VfjLiveName(ByVal item As Object) As Word.Range
    Dim named As Word.Range
    On Error GoTo Unavailable
    If mVfjDoc Is Nothing Then Exit Function
    If Not VftOpenDoc(mVfjDoc) Then Exit Function
    Set named = item("Range")
    If Not (named.Document Is mVfjDoc) Then Exit Function
    If StrComp(named.Text, CStr(item("Alias")), vbBinaryCompare) <> 0 Then Exit Function
    Set VfjLiveName = named
Unavailable:
End Function

Private Function VfjDateState(ByVal named As Word.Range) As Object
    Dim result As Object, following As Word.Range, paragraph As Word.Range, field As Word.Field
    Dim finish As Long, id As String, blocked As Boolean, text As String
    Set result = CreateObject("Scripting.Dictionary")
    result("Text") = "": result("Blocked") = False: result("Id") = ""
    result("Code") = "": result("Result") = "": result("Meta") = ""
    Set following = named.Duplicate: Set paragraph = named.Paragraphs(1).Range
    finish = named.End + VfaMaxDate + 4
    If finish > paragraph.End Then finish = paragraph.End
    following.SetRange named.End, finish
    For Each field In following.Fields
        If field.Code.Start - 1 >= named.End And field.Code.Start - 1 <= named.End + 2 And Not VfdIsIndexMark(field) Then
            result("Blocked") = True
            If field.Code.End - field.Code.Start > 512 Or field.Result.End - field.Result.Start > 2000 Then
                result("Code") = "protected:" & CStr(field.Code.End - field.Code.Start) & ":" & CStr(field.Result.End - field.Result.Start)
                Set VfjDateState = result
                Exit Function
            End If
            result("Code") = field.Code.Text: result("Result") = field.Result.Text
            If field.Type = wdFieldAddin Then id = VfaFieldBinding(mVfjDoc, field)
            If Len(id) > 0 Then
                Set result("Field") = field
                result("Id") = id: result("Text") = VfaTrimSpaces(field.Result.Text)
                result("Meta") = VftMeta(mVfjDoc, id)
            End If
            Set VfjDateState = result
            Exit Function
        End If
    Next field
    following.TextRetrievalMode.IncludeFieldCodes = False
    text = VfjPlainDate(following.Text, blocked)
    result("Text") = text: result("Blocked") = blocked
    Set VfjDateState = result
End Function

Private Function VfjDateStamp(ByVal state As Object) As String
    VfjDateStamp = VfpKeyPart(CStr(state("Text"))) & VfpKeyPart(CStr(state("Blocked"))) & _
        VfpKeyPart(CStr(state("Code"))) & VfpKeyPart(CStr(state("Result"))) & VfpKeyPart(CStr(state("Meta")))
End Function

Private Sub VfjRememberFirst(ByVal item As Object, ByVal index As Long)
    Dim row As Variant, person As String
    If Not CBool(item("ChoiceReady")) Then Exit Sub
    row = item("Choice"): person = VfpPersonKey(row)
    If Not mVfjFirst.Exists(person) Then
        mVfjFirst(person) = index
    ElseIf index < CLng(mVfjFirst(person)) Then
        mVfjFirst(person) = index
    End If
End Sub

Private Sub VfjRefreshItem(ByVal item As Object)
    Dim named As Word.Range, dateState As Object, rows As Collection, row As Variant, candidate As Variant
    Dim key As String, catalogName As String, status As String, canChange As Boolean, canAdd As Boolean
    Dim context As Word.Range, detail As String, stored As String
    Set named = VfjLiveName(item)
    item("CanChange") = False: item("CanAdd") = False: item("CanIgnore") = False
    If named Is Nothing Then
        item("Status") = VfText("Metin de\u011fi\u015fti; d\u00f6k\u00fcm\u00fc yenileyin.")
        item("Detail") = CStr(item("Status"))
        Exit Sub
    End If
    item("CanIgnore") = True
    Set dateState = VfjDateState(named)
    item("Date") = CStr(dateState("Text"))
    Set rows = item("Rows")
    If dateState.Exists("Field") And Not CBool(item("ChoiceReady")) Then
        stored = VfaReadVariable(mVfjDoc, VfaVariableName(CStr(dateState("Id")), "Key"))
        For Each candidate In rows
            If VfpPersonKey(candidate) = stored Then item("Choice") = candidate: item("ChoiceReady") = True: Exit For
        Next candidate
    End If
    catalogName = VfText("Katalogdan se\u00e7ilmeli")
    status = VfText("Ad aday\u0131; kimlik do\u011frulanmad\u0131.")
    canChange = True
    If CBool(item("ChoiceReady")) Then
        row = item("Choice"): catalogName = VfnTextName(CStr(row(0)))
        canChange = Not VfjSpellingMatches(row, CStr(item("Alias")))
        If canChange Then status = VfText("Uyar\u0131: kaynak yaz\u0131m\u0131 farkl\u0131.") Else status = VfText("Kaynak yaz\u0131m\u0131yla uyumlu.")
        detail = VfText("Katalogdaki tarih: ") & CStr(row(3)) & vbCrLf & _
            VfText("Tam ad: ") & CStr(row(1)) & vbCrLf & CStr(row(4)) & vbCrLf & CStr(row(5))
    ElseIf rows.Count = 1 Then
        row = rows(1)
        catalogName = VfnTextName(CStr(row(0))) & " (?)"
        status = VfText("Uyar\u0131: yak\u0131n e\u015fle\u015fme; do\u011frulay\u0131n.")
        detail = CStr(row(1)) & vbCrLf & CStr(row(4))
    ElseIf rows.Count > 1 Then
        catalogName = CStr(rows.Count) & VfText(" aday; se\u00e7im gerekli")
        status = VfText("Uyar\u0131: birden fazla \u015fah\u0131s e\u015fle\u015fiyor.")
    End If
    canAdd = Not CBool(dateState("Blocked"))
    If dateState.Exists("Field") And Len(CStr(dateState("Text"))) = 0 Then canAdd = True
    If CBool(item("ChoiceReady")) Then
        If Not VfaValidDate(CStr(row(3))) Then canAdd = False: status = status & VfText(" Kaynakta tarih yok.")
    End If
    If Not canAdd And Len(CStr(dateState("Text"))) = 0 Then status = status & VfText(" Alan var; korunuyor.")
    Set context = VfwContextRange(named)
    detail = VfText("Metindeki ad: ") & CStr(item("Alias")) & vbCrLf & _
        VfText("Metindeki tarih: ") & CStr(dateState("Text")) & vbCrLf & status & vbCrLf & detail & vbCrLf & vbCrLf & _
        context.Text & vbCrLf & vbCrLf & VfText("De\u011fi\u015ftir ve Ekle yaln\u0131z bu ge\u00e7i\u015fe uygulan\u0131r. E\u015fle\u015fmeyen adlarda katalogdan aray\u0131p \u015fahs\u0131 do\u011frulay\u0131n.")
    item("Catalog") = catalogName: item("Status") = status: item("Detail") = detail
    If Not VfpDateAllowed(mVfjDoc) Then canChange = False: canAdd = False
    item("CanChange") = canChange: item("CanAdd") = canAdd
End Sub

Private Sub VfjPrepare()
    Dim patterns As Object, heads As Object, cache As Object, snapshots As New Collection
    Dim hits As Collection, all As New Collection, seen As Object, ranked As Object, aliases As Object
    Dim hit As Object, item As Object, snapshot As Object, candidates As Collection, match As Variant, row As Variant
    Dim named As Word.Range, source As Word.Range, note As Word.Footnote, anchor As Word.Range
    Dim key As String, i As Long, epoch As Long, count As Long
    mVfjReady = False: epoch = mVfjEpoch
    If Not VefeyatEnsureData() Then Exit Sub
    Set mVfjCatalog = mVfRows
    Set patterns = CreateObject("Scripting.Dictionary"): Set heads = CreateObject("Scripting.Dictionary")
    Set cache = CreateObject("Scripting.Dictionary"): Set ranked = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    Set aliases = VfiRead(mVfjDoc)
    VfxPrepareDocument mVfjDoc, patterns, heads, cache
    VfzBegin mVfjDoc, patterns, heads, cache
    Set hits = VfbOrderedHits(mVfjDoc, patterns, heads, cache, snapshots, True)
    VfxVerifySnapshots snapshots, cache
    For Each hit In hits
        Set named = hit("Range")
        If Len(CStr(hit("Key"))) > 0 And Not VfiContains(aliases, CStr(hit("Alias"))) Then
            seen(VfjRangeKey(named)) = True
            all.Add hit
        End If
    Next hit
    For Each snapshot In snapshots
        Set source = snapshot("Range")
        Set anchor = Nothing
        If snapshot.Exists("Anchor") Then Set anchor = snapshot("Anchor")
        For Each match In VfjUnknownMatches(CStr(snapshot("Text")), patterns, cache)
            Set named = source.Duplicate
            named.SetRange source.Start + CLng(match(0)), source.Start + CLng(match(0)) + CLng(match(1))
            If StrComp(named.Text, CStr(match(2)), vbBinaryCompare) = 0 Then
                If Not seen.Exists(VfjRangeKey(named)) And Not VfiContains(aliases, CStr(match(2))) Then
                    If Not VfzReferenceAt(named) Then
                        Set hit = CreateObject("Scripting.Dictionary")
                        Set hit("Range") = named: hit("Alias") = CStr(match(2)): hit("Key") = ""
                        If Not anchor Is Nothing Then Set hit("Anchor") = anchor
                        seen(VfjRangeKey(named)) = True: all.Add hit
                    End If
                End If
            End If
        Next match
        count = count + 1
        VfbPulse VfText("\u015eah\u0131s d\u00f6k\u00fcm\u00fc haz\u0131rlan\u0131yor"), count, snapshots.Count
        If epoch <> mVfjEpoch Then Err.Raise VfbCancelError, "Vefeyat", "Inventory closed."
    Next snapshot
    Set mVfjFirst = CreateObject("Scripting.Dictionary")
    Set mVfjCache = CreateObject("Scripting.Dictionary")
    For Each hit In all
        Set named = hit("Range"): Set candidates = New Collection
        key = CStr(hit("Key")) & vbNullChar & VfNameQuery(CStr(hit("Alias")))
        If Len(CStr(hit("Key"))) > 0 Then
            If Not ranked.Exists(key) Then Set ranked(key) = VfbRankRows(patterns(CStr(hit("Key"))), VfNameQuery(CStr(hit("Alias"))))
            Set candidates = ranked(key)
        End If
        Set item = CreateObject("Scripting.Dictionary")
        Set item("Range") = named.Duplicate: item("Alias") = CStr(hit("Alias"))
        Set item("Rows") = candidates
        item("ChoiceReady") = False: item("Date") = "": item("Catalog") = ""
        item("Status") = "": item("Detail") = "": item("CanChange") = False: item("CanAdd") = False
        item("Story") = 0&: item("Order") = named.Start: item("Offset") = named.Start
        item("Location") = VfText("Metin")
        If named.StoryType = wdFootnotesStory Then
            item("Story") = 1&: item("Location") = VfText("Dipnot")
            If hit.Exists("Anchor") Then
                Set anchor = hit("Anchor"): item("Order") = anchor.Start
                Set item("Anchor") = anchor.Duplicate
            Else
                Err.Raise VfaInvalidData, "Vefeyat", "Footnote inventory anchor is unavailable."
            End If
        End If
        If VfqSingleSafe(candidates, CStr(item("Alias"))) Then row = candidates(1): item("Choice") = row: item("ChoiceReady") = True
        ' Replace collection elements after enumeration by building a new array.
        Set hit("Inventory") = item
    Next hit
    mVfjCount = all.Count
    ReDim mVfjItems(0 To mVfjCount)
    i = 0
    For Each hit In all
        i = i + 1: Set mVfjItems(i) = hit("Inventory")
    Next hit
    If mVfjCount > 1 Then VfjSort mVfjItems, 1, mVfjCount
    For i = 1 To mVfjCount
        Set item = mVfjItems(i)
        VfjRefreshItem item
        VfjRememberFirst item, i
        If i Mod 32 = 0 Then VfbPulse VfText("Ad ve tarihler kar\u015f\u0131la\u015ft\u0131r\u0131l\u0131yor"), i, mVfjCount
    Next i
    If Not (mVfjCatalog Is mVfRows) Or epoch <> mVfjEpoch Then Err.Raise VfaInvalidData, "Vefeyat", "Inventory source changed."
    mVfjPage = 1: mVfjReady = True
End Sub

Private Sub VfjDisplay(Optional ByVal selected As Long = 1)
    Dim rows As New Collection, item As Object, i As Long, last As Long, firstShown As Long, edit As Boolean, heading As String
    If mVfjPanel Is Nothing Or Not mVfjReady Then Exit Sub
    last = mVfjPage + VfjPageSize - 1: If last > mVfjCount Then last = mVfjCount
    For i = mVfjPage To last
        Set item = mVfjItems(i)
        rows.Add Array(CStr(item("Alias")), CStr(item("Catalog")), CStr(item("Date")), CStr(item("Status")), _
            CStr(item("Detail")), CBool(item("CanChange")), CBool(item("CanAdd")), CStr(item("Location")))
    Next i
    edit = Not mVfjDoc.ReadOnly And mVfjDoc.ProtectionType = wdNoProtection And Not mVfjFaulted
    firstShown = mVfjPage: If mVfjCount = 0 Then firstShown = 0
    heading = CStr(mVfjCount) & VfText(" ad ge\u00e7i\u015fi | ") & CStr(mVfjCatalog.Count) & VfText(" yerel D\u0130A kayd\u0131 | ") & _
        CStr(firstShown) & "-" & CStr(last) & vbCrLf & VfText("Alg\u0131lanan adlar\u0131 do\u011frulay\u0131n. Sat\u0131r i\u015flemleri yaln\u0131z se\u00e7ili ge\u00e7i\u015fe uygulan\u0131r.")
    mVfjPanel.InitInventory rows, heading, edit, selected
    mVfjPanel.InventoryNavigation (mVfjPage > 1), (last < mVfjCount)
End Sub

Private Function VfjIndex(ByVal rowIndex As Long) As Long
    If rowIndex < 1 Or rowIndex > VfjPageSize Then Exit Function
    If mVfjPage < 1 Or mVfjPage > mVfjCount Then Exit Function
    If rowIndex > mVfjCount - mVfjPage + 1 Then Exit Function
    VfjIndex = mVfjPage + rowIndex - 1
End Function

Private Sub VfjSelection(ByVal rowIndex As Long)
    Dim index As Long, item As Object, edit As Boolean
    If mVfjBusy Or mVfjPanel Is Nothing Then Exit Sub
    index = VfjIndex(rowIndex): If index = 0 Then Exit Sub
    Set item = mVfjItems(index)
    If Not VftOpenDoc(mVfjDoc) Then VfjClose: Exit Sub
    If Not (mVfjCatalog Is mVfRows) Then
        mVfjPanel.InventoryState rowIndex, CStr(item("Alias")), "", CStr(item("Date")), VfText("Katalog de\u011fi\u015fti"), VfText("D\u00f6k\u00fcm\u00fc yenileyin; eski e\u015fle\u015fme uygulanmaz."), False, False
        Exit Sub
    End If
    VfjRefreshItem item
    edit = Not mVfjDoc.ReadOnly And mVfjDoc.ProtectionType = wdNoProtection And Not mVfjFaulted
    If mVfjFaulted Then item("Detail") = VfText("\u0130\u015flem hatas\u0131ndan sonra yazma durduruldu. Belgeyi kontrol edip d\u00f6k\u00fcm\u00fc yenileyin.")
    mVfjPanel.InventoryState rowIndex, CStr(item("Alias")), CStr(item("Catalog")), CStr(item("Date")), CStr(item("Status")), CStr(item("Detail")), edit And CBool(item("CanChange")), edit And CBool(item("CanAdd")), edit And CBool(item("CanIgnore"))
End Sub

Private Sub VfjLocateIndex(ByVal index As Long)
    Dim item As Object, named As Word.Range
    If index < 1 Or index > mVfjCount Then Exit Sub
    Set item = mVfjItems(index): Set named = VfjLiveName(item)
    If named Is Nothing Then VfaWarn VfText("Ad de\u011fi\u015fti veya silindi. D\u00f6k\u00fcm\u00fc yenileyin."): Exit Sub
    If Not mVfjPanel Is Nothing Then
        If index < mVfjPage Or index >= mVfjPage + VfjPageSize Then
            mVfjPage = ((index - 1) \ VfjPageSize) * VfjPageSize + 1
            VfjDisplay index - mVfjPage + 1
        End If
        mVfjPanel.CollapseInventory
    End If
    mVfjDoc.Activate: named.Select
    Application.ActiveWindow.ScrollIntoView named, True
End Sub

Private Function VfjChoose(ByVal item As Object) As Variant
    Dim picker As VefeyatKisiSec0341, row As Variant, number As Long, detail As String, rows As Collection, choice As Long, named As Word.Range
    On Error GoTo Failed
    Set rows = item("Rows")
    If rows.Count > 1 Then
        choice = VfePickPerson(rows, CStr(item("Alias")))
        If choice > 0 And choice <= rows.Count Then VfjChoose = rows(choice)
        Exit Function
    End If
    If CBool(item("ChoiceReady")) Then VfjChoose = item("Choice"): Exit Function
    Set picker = New VefeyatKisiSec0341
    picker.InitCatalogChoice CStr(item("Alias"))
    picker.Show vbModal
    row = picker.ChosenCatalogPerson()
    If Not IsEmpty(row) Then VfjChoose = row
    Unload picker
    Exit Function
Failed:
    number = Err.Number: detail = Err.Description
    On Error Resume Next
    Unload picker
    On Error GoTo 0
    Err.Raise number, "Vefeyat", detail
End Function

Private Function VfjSettings() As String
    VfjSettings = VfaReadVariable(mVfjDoc, VfaNamespaceVariable) & vbNullChar & VfpGetScope(mVfjDoc) & vbNullChar & VfpGetPrefix(mVfjDoc) & vbNullChar & _
        VfiPayload(VfiRead(mVfjDoc)) & vbNullChar & VfuNameMode(mVfjDoc) & vbNullChar & VfuNamesPayload(VfuNames(mVfjDoc))
End Function

Private Function VfjCurrent(ByVal item As Object, ByVal epoch As Long, ByVal context As Word.Range, ByVal originalContext As String, _
                            ByVal settings As String, ByVal dateStamp As String) As Boolean
    Dim named As Word.Range, state As Object
    If epoch <> mVfjEpoch Or mVfjPanel Is Nothing Or mVfjDoc Is Nothing Then Exit Function
    If Not VftOpenDoc(mVfjDoc) Then Exit Function
    If Not (ActiveDocument Is mVfjDoc) Or Not (mVfjCatalog Is mVfRows) Then Exit Function
    If mVfjDoc.ReadOnly Or mVfjDoc.ProtectionType <> wdNoProtection Then Exit Function
    Set named = VfjLiveName(item): If named Is Nothing Then Exit Function
    If StrComp(context.Text, originalContext, vbBinaryCompare) <> 0 Then Exit Function
    If VfjSettings() <> settings Then Exit Function
    If VfzReferenceAt(named) Or VfcRangeExcluded(named) Then Exit Function
    If VfiContains(VfiRead(mVfjDoc), named.Text) Then Exit Function
    If Not VfaSupportedRange(mVfjDoc, named) Then Exit Function
    If VfmTargetExcluded(mVfjDoc, named, CStr(item("Alias"))) Then Exit Function
    Set state = VfjDateState(named)
    If VfjDateStamp(state) <> dateStamp Then Exit Function
    VfjCurrent = True
End Function

Private Function VfjNewBinding(ByRef row As Variant, ByVal aliasText As String, ByVal raw As String, ByRef changed As Boolean) As String
    Dim number As Long, value As String, id As String
    number = 1: value = VfaReadVariable(mVfjDoc, VfaNextVariable)
    If IsNumeric(value) Then
        If CDbl(value) >= 1 And CDbl(value) < 2000000000# Then number = CLng(value)
    End If
    Do
        id = "P" & CStr(number): number = number + 1
    Loop While VfaVariableExists(mVfjDoc, VfaVariableName(id, "Key"))
    changed = True
    VfaWriteVariable mVfjDoc, VfaNextVariable, CStr(number)
    VfaWriteVariable mVfjDoc, VfaVariableName(id, "Alias"), aliasText
    VfaWriteVariable mVfjDoc, VfaVariableName(id, "Raw"), raw
    VfaWriteVariable mVfjDoc, VfaVariableName(id, "OccurrenceOnly"), "1"
    VfaWriteVariable mVfjDoc, VfaVariableName(id, "Key"), VfpPersonKey(row)
    VfjNewBinding = id
End Function

Private Function VfjSameRecorded(ByVal field As Word.Field, ByRef row As Variant) As Boolean
    Dim id As String, key As String, parts As Variant, recorded As Variant, index As Object, identity As String
    id = VfaFieldBinding(mVfjDoc, field)
    key = VfaReadVariable(mVfjDoc, VfaVariableName(id, "Key"))
    If key = VfpPersonKey(row) Then VfjSameRecorded = True: Exit Function
    parts = VfePersonParts(key): If IsEmpty(parts) Then Exit Function
    recorded = Array(CStr(parts(1)), CStr(parts(2)), CStr(parts(3)), "", "", CStr(parts(0)))
    identity = VfqIdentity(row)
    If Left$(identity, 4) <> "DIA:" Or identity <> VfqIdentity(recorded) Then Exit Function
    Set index = VfqIndex(mVfjCatalog)
    If index.Exists(identity) Then VfjSameRecorded = (index(identity).Count = 1)
End Function

Private Sub VfjApply(ByVal rowIndex As Long, ByVal changeName As Boolean)
    Dim index As Long, item As Object, named As Word.Range, context As Word.Range, state As Object, row As Variant
    Dim epoch As Long, original As String, contextText As String, settings As String, dateStamp As String, proposal As String
    Dim undoObject As Object, started As Boolean, changed As Boolean, detail As String, id As String, raw As String, oldId As String
    Dim field As Word.Field, first As Long, earlier As Long, decision As Long, person As String, repair As Boolean, registrySlot As Long
    If mVfjBusy Or mVfbRunning Or mVfrInsertBusy Or mVfjFaulted Or mVfjPanel Is Nothing Then Exit Sub
    index = VfjIndex(rowIndex): If index = 0 Then Exit Sub
    Set item = mVfjItems(index): Set named = VfjLiveName(item)
    If named Is Nothing Then VfjSelection rowIndex: Exit Sub
    If Not VftOpenDoc(mVfjDoc) Or Not (mVfjCatalog Is mVfRows) Then Exit Sub
    If Not VfaEditable(mVfjDoc) Or Not VfpDateAllowed(mVfjDoc) Then Exit Sub
    epoch = mVfjEpoch
    mVfjBusy = True: mVfbRunning = True: mVflSuspended = True
    On Error GoTo Failed
    VfjLocateIndex index
    If epoch <> mVfjEpoch Then GoTo Clean
    original = named.Text: Set context = VfwContextRange(named): contextText = context.Text
    settings = VfjSettings(): Set state = VfjDateState(named): dateStamp = VfjDateStamp(state)
    repair = state.Exists("Field") And Len(CStr(state("Text"))) = 0
    If Not changeName And CBool(state("Blocked")) And Not repair Then VfaWarn VfaDateBlockReason(named): GoTo Clean
    row = VfjChoose(item): If IsEmpty(row) Then GoTo Clean
    If Not VfjCurrent(item, epoch, context, contextText, settings, dateStamp) Then GoTo Stale
    If changeName Then
        proposal = VfnChooseName(row, VfNameQuery(original))
        If Len(proposal) = 0 Then GoTo Clean
        proposal = proposal & Mid$(original, Len(VfNameQuery(original)) + 1)
        If Not VfaValidAlias(proposal) Then GoTo Clean
        If StrComp(proposal, original, vbBinaryCompare) = 0 Then
            If item("Rows").Count > 1 Then
                VdiCheckStore named, row
                started = VfaBeginUndo(undoObject)
                VdiStore named, row
                VfaEndUndo undoObject, started
            End If
            item("Choice") = row: item("ChoiceReady") = True
            VfjRememberFirst item, index
            GoTo Clean
        End If
        If state.Exists("Field") Then
            Set field = state("Field")
            If Not VfjSameRecorded(field, row) Then VfaWarn VfText("Mevcut tarih ba\u015fka bir \u015fahsa ba\u011fl\u0131; ad ve tarih korunuyor."): GoTo Clean
        End If
        If VfMessage(VfText("Yaln\u0131z se\u00e7ili ge\u00e7i\u015fte ad de\u011fi\u015ftirilsin mi?") & vbCrLf & original & vbCrLf & proposal & vbCrLf & _
            VfText("Mevcut tarih korunur."), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") <> vbYes Then GoTo Clean
    Else
        If repair Then
            Set field = state("Field")
            If Not VfjSameRecorded(field, row) Then VfaWarn VfText("Bo\u015f tarih alan\u0131 ba\u015fka bir \u015fahsa ba\u011fl\u0131; otomatik doldurulmad\u0131."): GoTo Clean
        End If
        If Not VfaValidDate(CStr(row(3))) Then VfaWarn VfText("Katalogda kullan\u0131labilir \u00f6l\u00fcm tarihi yok."): GoTo Clean
        person = VfpPersonKey(row)
        If VfpGetScope(mVfjDoc) = "first" And mVfjFirst.Exists(person) Then
            earlier = CLng(mVfjFirst(person))
            If earlier < index Then
                If Not VfjLiveName(mVfjItems(earlier)) Is Nothing Then
                    decision = Vefeyat0341UiEarlierChoice(VfText("Bu \u015fah\u0131s d\u00f6k\u00fcmde daha \u00f6nce ge\u00e7iyor. \u0130sterseniz bu ge\u00e7i\u015fe yine de tarih ekleyebilirsiniz."))
                    If decision = 2 Then VfjLocateIndex earlier: GoTo Clean
                    If decision <> 1 Then GoTo Clean
                End If
            End If
        End If
    End If
    If Not VfjCurrent(item, epoch, context, contextText, settings, dateStamp) Then GoTo Stale
    If Not VdiRowCurrent(row) Then GoTo Stale
    Set named = VfjLiveName(item): first = named.Start
    If item("Rows").Count > 1 Then VdiCheckStore named, row
    started = VfaBeginUndo(undoObject)
    If changeName Then
        registrySlot = VrgNameSlot(mVfjDoc, named)
        changed = True: named.Text = proposal
        If Not field Is Nothing Then
            oldId = CStr(state("Id")): raw = VfaReadVariable(mVfjDoc, VfaVariableName(oldId, "Raw"))
            id = VfjNewBinding(row, proposal, raw, changed)
            If VfaReadVariable(mVfjDoc, VfaVariableName(oldId, "Pinned")) = "1" Then VfaWriteVariable mVfjDoc, VfaVariableName(id, "Pinned"), "1"
            field.Locked = False: field.Code.Text = " ADDIN " & VfaFieldTag(mVfjDoc, id) & " "
            field.ShowCodes = False: field.Locked = True
        End If
        named.SetRange first, first + Len(proposal)
        item("Alias") = proposal
    Else
        If Not repair And VfbAnyDate(named) Then VfaWarn VfaDateBlockReason(named): GoTo Clean
        VfaEnsureNamespace mVfjDoc, changed
        id = VfjNewBinding(row, original, CStr(row(3)), changed)
        If repair Then
            field.Locked = False
            field.Result.Text = " " & VfpFormatDate(CStr(row(3)), VfpGetPrefix(mVfjDoc))
            field.Code.Text = " ADDIN " & VfaFieldTag(mVfjDoc, id) & " "
            field.ShowCodes = False: field.Locked = True
            If StrComp(field.Result.Text, " " & VfpFormatDate(CStr(row(3)), VfpGetPrefix(mVfjDoc)), vbBinaryCompare) <> 0 Then Err.Raise VfaInvalidData, "Vefeyat", "Word did not retain the repaired date."
        Else
            VfaInsertManaged mVfjDoc, named, VfpFormatDate(CStr(row(3)), VfpGetPrefix(mVfjDoc)), id, changed, True
        End If
        named.SetRange first, first + Len(original)
    End If
    If changeName Then VrgCapture mVfjDoc, named, row, "name", registrySlot
    Set item("Range") = named.Duplicate
    item("Choice") = row: item("ChoiceReady") = True
    If item("Rows").Count > 1 Then VdiStore named, row
    VfjRememberFirst item, index
    VfaEndUndo undoObject, started
    VftLog mVfjDoc, VfText("\u015eah\u0131s d\u00f6k\u00fcm\u00fc"), named, original, named.Text
    VfmInvalidate mVfjDoc
    GoTo Clean
Stale:
    VfaWarn VfText("Ad, ba\u011flam, katalog veya ayarlar de\u011fi\u015fti. Eski \u00f6neri uygulanmad\u0131; d\u00f6k\u00fcm\u00fc yenileyin.")
    GoTo Clean
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not field Is Nothing Then field.Locked = True
    VfaEndUndo undoObject, started
    If changed And epoch = mVfjEpoch Then mVfjFaulted = True
    VfaReportFailure detail, changed
Clean:
    VfaEndUndo undoObject, started
    If epoch = mVfjEpoch Then
        mVfjBusy = False: mVfbRunning = False: mVflSuspended = False
        If Not mVfjPanel Is Nothing Then VfjSelection mVfjPanel.LiveReviewChoice()
    End If
    Vefeyat0341UiInvalidate
End Sub

Private Sub VfjPruneExcluded(ByVal exclusions As Object)
    Dim readIndex As Long, writeIndex As Long, item As Object
    For readIndex = 1 To mVfjCount
        Set item = mVfjItems(readIndex)
        If Not VfiContains(exclusions, CStr(item("Alias"))) Then
            writeIndex = writeIndex + 1
            Set mVfjItems(writeIndex) = item
        End If
    Next readIndex
    mVfjCount = writeIndex
    ReDim Preserve mVfjItems(0 To mVfjCount)
    Set mVfjFirst = CreateObject("Scripting.Dictionary")
    For readIndex = 1 To mVfjCount
        VfjRememberFirst mVfjItems(readIndex), readIndex
    Next readIndex
    If mVfjCount = 0 Then
        mVfjPage = 1
    ElseIf mVfjPage > mVfjCount Then
        mVfjPage = ((mVfjCount - 1) \ VfjPageSize) * VfjPageSize + 1
    End If
End Sub

Private Sub VfjIgnore(ByVal rowIndex As Long)
    Dim index As Long, item As Object, named As Word.Range, context As Word.Range, state As Object
    Dim epoch As Long, literal As String, contextText As String, settings As String, dateStamp As String
    Dim exclusions As Object, updated As Object, payload As String, key As String
    Dim undoObject As Object, started As Boolean, changed As Boolean, refreshList As Boolean, detail As String
    If mVfjBusy Or mVfbRunning Or mVfrInsertBusy Or mVfjFaulted Or mVfjPanel Is Nothing Then Exit Sub
    index = VfjIndex(rowIndex): If index = 0 Then Exit Sub
    Set item = mVfjItems(index): Set named = VfjLiveName(item)
    If named Is Nothing Then VfjSelection rowIndex: Exit Sub
    If Not VfaEditable(mVfjDoc) Then Exit Sub
    epoch = mVfjEpoch
    mVfjBusy = True: mVfbRunning = True: mVflSuspended = True
    On Error GoTo Failed
    mVfjDoc.Activate
    If epoch <> mVfjEpoch Then GoTo Clean
    literal = named.Text
    Set context = VfwContextRange(named): contextText = context.Text
    settings = VfjSettings(): Set state = VfjDateState(named): dateStamp = VfjDateStamp(state)
    Set exclusions = VfiRead(mVfjDoc)
    If VfiContains(exclusions, literal) Then
        VfjPruneExcluded exclusions
        refreshList = True
        GoTo Clean
    End If
    If Not VfjCurrent(item, epoch, context, contextText, settings, dateStamp) Then GoTo Stale
    key = VfiKey(literal)
    If Len(key) = 0 Then GoTo Clean
    Set updated = VfiDecode(VfiPayload(exclusions))
    updated(key) = True
    payload = VfiPayload(updated)
    If VfMessage(VfText("Bu ifade bu belgede \u015fah\u0131s olarak de\u011ferlendirilmesin mi?") & vbCrLf & vbCrLf & _
        literal & vbCrLf & vbCrLf & VfText("Ayn\u0131 yaz\u0131mdaki t\u00fcm ge\u00e7i\u015fler \u015fah\u0131s d\u00f6k\u00fcm\u00fc ve toplu denetim d\u0131\u015f\u0131nda kal\u0131r. Metin ve tarihler korunur. Tercih yaln\u0131z bu belgeye kaydedilir; belge istisnalar\u0131ndan kald\u0131r\u0131labilir."), _
        vbYesNo Or vbQuestion Or vbDefaultButton2, VfText("\u015eah\u0131s de\u011fil")) <> vbYes Then GoTo Clean
    If Not VfjCurrent(item, epoch, context, contextText, settings, dateStamp) Then GoTo Stale
    started = VfaBeginUndo(undoObject)
    changed = True
    VfaWriteVariable mVfjDoc, VfiVariable, payload
    VfaEndUndo undoObject, started
    VfjPruneExcluded updated
    refreshList = True
    VftLog mVfjDoc, VfText("\u015eah\u0131s de\u011fil"), named, literal, literal
    VfmInvalidate mVfjDoc
    GoTo Clean
Stale:
    VfaWarn VfText("Ad, ba\u011flam, katalog veya ayarlar de\u011fi\u015fti. \u0130stisna kaydedilmedi; d\u00f6k\u00fcm\u00fc yenileyin.")
    GoTo Clean
Failed:
    detail = Err.Description
    On Error Resume Next
    VfaEndUndo undoObject, started
    If changed And epoch = mVfjEpoch Then mVfjFaulted = True
    VfaWarn VfText("\u015eah\u0131s de\u011fil tercihi kaydedilemedi: ") & detail
Clean:
    VfaEndUndo undoObject, started
    If epoch = mVfjEpoch Then
        mVfjBusy = False: mVfbRunning = False: mVflSuspended = False
        If Not mVfjPanel Is Nothing Then
            If refreshList Then
                If mVfjCount > 0 Then
                    If rowIndex > mVfjCount - mVfjPage + 1 Then rowIndex = mVfjCount - mVfjPage + 1
                End If
                VfjDisplay rowIndex
            Else
                VfjSelection mVfjPanel.LiveReviewChoice()
            End If
        End If
    End If
    Vefeyat0341UiInvalidate
End Sub

Public Function Vefeyat0341UiNativeChecks() As String
    Dim doc As Word.Document, original As Word.Document, named As Word.Range, field As Word.Field, note As Word.Footnote
    Dim row As Variant, id As String, changed As Boolean, count As Long, started As Single, elapsed As Double
    Dim picker As VefeyatKisiSec0341, undoObject As Object, undoStarted As Boolean, before As String, hadUndo As Boolean, undoResult As String
    Dim i As Long, text As String, maximum As Double, shading As Long
    On Error GoTo Failed
    If mVfbRunning Or mVfrInsertBusy Then Err.Raise VfuInvalid, "Vefeyat", "Another command is running."
    If Documents.Count > 0 Then Set original = ActiveDocument
    mVfrInsertBusy = True
    Set doc = Documents.Add(Visible:=True)
    doc.Content.Text = "Vefeyat Acceptance Person" & vbCr
    row = Array("Vefeyat Acceptance Person", "Synthetic Full Name", "", "(v. 500/1106'dan sonra)", "Synthetic", "")
    Set named = doc.Range(0, Len(CStr(row(0))))
    VfaEnsureNamespace doc, changed
    id = VfaGetBinding(doc, row, CStr(row(0)), changed)
    before = doc.Content.Text
    named.Select
    shading = ActiveWindow.View.FieldShading
    VfvBeginCommand
    undoStarted = VfaBeginUndo(undoObject)
    hadUndo = undoStarted
    started = Timer
    VfaInsertManaged doc, named, CStr(row(3)), id, changed
    VfaEndUndo undoObject, undoStarted
    elapsed = VfsElapsed(started) * 1000#
    VftAssert doc.Fields.Count = 1, "managed field count", count
    Set field = doc.Fields(1)
    VftAssert field.Type = wdFieldAddin And field.Locked And Not field.ShowCodes, "managed result contract", count
    VftAssert field.Result.Text = " " & CStr(row(3)), "uncertainty retained", count
    VfvFinishCommand
    VftAssert Selection.Range.Start = field.Result.End + 1 And Selection.Range.Start = Selection.Range.End, "caret outside the inserted managed field", count
    VftAssert ActiveWindow.View.FieldShading = shading, "normal insertion preserves Word field display preference", count
    VftAssert VftProblem(CStr(row(3)), CStr(row(0)), VfpPersonKey(row), field.Result.Text, "original") = "", "healthy field", count
    undoResult = "UNSUPPORTED"
    If hadUndo Then
        VftAssert doc.Undo(1), "custom Undo available", count
        VftAssert doc.Content.Text = before And doc.Fields.Count = 0, "custom Undo restores visible text and fields", count
        undoResult = "PASS"
        Set named = doc.Range(0, Len(CStr(row(0))))
        VfaInsertManaged doc, named, CStr(row(3)), id, changed
        Set field = doc.Fields(1)
    End If
    field.Locked = False: field.Result.Text = "": field.Locked = True
    VftAssert VftProblem(CStr(row(3)), CStr(row(0)), VfpPersonKey(row), field.Result.Text, "original") = "empty", "empty result diagnosis", count
    field.Locked = False: field.Result.Text = " " & CStr(row(3)): field.Locked = True
    Set picker = New VefeyatKisiSec0341
    VftAssert picker.ConfirmationVersion = 15, "form contract", count
    picker.InitLiveReview
    picker.Show vbModeless
    Unload picker: Set picker = Nothing
    VftAssert doc.Content.Text <> before, "form close preserves text", count
    text = "a" & Chr$(1) & Chr$(2) & Chr$(7) & Chr$(11) & Chr$(12) & Chr$(19) & Chr$(21) & Chr$(30) & Chr$(31) & Chr$(127) & "%b" & vbTab & vbCr & vbLf
    VftAssert VdiDecode(VdiEncode(text)) = text, "identity metadata keeps Word control marks", count
    VftAssert InStr(1, VdiEncode(text), Chr$(2), vbBinaryCompare) = 0, "identity metadata stores no raw control marks", count
    VftInventoryNative count
    VftBulkNative count
    VftRegistryNative count
    Set note = doc.Footnotes.Add(Range:=doc.Range(0, 0), Text:="Synthetic Person")
    Set named = note.Range.Duplicate
    named.End = named.Start + Len("Synthetic Person")
    VfaInsertManaged doc, named, "(v. 1)", id, changed
    VftAssert note.Range.Fields.Count = 1, "footnote field insertion", count
    VftIndexNative doc, count
    VftMarkNative doc, count
    VftPrefixNative doc, count
    text = ""
    For i = 1 To 200
        text = text & "Synthetic Person synthetic acceptance paragraph." & vbCr
        If i < 200 Then text = text & Chr$(12)
    Next i
    doc.Content.Text = text
    doc.Repaginate
    maximum = 0
    For i = 1 To 10
        Set named = doc.Range(0, Len("Synthetic Person"))
        started = Timer
        VfaInsertManaged doc, named, "(v. 1)", id, changed
        If VfsElapsed(started) * 1000# > maximum Then maximum = VfsElapsed(started) * 1000#
    Next i
    VftAssert doc.Fields.Count = 10, "large synthetic document field writes", count
    VftFinalizeNative doc, count
    VfaEnsureNamespace doc, changed
    id = VfaGetBinding(doc, row, CStr(row(0)), changed)
    doc.TrackRevisions = True
    Set named = doc.Range(0, Len("Synthetic Person"))
    VfaInsertManaged doc, named, "(v. 2)", id, changed
    VftAssert doc.Revisions.Count > 0, "tracked managed field creates a revision", count
    Vefeyat0341UiNativeChecks = "PASS " & CStr(count) & "; field_ms=" & CStr(CLng(elapsed)) & "; large_field_max_ms=" & CStr(CLng(maximum)) & _
        "; Word=" & Application.Version & "; customUndo=" & undoResult & "; control_marks=PASS; person_index=PASS; index_marks_neutral=PASS; prefix_reformat=PASS; finalize=PASS; bulk_scope_undo=PASS; bulk_tracking=preview-only; registry_save_reopen=PASS; other_computer=PENDING; end_to_end_author_review=PENDING; coauthoring=PENDING"
    GoTo Clean
Failed:
    Vefeyat0341UiNativeChecks = "FAIL " & CStr(Err.Number) & ": " & Err.Description
Clean:
    On Error Resume Next
    VfvCancelCommand
    VfaEndUndo undoObject, undoStarted
    If Not picker Is Nothing Then Unload picker
    If Not doc Is Nothing Then doc.Close SaveChanges:=wdDoNotSaveChanges
    If Not original Is Nothing Then original.Activate
    mVfrInsertBusy = False
End Function

Private Sub VftAssert(ByVal condition As Boolean, ByVal label As String, ByRef count As Long)
    If Not condition Then Err.Raise VfuInvalid, "Vefeyat native checks", label
    count = count + 1
End Sub

Private Sub VftInventoryNative(ByRef count As Long)
    Dim picker As VefeyatKisiSec0341, rows As New Collection, number As Long, detail As String
    On Error GoTo Failed
    Set picker = New VefeyatKisiSec0341
    picker.InitInventory rows, "Synthetic inventory", False
    VftAssert picker.Controls("Results").ColumnCount = 5, "inventory columns", count
    VftAssert Not picker.Controls("ChooseButton").Enabled And Not picker.Controls("StopScanButton").Enabled, "empty inventory cannot write", count
    picker.InventoryNavigation False, False
    picker.CollapseInventory
    picker.Show vbModeless
    Unload picker
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    On Error Resume Next
    Unload picker
    On Error GoTo 0
    Err.Raise number, "Vefeyat", detail
End Sub

' Disposable native scenarios. This procedure is never run at startup.
Private Sub VftBulkNative(ByRef count As Long)
    Dim doc As Word.Document, owner As Word.Document, oldCatalog As Collection, rows As New Collection, row As Variant
    Dim records As Collection, item As Object, plan As Collection, selected As Object, change As Object
    Dim named As Word.Range, second As Word.Range, note As Word.Footnote, table As Word.Table, shape As Word.Shape
    Dim before As String, cellBefore As String, marker As String, number As Long, detail As String, oldFields As Long
    On Error GoTo Failed
    If Documents.Count > 0 Then Set owner = ActiveDocument
    Set oldCatalog = mVfRows
    row = Array("Vefeyat Bulk Person", "Vefeyat Bulk Person", "", "(v. 500/1106)", "Synthetic native acceptance fixture", "")
    rows.Add row: Set mVfRows = rows
    Set doc = Documents.Add(Visible:=True)
    doc.Content.Text = CStr(row(0)) & vbCr & CStr(row(0)) & vbCr
    doc.Paragraphs(1).Style = wdStyleHeading1
    Set named = doc.Range(0, Len(CStr(row(0))))
    named.Font.Bold = True
    Set second = doc.Range(named.End + 1, named.End + 1 + Len(CStr(row(0))))
    before = doc.Content.Text
    Set records = New Collection
    records.Add VftBulkItem(named, rows, "1")
    records.Add VftBulkItem(second, rows, "2")
    VftAssert VfpGetScope(doc) = "first", "bulk default first use", count
    Set plan = VbpDecide(records, "first"): Set selected = VftBulkSelection(plan)
    VftAssert selected.Count = 1, "native first use proposal", count
    VbpApply doc, records, plan, selected
    VftAssert doc.Fields.Count = 1, "native first use inserts one date", count
    VftAssert doc.Paragraphs(1).Range.Style.NameLocal = doc.Styles(wdStyleHeading1).NameLocal, "heading paragraph style preserved", count
    VftAssert doc.Range(0, Len(CStr(row(0)))).Font.Bold = True, "original name font preserved", count
    VftAssert doc.Undo(1), "bulk native Undo available", count
    VftAssert doc.Content.Text = before And doc.Fields.Count = 0, "bulk Undo restores text and fields", count
    VftAssert VfaActiveBindingIds(doc).Count = 0, "orphan metadata inactive after Undo", count
    VftAssert doc.Redo(1), "bulk native Redo available", count
    VftAssert doc.Fields.Count = 1, "bulk Redo restores date field", count
    VftAssert doc.Undo(1), "second bulk Undo", count
    VfpSetScope doc, "all"
    Set records = New Collection
    Set named = doc.Range(0, Len(CStr(row(0))))
    Set second = doc.Range(named.End + 1, named.End + 1 + Len(CStr(row(0))))
    records.Add VftBulkItem(named, rows, "1")
    records.Add VftBulkItem(second, rows, "2")
    Set plan = VbpDecide(records, "all"): Set selected = VftBulkSelection(plan)
    VbpApply doc, records, plan, selected
    VftAssert doc.Fields.Count = 2, "native all uses inserts two dates", count
    Set table = doc.Tables.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), 1, 2)
    table.Cell(1, 1).Range.Text = CStr(row(0))
    table.Cell(1, 2).Range.Text = "Keep citation: Work, 1/13."
    cellBefore = table.Cell(1, 2).Range.Text
    Set named = table.Cell(1, 1).Range.Duplicate
    named.End = named.Start + Len(CStr(row(0)))
    Set records = New Collection: records.Add VftBulkItem(named, rows, "1")
    Set plan = VbpDecide(records, "all"): Set selected = VftBulkSelection(plan)
    VbpApply doc, records, plan, selected
    VftAssert table.Rows.Count = 1 And table.Columns.Count = 2, "table dimensions preserved", count
    VftAssert table.Cell(1, 2).Range.Text = cellBefore, "other table cell citation preserved", count
    VftAssert table.Cell(1, 1).Range.Fields.Count = 1, "date inside selected table cell", count
    Set note = doc.Footnotes.Add(doc.Range(0, 0), Text:=CStr(row(0)) & " explanation.")
    marker = note.Reference.Text
    Set named = note.Range.Duplicate: named.End = named.Start + Len(CStr(row(0)))
    Set records = New Collection: records.Add VftBulkItem(named, rows, "1", note.Reference)
    Set plan = VbpDecide(records, "all"): Set selected = VftBulkSelection(plan)
    VbpApply doc, records, plan, selected
    VftAssert note.Range.Fields.Count = 1 And note.Reference.Text = marker And doc.Footnotes.Count = 1, "footnote date and reference preserved", count
    doc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = CStr(row(0))
    Set shape = doc.Shapes.AddTextbox(1, 10, 10, 180, 40)
    shape.TextFrame.TextRange.Text = CStr(row(0))
    doc.Endnotes.Add doc.Range(0, 0), Text:=CStr(row(0))
    detail = VbsReport(doc)
    VftAssert InStr(detail, "sonnotlar") > 0 And InStr(detail, VfText("metin kutular\u0131")) > 0, "unsupported stories reported", count
    VftAssert shape.TextFrame.TextRange.Fields.Count = 0 And doc.Endnotes(1).Range.Fields.Count = 0, "unsupported stories untouched", count
    doc.Content.InsertAfter vbCr & CStr(row(0)) & vbCr
    Set named = doc.Content.Duplicate
    named.Find.ClearFormatting
    VftAssert named.Find.Execute(FindText:=CStr(row(0)), Forward:=False, Wrap:=wdFindStop), "native final name found", count
    VftAssert named.Text = CStr(row(0)), "final native test range", count
    Set records = New Collection: records.Add VftBulkItem(named, rows, "1")
    Set plan = VbpDecide(records, "all"): Set selected = VftBulkSelection(plan)
    VftAssert selected.Count = 1, "stale native proposal is writable before none mode", count
    before = doc.Content.Text: oldFields = doc.Fields.Count
    VfpSetScope doc, "none"
    VbpApply doc, records, plan, selected
    VftAssert doc.Content.Text = before And doc.Fields.Count = oldFields, "review-only mode rejects stale writes", count
    VfpSetScope doc, "all": doc.TrackRevisions = True
    On Error Resume Next
    Err.Clear
    VbpApply doc, records, plan, selected
    number = Err.Number
    Err.Clear
    On Error GoTo Failed
    VftAssert number = VbpInvalid And doc.Content.Text = before And doc.Fields.Count = oldFields, "tracking bulk blocked without mutations", count
    VftAssert doc.TrackRevisions, "tracking preference preserved", count
    named.InsertBefore "Tracked fixture "
    doc.TrackRevisions = False
    VftAssert doc.Revisions.Count > 0, "pending native revision exists", count
    Set named = doc.Content.Duplicate: named.Find.ClearFormatting
    VftAssert named.Find.Execute(FindText:=CStr(row(0)), Forward:=False, Wrap:=wdFindStop), "pending revision fixture name", count
    Set records = New Collection: records.Add VftBulkItem(named, rows, "1")
    Set plan = VbpDecide(records, "all"): Set selected = VftBulkSelection(plan)
    VftAssert selected.Count = 1, "pending revision fixture proposal", count
    before = doc.Content.Text: oldFields = doc.Fields.Count
    On Error Resume Next
    Err.Clear
    VbpApply doc, records, plan, selected
    number = Err.Number
    Err.Clear
    On Error GoTo Failed
    VftAssert number = VbpInvalid And doc.Content.Text = before And doc.Fields.Count = oldFields, "pending revisions block even with tracking off", count
    number = 0
    GoTo Clean
Failed:
    number = Err.Number: detail = Err.Description
Clean:
    On Error Resume Next
    Set mVfRows = oldCatalog
    If Not doc Is Nothing Then doc.Close SaveChanges:=wdDoNotSaveChanges
    If Not owner Is Nothing Then owner.Activate
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "Vefeyat bulk native checks", detail
End Sub

Private Function VftBulkItem(ByVal named As Word.Range, ByVal rows As Collection, ByVal id As String, _
                             Optional ByVal anchor As Word.Range = Nothing) As Object
    Dim hit As Object, item As Object
    Set hit = CreateObject("Scripting.Dictionary")
    Set hit("Range") = named: hit("Alias") = named.Text
    If Not anchor Is Nothing Then Set hit("Anchor") = anchor
    Set item = VbpRecord(hit, rows): item("Id") = id
    Set VftBulkItem = item
End Function

Private Function VftBulkSelection(ByVal plan As Collection) As Object
    Dim result As Object, change As Object
    Set result = CreateObject("Scripting.Dictionary")
    For Each change In plan
        If CStr(change("Kind")) <> "info" Then result(CStr(change("Key"))) = True
    Next change
    Set VftBulkSelection = result
End Function

' Explicit only. Save/reopen does not certify another physical computer.
Private Sub VftRegistryNative(ByRef count As Long)
    Dim doc As Word.Document, owner As Word.Document, oldCatalog As Collection, rows As New Collection, row As Variant
    Dim named As Word.Range, records As Object, before As String, file As String, fs As Object, created As Boolean
    Dim number As Long, detail As String, slot As Long
    On Error GoTo Failed
    If Documents.Count > 0 Then Set owner = ActiveDocument
    Set oldCatalog = mVfRows
    row = Array("Vefeyat Registry Fixture", "Synthetic Registry Full Name", "", "(v. 500/1106)", "Synthetic acceptance fixture, not a biographical claim.", "https://islamansiklopedisi.org.tr/vefeyat-test-registry-fixture")
    rows.Add row: Set mVfRows = rows
    Set doc = Documents.Add(Visible:=True)
    doc.Content.Text = CStr(row(0)) & " (v. 999)." & vbCr
    Set named = doc.Range(0, Len(CStr(row(0))))
    before = doc.Content.Text
    VrgCapture doc, named, row, "register"
    VftAssert VrgRead(doc).Count = 0 And doc.Content.Text = before, "registry default off", count
    VfaWriteVariable doc, VrgMode, "1"
    VftAssert VrgApi(doc), "native bookmark and variable capability", count
    VrgStore doc, named, row, "register"
    VrgStore doc, named, row, "register"
    Set records = VrgRead(doc)
    VftAssert records.Count = 1 And doc.Bookmarks.Exists("_Vrg_1"), "registry repeat has one anchor", count
    VftAssert doc.Content.Text = before And doc.Fields.Count = 0, "legacy plain text preserved", count
    VftAssert doc.ContentControls.Count = 0, "registry creates no content control", count
    Set fs = CreateObject("Scripting.FileSystemObject")
    file = fs.BuildPath(fs.GetSpecialFolder(2), fs.GetTempName & ".docx")
    If fs.FileExists(file) Then Err.Raise VrgInvalid, "Vefeyat", "Temporary test path already exists."
    On Error Resume Next
    Err.Clear
    CallByName doc, "SaveAs2", VbMethod, file, 12
    If Err.Number <> 0 Then
        Err.Clear
        CallByName doc, "SaveAs", VbMethod, file, 12
    End If
    number = Err.Number: detail = Err.Description
    created = fs.FileExists(file)
    On Error GoTo Failed
    If number <> 0 Then Err.Raise number, "Vefeyat", detail
    doc.Close SaveChanges:=wdDoNotSaveChanges: Set doc = Nothing
    Set mVfRows = Nothing: Set mVrgCatalogOwner = Nothing
    Set mVrgCatalogIndex = Nothing: Set mVrgLegacyIndex = Nothing
    Set doc = Documents.Open(FileName:=file, ReadOnly:=False, AddToRecentFiles:=False)
    Set records = VrgRead(doc)
    VftAssert records.Count = 1 And doc.Content.Text = before, "DOCX registry source/text survives save/reopen", count
    Set named = VrgLive(doc, records("1"))
    VftAssert Not named Is Nothing, "reopened registry bookmark resolves without local catalogue", count
    VftAssert named.Text = CStr(row(0)) And CStr(records("1")(3)) = VfqIdentity(row), "reopened stable identity and spelling", count
    VfaWriteVariable doc, VrgMode, "0"
    VftAssert VrgRead(doc).Count = 1 And doc.Content.Text = before, "disable retains portable registry", count
    number = 0
    GoTo Clean
Failed:
    number = Err.Number: detail = Err.Description
Clean:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=wdDoNotSaveChanges
    If created Then fs.DeleteFile file, True
    Set mVfRows = oldCatalog: Set mVrgCatalogOwner = Nothing
    Set mVrgCatalogIndex = Nothing: Set mVrgLegacyIndex = Nothing
    If Not owner Is Nothing Then owner.Activate
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "Vefeyat registry native checks", detail
End Sub

Public Sub Vefeyat0341UiOnLoad(ByVal ribbon As Object)
    mVflSuspended = False
    mVflPending = False
    Set mVflStates = Nothing
    If Not mVfrEvents Is Nothing Then mVfrEvents.Detach
    Set mVfrRibbon = ribbon
    Set mVfrEvents = New VefeyatOlaylar0341
    mVfrEvents.Attach Application, ribbon
    VflArm
End Sub

Public Sub Vefeyat0341UiInvalidate()
    If Not mVfrRibbon Is Nothing Then mVfrRibbon.Invalidate
End Sub

Public Function Vefeyat0341UiReady() As Boolean
    Vefeyat0341UiReady = Not mVfrRibbon Is Nothing
End Function

Public Sub Vefeyat0341UiInsert(ByVal control As Object)
    Vefeyat0341OlumTarihiEkle
End Sub

Public Sub Vefeyat0341OlumTarihiEkle()
    Dim tick As Single, number As Long, detail As String
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    tick = VfgClock()
    On Error GoTo Failed
    mVfrInsertBusy = True
    mVflSuspended = True
    Application.StatusBar = VfText("Vefey\u00e2t: se\u00e7ili ad aran\u0131yor...")
    VfvBeginCommand
    VfaInsertSelected
    If mVfaCommandChanged Then VfvFinishCommand Else VfvCancelCommand
    mVflSuspended = False
    If mVfaCommandChanged And Application.Documents.Count > 0 Then VfmObserve ActiveDocument, Selection.Range
    VflCaptureCommand
    VflArm
    mVfrInsertBusy = False
    VfgEnd "insert", tick
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    VfvCancelCommand
    mVflSuspended = False
    mVfrInsertBusy = False
    VfgEnd "insert", tick, number, mVfaCommandChanged
    VfrReport detail
End Sub

Public Sub Vefeyat0341UiCanInsert(ByVal control As Object, ByRef enabled As Variant)
    On Error GoTo Unavailable
    ' Keep the command available in an editable document. The action validates
    ' the live selection and explains a rejected selection before any change.
    enabled = VfrEditableDocument()
    If mVfbRunning Or mVfrInsertBusy Then enabled = False
    VflArm
    Exit Sub
Unavailable:
    enabled = False
End Sub

Public Sub Vefeyat0341UiDisconnect()
    VfvCancelCommand
    Set mVfvViews = Nothing
    VftClose
    VfjClose
    VfwClose
    mVflSuspended = True
    If Not mVfrEvents Is Nothing Then mVfrEvents.Detach
    Set mVfrEvents = Nothing
    Set mVfrRibbon = Nothing
    Set mVfmStates = Nothing
End Sub

Public Function Vefeyat0341UiExplain(ByRef row As Variant, ByVal query As String) As String
    If Len(query) = 0 Then
        If Not mVfwHits Is Nothing Then
            If mVfwIndex >= 1 And mVfwIndex <= mVfwHits.Count Then query = CStr(mVfwHits(mVfwIndex)("Alias"))
        End If
    End If
    Vefeyat0341UiExplain = VfqExplain(row, query)
End Function

Public Sub Vefeyat0341UiTools(ByVal control As Object)
    Dim number As Long, detail As String
    Dim doc As Word.Document, stamp As String
    If mVfbRunning Or mVfrInsertBusy Or Documents.Count = 0 Then Exit Sub
    On Error GoTo Failed
    Set doc = ActiveDocument
    mVfrInsertBusy = True
    Select Case CStr(control.Tag)
        Case "health": VftHealth False
        Case "catalog": VftHealth True
        Case "history": VftHistoryShow
        Case "export": VfuProject False
        Case "import": VfuProject True
        Case "remember_name": VfuRememberName
        Case "clear_names"
            If VfrEditableDocument() Then
                stamp = VfaReadVariable(doc, VfuNamesVariable)
                If VfMessage(VfText("Bu belgenin tercih edilen \u015fah\u0131s yaz\u0131mlar\u0131 temizlensin mi? Metin de\u011fi\u015fmez."), vbYesNo Or vbQuestion Or vbDefaultButton2, "Vefeyat") = vbYes Then
                    If ActiveDocument Is doc Then
                        If VfaEditable(doc) And stamp = VfaReadVariable(doc, VfuNamesVariable) Then VfaWriteVariable doc, VfuNamesVariable, "Vfu1:"
                    End If
                End If
            End If
        Case "not_person", "work_title", "other_person": VfuFeedback CStr(control.Tag)
    End Select
Clean:
    mVfrInsertBusy = False
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    number = Err.Number: detail = Err.Description
    mVftBusy = False
    VfrReport detail & " (" & CStr(number) & ")"
    GoTo Clean
End Sub

Public Sub Vefeyat0341UiToolApply(ByVal index As Long, ByVal unlink As Boolean)
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Clean
    mVfrInsertBusy = True
    VftApply index, unlink
Clean:
    mVfrInsertBusy = False
End Sub

Public Sub Vefeyat0341UiToolLocate(ByVal index As Long)
    On Error GoTo Failed
    VftLocate index
    Exit Sub
Failed:
    VfaWarn VfText("Kay\u0131t silinmi\u015f veya de\u011fi\u015fmi\u015f olabilir. Denetimi yenileyin.")
End Sub

Public Sub Vefeyat0341UiToolClose(Optional ByVal panelClosing As Boolean = False)
    VftClose panelClosing
End Sub

Public Sub Vefeyat0341UiToolDocumentClose(ByVal doc As Word.Document)
    Dim i As Long, item As Object
    VfvForgetDocument doc
    If Not mVftDoc Is Nothing Then
        If mVftDoc Is doc Then VftClose
    End If
    If Not mVftHistory Is Nothing Then
        For i = mVftHistory.Count To 1 Step -1
            Set item = mVftHistory(i)
            If item("Doc") Is doc Then mVftHistory.Remove i
        Next i
    End If
End Sub

Public Sub Vefeyat0341UiGetNameMode(ByVal control As Object, ByRef selectedId As Variant)
    selectedId = "name_ask"
    On Error Resume Next
    If Documents.Count > 0 Then selectedId = "name_" & VfuNameMode(ActiveDocument)
End Sub

Public Sub Vefeyat0341UiSetNameMode(ByVal control As Object, ByVal selectedId As String, ByVal index As Integer)
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    If Left$(selectedId, 5) <> "name_" Then Exit Sub
    selectedId = Mid$(selectedId, 6)
    If Not VfrEditableDocument() Or Not VfuKnownMode(selectedId) Then Exit Sub
    VfaWriteVariable ActiveDocument, VfuModeVariable, selectedId
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiDocumentEnabled(ByVal control As Object, ByRef enabled As Variant)
    enabled = VfrEditableDocument()
    If mVfbRunning Or mVfrInsertBusy Then enabled = False
End Sub

Public Sub Vefeyat0341UiGetPrefix(ByVal control As Object, ByRef selectedId As Variant)
    On Error GoTo DefaultValue
    If Application.Documents.Count > 0 Then
        selectedId = VfpGetPrefix(ActiveDocument)
        Exit Sub
    End If
DefaultValue:
    selectedId = "original"
End Sub

Public Sub Vefeyat0341UiGetScope(ByVal control As Object, ByRef selectedId As Variant)
    On Error GoTo DefaultValue
    If Application.Documents.Count > 0 Then
        selectedId = VfpGetScope(ActiveDocument)
        Exit Sub
    End If
DefaultValue:
    selectedId = "first"
End Sub

Public Sub Vefeyat0341UiGetAuditMode(ByVal control As Object, ByRef selectedId As Variant)
    On Error GoTo DefaultValue
    If Application.Documents.Count > 0 Then
        selectedId = VfpGetAuditMode(ActiveDocument)
        Exit Sub
    End If
DefaultValue:
    selectedId = "audit_all"
End Sub

Public Sub Vefeyat0341UiSetAuditMode(ByVal control As Object, ByVal selectedId As String, ByVal index As Integer)
    Dim undoObject As Object, undoStarted As Boolean, errorText As String
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    If Not VfrEditableDocument() Then Exit Sub
    undoStarted = VfaBeginUndo(undoObject)
    VfpSetAuditMode ActiveDocument, selectedId
    VfaEndUndo undoObject, undoStarted
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    errorText = Err.Description
    VfaEndUndo undoObject, undoStarted
    VfrReport errorText
End Sub

Public Sub Vefeyat0341UiSetPrefix(ByVal control As Object, ByVal selectedId As String, ByVal index As Integer)
    Dim undoObject As Object, undoStarted As Boolean, errorText As String
    Dim doc As Word.Document, newKey As String, total As Long, kept As Long, convert As Boolean, converted As Long, message As String
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    If Not VfrEditableDocument() Then Exit Sub
    Set doc = ActiveDocument
    newKey = VfpRequirePrefix(selectedId)
    ' Existing Vefeyat dates follow the new prefix after one confirmation.
    total = VfpReformatDates(doc, newKey, True, kept)
    If total > 0 Then
        If doc.TrackRevisions Then
            VfaWarn VfText("K\u0131saltma de\u011fi\u015ftirildi, fakat De\u011fi\u015fiklik \u0130zleme a\u00e7\u0131k oldu\u011fu i\u00e7in belgedeki ") & CStr(total) & VfText(" Vefey\u00e2t tarihi \u00e7evrilmedi. \u0130zlemeyi kapat\u0131p k\u0131saltmay\u0131 yeniden se\u00e7in.")
        Else
            message = VfText("K\u0131saltma """) & VfpPrefixLabel(newKey) & """ olarak ayarlanacak." & vbCrLf & vbCrLf & _
                "Belgedeki " & CStr(total) & VfText(" Vefey\u00e2t tarihi de """) & VfpPrefixLabel(newKey) & VfText(""" k\u0131saltmas\u0131na \u00e7evrilsin mi?") & vbCrLf & vbCrLf
            If kept > 0 Then message = message & CStr(kept) & VfText(" tarih elle de\u011fi\u015ftirildi\u011fi i\u00e7in korunacak.") & vbCrLf
            message = message & VfText("Elle yaz\u0131lm\u0131\u015f d\u00fcz metin tarihlere dokunulmaz; onlar i\u00e7in Teslim ve dizin \u2192 Tutarl\u0131l\u0131k raporuna bak\u0131n. Ctrl+Z ile geri al\u0131nabilir.")
            convert = (VfMessage(message, vbYesNo Or vbQuestion, VfText("Vefey\u00e2t: k\u0131saltma")) = vbYes)
            If Documents.Count = 0 Then Exit Sub
            If Not (ActiveDocument Is doc) Then Exit Sub
        End If
    End If
    undoStarted = VfaBeginUndo(undoObject)
    VfpSetPrefix doc, selectedId
    If convert Then converted = VfpReformatDates(doc, newKey, False, kept)
    VflCaptureCommand
    VfaEndUndo undoObject, undoStarted
    Vefeyat0341UiInvalidate
    If converted > 0 Then Application.StatusBar = CStr(converted) & VfText(" Vefey\u00e2t tarihi """) & VfpPrefixLabel(newKey) & VfText(""" k\u0131saltmas\u0131na \u00e7evrildi.")
    Exit Sub
Failed:
    errorText = Err.Description
    VfaEndUndo undoObject, undoStarted
    VfrReport errorText
End Sub

Public Sub Vefeyat0341UiSetScope(ByVal control As Object, ByVal selectedId As String, ByVal index As Integer)
    Dim undoObject As Object, undoStarted As Boolean, errorText As String
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    If Not VfrEditableDocument() Then Exit Sub
    undoStarted = VfaBeginUndo(undoObject)
    VfpSetScope ActiveDocument, selectedId
    VflCaptureCommand
    VflArm
    VfaEndUndo undoObject, undoStarted
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    errorText = Err.Description
    VfaEndUndo undoObject, undoStarted
    VfrReport errorText
End Sub

Public Sub Vefeyat0341UiRefresh(ByVal control As Object)
    Vefeyat0341TarihleriGuncelle
End Sub

Public Sub Vefeyat0341TarihleriGuncelle()
    VbpRun
End Sub

Public Sub Vefeyat0341UiChooseCsv(ByVal control As Object)
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    VefeyatChooseData
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Function Vefeyat0341UiFindPeople(ByVal query As String) As Collection
    Dim tick As Single, number As Long, detail As String
    Set Vefeyat0341UiFindPeople = New Collection
    If Len(query) = 0 Or Len(query) > 200 Then Exit Function
    If Not VefeyatEnsureData() Then Exit Function
    tick = VfgClock()
    On Error GoTo Failed
    Set Vefeyat0341UiFindPeople = VefeyatFind(query)
    VfgEnd "search", tick
    Exit Function
Failed:
    number = Err.Number: detail = Err.Description
    VfgEnd "search", tick, number
    Err.Raise number, "Vefeyat", detail
End Function

Public Sub Vefeyat0341UiSearch(ByVal control As Object)
    Dim picker As VefeyatKisiSec0341, errorText As String
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    mVflSuspended = True
    Set picker = New VefeyatKisiSec0341
    picker.InitSearch
    picker.Show vbModal
    Unload picker
    mVflSuspended = False
    VflResetWait
    VflArm
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    Unload picker
    On Error GoTo 0
    mVflSuspended = False
    VflArm
    VfrReport errorText
End Sub

Public Function Vefeyat0341UiEarlierChoice(ByVal message As String) As Long
    Dim picker As VefeyatKisiSec0341, detail As String
    On Error GoTo Failed
    Set picker = New VefeyatKisiSec0341
    picker.InitEarlier message
    If Not mVfwPanel Is Nothing Then
        picker.SetSideWarning
        VfwPosition picker
    End If
    picker.Show vbModal
    Vefeyat0341UiEarlierChoice = picker.ChosenIndex
    Unload picker
    Exit Function
Failed:
    detail = Err.Description
    On Error Resume Next
    Unload picker
    On Error GoTo 0
    VfrReport detail
End Function

Public Sub Vefeyat0341UiDia(ByVal control As Object)
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    VfdOpenSelected
End Sub

Public Sub Vefeyat0341UiCanOpenDia(ByVal control As Object, ByRef enabled As Variant)
    Dim aliasText As String
    enabled = False
    On Error GoTo Finished
    If mVfbRunning Or Documents.Count = 0 Then Exit Sub
    If Selection.Range.End - Selection.Range.Start > VfaMaxName Then Exit Sub
    enabled = VfaSelectionAlias(Selection.Range.Text, aliasText)
Finished:
End Sub

Public Sub Vefeyat0341UiProbe()
    ' Only acknowledge this code host in memory; never edit a document or timer.
    mVflProbePath = ThisDocument.FullName
End Sub

Public Sub Vefeyat0341UiInfo(ByVal control As Object)
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    If Not VefeyatEnsureData() Then Exit Sub
    VfMessage "Vefeyat " & VfrVersion & " - Belge taramasi ve isim bicimleri" & vbCrLf & _
           CStr(VefeyatRecordCount()) & VfText(" yerel kay\u0131t") & vbCrLf & vbCrLf & _
           "CSV: " & VefeyatDataPath() & vbCrLf & vbCrLf & _
           "Eklenti: " & ThisDocument.FullName & vbCrLf & vbCrLf & VfmDiagnostic() & vbCrLf & VflDiagnostic() & vbCrLf & VfxDiagnostic(), vbInformation, VfText("Vefey\u00e2t")
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiFlatDates(ByVal control As Object, ByVal pressed As Boolean)
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    On Error GoTo Failed
    VfvSetFlat pressed
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiGetFlatDates(ByVal control As Object, ByRef pressed As Variant)
    pressed = False
    On Error GoTo Finished
    If Documents.Count = 0 Then Exit Sub
    pressed = (ActiveWindow.View.FieldShading = wdFieldShadingNever)
Finished:
End Sub

Public Sub Vefeyat0341UiSpelling(ByVal control As Object)
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    mVflSuspended = True
    VfnCheckSelected
    mVflSuspended = False
    If Application.Documents.Count > 0 Then VfmObserve ActiveDocument, Selection.Range
    VflArm
    Exit Sub
Failed:
    mVflSuspended = False
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiAudit(ByVal control As Object)
    Dim detail As String
    If Not mVfwPanel Is Nothing Then
        If Not mVfwPanel.Visible Then mVfwPanel.Show vbModeless
        VfwShowCurrent
        Exit Sub
    End If
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    If mVfwPaused Then VfwClose
    Set mVfmWorkDoc = Nothing
    Set mVfmWorkState = Nothing
    mVfbRunning = True
    mVflSuspended = True
    mVfbLastStage = vbNullString
    VfbResetProgress
    Set mVfbProgress = New VefeyatKisiSec0341
    mVfbProgress.InitProgress
    mVfbProgress.Show vbModal
    Unload mVfbProgress
    Set mVfbProgress = Nothing
    mVfbRunning = False
    mVflSuspended = False
    If mVfwPrepared Then
        mVfbRunning = True
        mVflSuspended = True
        VfwOpen
    Else
        VfwClose
    End If
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    Unload mVfbProgress
    Set mVfbProgress = Nothing
    On Error GoTo 0
    mVfbRunning = False
    mVflSuspended = False
    VfwClose
    VfrReport detail
End Sub

Public Sub Vefeyat0341UiRunAudit()
    Dim detail As String, number As Long, tick As Single
    If Not mVfbRunning Or mVfbExecuting Or mVfbExecuted Then Exit Sub
    If mVfbProgress Is Nothing Then Exit Sub
    mVfbExecuting = True: mVfbExecuted = True
    tick = VfgClock()
    On Error GoTo Failed
    mVfbLastStage = VfText("Toplu i\u015flem haz\u0131rl\u0131\u011f\u0131")
    If VbpDispatchPrepare() Then
        ' The same one-shot progress gate prepares the immutable preview.
    ElseIf mVfjPreparing Then
        VfjPrepare
    ElseIf mVfmWorkDoc Is Nothing Then
        VfwPrepare
    ElseIf mVfmWorkState Is Nothing Then
        VfmWarm mVfmWorkDoc
    Else
        VfmFlushPending mVfmWorkState
    End If
    mVfbProgress.Hide
    VfgEnd "prepare", tick
    mVfbExecuting = False
    Exit Sub
Failed:
    number = Err.Number
    detail = Err.Description
    VfgEnd "prepare", tick, number
    detail = VfText("Tarama tamamlanamad\u0131. Hata ") & CStr(number) & ": " & detail
    If Len(mVfbLastStage) > 0 Then detail = detail & vbCrLf & VfText("A\u015fama: ") & mVfbLastStage
    mVfbProgress.Hide
    If number = VfbCancelError Then
        Application.StatusBar = VfText("Tarama durduruldu; belge de\u011fi\u015ftirilmedi.")
    Else
        VfrReport detail
    End If
    mVfbExecuting = False
End Sub

Public Sub Vefeyat0341UiCanAudit(ByVal control As Object, ByRef enabled As Variant)
    enabled = Not (mVfwPanel Is Nothing)
    If Not enabled Then enabled = VfrEditableDocument() And Not mVfbRunning And Not mVfrInsertBusy
End Sub

Public Sub Vefeyat0341UiReviewApply(ByVal choice As Long)
    VfwApply choice
End Sub

Public Sub Vefeyat0341UiReviewMove(ByVal direction As Long)
    VfwMove direction
End Sub

Public Sub Vefeyat0341UiReviewLocate()
    If Not mVfwBusy Then VfwShowCurrent
End Sub

Public Sub Vefeyat0341UiReviewChoice()
    VfwChoiceChanged
End Sub

Public Sub Vefeyat0341UiReviewClose(Optional ByVal panelClosing As Boolean = False)
    VfwClose panelClosing
End Sub

Public Sub Vefeyat0341UiReviewPause()
    VfwPause
End Sub

Public Sub Vefeyat0341UiResume(ByVal control As Object)
    VfwResume
End Sub

Public Sub Vefeyat0341UiCanResume(ByVal control As Object, ByRef enabled As Variant)
    enabled = False
    On Error GoTo Finished
    If mVfbRunning Or Not mVfwPaused Then Exit Sub
    If VfwOpenDocument() Then enabled = (ActiveDocument Is mVfwDoc)
Finished:
End Sub

Public Sub Vefeyat0341UiReviewIgnore()
    VfwIgnoreCurrent
End Sub

Public Sub Vefeyat0341UiExclusions(ByVal control As Object)
    If mVfbRunning Or Documents.Count = 0 Then Exit Sub
    On Error GoTo Failed
    VfiShow ActiveDocument
    Exit Sub
Failed:
    VfaWarn VfText("Belge istisnalar\u0131 okunamad\u0131. \u0130stisnalar\u0131 temizle ile listeyi s\u0131f\u0131rlayabilirsiniz.")
End Sub

Public Sub Vefeyat0341UiClearExclusions(ByVal control As Object)
    Dim undoObject As Object, started As Boolean, number As Long, doc As Word.Document, wasSuspended As Boolean
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    wasSuspended = mVflSuspended
    On Error GoTo Failed
    If Not VfrEditableDocument() Then Exit Sub
    Set doc = ActiveDocument
    mVfbRunning = True: mVflSuspended = True
    If VfMessage(VfText("Bu belgenin toplu denetim istisnalar\u0131 temizlensin mi? Metin ve tarihler de\u011fi\u015fmez. Sonraki denetimde bu ifadeler yeniden sorulabilir."), _
        vbYesNo Or vbQuestion Or vbDefaultButton2, VfText("\u0130stisnalar\u0131 temizle")) <> vbYes Then GoTo Clean
    If Not (ActiveDocument Is doc) Then GoTo Clean
    If Not VfaEditable(doc) Then GoTo Clean
    started = VfaBeginUndo(undoObject)
    VfaWriteVariable doc, VfiVariable, "Vfi1"
    VfaEndUndo undoObject, started
    VfaWarn VfText("Belge istisnalar\u0131 temizlendi. Yeni denetim ba\u015flatabilirsiniz.")
    GoTo Clean
Failed:
    number = Err.Number
    VfaEndUndo undoObject, started
    VfaWarn VfText("\u0130stisnalar temizlenemedi. Hata: ") & CStr(number)
Clean:
    mVfbRunning = False: mVflSuspended = wasSuspended
    Vefeyat0341UiInvalidate
End Sub

Public Sub Vefeyat0341UiEditExclusions(ByVal control As Object)
    Dim wasSuspended As Boolean
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    wasSuspended = mVflSuspended
    On Error GoTo Failed
    If Not VfrEditableDocument() Then Exit Sub
    mVfbRunning = True: mVflSuspended = True
    VfiEdit ActiveDocument
    GoTo Clean
Failed:
    VfaWarn VfText("\u0130stisna listesi a\u00e7\u0131lamad\u0131.")
Clean:
    mVfbRunning = False: mVflSuspended = wasSuspended
    Vefeyat0341UiInvalidate
End Sub

Public Sub Vefeyat0341UiCanInspectDocument(ByVal control As Object, ByRef enabled As Variant)
    enabled = (Documents.Count > 0 And Not mVfbRunning)
End Sub

Public Sub Vefeyat0341UiSourceInfo(ByVal control As Object)
    If mVfbRunning Or mVfrInsertBusy Then Exit Sub
    On Error GoTo Failed
    VfeShowSelected
    Exit Sub
Failed:
    VfaWarn VfText("Kaynak bilgisi okunamad\u0131; belge de\u011fi\u015ftirilmedi.")
End Sub

Public Sub Vefeyat0341UiCanReadSource(ByVal control As Object, ByRef enabled As Variant)
    Dim selected As Word.Range
    enabled = False
    On Error GoTo Finished
    If mVfbRunning Or Documents.Count = 0 Then Exit Sub
    Set selected = Selection.Range
    If selected.End - selected.Start > 512 Then Exit Sub
    enabled = (selected.StoryType = wdMainTextStory Or selected.StoryType = wdFootnotesStory)
Finished:
End Sub

Public Sub Vefeyat0341UiDiagnostics(ByVal control As Object)
    VfMessage "Vefeyat " & VfrVersion & vbCrLf & "Word " & Application.Version & vbCrLf & vbCrLf & VfgReport(), vbInformation, VfText("Tan\u0131 ve s\u00fcreler")
End Sub

Public Sub Vefeyat0341UiReviewDocumentClose(ByVal doc As Word.Document)
    If Not mVfwDoc Is Nothing Then
        If mVfwDoc Is doc Then VfwClose
    End If
End Sub

Public Sub Vefeyat0341UiWarmMemory(ByVal doc As Word.Document, Optional ByVal repairOnly As Boolean = False)
    Dim detail As String, state As Object, wasSuspended As Boolean
    If mVfbRunning Or mVfmBusy Then Exit Sub
    On Error GoTo Failed
    wasSuspended = mVflSuspended
    If mVfmStates Is Nothing Then Set mVfmStates = New Collection
    Set state = VfmState(doc)
    If state Is Nothing Then
        Set state = VfmNewState(doc)
        mVfmStates.Add state
    End If
    mVfmBusy = True
    mVfbRunning = True
    mVflSuspended = True
    Set mVfmWorkDoc = doc
    Set mVfmWorkState = Nothing
    If repairOnly Then Set mVfmWorkState = VfmState(doc)
    If doc.Content.End <= 1 And doc.Footnotes.Count = 0 Then
        VfmWarm doc
    Else
        mVfbLastStage = vbNullString
        VfbResetProgress
    Set mVfbProgress = New VefeyatKisiSec0341
        mVfbProgress.InitProgress
        mVfbProgress.Show vbModal
        Unload mVfbProgress
    End If
Clean:
    Set mVfbProgress = Nothing
    Set mVfmWorkDoc = Nothing
    Set mVfmWorkState = Nothing
    mVfmBusy = False
    mVfbRunning = False
    mVflSuspended = wasSuspended
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    Unload mVfbProgress
    On Error GoTo 0
    Application.StatusBar = VfText("Belge belle\u011fi haz\u0131rlanamad\u0131: ") & detail
    GoTo Clean
End Sub

Public Sub Vefeyat0341UiRefreshMemory(ByVal control As Object)
    Dim state As Object
    If Application.Documents.Count = 0 Then Exit Sub
    If mVfbRunning Or mVfmBusy Then Exit Sub
    Set state = VfmState(ActiveDocument)
    If Not state Is Nothing Then state("Ready") = False
    Vefeyat0341UiWarmMemory ActiveDocument
End Sub

Public Sub Vefeyat0341UiMemoryActivate(ByVal doc As Word.Document)
    On Error GoTo Quiet
    ' Opening/activating a document must never read its body, catalog or
    ' metadata. Explicit first-insert/audit/refresh commands prepare memory.
    VfmPrune
Quiet:
End Sub

Public Sub Vefeyat0341UiMemorySelection(ByVal Sel As Word.Selection)
    On Error GoTo Quiet
    If mVfbRunning Or mVfmBusy Or mVflSuspended Then Exit Sub
    If mVfmNeedsPrune Then
        VfmPrune
        mVfmNeedsPrune = False
    End If
    VfmObserve Sel.Document, Sel.Range, False
Quiet:
End Sub

Public Sub Vefeyat0341UiMemoryClose(ByVal doc As Word.Document)
    ' BeforeClose can be cancelled. Prune again after the close has completed.
    mVfmNeedsPrune = True
    On Error Resume Next
    VfmPrune
    On Error GoTo 0
End Sub

Public Sub Vefeyat0341UiMemoryPrune()
    On Error Resume Next
    VfmPrune
    On Error GoTo 0
End Sub

Public Sub Vefeyat0341UiHelp(ByVal control As Object)
    VfMessage VfText("Ad\u0131 se\xe7ip \xd6l\xfcm tarihi ekle'ye bas\u0131n. Tek ve g\xfc\xe7l\xfc e\u015fle\u015fmede do\u011frudan eklenir; yak\u0131n yaz\u0131mda veya birden fazla adayda \u015fahs\u0131 do\u011frulay\u0131n.") & _
        vbCrLf & vbCrLf & VfText("Denetim: ilk/her ge\xe7i\u015f ve tarama kapsam\u0131 bu gruptad\u0131r. T\xfcm belgeyi denetle, ana metin ve dipnotlarda s\u0131rayla ilerler. Her Onayla ve sonraki yaln\u0131z g\xf6sterilen ge\xe7i\u015fe uygulan\u0131r; Atla/Sonraki ge\xe7i\u015fi korur.") & _
        vbCrLf & vbCrLf & VfText("\u0130sim yaz\u0131m\u0131 en soldaki b\xfcy\xfck d\xfc\u011fmedir. \xd6nizlemeden sonra yaln\u0131z se\xe7ili ad de\u011fi\u015fir. \u015eah\u0131s ve kaynak grubunda arama, D\u0130A ve kay\u0131tl\u0131 kaynak bulunur.") & _
        vbCrLf & vbCrLf & VfText("Belge sa\u011fl\u0131\u011f\u0131 ve ge\xe7mi\u015f Belge grubunda; istisnalar ve tarih g\xfcncelleme Di\u011fer ara\xe7lar men\xfcs\xfcndedir. Proje tercihleri Tercihler > Ayarlar alt\u0131ndad\u0131r.") & _
        vbCrLf & vbCrLf & VfText("Normal ekleme bitince imle\xe7 tarihin d\u0131\u015f\u0131na al\u0131n\u0131r; tarih ba\u011flant\u0131s\u0131 korunur. Alan g\xf6lgelerini gizle, Ayarlar > Geli\u015fmi\u015f alt\u0131ndad\u0131r ve Word penceresindeki b\xfct\xfcn alanlar\u0131 etkiler. A\xe7\u0131l\u0131\u015fta genel g\xf6lge ayar\u0131 de\u011fi\u015ftirilmez.") & _
        vbCrLf & vbCrLf & VfText("Yerel CSV kaynak dosyan\u0131z de\u011fi\u015ftirilmez. Yazarken otomatik ekleme kapal\u0131d\u0131r. Belge i\u015flemleri i\xe7in Word'\xfcn Geri Al komutunu kullan\u0131n."), _
        vbInformation, VfText("Vefey\u00e2t kullan\u0131m\u0131")
End Sub

Private Function VfePickPerson(ByVal people As Collection, ByVal query As String, _
                              Optional ByVal spelling As Boolean = False, Optional ByVal forceDialog As Boolean = False) As Long
    Dim picker As VefeyatKisiSec0341
    Dim errorNumber As Long, errorText As String, proof As Object, ordered As Collection, mapping As Variant, choice As Long
    On Error GoTo Failed
    Set proof = VdiBuild(people, query, Selection.Range)
    If CLng(proof("Remembered")) > 0 And Not forceDialog Then
        If VdiCurrent(proof) Then VfePickPerson = CLng(proof("Remembered"))
        Exit Function
    End If
    Set ordered = proof("Rows")
    Set picker = New VefeyatKisiSec0341
    picker.InitCandidates ordered, query
    picker.SetIdentityEvidence proof
    If spelling Then picker.SetSpellingMode
    picker.Show vbModal
    choice = picker.ChosenIndex
    If choice > 0 And choice <= people.Count Then
        If VdiCurrent(proof) Then
            mapping = proof("Indices"): VfePickPerson = CLng(mapping(choice))
        Else
            VfaWarn VfText("Ba\u011flam, belge, kaynak veya onay kayd\u0131 de\u011fi\u015fti; i\u015flem yap\u0131lmad\u0131.")
        End If
    End If
    Unload picker
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    Unload picker
    On Error GoTo 0
    Err.Raise errorNumber, "Vefeyat", errorText
End Function

Private Function VfrEditableDocument() As Boolean
    On Error GoTo Unavailable
    If Application.Documents.Count = 0 Then Exit Function
    If ActiveDocument.ReadOnly Then Exit Function
    If ActiveDocument.ProtectionType <> wdNoProtection Then Exit Function
    VfrEditableDocument = True
Unavailable:
End Function

Private Sub VfrReport(ByVal message As String)
    Application.StatusBar = VfText("Vefey\u00e2t i\u015flemi tamamlanamad\u0131.")
    VfMessage message, vbExclamation, VfText("Vefey\u00e2t")
End Sub


Public Sub Vefeyat0341UiInventory(ByVal control As Object)
    Dim detail As String
    If mVfbRunning Or mVfrInsertBusy Or Documents.Count = 0 Then Exit Sub
    On Error GoTo Failed
    If Not mVfjPanel Is Nothing Then
        If ActiveDocument Is mVfjDoc Then
            If Not mVfjPanel.Visible Then mVfjPanel.Show vbModeless
            Exit Sub
        End If
    End If
    VfjClose
    Set mVfjDoc = ActiveDocument
    mVfjBusy = True: mVfjPreparing = True
    mVfbRunning = True: mVflSuspended = True
    VfbResetProgress
    Set mVfbProgress = New VefeyatKisiSec0341
    mVfbProgress.InitProgress
    mVfbProgress.Show vbModal
    Unload mVfbProgress
    Set mVfbProgress = Nothing
    mVfjPreparing = False: mVfjBusy = False
    mVfbRunning = False: mVflSuspended = False
    If Not mVfjReady Then VfjClose: Exit Sub
    Set mVfjPanel = New VefeyatKisiSec0341
    VfjDisplay
    mVfjPanel.Show vbModeless
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    Unload mVfbProgress
    Set mVfbProgress = Nothing
    VfjClose
    mVfbRunning = False: mVflSuspended = False
    VfrReport detail
End Sub

Public Sub Vefeyat0341UiInventorySelection(ByVal index As Long)
    On Error GoTo Failed
    VfjSelection index
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiInventoryAction(ByVal index As Long, ByVal changeName As Boolean)
    VfjApply index, changeName
End Sub

Public Sub Vefeyat0341UiInventoryIgnore(ByVal rowIndex As Long)
    VfjIgnore rowIndex
End Sub

Public Sub Vefeyat0341UiInventoryLocate(ByVal rowIndex As Long)
    Dim index As Long
    If mVfjBusy Or mVfbRunning Then Exit Sub
    On Error GoTo Failed
    index = VfjIndex(rowIndex)
    If index > 0 Then VfjLocateIndex index
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiInventoryPage(ByVal direction As Long)
    If mVfjBusy Or mVfbRunning Or mVfjPanel Is Nothing Then Exit Sub
    On Error GoTo Failed
    If direction > 0 Then
        If mVfjPage + VfjPageSize > mVfjCount Then Exit Sub
        mVfjPage = mVfjPage + VfjPageSize
    Else
        If mVfjPage <= 1 Then Exit Sub
        mVfjPage = mVfjPage - VfjPageSize
    End If
    VfjDisplay
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiInventoryRefresh()
    If mVfjBusy Or mVfbRunning Or mVfjDoc Is Nothing Then Exit Sub
    On Error GoTo Failed
    If Not VftOpenDoc(mVfjDoc) Then VfjClose: Exit Sub
    mVfjDoc.Activate
    VfjClose
    Vefeyat0341UiInventory Nothing
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiInventoryClose(Optional ByVal panelClosing As Boolean = False, Optional ByVal sender As Object = Nothing)
    If Not sender Is Nothing Then
        If mVfjPanel Is Nothing Then Exit Sub
        If Not (sender Is mVfjPanel) Then Exit Sub
    End If
    VfjClose panelClosing
End Sub

Public Sub Vefeyat0341UiInventoryDocumentClose(ByVal doc As Word.Document)
    If mVfjDoc Is Nothing Then Exit Sub
    If doc Is mVfjDoc Then VfjClose
End Sub

Public Function Vefeyat0341UiFindInventoryPeople(ByVal query As String) As Collection
    Set Vefeyat0341UiFindInventoryPeople = New Collection
    If Len(query) = 0 Or Len(query) > VfaMaxName Then Exit Function
    If Not VefeyatEnsureData() Then Exit Function
    Set Vefeyat0341UiFindInventoryPeople = VfaConfirmationCandidates(VefeyatFind(query, False), query)
End Function

Public Function Vefeyat0341UiContextExplain(ByVal proof As Object, ByRef row As Variant) As String
    Vefeyat0341UiContextExplain = VdiDescribe(proof, row)
End Function

Public Sub Vefeyat0341UiIdentityManage(ByVal control As Object)
    On Error GoTo Failed
    VdiManage
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiIdentityForget(ByVal control As Object)
    On Error GoTo Failed
    VdiForgetSelected
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiBulkPreview(ByVal control As Object)
    VbpRun
End Sub

Public Sub Vefeyat0341UiCoverage(ByVal control As Object)
    If Documents.Count = 0 Then Exit Sub
    VfMessage VbsReport(ActiveDocument), vbInformation, VfText("Belge denetim kapsam\u0131")
End Sub

Public Sub Vefeyat0341UiRegistryGet(ByVal control As Object, ByRef pressed As Variant)
    pressed = False
    If Documents.Count > 0 Then pressed = VrgEnabled(ActiveDocument)
End Sub

Public Sub Vefeyat0341UiRegistryToggle(ByVal control As Object, ByVal pressed As Boolean)
    On Error GoTo Failed
    If Documents.Count = 0 Or mVfbRunning Or mVfrInsertBusy Then Exit Sub
    If Not VfaEditable(ActiveDocument) Then Exit Sub
    If pressed And Not VrgApi(ActiveDocument) Then VfaWarn VfText("Bu dosya/API dinamik kay\u0131tlar\u0131 korumak i\u00e7in desteklenmiyor; mevcut metin i\u015flemleri korunur."): Exit Sub
    If pressed Then
        VfaWriteVariable ActiveDocument, VrgMode, "1"
    Else
        VfaWriteVariable ActiveDocument, VrgMode, "0"
    End If
    Vefeyat0341UiInvalidate
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiRegistryRegister(ByVal control As Object)
    On Error GoTo Failed
    VrgRegisterSelected
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341UiRegistryManage(ByVal control As Object)
    VrgManage False
End Sub

Public Sub Vefeyat0341UiRegistryFind(ByVal control As Object)
    VrgManage True
End Sub

Public Sub Vefeyat0341UiRegistryForget(ByVal control As Object)
    On Error GoTo Failed
    VrgForgetSelected
    Exit Sub
Failed:
    VfrReport Err.Description
End Sub

Public Sub Vefeyat0341Kur()
    Dim fso As Object, picker As Object, sourceXml As String, stage As String
    Dim root As String, rawTemplate As String, sourceDocument As Document, originalDocument As Document
    Dim sourceZip As String, unpacked As String, outputZip As String, checked As String
    Dim startup As String, installedPath As String, readyPath As String
    Dim completed As Boolean
    Dim previousVba As Variant, errorText As String, errorNumber As Long
    Dim incomingManifest As Object, outgoingManifest As Object
    On Error GoTo Failed
    If mVfsBusy Then Exit Sub
    mVfsBusy = True
    If Documents.Count > 0 Then Set originalDocument = ActiveDocument
    Set sourceDocument = VfsRequireSource()
    VfsCheckDispatch sourceDocument
    If VfpRunPolicyChecks() < 1 Then VfsFail VfText("Tarih bi\u00e7imlendirme denetimi ba\u015far\u0131s\u0131z oldu.")
    If VfaRunLookupChecks() < 1 Then VfsFail VfText("Ki\u015fi arama denetimi ba\u015far\u0131s\u0131z oldu.")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set picker = Application.FileDialog(3)
    With picker
        .Title = VfText("Vefey\u00e2t kurulumu: paketteki VefeyatRibbon.xml dosyas\u0131n\u0131 se\u00e7in")
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Vefeyat Ribbon XML", "*.xml"
        If .Show <> -1 Then GoTo Finished
        sourceXml = CStr(.SelectedItems(1))
    End With
    VfsRequireLocal fso, sourceXml
    If StrComp(fso.GetFileName(sourceXml), VfsRibbonName, vbTextCompare) <> 0 Then
        VfsFail VfText("L\u00fctfen paketteki VefeyatRibbon.xml dosyas\u0131n\u0131 se\u00e7in.")
    End If
    VfsValidateRibbon sourceXml
    root = fso.GetParentFolderName(VefeyatDataPath())
    VfsEnsureFolder fso, root & "\Setup"
    stage = root & "\Setup\" & fso.GetTempName
    fso.CreateFolder stage
    fso.CopyFile sourceXml, stage & "\" & VfsRibbonName, False
    VfsValidateRibbon stage & "\" & VfsRibbonName

    startup = Options.DefaultFilePath(wdStartupPath)
    If Len(startup) = 0 Then VfsFail VfText("Word ba\u015flang\u0131\u00e7 klas\u00f6r\u00fc tan\u0131ml\u0131 de\u011fil.")
    VfsRequireLocal fso, startup
    If StrComp(fso.GetAbsolutePathName(fso.GetParentFolderName(sourceDocument.FullName)), _
               fso.GetAbsolutePathName(fso.BuildPath(startup, ".")), vbTextCompare) = 0 Then
        VfsFail VfText("Haz\u0131rl\u0131k \u015fablonunu Word ba\u015flang\u0131\u00e7 klas\u00f6r\u00fcne kaydetmeyin. \u0130ndirilenler gibi ayr\u0131 bir klas\u00f6r kullan\u0131n.")
    End If
    installedPath = fso.BuildPath(startup, VfsTemplateName)

    ' The user creates this separate template and imports the three components.
    ' Save only that open preparation template. Normal is never copied or saved.
    VfsCheckComponents
    ' VBA code edits do not always dirty the document's Saved flag.
    sourceDocument.Saved = False
    sourceDocument.Save
    rawTemplate = stage & "\VefeyatRaw.dotm"
    VfsCopyPreparedTemplate fso, sourceDocument, rawTemplate

    sourceZip = stage & "\source.zip"
    fso.CopyFile rawTemplate, sourceZip, False
    unpacked = stage & "\unpacked"
    Set incomingManifest = VfsExtractZip(fso, sourceZip, unpacked)
    VfsRequirePart fso, unpacked, "[Content_Types].xml"
    VfsRequirePart fso, unpacked, "_rels\.rels"
    VfsRequirePart fso, unpacked, "word\vbaProject.bin"
    previousVba = VfsReadBytes(unpacked & "\word\vbaProject.bin")
    VfsInjectRibbon fso, unpacked, stage & "\" & VfsRibbonName
    Set outgoingManifest = VfsFolderManifest(fso, unpacked)
    outputZip = stage & "\ready.zip"
    VfsCreateZip fso, unpacked, outputZip, outgoingManifest

    ' Re-extract the finished archive and compare the actual files, not only names.
    checked = stage & "\checked"
    Set incomingManifest = VfsExtractZip(fso, outputZip, checked)
    VfsVerifyPackage fso, checked, outgoingManifest, previousVba
    readyPath = stage & "\" & VfsTemplateName
    fso.CopyFile outputZip, readyPath, False

    ' An existing add-in stays loaded and its Startup file stays untouched.
    ' Prepare a verified replacement outside Startup for a closed-Word copy.
    If fso.FileExists(installedPath) Or VfsOwnedLoaded(fso, installedPath) Then
        VfsStageUpdate fso, root, readyPath, installedPath, startup
        completed = True
        GoTo Finished
    End If

    ' A first installation has no existing file or loaded owned add-in.
    ' MoveFile refuses an unexpected destination created after the check.
    VfsEnsureFolder fso, startup
    fso.MoveFile readyPath, installedPath
    completed = True
    ' Loading now would leave two projects with the same callback module name:
    ' the open preparation template and its installed copy. Load on next start.
    VfMessage "Vefeyat eklenti dosyasi hazirlandi." & vbCrLf & vbCrLf & _
        "Calismanizi kaydedin ve Word'un BUTUN pencerelerini kapatin. Word'u yeniden acinca Vefeyat sekmesi yuklenir." & _
        vbCrLf & vbCrLf & "Hazirlik sablonunu gunluk kullanimda acmayin. CSV verileriniz bilgisayarinizda kalir.", _
        vbInformation, "Vefeyat"
    GoTo Finished
Failed:
    errorText = Err.Description
    errorNumber = Err.Number
    If errorNumber = 70 Or errorNumber = 75 Then
        errorText = errorText & vbCrLf & VfsRestartMessage()
    End If
    VfMessage VfText("Kurulum tamamlanamad\u0131: ") & errorText & vbCrLf & vbCrLf & _
        VfText("CSV verileriniz ve haz\u0131rl\u0131k \u015fablonundaki kaynak kod korunur."), vbExclamation, VfText("Vefey\u00e2t")
Finished:
    On Error Resume Next
    If Not originalDocument Is Nothing Then originalDocument.Activate
    ' Keep a failed stage for diagnosis and recovery; delete only a successful one.
    If completed And Len(stage) > 0 Then fso.DeleteFolder stage, False
    On Error GoTo 0
    mVfsBusy = False
End Sub

Public Sub Vefeyat0341EklentiyiKaldir()
    Dim fso As Object, root As String, startup As String, installedPath As String
    Dim backupPath As String, wasInstalled As Boolean, removed As Boolean, sourceDocument As Document
    Dim errorText As String
    On Error GoTo Failed
    If mVfsBusy Then Exit Sub
    mVfsBusy = True
    Set sourceDocument = VfsRequireSource()
    Set fso = CreateObject("Scripting.FileSystemObject")
    startup = Options.DefaultFilePath(wdStartupPath)
    If Len(startup) = 0 Then VfsFail VfText("Word ba\u015flang\u0131\u00e7 klas\u00f6r\u00fc tan\u0131ml\u0131 de\u011fil.")
    VfsRequireLocal fso, startup
    installedPath = fso.BuildPath(startup, VfsTemplateName)
    If VfMessage(VfText("Vefey\u00e2t Word eklentisi kald\u0131r\u0131ls\u0131n m\u0131?") & vbCrLf & _
        installedPath & vbCrLf & vbCrLf & _
        VfText("CSV verileri, ayarlar ve haz\u0131rl\u0131k \u015fablonundaki kaynak kod korunacak."), _
        vbYesNo + vbQuestion + vbDefaultButton2, VfText("Vefey\u00e2t")) <> vbYes Then GoTo Finished
    wasInstalled = VfsUnloadOwnedAddin(fso, installedPath)
    If fso.FileExists(installedPath) Then
        If Not VfsCanReadExclusive(installedPath) Then VfsFail VfsRestartMessage()
        root = fso.GetParentFolderName(VefeyatDataPath())
        VfsEnsureFolder fso, root & "\Backups"
        backupPath = root & "\Backups\" & fso.GetBaseName(fso.GetTempName) & ".dotm"
        fso.MoveFile installedPath, backupPath
    End If
    removed = True
    VfMessage VfText("Vefey\u00e2t eklentisi kald\u0131r\u0131ld\u0131. CSV verileriniz, ayarlar\u0131n\u0131z ve kaynak kod korundu."), _
        vbInformation, VfText("Vefey\u00e2t")
    GoTo Finished
Failed:
    errorText = Err.Description
    On Error Resume Next
    If wasInstalled And Not removed Then errorText = errorText & vbCrLf & VfsRestartMessage()
    On Error GoTo 0
    VfMessage VfText("Eklenti kald\u0131r\u0131lamad\u0131: ") & errorText, vbExclamation, VfText("Vefey\u00e2t")
Finished:
    mVfsBusy = False
End Sub

Private Function VfsRequireSource() As Document
    Dim containerPath As String, candidate As Document
    ' Bind setup to the code's own host; MacroContainer is late-bound and can
    ' resolve differently when a global add-in and a preparation file coexist.
    containerPath = ThisDocument.FullName
    If StrComp(containerPath, NormalTemplate.FullName, vbTextCompare) = 0 Then
        VfsFail VfText("Kurulum Normal projesinden \u00e7al\u0131\u015ft\u0131r\u0131lamaz. Paketteki yeni haz\u0131rl\u0131k \u015fablonu ad\u0131mlar\u0131n\u0131 uygulay\u0131n.")
    End If
    For Each candidate In Documents
        If StrComp(candidate.FullName, containerPath, vbTextCompare) = 0 Then
            If StrComp(candidate.Name, VfsSourceTemplateName, vbTextCompare) <> 0 Or candidate.SaveFormat <> 15 Then
                VfsFail VfText("Bo\u015f bir belgeyi VefeyatHazirlik.dotm ad\u0131yla Word makro i\u00e7erebilen \u015fablonu olarak kaydedin ve bile\u015fenleri onun projesine al\u0131n.")
            End If
            If candidate.ReadOnly Or candidate.ProtectionType <> wdNoProtection Then
                VfsFail VfText("Haz\u0131rl\u0131k \u015fablonu d\u00fczenlenebilir olmal\u0131.")
            End If
            If candidate.Content.Text <> vbCr Or candidate.InlineShapes.Count > 0 Or _
               candidate.Shapes.Count > 0 Or candidate.Footnotes.Count > 0 Or candidate.Endnotes.Count > 0 Then
                VfsFail VfText("Haz\u0131rl\u0131k \u015fablonu bo\u015f bir belgeden olu\u015fturulmal\u0131; ki\u015fisel belge i\u00e7eri\u011fi i\u00e7ermemeli.")
            End If
            Set VfsRequireSource = candidate
            Exit Function
        End If
    Next candidate
    VfsFail VfText("VefeyatHazirlik.dotm dosyas\u0131n\u0131 Word'de Dosya > A\u00e7 ile d\u00fczenlemek i\u00e7in a\u00e7\u0131n. Kurulumu onun VefeyatEklenti0341 mod\u00fcl\u00fcnden \u00e7al\u0131\u015ft\u0131r\u0131n.")
End Function

Private Sub VfsCheckDispatch(ByVal sourceDocument As Document)
    Dim detail As String, callback As String
    On Error GoTo Failed
    ' Word limits macro lookup to the current document/template context.
    sourceDocument.Activate
    mVflProbePath = vbNullString
    callback = VflMacroPath("Vefeyat0341UiProbe")
    ' Probe a Public Sub, like the timer entry, without depending on a return value.
    Application.Run callback
    If StrComp(mVflProbePath, sourceDocument.FullName, vbTextCompare) <> 0 Then
        VfsFail "The callback did not acknowledge this preparation template."
    End If
    mVflProbePath = vbNullString
    Exit Sub
Failed:
    detail = Err.Description
    mVflProbePath = vbNullString
    VfsFail VfText("Haz\u0131rl\u0131k dosyas\u0131n\u0131n makro \u00e7a\u011fr\u0131s\u0131 do\u011frulanamad\u0131. Paketteki ana mod\u00fcl\u00fc haz\u0131rl\u0131k projesine bir kez i\u00e7e al\u0131n. \u00c7a\u011fr\u0131: ") & _
        callback & vbCrLf & VfText("Ayr\u0131nt\u0131: ") & detail
End Sub

Private Sub VfsCheckComponents()
    Dim picker As VefeyatKisiSec0341, events As VefeyatOlaylar0341
    Dim detail As String, confirmationVersion As Long
    On Error GoTo Failed
    Set events = New VefeyatOlaylar0341
    Set picker = New VefeyatKisiSec0341
    If picker.Controls.Count < 8 Then VfsFail "VefeyatKisiSec0341 controls are incomplete."
    On Error Resume Next
    confirmationVersion = CLng(CallByName(picker, "ConfirmationVersion", VbMethod))
    Err.Clear
    On Error GoTo Failed
    If confirmationVersion <> 15 Then VfsFail VfText("Paketteki VefeyatKisiSec0341.frm dosyas\u0131n\u0131 File > Import File ile haz\u0131rl\u0131k projesine al\u0131n. Form ve ana mod\u00fcl ayn\u0131 s\u00fcr\u00fcmden olmal\u0131.")
    Unload picker
    Set picker = Nothing
    Set events = Nothing
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    Unload picker
    Set picker = Nothing
    Set events = Nothing
    On Error GoTo 0
    VfsFail VfText("Haz\u0131rl\u0131k \u015fablonundaki ki\u015fi se\u00e7im penceresi y\u00fcklenemedi: ") & detail
End Sub

Private Sub VfsCopyPreparedTemplate(ByVal fso As Object, ByVal sourceDocument As Document, ByVal destination As String)
    Dim verified As Document
    Set verified = VfsRequireSource()
    If Not verified Is sourceDocument Then VfsFail "Preparation template changed."
    VfsRequireLocal fso, sourceDocument.FullName
    If Not sourceDocument.Saved Then VfsFail "Save the preparation template before copying."
    fso.CopyFile sourceDocument.FullName, destination, False
End Sub

Private Function VfsOwnedLoaded(ByVal fso As Object, ByVal fullName As String) As Boolean
    Dim addin As AddIn, actual As String, expected As String
    On Error GoTo Failed
    expected = fso.GetAbsolutePathName(fullName)
    For Each addin In AddIns
        If addin.Installed Then
            actual = fso.GetAbsolutePathName(fso.BuildPath(addin.Path, addin.Name))
            If StrComp(actual, expected, vbTextCompare) = 0 Then
                VfsOwnedLoaded = True
                Exit Function
            End If
        End If
    Next addin
    Exit Function
Failed:
    VfsFail "The existing Vefeyat add-in could not be checked. Close all Word windows before retrying."
End Function

Private Sub VfsStageUpdate(ByVal fso As Object, ByVal root As String, ByVal readyPath As String, _
                           ByVal installedPath As String, ByVal startup As String)
    Dim pendingFolder As String, pendingPath As String, notesPath As String
    Dim transactionFolder As String, stagedNotes As String, previousPending As String, previousNotes As String
    Dim backupPath As String, notes As String, expectedBytes As Variant
    Dim oldFileMoved As Boolean, oldNotesMoved As Boolean, newFileMoved As Boolean, newNotesMoved As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    pendingFolder = fso.BuildPath(root, "Guncelleme")
    If Not VfsFolderOutsideStartup(fso, pendingFolder, startup) Then
        VfsFail "The update folder must be outside the Word Startup folder."
    End If
    VfsEnsureFolder fso, pendingFolder
    pendingPath = fso.BuildPath(pendingFolder, VfsTemplateName)
    notesPath = fso.BuildPath(pendingFolder, "GUNCELLEME.txt")
    transactionFolder = fso.GetParentFolderName(readyPath)
    stagedNotes = fso.BuildPath(transactionFolder, "GUNCELLEME.txt")
    previousPending = fso.BuildPath(transactionFolder, "onceki-guncelleme.dotm")
    previousNotes = fso.BuildPath(transactionFolder, "onceki-guncelleme.txt")
    expectedBytes = VfsReadBytes(readyPath)

    ' Reading a loaded file for a backup is best-effort. Never unload it.
    backupPath = VfsTryBackupInstalled(fso, root, installedPath)
    notes = "VEFEYAT GUNCELLEMESI HAZIR; HENUZ KURULMADI." & vbCrLf & _
        "SURUM: " & VfrVersion & vbCrLf & vbCrLf & _
        "Mevcut Word eklentisi degistirilmedi. Yeni dosya asagidaki klasorde hazirlandi." & vbCrLf & vbCrLf & _
        "KAYNAK DOSYA:" & vbCrLf & pendingPath & vbCrLf & vbCrLf & _
        "HEDEF KLASOR (Word'un gercek baslangic klasoru):" & vbCrLf & startup & vbCrLf & vbCrLf & _
        "HEDEF DOSYA:" & vbCrLf & installedPath & vbCrLf & vbCrLf & _
        "1. Acik belgelerinizi kaydedin ve Word'un BUTUN pencerelerini kapatin." & vbCrLf
    If Len(backupPath) > 0 Then
        notes = notes & "2. Onceki eklentinin dogrulanmis yedegi burada saklandi:" & vbCrLf & backupPath & vbCrLf
    Else
        notes = notes & "2. Otomatik yedek olusturulamadi. Word kapaliyken hedefte eski VefeyatEklenti.dotm varsa," & vbCrLf & _
            "   ONCE onu Word baslangic klasorunun DISINDA sececeginiz bir klasore yedekleyin." & vbCrLf
        If VfsFolderOutsideStartup(fso, fso.BuildPath(root, "Backups"), startup) Then
            notes = notes & "   Ornegin su yerel yedek klasorunu kullanabilirsiniz:" & vbCrLf & fso.BuildPath(root, "Backups") & vbCrLf
        End If
    End If
    notes = notes & "3. Dosya Gezgini ile YALNIZ yukaridaki KAYNAK DOSYAYI HEDEF KLASORE kopyalayin." & vbCrLf & _
        "   Hedefteki ayni adli VefeyatEklenti.dotm dosyasini degistirin." & vbCrLf & _
        "   GUNCELLEME.txt dosyasini veya hazirlik sablonunu baslangic klasorune kopyalamayin." & vbCrLf & _
        "4. Word'u yeniden acin ve normal bir belgeyle Vefeyat sekmesini deneyin." & vbCrLf & vbCrLf & _
        "CSV verileriniz ve hazirlik sablonunuz korunur. Bu dosya henuz otomatik olarak kurulmus degildir." & vbCrLf
    VfsWriteVerifiedUtf8 stagedNotes, notes

    ' Commit only the pending pair. Startup and its loaded add-in are untouched.
    If fso.FileExists(pendingPath) Then
        fso.MoveFile pendingPath, previousPending
        oldFileMoved = True
    End If
    If fso.FileExists(notesPath) Then
        fso.MoveFile notesPath, previousNotes
        oldNotesMoved = True
    End If
    fso.MoveFile readyPath, pendingPath
    newFileMoved = True
    fso.MoveFile stagedNotes, notesPath
    newNotesMoved = True
    If Not VfsBytesEqual(expectedBytes, VfsReadBytes(pendingPath)) Then VfsFail "The pending update file changed."
    If VfsReadUtf8(notesPath) <> notes Then VfsFail "The pending update instructions changed."
    VfMessage "Guncelleme hazirlandi; HENUZ KURULMADI." & vbCrLf & vbCrLf & _
        "Mevcut eklenti acik Word icinde degistirilmedi." & vbCrLf & vbCrLf & _
        "Word'un BUTUN pencerelerini kapatin. Sonra GUNCELLEME.txt icindeki yedekleme ve dosya kopyalama adimlarini uygulayin." & _
        vbCrLf & vbCrLf & "Guncelleme klasoru:" & vbCrLf & pendingFolder & vbCrLf & vbCrLf & _
        "Hedef klasor:" & vbCrLf & startup, vbInformation, "Vefeyat"
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If newNotesMoved Then fso.DeleteFile notesPath, False
    If newFileMoved Then fso.MoveFile pendingPath, readyPath
    If oldFileMoved Then fso.MoveFile previousPending, pendingPath
    If oldNotesMoved Then fso.MoveFile previousNotes, notesPath
    On Error GoTo 0
    Err.Raise errorNumber, "VefeyatSetup", errorText
End Sub

Private Function VfsFolderOutsideStartup(ByVal fso As Object, ByVal folderPath As String, ByVal startup As String) As Boolean
    Dim startupPrefix As String, candidatePrefix As String
    startupPrefix = fso.GetAbsolutePathName(fso.BuildPath(startup, "."))
    If Right$(startupPrefix, 1) <> "\" Then startupPrefix = startupPrefix & "\"
    candidatePrefix = fso.GetAbsolutePathName(fso.BuildPath(folderPath, "."))
    If Right$(candidatePrefix, 1) <> "\" Then candidatePrefix = candidatePrefix & "\"
    VfsFolderOutsideStartup = (StrComp(Left$(candidatePrefix, Len(startupPrefix)), startupPrefix, vbTextCompare) <> 0)
End Function

Private Function VfsTryBackupInstalled(ByVal fso As Object, ByVal root As String, ByVal installedPath As String) As String
    Dim backupPath As String
    On Error GoTo NotBackedUp
    If Not fso.FileExists(installedPath) Then Exit Function
    If Not VfsFolderOutsideStartup(fso, fso.BuildPath(root, "Backups"), fso.GetParentFolderName(installedPath)) Then Exit Function
    VfsEnsureFolder fso, fso.BuildPath(root, "Backups")
    backupPath = fso.BuildPath(fso.BuildPath(root, "Backups"), fso.GetBaseName(fso.GetTempName) & ".dotm")
    fso.CopyFile installedPath, backupPath, False
    If Not VfsBytesEqual(VfsReadBytes(installedPath), VfsReadBytes(backupPath)) Then GoTo NotBackedUp
    VfsTryBackupInstalled = backupPath
    Exit Function
NotBackedUp:
    On Error Resume Next
    If Len(backupPath) > 0 Then
        If fso.FileExists(backupPath) Then fso.DeleteFile backupPath, False
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub VfsWriteVerifiedUtf8(ByVal path As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.SaveToFile path, 2
    stream.Close
    If VfsReadUtf8(path) <> value Then VfsFail "The update instructions could not be verified."
End Sub

Private Function VfsReadUtf8(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    VfsReadUtf8 = stream.ReadText(-1)
    stream.Close
End Function

Private Function VfsUnloadOwnedAddin(ByVal fso As Object, ByVal fullName As String) As Boolean
    Dim addin As AddIn, actual As String
    On Error GoTo Failed
    For Each addin In AddIns
        actual = fso.BuildPath(addin.Path, addin.Name)
        If StrComp(fso.GetAbsolutePathName(actual), fso.GetAbsolutePathName(fullName), vbTextCompare) = 0 Then
            If addin.Installed Then
                addin.Installed = False
                VfsUnloadOwnedAddin = True
            End If
        End If
    Next addin
    Exit Function
Failed:
    VfsFail VfsRestartMessage()
End Function

Private Sub VfsRequireLocal(ByVal fso As Object, ByVal path As String)
    Dim driveName As String, drive As Object
    If Len(path) < 3 Then VfsFail VfText("Yerel dosya yolu gerekli.")
    If Mid$(path, 2, 2) <> ":\" Then VfsFail VfText("A\u011f yolu kullan\u0131lamaz; bilgisayar\u0131n\u0131zdaki bir dosyay\u0131 se\u00e7in.")
    driveName = fso.GetDriveName(path)
    If Len(driveName) = 0 Then VfsFail VfText("Yerel s\u00fcr\u00fcc\u00fc bulunamad\u0131.")
    Set drive = fso.GetDrive(driveName)
    If drive.DriveType = 3 Then VfsFail VfText("A\u011f s\u00fcr\u00fcc\u00fcs\u00fc kullan\u0131lamaz.")
End Sub

Private Sub VfsEnsureFolder(ByVal fso As Object, ByVal path As String)
    Dim parent As String
    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) = 0 Or parent = path Then VfsFail VfText("Klas\u00f6r olu\u015fturulamad\u0131: ") & path
    If Not fso.FolderExists(parent) Then VfsEnsureFolder fso, parent
    fso.CreateFolder path
End Sub

Private Function VfsLoadXml(ByVal path As String) As Object
    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.validateOnParse = False
    dom.resolveExternals = False
    dom.setProperty "ProhibitDTD", True
    If Not dom.Load(path) Then VfsFail VfText("XML dosyas\u0131 okunamad\u0131: ") & dom.parseError.reason
    Set VfsLoadXml = dom
End Function

Private Sub VfsValidateRibbon(ByVal path As String)
    Dim fso As Object, dom As Object, attr As Object, element As Object, value As String
    Dim callbackCount As Long, name As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then VfsFail VfText("VefeyatRibbon.xml bulunamad\u0131.")
    If fso.GetFile(path).Size > 131072 Then VfsFail VfText("\u015eerit XML dosyas\u0131 beklenenden b\u00fcy\u00fck.")
    Set dom = VfsLoadXml(path)
    If dom.documentElement.baseName <> "customUI" Then VfsFail VfText("Ge\u00e7erli bir Vefey\u00e2t \u015ferit dosyas\u0131 gerekli.")
    If dom.documentElement.namespaceURI <> VfsUiNamespace Then VfsFail VfText("\u015eerit XML ad alan\u0131 ge\u00e7ersiz.")
    For Each element In dom.SelectNodes("//*")
        If element.namespaceURI <> VfsUiNamespace Then VfsFail VfText("\u015eerit dosyas\u0131 beklenmeyen bir XML ad alan\u0131 i\u00e7eriyor.")
    Next element
    For Each attr In dom.SelectNodes("//@*")
        name = attr.baseName
        If Left$(name, 3) = "get" Or Left$(name, 2) = "on" Or name = "loadImage" Then
            value = CStr(attr.nodeValue)
            If Not VfsValidCallback(value) Then VfsFail VfText("Beklenmeyen \u015ferit geri \u00e7a\u011fr\u0131s\u0131: ") & value
            callbackCount = callbackCount + 1
        End If
    Next attr
    If callbackCount = 0 Then VfsFail VfText("\u015eerit dosyas\u0131nda Vefey\u00e2t komutlar\u0131 bulunamad\u0131.")
End Sub

Private Function VfsValidCallback(ByVal value As String) As Boolean
    Dim prefix As String, suffix As String, i As Long, ch As String
    prefix = VfsInstalledModule & ".Vefeyat0341Ui"
    If Left$(value, Len(prefix)) <> prefix Then Exit Function
    suffix = Mid$(value, Len(prefix) + 1)
    If Len(suffix) = 0 Then Exit Function
    For i = 1 To Len(suffix)
        ch = Mid$(suffix, i, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    VfsValidCallback = True
End Function

Private Sub VfsInjectRibbon(ByVal fso As Object, ByVal unpacked As String, ByVal ribbonPath As String)
    Dim rels As Object, types As Object, node As Object, found As Object
    Set rels = VfsLoadXml(unpacked & "\_rels\.rels")
    If rels.documentElement.namespaceURI <> VfsRelationshipsNamespace Then VfsFail "Invalid package relationships."
    rels.setProperty "SelectionNamespaces", "xmlns:r='" & VfsRelationshipsNamespace & "'"
    Set found = rels.SelectSingleNode("/r:Relationships/r:Relationship[@Type='" & VfsUiRelationship & "']")
    If Not found Is Nothing Then VfsFail "The fresh template already contains Ribbon customizations."
    Set found = rels.SelectSingleNode("/r:Relationships/r:Relationship[@Id='rIdVefeyatUI']")
    If Not found Is Nothing Then VfsFail "Package relationship identifier collision."
    Set node = rels.createNode(1, "Relationship", VfsRelationshipsNamespace)
    node.setAttribute "Id", "rIdVefeyatUI"
    node.setAttribute "Type", VfsUiRelationship
    node.setAttribute "Target", "/customUI/customUI14.xml"
    rels.documentElement.appendChild node
    rels.Save unpacked & "\_rels\.rels"
    Set types = VfsLoadXml(unpacked & "\[Content_Types].xml")
    If types.documentElement.namespaceURI <> VfsContentTypesNamespace Then VfsFail "Invalid package content types."
    types.setProperty "SelectionNamespaces", "xmlns:t='" & VfsContentTypesNamespace & "'"
    Set found = types.SelectSingleNode("/t:Types/t:Override[@PartName='/customUI/customUI14.xml']")
    If Not found Is Nothing Then VfsFail "The fresh template already declares the Ribbon part."
    Set node = types.createNode(1, "Override", VfsContentTypesNamespace)
    node.setAttribute "PartName", "/customUI/customUI14.xml"
    node.setAttribute "ContentType", "application/xml"
    types.documentElement.appendChild node
    types.Save unpacked & "\[Content_Types].xml"
    VfsEnsureFolder fso, unpacked & "\customUI"
    fso.CopyFile ribbonPath, unpacked & "\customUI\customUI14.xml", False
End Sub

Private Function VfsShellFolder(ByVal shell As Object, ByVal path As String) As Object
    Dim variantPath As Variant
    variantPath = path
    Set VfsShellFolder = shell.NameSpace(variantPath)
End Function

Private Function VfsExtractZip(ByVal fso As Object, ByVal zipPath As String, ByVal folderPath As String) As Object
    Dim shell As Object, archive As Object, folder As Object, manifest As Object
    Set shell = CreateObject("Shell.Application")
    Set archive = VfsShellFolder(shell, zipPath)
    If archive Is Nothing Then VfsFail VfText("Windows ZIP dosyas\u0131n\u0131 a\u00e7amad\u0131.")
    Set manifest = CreateObject("Scripting.Dictionary")
    VfsShellManifest archive, "", manifest
    If manifest.Count = 0 Then VfsFail "The Word ZIP package is empty."
    fso.CreateFolder folderPath
    Set folder = VfsShellFolder(shell, folderPath)
    If folder Is Nothing Then VfsFail "Windows could not open the staging folder."
    folder.CopyHere archive.Items, 1556
    VfsWaitForExtraction fso, folderPath, manifest
    Set folder = Nothing
    Set archive = Nothing
    Set shell = Nothing
    Set VfsExtractZip = manifest
End Function

Private Sub VfsShellManifest(ByVal folder As Object, ByVal relative As String, ByVal manifest As Object)
    Dim item As Object, child As Object, name As String, leaf As String
    For Each item In folder.Items
        ' Path preserves file extensions even when Explorer hides known extensions.
        leaf = CStr(item.Path)
        leaf = Mid$(leaf, InStrRev(leaf, "\") + 1)
        If Len(leaf) = 0 Or leaf = "." Or leaf = ".." Then VfsFail "Invalid ZIP package part name."
        If InStr(leaf, ":") > 0 Or InStr(leaf, "/") > 0 Then VfsFail "Invalid ZIP package part name."
        name = relative & leaf
        If item.IsFolder Then
            Set child = item.GetFolder
            VfsShellManifest child, name & "\", manifest
        Else
            If manifest.Exists(name) Then VfsFail "Duplicate ZIP package part."
            If CDbl(item.Size) > 33554432# Then VfsFail "Unexpectedly large package part."
            manifest.Add name, CDbl(item.Size)
        End If
    Next item
End Sub

Private Function VfsFolderManifest(ByVal fso As Object, ByVal folderPath As String) As Object
    Dim manifest As Object
    Set manifest = CreateObject("Scripting.Dictionary")
    VfsAddFolderFiles fso.GetFolder(folderPath), "", manifest
    Set VfsFolderManifest = manifest
End Function

Private Sub VfsAddFolderFiles(ByVal folder As Object, ByVal relative As String, ByVal manifest As Object)
    Dim file As Object, child As Object
    For Each file In folder.Files
        If CDbl(file.Size) > 33554432# Then VfsFail "Unexpectedly large package part."
        manifest.Add relative & CStr(file.Name), CDbl(file.Size)
    Next file
    For Each child In folder.SubFolders
        VfsAddFolderFiles child, relative & CStr(child.Name) & "\", manifest
    Next child
End Sub

Private Function VfsFilesReady(ByVal fso As Object, ByVal folderPath As String, ByVal manifest As Object) As Boolean
    Dim key As Variant, path As String
    On Error GoTo NotReady
    For Each key In manifest.Keys
        path = folderPath & "\" & CStr(key)
        If Not fso.FileExists(path) Then Exit Function
        If CDbl(fso.GetFile(path).Size) <> CDbl(manifest(key)) Then Exit Function
        If Not VfsCanReadExclusive(path) Then Exit Function
    Next key
    VfsFilesReady = True
NotReady:
End Function

Private Sub VfsWaitForExtraction(ByVal fso As Object, ByVal folderPath As String, ByVal manifest As Object)
    Dim started As Single, stable As Single, lastPoll As Single, matched As Boolean
    started = Timer
    lastPoll = started
    Do
        If VfsElapsed(lastPoll) >= 0.1 Then
            lastPoll = Timer
            If VfsFilesReady(fso, folderPath, manifest) Then
                If Not matched Then stable = Timer
                matched = True
                If VfsElapsed(stable) >= 0.3 Then Exit Sub
            Else
                matched = False
            End If
        End If
        If VfsElapsed(started) > 60 Then VfsFail VfText("ZIP dosyas\u0131 zaman\u0131nda a\u00e7\u0131lamad\u0131. L\u00fctfen yeniden deneyin.")
        DoEvents
    Loop
End Sub

Private Sub VfsCreateZip(ByVal fso As Object, ByVal folderPath As String, ByVal zipPath As String, ByVal manifest As Object)
    Dim header(0 To 21) As Byte, fileNumber As Integer, opened As Boolean
    Dim shell As Object, archive As Object, folder As Object, errorText As String
    On Error GoTo Failed
    header(0) = 80
    header(1) = 75
    header(2) = 5
    header(3) = 6
    fileNumber = FreeFile
    Open zipPath For Binary Access Write As #fileNumber
    opened = True
    Put #fileNumber, , header
    Close #fileNumber
    opened = False
    Set shell = CreateObject("Shell.Application")
    Set archive = VfsShellFolder(shell, zipPath)
    Set folder = VfsShellFolder(shell, folderPath)
    If archive Is Nothing Or folder Is Nothing Then VfsFail "Windows could not create the ZIP package."
    archive.CopyHere folder.Items, 1556
    VfsWaitForArchive fso, shell, zipPath, manifest
    Set archive = Nothing
    Set folder = Nothing
    Set shell = Nothing
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If opened Then Close #fileNumber
    On Error GoTo 0
    VfsFail errorText
End Sub

Private Sub VfsWaitForArchive(ByVal fso As Object, ByVal shell As Object, ByVal zipPath As String, ByVal manifest As Object)
    Dim started As Single, stable As Single, lastPoll As Single, matched As Boolean
    Dim previousSize As Double, currentSize As Double
    started = Timer
    lastPoll = started
    Do
        If VfsElapsed(lastPoll) >= 0.1 Then
            lastPoll = Timer
            currentSize = CDbl(fso.GetFile(zipPath).Size)
            If VfsArchiveMatches(shell, zipPath, manifest) And VfsCanReadExclusive(zipPath) Then
                If Not matched Or currentSize <> previousSize Then stable = Timer
                matched = True
                If VfsElapsed(stable) >= 0.5 Then Exit Sub
            Else
                matched = False
            End If
            previousSize = currentSize
        End If
        If VfsElapsed(started) > 60 Then VfsFail VfText("ZIP paketi tamamlanamad\u0131. Mevcut eklenti de\u011fi\u015ftirilmedi.")
        DoEvents
    Loop
End Sub

Private Function VfsArchiveMatches(ByVal shell As Object, ByVal zipPath As String, ByVal expected As Object) As Boolean
    Dim archive As Object, actual As Object, key As Variant
    On Error GoTo NotReady
    Set archive = VfsShellFolder(shell, zipPath)
    If archive Is Nothing Then Exit Function
    Set actual = CreateObject("Scripting.Dictionary")
    VfsShellManifest archive, "", actual
    If actual.Count <> expected.Count Then Exit Function
    For Each key In expected.Keys
        If Not actual.Exists(key) Then Exit Function
        If CDbl(actual(key)) <> CDbl(expected(key)) Then Exit Function
    Next key
    VfsArchiveMatches = True
NotReady:
End Function

Private Function VfsElapsed(ByVal started As Single) As Single
    Dim value As Single
    value = Timer - started
    If value < 0 Then value = value + 86400
    VfsElapsed = value
End Function

Private Function VfsCanReadExclusive(ByVal path As String) As Boolean
    Dim fileNumber As Integer, opened As Boolean
    On Error GoTo Failed
    fileNumber = FreeFile
    Open path For Binary Access Read Lock Read Write As #fileNumber
    opened = True
    Close #fileNumber
    opened = False
    VfsCanReadExclusive = True
    Exit Function
Failed:
    On Error Resume Next
    If opened Then Close #fileNumber
    On Error GoTo 0
End Function

Private Function VfsReadBytes(ByVal path As String) As Variant
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile path
    VfsReadBytes = stream.Read
    stream.Close
End Function

Private Function VfsBytesEqual(ByVal first As Variant, ByVal second As Variant) As Boolean
    Dim i As Long
    If UBound(first) <> UBound(second) Then Exit Function
    If LBound(first) <> LBound(second) Then Exit Function
    For i = LBound(first) To UBound(first)
        If first(i) <> second(i) Then Exit Function
    Next i
    VfsBytesEqual = True
End Function

Private Sub VfsRequirePart(ByVal fso As Object, ByVal folderPath As String, ByVal relative As String)
    Dim path As String
    path = folderPath & "\" & relative
    If Not fso.FileExists(path) Then VfsFail "Missing Word package part: " & relative
    If fso.GetFile(path).Size = 0 Then VfsFail "Empty Word package part: " & relative
End Sub

Private Sub VfsVerifyPackage(ByVal fso As Object, ByVal folderPath As String, ByVal expected As Object, ByVal originalVba As Variant)
    Dim actual As Object, key As Variant, dom As Object, node As Object
    Set actual = VfsFolderManifest(fso, folderPath)
    If actual.Count <> expected.Count Then VfsFail "The ZIP package file count changed."
    For Each key In expected.Keys
        If Not actual.Exists(key) Then VfsFail "A verified ZIP package part is missing."
        If CDbl(actual(key)) <> CDbl(expected(key)) Then VfsFail "A verified ZIP package part has changed size."
    Next key
    VfsRequirePart fso, folderPath, "[Content_Types].xml"
    VfsRequirePart fso, folderPath, "_rels\.rels"
    VfsRequirePart fso, folderPath, "word\vbaProject.bin"
    VfsRequirePart fso, folderPath, "customUI\customUI14.xml"
    If Not VfsBytesEqual(originalVba, VfsReadBytes(folderPath & "\word\vbaProject.bin")) Then
        VfsFail "The VBA project changed while building the ZIP package."
    End If
    VfsValidateRibbon folderPath & "\customUI\customUI14.xml"
    Set dom = VfsLoadXml(folderPath & "\_rels\.rels")
    dom.setProperty "SelectionNamespaces", "xmlns:r='" & VfsRelationshipsNamespace & "'"
    Set node = dom.SelectSingleNode("/r:Relationships/r:Relationship[@Type='" & VfsUiRelationship & "']")
    If node Is Nothing Then VfsFail "The Ribbon package relationship is missing."
    If node.getAttribute("Target") <> "/customUI/customUI14.xml" Then VfsFail "The Ribbon package relationship is invalid."
    Set dom = VfsLoadXml(folderPath & "\[Content_Types].xml")
    dom.setProperty "SelectionNamespaces", "xmlns:t='" & VfsContentTypesNamespace & "'"
    Set node = dom.SelectSingleNode("/t:Types/t:Override[@PartName='/customUI/customUI14.xml']")
    If node Is Nothing Then VfsFail "The Ribbon content type is missing."
    If node.getAttribute("ContentType") <> "application/xml" Then VfsFail "The Ribbon content type is invalid."
End Sub

Private Function VfsRestartMessage() As String
    VfsRestartMessage = VfText("Eklenti dosyas\u0131 Word taraf\u0131ndan kullan\u0131l\u0131yor. \u00c7al\u0131\u015fman\u0131z\u0131 kaydedip t\u00fcm Word pencerelerini kapat\u0131n. Haz\u0131rl\u0131k \u015fablonunu daha sonra Dosya > A\u00e7 ile a\u00e7abilirsiniz.")
End Function

Private Sub VfsFail(ByVal message As String)
    Err.Raise vbObjectError + 2480, "VefeyatSetup", message
End Sub
' ---------------------------------------------------------------------------
' 0.3.37: Teslim oncesi temizlik ve tarih tutarliligi raporu.
' Temizlik yalniz Vefeyat'a ait ADDIN tarih alanlarini duz metne cevirir,
' Vefeyat gizli yer imlerini ve Vefeyat belge degiskenlerini siler. Diger
' alanlara (Zotero, EndNote, Mendeley, sayfa numarasi vb.) dokunulmaz.
' Rapor salt okunurdur; belgeye yazmaz.
' ---------------------------------------------------------------------------

Private Function VfcOwnedField(ByVal field As Word.Field) As Boolean
    Dim code As String
    If field.Type <> wdFieldAddin Then Exit Function
    code = LTrim$(field.Code.Text)
    If StrComp(Left$(code, 6), "ADDIN ", vbTextCompare) <> 0 Then Exit Function
    code = LTrim$(Mid$(code, 7))
    VfcOwnedField = (StrComp(Left$(code, Len(VfaTagPrefix)), VfaTagPrefix, vbTextCompare) = 0)
End Function

Private Function VfcOwnedBookmark(ByVal name As String) As Boolean
    If StrComp(Left$(name, Len(VfaManualBookmarkPrefix)), VfaManualBookmarkPrefix, vbTextCompare) = 0 Then
        VfcOwnedBookmark = True
    ElseIf StrComp(Left$(name, 5), "_Vdc_", vbTextCompare) = 0 Then
        VfcOwnedBookmark = True
    ElseIf StrComp(Left$(name, 5), "_Vrg_", vbTextCompare) = 0 Then
        VfcOwnedBookmark = True
    End If
End Function

Private Function VfcOwnedVariable(ByVal name As String) As Boolean
    VfcOwnedVariable = (StrComp(Left$(name, 7), "Vefeyat", vbTextCompare) = 0)
End Function

Private Function VfcCountFields(ByVal doc As Word.Document) As Long
    Dim story As Word.Range, field As Word.Field, total As Long
    For Each story In doc.StoryRanges
        Do While Not story Is Nothing
            For Each field In story.Fields
                If VfcOwnedField(field) Then total = total + 1
            Next field
            Set story = story.NextStoryRange
        Loop
    Next story
    VfcCountFields = total
End Function

Private Function VfcBookmarkNames(ByVal doc As Word.Document) As Collection
    Dim result As New Collection, mark As Word.Bookmark, previous As Boolean
    previous = doc.Bookmarks.ShowHidden
    doc.Bookmarks.ShowHidden = True
    For Each mark In doc.Bookmarks
        If VfcOwnedBookmark(mark.Name) Then result.Add mark.Name
    Next mark
    doc.Bookmarks.ShowHidden = previous
    Set VfcBookmarkNames = result
End Function

Private Function VfcVariableNames(ByVal doc As Word.Document) As Collection
    Dim result As New Collection, variable As Word.Variable
    For Each variable In doc.Variables
        If VfcOwnedVariable(variable.Name) Then result.Add variable.Name
    Next variable
    Set VfcVariableNames = result
End Function

Private Function VfcIdle() As Boolean
    If mVfcBusy Or mVfbRunning Or mVfmBusy Or mVfsBusy Or mVfwBusy Or mVftBusy Then Exit Function
    If mVfjBusy Or mVfjPreparing Or mVfrInsertBusy Or mVflBusy Or mVfxBusy Then Exit Function
    If Not mVfwPanel Is Nothing Then Exit Function
    If Not mVftPanel Is Nothing Then Exit Function
    If Not mVfjPanel Is Nothing Then Exit Function
    If Not mVfbProgress Is Nothing Then Exit Function
    If mVfwPaused Then Exit Function
    VfcIdle = True
End Function

' Non-interactive core. The caller owns all checks, confirmation and Undo.
Private Sub VfcCleanDocument(ByVal doc As Word.Document, ByRef unlinked As Long, _
                             ByRef removedBookmarks As Long, ByRef removedVariables As Long)
    Dim story As Word.Range, field As Word.Field, i As Long, name As Variant
    Dim names As Collection, previous As Boolean, errorNumber As Long, errorText As String
    unlinked = 0: removedBookmarks = 0: removedVariables = 0
    For Each story In doc.StoryRanges
        Do While Not story Is Nothing
            ' Reverse order keeps the remaining indices valid after each Unlink.
            For i = story.Fields.Count To 1 Step -1
                Set field = story.Fields(i)
                If VfcOwnedField(field) Then
                    field.Locked = False
                    field.Unlink
                    unlinked = unlinked + 1
                End If
            Next i
            Set story = story.NextStoryRange
        Loop
    Next story
    Set names = VfcBookmarkNames(doc)
    previous = doc.Bookmarks.ShowHidden
    doc.Bookmarks.ShowHidden = True
    On Error GoTo Failed
    For Each name In names
        If doc.Bookmarks.Exists(CStr(name)) Then
            doc.Bookmarks(CStr(name)).Delete
            removedBookmarks = removedBookmarks + 1
        End If
    Next name
    doc.Bookmarks.ShowHidden = previous
    On Error GoTo 0
    Set names = VfcVariableNames(doc)
    For Each name In names
        doc.Variables(CStr(name)).Delete
        removedVariables = removedVariables + 1
    Next name
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    doc.Bookmarks.ShowHidden = previous
    On Error GoTo 0
    Err.Raise errorNumber, "VefeyatFinalize", errorText
End Sub

Private Sub VfcForgetDocument(ByVal doc As Word.Document)
    On Error Resume Next
    VfvCancelCommand
    VfmMetadataChanged doc
    VfmInvalidate doc
    Vefeyat0341UiInvalidate
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub Vefeyat0341UiFinalize(ByVal control As Object)
    Dim doc As Word.Document, fieldCount As Long, bookmarkCount As Long, variableCount As Long
    Dim message As String, undoObject As Object, undoStarted As Boolean, changed As Boolean, errorText As String
    Dim unlinked As Long, removedBookmarks As Long, removedVariables As Long
    If Documents.Count = 0 Then Exit Sub
    If Not VfcIdle() Then
        VfaWarn VfText("\u00d6nce a\u00e7\u0131k Vefey\u00e2t pencerelerini (inceleme, \u015fah\u0131s d\u00f6k\u00fcm\u00fc, belge sa\u011fl\u0131\u011f\u0131, \u00f6nizleme) kapat\u0131n ve duraklat\u0131lm\u0131\u015f incelemeyi bitirin. Sonra temizli\u011fi yeniden ba\u015flat\u0131n.")
        Exit Sub
    End If
    Set doc = ActiveDocument
    If Not VfaEditable(doc) Then Exit Sub
    If doc.TrackRevisions Or doc.Revisions.Count > 0 Then
        VfaWarn VfText("De\u011fi\u015fiklik \u0130zleme a\u00e7\u0131k veya belgede bekleyen izlenen de\u011fi\u015fiklikler var. Temizlik yap\u0131lmad\u0131. \u00d6nce de\u011fi\u015fiklikleri kabul veya reddedip izlemeyi kapat\u0131n.")
        Exit Sub
    End If
    If Len(doc.Path) = 0 Or Not doc.Saved Then
        VfaWarn VfText("\u00d6nce belgeyi kaydedin. Temizlik diskteki kay\u0131tl\u0131 dosyaya dokunmaz; bir sorun olursa belgeyi kaydetmeden kapat\u0131p yeniden a\u00e7arak geri d\u00f6nebilirsiniz.")
        Exit Sub
    End If
    On Error GoTo Failed
    fieldCount = VfcCountFields(doc)
    bookmarkCount = VfcBookmarkNames(doc).Count
    variableCount = VfcVariableNames(doc).Count
    If fieldCount + bookmarkCount + variableCount = 0 Then
        VfMessage VfText("Bu belgede Vefey\u00e2t'a ait tarih alan\u0131, gizli yer imi veya belge de\u011fi\u015fkeni yok. Temizlenecek bir \u015fey bulunmad\u0131."), vbInformation, VfText("Vefey\u00e2t")
        Exit Sub
    End If
    message = VfText("Teslim \u00f6ncesi temizlik bu belgede \u015funlar\u0131 yapacak:") & vbCrLf & vbCrLf & _
        "- " & CStr(fieldCount) & VfText(" Vefey\u00e2t tarih alan\u0131 d\u00fcz metne \u00e7evrilecek (g\u00f6r\u00fcnen tarih aynen kal\u0131r).") & vbCrLf & _
        "- " & CStr(bookmarkCount) & VfText(" gizli Vefey\u00e2t yer imi silinecek.") & vbCrLf & _
        "- " & CStr(variableCount) & VfText(" Vefey\u00e2t belge de\u011fi\u015fkeni silinecek (kimlik onaylar\u0131, istisnalar, ad tercihleri, belge ayarlar\u0131 ve dinamik kay\u0131tlar).") & vbCrLf & vbCrLf & _
        VfText("Zotero, EndNote, Mendeley, sayfa numaras\u0131 ve di\u011fer alanlara dokunulmaz. Metindeki adlar ve tarihler de\u011fi\u015fmez.") & vbCrLf & _
        VfText("Bundan sonra Vefey\u00e2t bu belgedeki eski tarihleri tan\u0131maz ve g\u00fcncelleyemez.") & vbCrLf & vbCrLf & _
        VfText("\u00d6nerilen kullan\u0131m: Temizlikten sonra Dosya > Farkl\u0131 Kaydet ile ayr\u0131 bir teslim kopyas\u0131 kaydedin (\u00f6rne\u011fin Tez_teslim.docx). As\u0131l \u00e7al\u0131\u015fma dosyan\u0131z\u0131n \u00fczerine kaydetmeyin.") & vbCrLf & vbCrLf & _
        "Devam edilsin mi?"
    If VfMessage(message, vbYesNo Or vbQuestion Or vbDefaultButton2, VfText("Vefey\u00e2t: teslim \u00f6ncesi temizlik")) <> vbYes Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    If Not (ActiveDocument Is doc) Or Not VfcIdle() Or doc.TrackRevisions Then
        VfaWarn VfText("Belge veya Vefey\u00e2t durumu onaydan sonra de\u011fi\u015fti. Temizlik yap\u0131lmad\u0131.")
        Exit Sub
    End If
    mVfcBusy = True
    undoStarted = VfaBeginUndo(undoObject)
    Application.ScreenUpdating = False
    changed = True
    VfcCleanDocument doc, unlinked, removedBookmarks, removedVariables
    VfaEndUndo undoObject, undoStarted
    Application.ScreenUpdating = True
    VfcForgetDocument doc
    mVfcBusy = False
    VfMessage VfText("Temizlik tamamland\u0131.") & vbCrLf & vbCrLf & _
        CStr(unlinked) & VfText(" tarih alan\u0131 d\u00fcz metne \u00e7evrildi, ") & CStr(removedBookmarks) & " yer imi ve " & _
        CStr(removedVariables) & VfText(" belge de\u011fi\u015fkeni silindi.") & vbCrLf & vbCrLf & _
        VfText("\u015eimdi Dosya > Farkl\u0131 Kaydet ile teslim kopyas\u0131n\u0131 ayr\u0131 bir adla kaydedin. \u00c7al\u0131\u015fmaya devam etmek i\u00e7in as\u0131l dosyan\u0131z\u0131 (temizlenmemi\u015f h\u00e2lini) kullan\u0131n."), _
        vbInformation, VfText("Vefey\u00e2t: teslim \u00f6ncesi temizlik")
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    VfaEndUndo undoObject, undoStarted
    Application.ScreenUpdating = True
    VfcForgetDocument doc
    mVfcBusy = False
    On Error GoTo 0
    If changed Then
        VfaWarn VfText("Temizlik tamamlanamad\u0131: ") & Left$(errorText, 300) & vbCrLf & vbCrLf & _
            VfText("Belge k\u0131smen temizlenmi\u015f olabilir. Belgeyi kaydetmeden kapat\u0131p diskteki s\u00fcr\u00fcm\u00fc yeniden a\u00e7\u0131n.")
    Else
        VfaWarn VfText("Temizlik ba\u015flat\u0131lamad\u0131; belge de\u011fi\u015ftirilmedi: ") & Left$(errorText, 300)
    End If
End Sub

' ---- Tarih tutarliligi raporu (salt okunur) ----

Private Function VfcDateKind(ByVal body As String) As String
    Dim i As Long, ch As String, digitsOnly As Boolean
    body = Trim$(body)
    If InStr(body, ";") > 0 Then body = Trim$(Left$(body, InStr(body, ";") - 1))
    If Len(body) = 0 Then Exit Function
    If Not (Left$(body, 1) Like "[0-9]") Then Exit Function
    If InStr(body, "/") > 0 Then
        VfcDateKind = "hm"
        Exit Function
    End If
    If InStr(body, "-") > 0 Or InStr(body, ChrW(&H2013)) > 0 Then
        VfcDateKind = "range"
        Exit Function
    End If
    digitsOnly = True
    For i = 1 To Len(body)
        ch = Mid$(body, i, 1)
        If Not (ch Like "[0-9]") Then digitsOnly = False: Exit For
    Next i
    If digitsOnly Then VfcDateKind = "single" Else VfcDateKind = "other"
End Function

Private Sub VfcCount(ByVal counts As Object, ByVal key As String)
    If counts.Exists(key) Then counts(key) = CLng(counts(key)) + 1 Else counts.Add key, 1&
End Sub

Private Sub VfcExample(ByVal examples As Object, ByVal key As String, ByVal sample As String, ByVal place As String)
    Dim list As Collection
    If Not examples.Exists(key) Then
        Set list = New Collection
        examples.Add key, list
    Else
        Set list = examples(key)
    End If
    If list.Count < 3 Then list.Add sample & "  [" & place & "]"
End Sub

Private Sub VfcScanDates(ByVal text As String, ByVal place As String, ByVal counts As Object, ByVal examples As Object)
    Dim markers As Variant, labels As Variant, m As Long, position As Long, closing As Long
    Dim body As String, kind As String, sample As String
    markers = Array(VfText("(\u00f6."), VfText("(\u00f6l."), "(v.")
    labels = Array("marker:" & VfText("\u00f6."), "marker:" & VfText("\u00f6l."), "marker:v.")
    For m = 0 To 2
        position = InStr(1, text, CStr(markers(m)), vbBinaryCompare)
        Do While position > 0
            closing = InStr(position, text, ")", vbBinaryCompare)
            If closing > 0 And closing - position <= 60 Then
                body = Mid$(text, position + Len(CStr(markers(m))), closing - position - Len(CStr(markers(m))))
                kind = VfcDateKind(body)
                If Len(kind) > 0 Then
                    sample = Mid$(text, position, closing - position + 1)
                    VfcCount counts, CStr(labels(m))
                    VfcExample examples, CStr(labels(m)), sample, place
                    VfcCount counts, "kind:" & kind
                    VfcExample examples, "kind:" & kind, sample, place
                End If
            End If
            position = InStr(position + 1, text, CStr(markers(m)), vbBinaryCompare)
        Loop
    Next m
End Sub

Private Function VfcLines(ByVal counts As Object, ByVal examples As Object, ByVal keys As Variant, ByVal labels As Variant, ByRef mixed As Boolean) As String
    Dim i As Long, used As Long, maximum As Long, result As String, item As Variant, key As String
    For i = LBound(keys) To UBound(keys)
        key = CStr(keys(i))
        If counts.Exists(key) Then
            used = used + 1
            If CLng(counts(key)) > maximum Then maximum = CLng(counts(key))
        End If
    Next i
    mixed = (used > 1)
    For i = LBound(keys) To UBound(keys)
        key = CStr(keys(i))
        If counts.Exists(key) Then
            result = result & "  " & CStr(labels(i)) & ": " & CStr(counts(key))
            If mixed And CLng(counts(key)) < maximum Then
                result = result & VfText("   \u2190 az\u0131nl\u0131kta") & vbCrLf
                For Each item In examples(key)
                    result = result & "      " & CStr(item) & vbCrLf
                Next item
            Else
                result = result & vbCrLf
            End If
        End If
    Next i
    VfcLines = result
End Function

Public Sub Vefeyat0341UiConsistency(ByVal control As Object)
    Dim doc As Word.Document, counts As Object, examples As Object, report As String
    Dim mixedMarkers As Boolean, mixedKinds As Boolean, total As Long, key As Variant
    On Error GoTo Failed
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    Set counts = CreateObject("Scripting.Dictionary")
    Set examples = CreateObject("Scripting.Dictionary")
    VfcScanDates doc.StoryRanges(wdMainTextStory).Text, "Metin", counts, examples
    If doc.Footnotes.Count > 0 Then VfcScanDates doc.StoryRanges(wdFootnotesStory).Text, "Dipnot", counts, examples
    For Each key In counts.Keys
        If Left$(CStr(key), 7) = "marker:" Then total = total + CLng(counts(key))
    Next key
    If total = 0 Then
        VfMessage VfText("Ana metin ve dipnotlarda (\u00f6. \u2026), (\u00f6l. \u2026) veya (v. \u2026) bi\u00e7iminde tarih bulunamad\u0131."), vbInformation, VfText("Vefey\u00e2t: tarih tutarl\u0131l\u0131\u011f\u0131")
        Exit Sub
    End If
    report = "Ana metin ve dipnotlarda " & CStr(total) & VfText(" tarih bulundu. Bu rapor belgeyi de\u011fi\u015ftirmez.") & vbCrLf & vbCrLf
    report = report & VfText("K\u0131saltma:") & vbCrLf & _
        VfcLines(counts, examples, Array("marker:" & VfText("\u00f6."), "marker:" & VfText("\u00f6l."), "marker:v."), _
                 Array(VfText("\u00f6."), VfText("\u00f6l."), "v."), mixedMarkers) & vbCrLf
    report = report & VfText("Bi\u00e7im:") & vbCrLf & _
        VfcLines(counts, examples, Array("kind:hm", "kind:single", "kind:range", "kind:other"), _
                 Array(VfText("hicr\u00ee/mil\u00e2d\u00ee (505/1111)"), VfText("tek y\u0131l (1111)"), VfText("aral\u0131k (1100-1111)"), VfText("di\u011fer")), mixedKinds) & vbCrLf
    If mixedMarkers Or mixedKinds Then
        report = report & VfText("Belgede birden fazla k\u0131saltma veya bi\u00e7im var. Az\u0131nl\u0131kta kalan \u00f6rnekleri Word'\u00fcn Bul (Ctrl+F) komutuyla bulup tek bi\u00e7ime getirmeyi d\u00fc\u015f\u00fcn\u00fcn. Aral\u0131k ve belirsiz tarihler bilin\u00e7li olarak farkl\u0131 olabilir.")
    Else
        report = report & VfText("K\u0131saltma ve bi\u00e7im tutarl\u0131 g\u00f6r\u00fcn\u00fcyor.")
    End If
    report = report & vbCrLf & vbCrLf & VfText("Not: Ayn\u0131 \u015fahs\u0131n farkl\u0131 yaz\u0131mlar\u0131 i\u00e7in Belge > \u015eah\u0131s d\u00f6k\u00fcm\u00fc listesindeki ""kaynak yaz\u0131m\u0131 farkl\u0131"" uyar\u0131lar\u0131na bak\u0131n. Ba\u015fl\u0131k/altbilgi, sonnot ve metin kutular\u0131 bu rapora dahil de\u011fildir.")
    VfMessage report, vbInformation, VfText("Vefey\u00e2t: tarih tutarl\u0131l\u0131\u011f\u0131")
    Exit Sub
Failed:
    VfaWarn VfText("Tutarl\u0131l\u0131k raporu haz\u0131rlanamad\u0131: ") & Left$(Err.Description, 300)
End Sub

' Native acceptance helper, called from Vefeyat0341UiNativeChecks on its own
' synthetic document only.
Private Sub VftFinalizeNative(ByVal doc As Word.Document, ByRef count As Long)
    Dim other As Word.Field, unlinked As Long, removedBookmarks As Long, removedVariables As Long, mark As String
    Dim owned As Long
    owned = VfcCountFields(doc)
    VftAssert owned > 0, "finalize sees owned date fields", count
    Set other = doc.Fields.Add(Range:=doc.Range(0, 0), Type:=wdFieldEmpty, _
                               Text:="ADDIN ZOTERO_ITEM synthetic", PreserveFormatting:=False)
    mark = VfaManualBookmarkPrefix & "NativeFinalize"
    doc.Bookmarks.Add Name:=mark, Range:=doc.Range(0, 1)
    VfcCleanDocument doc, unlinked, removedBookmarks, removedVariables
    VftAssert unlinked = owned, "finalize unlinks every owned date field", count
    VftAssert VfcCountFields(doc) = 0, "no owned date field remains", count
    VftAssert doc.Fields.Count = 1, "foreign ADDIN field preserved", count
    VftAssert InStr(1, doc.Fields(1).Code.Text, "ZOTERO_ITEM", vbBinaryCompare) > 0, "foreign field code unchanged", count
    VftAssert InStr(1, doc.Content.Text, "(v. 1)", vbBinaryCompare) > 0, "visible date text kept", count
    VftAssert removedBookmarks >= 1 And VfcBookmarkNames(doc).Count = 0, "owned hidden bookmarks removed", count
    VftAssert removedVariables >= 1 And VfcVariableNames(doc).Count = 0, "owned document variables removed", count
    doc.Fields(1).Delete
End Sub
' ---------------------------------------------------------------------------
' 0.3.39: Sahislar dizini ve sahis raporu.
' Ikisi de Sahis dokumunun ayni taramasini kullanir ve yalniz kimligi kesin
' (tek guclu katalog eslesmesi veya mevcut Vefeyat tarih bagi) olan gecisleri
' kullanir. Belirsiz adlar tahminle dizine veya rapora sahis olarak girmez.
' Dizin isaretleri Word'un kendi XE alanlaridir: XE "Giris" \f "v". \f "v"
' Vefeyat'in dizin turudur; kullanicinin kendi XE alanlarina dokunulmaz.
' Isaret adin ONUNE konur; boylece adin hemen arkasindaki tarih alani ve duz
' tarih algilamasi degismez.
' ---------------------------------------------------------------------------

Private Function VfdOwnedCode(ByVal code As String, ByVal kind As String) As Boolean
    code = LCase$(Trim$(code))
    If Left$(code, Len(kind) + 1) <> kind & " " Then Exit Function
    VfdOwnedCode = (InStr(1, code, "\f ""v""", vbBinaryCompare) > 0)
End Function

Private Function VfdEntryText(ByVal value As String) As String
    Dim i As Long, code As Long, result As String, ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1): code = AscW(ch) And &HFFFF&
        If code < 32 Or code = 127 Then ch = " "
        If ch = "\" Or ch = """" Then ch = ""
        If ch = ";" Then ch = ","
        result = result & ch
    Next i
    result = Trim$(VfaNormalizeWhitespace(result))
    ' A colon starts a sub-entry in Word indexes; keep it as text.
    VfdEntryText = Replace(result, ":", "\:")
End Function

Private Function VfdAddMark(ByVal named As Word.Range, ByVal entry As String) As Word.Field
    Dim point As Word.Range
    Set point = named.Duplicate
    point.Collapse wdCollapseStart
    Set VfdAddMark = point.Fields.Add(Range:=point, Type:=wdFieldEmpty, _
        Text:="XE """ & entry & """ \f ""v""", PreserveFormatting:=False)
End Function

' Counts (countOnly) or deletes Vefeyat index marks, optionally with Vefeyat
' INDEX fields. Other XE and INDEX fields are never touched.
Private Function VfdRemoveMarks(ByVal doc As Word.Document, ByVal includeIndex As Boolean, ByVal countOnly As Boolean) As Long
    Dim story As Word.Range, field As Word.Field, i As Long, total As Long, owned As Boolean
    For Each story In doc.StoryRanges
        Do While Not story Is Nothing
            For i = story.Fields.Count To 1 Step -1
                Set field = story.Fields(i)
                owned = VfdOwnedCode(field.Code.Text, "xe")
                If includeIndex And Not owned Then owned = VfdOwnedCode(field.Code.Text, "index")
                If owned Then
                    total = total + 1
                    If Not countOnly Then field.Delete
                End If
            Next i
            Set story = story.NextStoryRange
        Loop
    Next story
    VfdRemoveMarks = total
End Function

Private Function VfdIndexField(ByVal doc As Word.Document) As Word.Field
    Dim field As Word.Field
    For Each field In doc.Content.Fields
        If VfdOwnedCode(field.Code.Text, "index") Then
            Set VfdIndexField = field
            Exit Function
        End If
    Next field
End Function

' Runs the Sahis dokumu preparation behind the usual cancellable progress
' window without opening the inventory panel. The caller calls VfjClose.
Private Function VfdPrepare(ByVal doc As Word.Document) As Boolean
    Dim detail As String
    On Error GoTo Failed
    VfjClose
    Set mVfjDoc = doc
    mVfjBusy = True: mVfjPreparing = True
    mVfbRunning = True: mVflSuspended = True
    VfbResetProgress
    Set mVfbProgress = New VefeyatKisiSec0341
    mVfbProgress.InitProgress
    mVfbProgress.Show vbModal
    Unload mVfbProgress
    Set mVfbProgress = Nothing
    mVfjPreparing = False: mVfjBusy = False
    mVfbRunning = False: mVflSuspended = False
    VfdPrepare = mVfjReady
    Exit Function
Failed:
    detail = Err.Description
    On Error Resume Next
    Unload mVfbProgress
    Set mVfbProgress = Nothing
    mVfjPreparing = False: mVfjBusy = False
    mVfbRunning = False: mVflSuspended = False
    VfjClose
    On Error GoTo 0
    Err.Raise VfaInvalidData, "Vefeyat", detail
End Function

' Groups the prepared inventory by confirmed person, in first-use order.
Private Function VfdPeople(ByRef unresolved As Long, ByRef unknown As Long, ByRef stale As Long, _
                           Optional ByRef confirmedCount As Long) As Collection
    Dim result As New Collection, groups As Object, group As Object, item As Object, rows As Collection
    Dim row As Variant, key As String, i As Long, confirmed As Object, resolved As Boolean
    Set groups = CreateObject("Scripting.Dictionary")
    unresolved = 0: unknown = 0: stale = 0: confirmedCount = 0
    Set confirmed = VfdConfirmedMap(mVfjDoc)
    For i = 1 To mVfjCount
        Set item = mVfjItems(i)
        resolved = False
        If item.Exists("VfdResolved") Then item.Remove "VfdResolved"
        If CBool(item("ChoiceReady")) Then
            row = item("Choice"): resolved = True
        ElseIf confirmed.Count > 0 Then
            If VfdConfirmedRow(item, confirmed, row) Then resolved = True: confirmedCount = confirmedCount + 1
        End If
        If Not resolved Then
            Set rows = item("Rows")
            If rows.Count = 0 Then unknown = unknown + 1 Else unresolved = unresolved + 1
        ElseIf VfjLiveName(item) Is Nothing Then
            stale = stale + 1
        Else
            item("VfdResolved") = True: item("VfdRow") = row
            key = VfpPersonKey(row)
            If groups.Exists(key) Then
                Set group = groups(key)
            Else
                Set group = CreateObject("Scripting.Dictionary")
                group("Key") = key: group("Row") = row: group("Entry") = ""
                Set group("Items") = New Collection
                groups.Add key, group
                result.Add group
            End If
            group("Items").Add item
        End If
    Next i
    Set VfdPeople = result
End Function

' Index entry: catalog short name as written in the catalog ("Razi, Fahreddin"
' order is kept because it suits an index). Different people that share one
' entry are told apart by the catalog date, else by the full name.
Private Sub VfdAssignEntries(ByVal people As Collection)
    Dim names As Object, group As Object, row As Variant, entry As String
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbTextCompare
    For Each group In people
        row = group("Row")
        entry = VfdEntryText(CStr(row(0)))
        If Len(entry) = 0 Then entry = VfdEntryText(VfnTextName(CStr(row(1))))
        group("Entry") = entry
        If names.Exists(entry) Then names(entry) = CLng(names(entry)) + 1 Else names.Add entry, 1&
    Next group
    For Each group In people
        If CLng(names(CStr(group("Entry")))) > 1 Then
            row = group("Row")
            If Len(VfdEntryText(CStr(row(3)))) > 0 Then
                group("Entry") = CStr(group("Entry")) & " " & VfdEntryText(CStr(row(3)))
            ElseIf Len(VfdEntryText(VfnTextName(CStr(row(1))))) > 0 Then
                group("Entry") = VfdEntryText(VfnTextName(CStr(row(1))))
            End If
        End If
    Next group
End Sub

Private Function VfdOccurrenceCount(ByVal people As Collection) As Long
    Dim group As Object, total As Long
    For Each group In people
        total = total + group("Items").Count
    Next group
    VfdOccurrenceCount = total
End Function

Private Function VfdSkippedText(ByVal unresolved As Long, ByVal unknown As Long, ByVal stale As Long) As String
    Dim text As String
    If unresolved > 0 Then text = text & "- " & CStr(unresolved) & VfText(" ge\u00e7i\u015fin kimli\u011fi belirsiz (birden fazla aday veya yak\u0131n e\u015fle\u015fme); \u015eah\u0131s d\u00f6k\u00fcm\u00fcnden do\u011frulay\u0131n.") & vbCrLf
    If unknown > 0 Then text = text & "- " & CStr(unknown) & VfText(" ad aday\u0131 katalogda yok.") & vbCrLf
    If stale > 0 Then text = text & "- " & CStr(stale) & VfText(" ge\u00e7i\u015f tarama s\u0131ras\u0131nda de\u011fi\u015fti.") & vbCrLf
    If Len(text) > 0 Then text = "Dahil edilmeyenler:" & vbCrLf & text
    VfdSkippedText = text
End Function

Private Sub VfdBusyWarning()
    VfaWarn VfText("\u00d6nce a\u00e7\u0131k Vefey\u00e2t pencerelerini (inceleme, \u015fah\u0131s d\u00f6k\u00fcm\u00fc, belge sa\u011fl\u0131\u011f\u0131, \u00f6nizleme) kapat\u0131n ve duraklat\u0131lm\u0131\u015f incelemeyi bitirin. Sonra komutu yeniden ba\u015flat\u0131n.")
End Sub

Public Sub Vefeyat0341UiIndex(ByVal control As Object)
    Dim doc As Word.Document, people As Collection, group As Object, item As Object, entries As Object
    Dim named As Word.Range, at As Word.Range, indexField As Word.Field, row As Variant
    Dim unresolved As Long, unknown As Long, stale As Long, occurrences As Long, oldMarks As Long, inserted As Long, i As Long
    Dim placeIndex As Boolean, indexNote As String, message As String, errorText As String, confirmedCount As Long
    Dim undoObject As Object, undoStarted As Boolean, changed As Boolean, prepared As Boolean
    If Documents.Count = 0 Then Exit Sub
    If Not VfcIdle() Then VfdBusyWarning: Exit Sub
    Set doc = ActiveDocument
    If Not VfaEditable(doc) Then Exit Sub
    If doc.TrackRevisions Then
        VfaWarn VfText("De\u011fi\u015fiklik \u0130zleme a\u00e7\u0131k. Dizin i\u015faretleri izlenen de\u011fi\u015fiklik olarak eklenmesin diye i\u015flem yap\u0131lmad\u0131. \u0130zlemeyi kapat\u0131p yeniden deneyin.")
        Exit Sub
    End If
    On Error GoTo Failed
    prepared = True
    If Not VfdPrepare(doc) Then GoTo Clean
    Set people = VfdPeople(unresolved, unknown, stale, confirmedCount)
    If people.Count = 0 Then
        VfMessage VfText("Kimli\u011fi kesin olarak belirlenmi\u015f \u015fah\u0131s bulunamad\u0131. Dizine eklenecek ge\u00e7i\u015f yok.") & vbCrLf & vbCrLf & _
            VfdSkippedText(unresolved, unknown, stale), vbInformation, VfText("Vefey\u00e2t: \u015fah\u0131slar dizini")
        GoTo Clean
    End If
    VfdAssignEntries people
    occurrences = VfdOccurrenceCount(people)
    oldMarks = VfdRemoveMarks(doc, False, True)
    Set indexField = VfdIndexField(doc)
    If indexField Is Nothing Then
        If Selection.StoryType = wdMainTextStory And Selection.Range.Fields.Count = 0 Then
            placeIndex = True
            Set at = Selection.Range.Duplicate
            at.Collapse wdCollapseStart
            indexNote = VfText("Dizin, imlecin \u015fu an bulundu\u011fu yere eklenecek. \u00d6nce imleci dizin sayfas\u0131na (\u00f6rne\u011fin ""\u015eah\u0131slar Dizini"" ba\u015fl\u0131\u011f\u0131n\u0131n alt\u0131na) koydu\u011funuzdan emin olun.")
        Else
            indexNote = VfText("Dizin alan\u0131 eklenmeyecek: imle\u00e7 ana metinde de\u011fil. Yaln\u0131z i\u015faretler eklenecek; dizin i\u00e7in imleci dizin sayfas\u0131na koyup komutu yeniden \u00e7al\u0131\u015ft\u0131r\u0131n.")
        End If
    Else
        indexNote = VfText("Belgedeki mevcut Vefey\u00e2t dizini g\u00fcncellenecek.")
    End If
    message = VfText("\u015eah\u0131slar dizini:") & vbCrLf & vbCrLf & _
        "- " & CStr(people.Count) & VfText(" \u015fah\u0131s, ") & CStr(occurrences) & VfText(" ge\u00e7i\u015f dizine i\u015faretlenecek (ana metin ve dipnotlar).") & vbCrLf
    If confirmedCount > 0 Then message = message & "- " & CStr(confirmedCount) & VfText(" ge\u00e7i\u015f, \u015eah\u0131s d\u00f6k\u00fcm\u00fcnde veya incelemede elle onaylad\u0131\u011f\u0131n\u0131z kimlikle dahil edildi.") & vbCrLf
    If oldMarks > 0 Then message = message & "- " & CStr(oldMarks) & VfText(" eski Vefey\u00e2t dizin i\u015fareti kald\u0131r\u0131l\u0131p yeniden olu\u015fturulacak.") & vbCrLf
    message = message & "- " & indexNote & vbCrLf & vbCrLf & VfdSkippedText(unresolved, unknown, stale)
    message = message & vbCrLf & VfText("\u0130\u015faretler gizli metindir; g\u00f6r\u00fcnen metin de\u011fi\u015fmez. Kendi XE dizin i\u015faretlerinize dokunulmaz. Ctrl+Z ile geri al\u0131nabilir.") & vbCrLf & vbCrLf & "Devam edilsin mi?"
    If VfMessage(message, vbYesNo Or vbQuestion Or vbDefaultButton2, VfText("Vefey\u00e2t: \u015fah\u0131slar dizini")) <> vbYes Then GoTo Clean
    If Documents.Count = 0 Then GoTo Clean
    If Not (ActiveDocument Is doc) Or doc.TrackRevisions Or mVfcBusy Then
        VfaWarn VfText("Belge veya Vefey\u00e2t durumu onaydan sonra de\u011fi\u015fti. Dizin olu\u015fturulmad\u0131.")
        GoTo Clean
    End If
    Set entries = CreateObject("Scripting.Dictionary")
    For Each group In people
        entries(CStr(group("Key"))) = CStr(group("Entry"))
    Next group
    mVfcBusy = True
    undoStarted = VfaBeginUndo(undoObject)
    Application.ScreenUpdating = False
    changed = True
    Call VfdRemoveMarks(doc, False, False)
    ' Reverse document order: each insertion leaves the earlier names in place.
    For i = mVfjCount To 1 Step -1
        Set item = mVfjItems(i)
        If item.Exists("VfdResolved") Then
            Set named = VfjLiveName(item)
            If Not named Is Nothing Then
                row = item("VfdRow")
                If entries.Exists(VfpPersonKey(row)) Then
                    VfdAddMark named, CStr(entries(VfpPersonKey(row)))
                    inserted = inserted + 1
                End If
            End If
        End If
    Next i
    If placeIndex Then
        Set indexField = at.Fields.Add(Range:=at, Type:=wdFieldEmpty, _
            Text:="INDEX \f ""v"" \z ""1055""", PreserveFormatting:=False)
    End If
    If Not indexField Is Nothing Then indexField.Update
    VfaEndUndo undoObject, undoStarted
    Application.ScreenUpdating = True
    VfcForgetDocument doc
    mVfcBusy = False
    message = CStr(inserted) & VfText(" ge\u00e7i\u015f, ") & CStr(people.Count) & VfText(" \u015fah\u0131s ad\u0131yla dizine i\u015faretlendi.") & vbCrLf & vbCrLf
    If indexField Is Nothing Then
        message = message & VfText("Belgede Vefey\u00e2t dizini yok. \u0130mleci dizin sayfas\u0131na koyup bu komutu yeniden \u00e7al\u0131\u015ft\u0131r\u0131n.")
    Else
        message = message & VfText("Dizin g\u00fcncellendi. Metin veya sayfa d\u00fczeni de\u011fi\u015firse dizine sa\u011f t\u0131klay\u0131p Alan\u0131 G\u00fcncelle\u015ftir (F9) se\u00e7in ya da bu komutu yeniden \u00e7al\u0131\u015ft\u0131r\u0131n.")
    End If
    If unresolved > 0 Then message = message & vbCrLf & vbCrLf & CStr(unresolved) & VfText(" belirsiz ge\u00e7i\u015f dizine girmedi; \u015eah\u0131s d\u00f6k\u00fcm\u00fcnden do\u011frulay\u0131p komutu yeniden \u00e7al\u0131\u015ft\u0131rabilirsiniz.")
    VfMessage message, vbInformation, VfText("Vefey\u00e2t: \u015fah\u0131slar dizini")
    GoTo Clean
Failed:
    errorText = Err.Description
    On Error Resume Next
    VfaEndUndo undoObject, undoStarted
    Application.ScreenUpdating = True
    If Not doc Is Nothing Then VfcForgetDocument doc
    mVfcBusy = False
    On Error GoTo 0
    If changed Then
        VfaWarn VfText("Dizin tamamlanamad\u0131: ") & Left$(errorText, 300) & vbCrLf & vbCrLf & _
            VfText("Belgede k\u0131smi dizin i\u015faretleri olabilir. Ctrl+Z ile geri al\u0131n veya Dizin i\u015faretlerini kald\u0131r komutunu kullan\u0131n.")
    Else
        VfaWarn VfText("Dizin haz\u0131rlanamad\u0131; belge de\u011fi\u015ftirilmedi: ") & Left$(errorText, 300)
    End If
Clean:
    On Error Resume Next
    If prepared Then VfjClose
    On Error GoTo 0
End Sub

Public Sub Vefeyat0341UiIndexRemove(ByVal control As Object)
    Dim doc As Word.Document, marks As Long, total As Long, removed As Long, errorText As String
    Dim undoObject As Object, undoStarted As Boolean, changed As Boolean
    If Documents.Count = 0 Then Exit Sub
    If Not VfcIdle() Then VfdBusyWarning: Exit Sub
    Set doc = ActiveDocument
    If Not VfaEditable(doc) Then Exit Sub
    If doc.TrackRevisions Then
        VfaWarn VfText("De\u011fi\u015fiklik \u0130zleme a\u00e7\u0131k. \u0130\u015flem yap\u0131lmad\u0131. \u0130zlemeyi kapat\u0131p yeniden deneyin.")
        Exit Sub
    End If
    On Error GoTo Failed
    marks = VfdRemoveMarks(doc, False, True)
    total = VfdRemoveMarks(doc, True, True)
    If total = 0 Then
        VfMessage VfText("Bu belgede Vefey\u00e2t dizin i\u015fareti veya dizini yok."), vbInformation, VfText("Vefey\u00e2t: \u015fah\u0131slar dizini")
        Exit Sub
    End If
    If VfMessage(CStr(marks) & VfText(" Vefey\u00e2t dizin i\u015fareti ve ") & CStr(total - marks) & VfText(" Vefey\u00e2t dizin alan\u0131 silinecek. Kendi dizin i\u015faretlerinize ve g\u00f6r\u00fcnen metne dokunulmaz.") & vbCrLf & vbCrLf & "Devam edilsin mi?", _
                 vbYesNo Or vbQuestion Or vbDefaultButton2, VfText("Vefey\u00e2t: \u015fah\u0131slar dizini")) <> vbYes Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    If Not (ActiveDocument Is doc) Or doc.TrackRevisions Or Not VfcIdle() Then
        VfaWarn VfText("Belge veya Vefey\u00e2t durumu onaydan sonra de\u011fi\u015fti. \u0130\u015flem yap\u0131lmad\u0131.")
        Exit Sub
    End If
    mVfcBusy = True
    undoStarted = VfaBeginUndo(undoObject)
    changed = True
    removed = VfdRemoveMarks(doc, True, False)
    VfaEndUndo undoObject, undoStarted
    VfcForgetDocument doc
    mVfcBusy = False
    VfMessage CStr(removed) & VfText(" Vefey\u00e2t dizin \u00f6\u011fesi silindi."), vbInformation, VfText("Vefey\u00e2t: \u015fah\u0131slar dizini")
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    VfaEndUndo undoObject, undoStarted
    VfcForgetDocument doc
    mVfcBusy = False
    On Error GoTo 0
    VfaWarn VfText("Dizin i\u015faretleri kald\u0131r\u0131lamad\u0131: ") & Left$(errorText, 300)
End Sub

' ---- Sahis raporu (yeni belgeye yazar; incelenen belgeyi degistirmez) ----

Private Function VfdTokenCount(ByVal value As String) As Long
    value = VefeyatNormalize(value)
    If Len(value) = 0 Then Exit Function
    VfdTokenCount = UBound(Split(value, " ")) + 1
End Function

Private Function VfdCell(ByVal value As String) As String
    Dim i As Long, code As Long, ch As String, result As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1): code = AscW(ch) And &HFFFF&
        If code < 32 Or code = 127 Then ch = " "
        result = result & ch
    Next i
    VfdCell = Trim$(VfaNormalizeWhitespace(result))
End Function

Private Function VfdPage(ByVal item As Object) As String
    Dim named As Word.Range
    On Error GoTo Unavailable
    Set named = VfjLiveName(item)
    If named Is Nothing Then Exit Function
    VfdPage = CStr(named.Information(wdActiveEndAdjustedPageNumber))
Unavailable:
End Function

Private Function VfdNotes(ByVal group As Object) As String
    Dim items As Collection, first As Object, row As Variant, fullName As String, notes As String
    Dim j As Long, firstDate As String, later As Object
    Set items = group("Items"): Set first = items(1): row = group("Row")
    fullName = VfnTextName(CStr(row(1)))
    If Len(fullName) > 0 And VfdTokenCount(CStr(first("Alias"))) < VfdTokenCount(fullName) Then
        notes = notes & VfText("\u0130lk ge\u00e7i\u015fte k\u0131sa ad; tam ad \u00f6nerisi: ") & fullName & ". "
    End If
    firstDate = VfaTrimSpaces(CStr(first("Date")))
    If Len(firstDate) = 0 And VfaValidDate(CStr(row(3))) Then
        For j = 2 To items.Count
            Set later = items(j)
            If Len(VfaTrimSpaces(CStr(later("Date")))) > 0 Then Exit For
        Next j
        If j <= items.Count Then
            notes = notes & VfText("Tarih ilk ge\u00e7i\u015fte de\u011fil, ") & CStr(j) & VfText(". ge\u00e7i\u015fte. ")
        Else
            notes = notes & VfText("\u0130lk ge\u00e7i\u015fte tarih yok. ")
        End If
    End If
    VfdNotes = Trim$(notes)
End Function

Public Sub Vefeyat0341UiPeopleReport(ByVal control As Object)
    Dim doc As Word.Document, report As Word.Document, people As Collection, group As Object, first As Object
    Dim row As Variant, lines() As String, n As Long, unresolved As Long, unknown As Long, stale As Long
    Dim flagged As Long, notes As String, target As Word.Range, tbl As Word.Table, prepared As Boolean
    Dim title As String, summary As String, errorText As String, items As Collection, confirmedCount As Long
    Dim chrono As String, startPos As Long, heading As String
    If Documents.Count = 0 Then Exit Sub
    If Not VfcIdle() Then VfdBusyWarning: Exit Sub
    Set doc = ActiveDocument
    On Error GoTo Failed
    prepared = True
    If Not VfdPrepare(doc) Then GoTo Clean
    Set people = VfdPeople(unresolved, unknown, stale, confirmedCount)
    If people.Count = 0 Then
        VfMessage VfText("Kimli\u011fi kesin olarak belirlenmi\u015f \u015fah\u0131s bulunamad\u0131.") & vbCrLf & vbCrLf & _
            VfdSkippedText(unresolved, unknown, stale), vbInformation, VfText("Vefey\u00e2t: \u015fah\u0131s raporu")
        GoTo Clean
    End If
    ReDim lines(0 To people.Count)
    lines(0) = Join(Array(VfText("\u015eah\u0131s (D\u0130A)"), VfText("\u0130lk ge\u00e7i\u015fteki yaz\u0131m"), "Yer", "Sayfa", VfText("Ge\u00e7i\u015f"), _
                          VfText("\u00d6l\u00fcm tarihi (D\u0130A)"), "Notlar", VfText("D\u0130A ba\u011flant\u0131s\u0131")), vbTab)
    For Each group In people
        n = n + 1
        Set items = group("Items"): Set first = items(1): row = group("Row")
        notes = VfdNotes(group)
        If Len(notes) > 0 Then flagged = flagged + 1
        lines(n) = Join(Array(VfdCell(VfnTextName(CStr(row(0)))), VfdCell(CStr(first("Alias"))), VfdCell(CStr(first("Location"))), _
                              VfdPage(first), CStr(group("Items").Count), VfdCell(CStr(row(3))), VfdCell(notes), VfdCell(CStr(row(5)))), vbTab)
    Next group
    title = VfText("Vefey\u00e2t \u015fah\u0131s raporu: ") & doc.Name
    summary = CStr(people.Count) & VfText(" \u015fah\u0131s, ") & CStr(VfdOccurrenceCount(people)) & VfText(" ge\u00e7i\u015f. ") & _
        CStr(flagged) & VfText(" \u015fah\u0131s i\u00e7in not var. ") & _
        VfText("Yaln\u0131z kimli\u011fi kesin ge\u00e7i\u015fler listelenir; s\u0131ra ilk ge\u00e7i\u015f s\u0131ras\u0131d\u0131r. Bu rapor tezi de\u011fi\u015ftirmez.")
    If confirmedCount > 0 Then summary = summary & " " & CStr(confirmedCount) & VfText(" ge\u00e7i\u015f elle onaylanm\u0131\u015f kimlikle dahil edildi.")
    If unresolved + unknown + stale > 0 Then summary = summary & " " & Replace(VfdSkippedText(unresolved, unknown, stale), vbCrLf, " ")
    Set report = Documents.Add
    report.PageSetup.Orientation = wdOrientLandscape
    Set target = report.Content
    target.Text = title & vbCr & summary & vbCr
    report.Paragraphs(1).Range.Font.Bold = True
    report.Paragraphs(1).Range.Font.Size = 14
    Set target = report.Content
    target.Collapse wdCollapseEnd
    target.InsertAfter Join(lines, vbCr)
    Set tbl = target.ConvertToTable(Separator:=wdSeparateByTabs, NumColumns:=8)
    tbl.Range.Font.Size = 9
    tbl.Borders.Enable = True
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(1).HeadingFormat = True
    tbl.AutoFitBehavior wdAutoFitWindow
    ' Kronoloji: second table, ordered by Gregorian death year.
    heading = VfText("Kronoloji (mil\u00e2d\u00ee \u00f6l\u00fcm y\u0131l\u0131na g\u00f6re; (?) i\u015faretli y\u0131llar tek say\u0131dan okunmu\u015ftur, hicr\u00ee olabilir)")
    chrono = VfdChronoText(people)
    startPos = report.Content.End - 1
    report.Range(startPos, startPos).InsertAfter heading & vbCr & chrono
    report.Range(startPos, startPos + Len(heading)).Font.Bold = True
    report.Range(startPos, startPos + Len(heading)).Font.Size = 12
    Set target = report.Range(startPos + Len(heading) + 1, report.Content.End)
    Set tbl = target.ConvertToTable(Separator:=wdSeparateByTabs, NumColumns:=3)
    tbl.Range.Font.Size = 9
    tbl.Range.Font.Bold = False
    tbl.Borders.Enable = True
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(1).HeadingFormat = True
    tbl.AutoFitBehavior wdAutoFitContent
    VfMessage VfText("\u015eah\u0131s raporu yeni bir belgede haz\u0131rland\u0131. Notlar s\u00fctunu ilk ge\u00e7i\u015fte k\u0131sa ad ve eksik tarih durumlar\u0131n\u0131 g\u00f6sterir. Tabloyu kopyalay\u0131p Excel'e yap\u0131\u015ft\u0131rabilirsiniz."), _
        vbInformation, VfText("Vefey\u00e2t: \u015fah\u0131s raporu")
    GoTo Clean
Failed:
    errorText = Err.Description
    VfaWarn VfText("\u015eah\u0131s raporu haz\u0131rlanamad\u0131: ") & Left$(errorText, 300)
Clean:
    On Error Resume Next
    If prepared Then VfjClose
    On Error GoTo 0
End Sub

' Native acceptance helper on the synthetic acceptance document only.
Private Sub VftIndexNative(ByVal doc As Word.Document, ByRef count As Long)
    Dim point As Word.Range, named As Word.Range, indexField As Word.Field, field As Word.Field, ownedCode As String
    doc.Content.Text = "Synthetic Person met Other Person." & vbCr
    Set point = doc.Range(Len("Synthetic Person met "), Len("Synthetic Person met "))
    point.Fields.Add Range:=point, Type:=wdFieldEmpty, Text:="XE ""Konu""", PreserveFormatting:=False
    Set named = doc.Range(0, Len("Synthetic Person"))
    VfdAddMark named, VfdEntryText("Person: Synthetic ""x""")
    VftAssert VfdRemoveMarks(doc, False, True) = 1, "one owned index mark", count
    VftAssert doc.Fields.Count = 2, "user index mark kept beside owned mark", count
    For Each field In doc.Fields
        If VfdOwnedCode(field.Code.Text, "xe") Then ownedCode = field.Code.Text
    Next field
    VftAssert InStr(1, ownedCode, "Person\: Synthetic x", vbBinaryCompare) > 0, "index entry escaped", count
    Set point = doc.Content
    point.Collapse wdCollapseEnd
    Set indexField = point.Fields.Add(Range:=point, Type:=wdFieldEmpty, Text:="INDEX \f ""v"" \z ""1055""", PreserveFormatting:=False)
    indexField.Update
    VftAssert InStr(1, indexField.Result.Text, "Person: Synthetic", vbBinaryCompare) > 0, "index lists Vefeyat entry", count
    VftAssert InStr(1, indexField.Result.Text, "Konu", vbBinaryCompare) = 0, "index excludes user entries", count
    VftAssert Not VfdIndexField(doc) Is Nothing, "owned index field found", count
    VftAssert VfdRemoveMarks(doc, True, False) = 2, "owned mark and index removed", count
    VftAssert doc.Fields.Count = 1, "user index entry preserved", count
    VftAssert InStr(1, doc.Fields(1).Code.Text, "Konu", vbBinaryCompare) > 0, "user index entry unchanged", count
    VftAssert InStr(1, doc.Content.Text, "Synthetic Person met", vbBinaryCompare) > 0, "visible text unchanged", count
    doc.Fields(1).Delete
End Sub
' ---------------------------------------------------------------------------
' 0.3.40: Dizin isaretleri (XE) tarih veya kimlik onayini etkilemez; tarih
' engellendiginde neden gosterilir; elle onaylanmis belirsiz adlar dizine ve
' rapora girer; sahis raporuna kronoloji tablosu eklenir.
' ---------------------------------------------------------------------------

Private Function VfdIsIndexMark(ByVal field As Word.Field) As Boolean
    VfdIsIndexMark = (field.Type = wdFieldIndexEntry)
End Function

' An XE field has no result; its span ends after the code and end mark.
Private Function VfdFieldEnd(ByVal field As Word.Field) As Long
    If field.Type = wdFieldIndexEntry Then
        VfdFieldEnd = field.Code.End + 1
    Else
        VfdFieldEnd = field.Result.End + 1
    End If
End Function

Private Function VdiMarkLength(ByVal span As Word.Range) As Long
    Dim field As Word.Field, total As Long
    For Each field In span.Fields
        If VfdIsIndexMark(field) Then
            If field.Code.Start - 1 >= span.Start And field.Code.End + 1 <= span.End Then
                total = total + (field.Code.End + 1) - (field.Code.Start - 1)
            End If
        End If
    Next field
    VdiMarkLength = total
End Function

' Visible text of `count` characters before (backwards) or after a point.
' With no XE marks in reach this is exactly the old MoveStart/MoveEnd text.
' With marks, the window is widened by their length and their spans are
' left out, so the visible characters are the same as without the marks.
Private Function VdiSpanText(ByVal named As Word.Range, ByVal point As Long, ByVal count As Long, ByVal backwards As Boolean) As String
    Dim span As Word.Range, story As Word.Range, piece As Word.Range, field As Word.Field
    Dim extra As Long, previous As Long, guard As Long, first As Long, last As Long, cursor As Long, text As String
    Set span = named.Duplicate
    span.SetRange point, point
    If backwards Then Call span.MoveStart(wdCharacter, -count) Else Call span.MoveEnd(wdCharacter, count)
    extra = VdiMarkLength(span)
    If extra = 0 Then
        span.TextRetrievalMode.IncludeFieldCodes = False
        VdiSpanText = span.Text
        Exit Function
    End If
    Set story = named.Document.StoryRanges(named.StoryType)
    Do
        previous = extra
        If backwards Then
            first = point - count - extra: If first < story.Start Then first = story.Start
            span.SetRange first, point
        Else
            last = point + count + extra: If last > story.End Then last = story.End
            span.SetRange point, last
        End If
        extra = VdiMarkLength(span)
        guard = guard + 1
    Loop While extra <> previous And guard < 16
    cursor = span.Start
    For Each field In span.Fields
        If VfdIsIndexMark(field) Then
            If field.Code.Start - 1 >= cursor And field.Code.End + 1 <= span.End Then
                If field.Code.Start - 1 > cursor Then
                    Set piece = span.Duplicate
                    piece.SetRange cursor, field.Code.Start - 1
                    piece.TextRetrievalMode.IncludeFieldCodes = False
                    text = text & piece.Text
                End If
                cursor = field.Code.End + 1
            End If
        End If
    Next field
    If span.End > cursor Then
        Set piece = span.Duplicate
        piece.SetRange cursor, span.End
        piece.TextRetrievalMode.IncludeFieldCodes = False
        text = text & piece.Text
    End If
    VdiSpanText = text
End Function

' Why a date cannot go after this name. Read-only; used only for messages.
Private Function VfaDateBlockReason(ByVal named As Word.Range) As String
    Dim following As Word.Range, field As Word.Field
    On Error GoTo Generic
    If VfmDatePresent(named.Document, named) Then
        VfaDateBlockReason = VfText("Tarih eklenmedi: ad\u0131n hemen arkas\u0131nda zaten bir tarih var (Vefey\u00e2t tarihi veya (\u00f6. \u2026) bi\u00e7iminde yaz\u0131lm\u0131\u015f tarih). Mevcut tarihe dokunulmad\u0131.")
        Exit Function
    End If
    Set following = named.Duplicate
    following.SetRange named.End, named.End
    Call following.MoveEnd(wdCharacter, 100)
    For Each field In following.Fields
        If field.Code.Start - 1 <= named.End + 2 And Not VfdIsIndexMark(field) Then
            VfaDateBlockReason = VfText("Tarih eklenmedi: ad\u0131n hemen arkas\u0131nda ba\u015fka bir Word alan\u0131 var (") & _
                VfdCell(Left$(Trim$(field.Code.Text), 40)) & VfText("). At\u0131f veya ba\u015fka bir eklentinin alan\u0131n\u0131n i\u00e7ine ya da \u00f6n\u00fcne tarih yaz\u0131lmaz. Gerekirse tarihi elle ekleyin.")
            Exit Function
        End If
    Next field
Generic:
    VfaDateBlockReason = VfText("Tarih eklenmedi: ad\u0131n arkas\u0131nda tarih benzeri bir ifade bulundu (\u00f6./\u00f6l./v. \u2026). Mevcut metne dokunulmad\u0131.")
End Function

' Live contextual (Vdi) confirmations by range key. Corrupt or stale records
' are ignored here; the identity tools report them separately.
Private Function VfdConfirmedMap(ByVal doc As Word.Document) As Object
    Dim result As Object, records As Object, key As Variant, values As Collection, live As Word.Range
    Set result = CreateObject("Scripting.Dictionary")
    Set VfdConfirmedMap = result
    On Error GoTo Done
    Set records = VdiRead(doc)
    For Each key In records.Keys
        Set values = records(key)
        Set live = VdiLiveRecord(doc, values)
        If Not live Is Nothing Then Set result(VfjRangeKey(live)) = values
    Next key
Done:
End Function

Private Function VfdConfirmedRow(ByVal item As Object, ByVal confirmed As Object, ByRef row As Variant) As Boolean
    Dim named As Word.Range, values As Collection, candidate As Variant, rows As Collection
    Set named = VfjLiveName(item)
    If named Is Nothing Then Exit Function
    If Not confirmed.Exists(VfjRangeKey(named)) Then Exit Function
    Set values = confirmed(VfjRangeKey(named))
    Set rows = item("Rows")
    For Each candidate In rows
        If VfpPersonKey(candidate) = CStr(values(3)) And VfqRowStamp(candidate) = CStr(values(4)) Then
            row = candidate
            VfdConfirmedRow = True
            Exit Function
        End If
    Next candidate
End Function

' ---- Kronoloji ----

' Gregorian year for ordering: the number after "/" in "505/1111", else the
' first number. Before Christ ("m.o.") is negative. 0 when unknown.
Private Function VfdChronoYear(ByVal raw As String, ByRef certain As Boolean) As Long
    Dim slash As Long, i As Long, ch As String, digits As String, start As Long, key As String
    certain = False
    slash = InStr(1, raw, "/", vbBinaryCompare)
    If slash > 0 Then start = slash + 1 Else start = 1
    For i = start To Len(raw)
        ch = Mid$(raw, i, 1)
        If ch Like "[0-9]" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        End If
    Next i
    If Len(digits) = 0 Or Len(digits) > 4 Then Exit Function
    VfdChronoYear = CLng(digits)
    key = VefeyatNormalize(raw)
    If InStr(1, " " & key & " ", " m o ", vbBinaryCompare) > 0 Then
        VfdChronoYear = -VfdChronoYear
        certain = True
    Else
        certain = (slash > 0)
    End If
End Function

Private Function VfdChronoText(ByVal people As Collection) As String
    Dim years() As Long, lines() As String, flags() As Boolean, n As Long, i As Long, j As Long
    Dim group As Object, row As Variant, certain As Boolean, y As Long, line As String, f As Boolean, result As String
    ReDim years(1 To people.Count): ReDim lines(1 To people.Count): ReDim flags(1 To people.Count)
    For Each group In people
        row = group("Row")
        y = VfdChronoYear(CStr(row(3)), certain)
        n = n + 1
        years(n) = y: flags(n) = (y <> 0)
        If y = 0 Then
            line = "?"
        ElseIf y < 0 Then
            line = VfText("m.\u00f6. ") & CStr(-y)
        Else
            line = CStr(y)
        End If
        If y <> 0 And Not certain Then line = line & " (?)"
        lines(n) = line & vbTab & VfdCell(VfnTextName(CStr(row(0)))) & vbTab & VfdCell(CStr(row(3)))
    Next group
    ' Stable insertion sort; unknown years last.
    For i = 2 To n
        y = years(i): line = lines(i): f = flags(i): j = i - 1
        Do While j >= 1
            If Not VfdChronoAfter(years(j), flags(j), y, f) Then Exit Do
            years(j + 1) = years(j): lines(j + 1) = lines(j): flags(j + 1) = flags(j)
            j = j - 1
        Loop
        years(j + 1) = y: lines(j + 1) = line: flags(j + 1) = f
    Next i
    result = VfText("Mil\u00e2d\u00ee y\u0131l") & vbTab & VfText("\u015eah\u0131s (D\u0130A)") & vbTab & VfText("\u00d6l\u00fcm tarihi (D\u0130A)")
    For i = 1 To n
        result = result & vbCr & lines(i)
    Next i
    VfdChronoText = result
End Function

Private Function VfdChronoAfter(ByVal yearA As Long, ByVal knownA As Boolean, ByVal yearB As Long, ByVal knownB As Boolean) As Boolean
    If knownA And Not knownB Then Exit Function
    If knownB And Not knownA Then VfdChronoAfter = True: Exit Function
    VfdChronoAfter = (yearA > yearB)
End Function

' Native acceptance helper: index marks must not change identity signatures
' or block a neighbouring name's date.
Private Sub VftMarkNative(ByVal doc As Word.Document, ByRef count As Long)
    Dim first As Word.Range, second As Word.Range, sigA As String, sigB As String, marks As Long
    Dim field As Word.Field, positions As New Collection
    doc.Content.Text = "Alpha Synthetic Person, Other Person beta." & vbCr
    Set first = doc.Range(6, 6 + Len("Synthetic Person"))
    Set second = doc.Range(24, 24 + Len("Other Person"))
    VftAssert first.Text = "Synthetic Person" And second.Text = "Other Person", "mark test positions", count
    sigA = VdiSignature(first): sigB = VdiSignature(second)
    VfdAddMark second, "Other"
    VfdAddMark first, "Synthetic"
    For Each field In doc.Fields
        If VfdOwnedCode(field.Code.Text, "xe") Then positions.Add field.Code.End + 1
    Next field
    VftAssert positions.Count = 2, "two marks inserted", count
    Set first = doc.Range(CLng(positions(1)), CLng(positions(1)) + Len("Synthetic Person"))
    Set second = doc.Range(CLng(positions(2)), CLng(positions(2)) + Len("Other Person"))
    VftAssert first.Text = "Synthetic Person" And second.Text = "Other Person", "names follow their marks", count
    VftAssert VdiSignature(first) = sigA, "mark keeps first identity signature", count
    VftAssert VdiSignature(second) = sigB, "mark keeps second identity signature", count
    VftAssert Not VfbAnyDate(first), "neighbour mark is not a date", count
    VftAssert Not CBool(VfjDateStateFor(first)), "neighbour mark does not block inventory date", count
    marks = VfdRemoveMarks(doc, False, False)
    VftAssert marks = 2 And doc.Fields.Count = 0, "marks removed", count
End Sub

Private Function VfjDateStateFor(ByVal named As Word.Range) As Boolean
    Dim state As Object, previous As Word.Document
    Set previous = mVfjDoc
    Set mVfjDoc = named.Document
    Set state = VfjDateState(named)
    Set mVfjDoc = previous
    VfjDateStateFor = CBool(state("Blocked"))
End Function
' ---------------------------------------------------------------------------
' 0.3.41: Ilk gecis denetimi bellek eskidiginde de calisir; kisaltma
' degisince mevcut Vefeyat tarihleri onayla yeni kisaltmaya cevrilir.
' ---------------------------------------------------------------------------

' Literal spellings worth a native Word Find before the selected name.
Private Function VfmLiteralAliases(ByVal state As Object, ByRef row As Variant, ByVal query As String) As Collection
    Dim result As New Collection, seen As Object, i As Long, value As Variant, words As Variant, shortName As String
    Dim person As String, known As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    VfmAddLiteral result, seen, VfbBareAlias(query)
    For i = 0 To 2
        VfmAddLiteral result, seen, CStr(row(i))
        VfmAddLiteral result, seen, VfbWithoutArticles(CStr(row(i)))
        VfmAddLiteral result, seen, VfnTextName(CStr(row(i)))
    Next i
    person = VfpPersonKey(row)
    If state("Aliases").Exists(person) Then
        Set known = state("Aliases")(person)
        For Each value In known.Keys
            VfmAddLiteral result, seen, CStr(value)
        Next value
    End If
    shortName = CStr(row(0))
    If InStr(shortName, ",") > 0 Then
        shortName = Trim$(Left$(shortName, InStr(shortName, ",") - 1))
    Else
        words = Split(Trim$(shortName), " ")
        If UBound(words) >= 0 Then shortName = CStr(words(UBound(words)))
    End If
    If Len(VfbKey(shortName)) >= 3 And Not VfbNoise(VfbKey(shortName)) Then VfmAddLiteral result, seen, shortName
    Set VfmLiteralAliases = result
End Function

Private Sub VfmAddLiteral(ByVal result As Collection, ByVal seen As Object, ByVal value As String)
    value = Trim$(value)
    If Len(value) < 3 Or Len(value) > 200 Then Exit Sub
    If InStr(value, ",") > 0 Or InStr(value, "^") > 0 Then Exit Sub
    If Not VfaValidAlias(value) Then Exit Sub
    If seen.Exists(value) Then Exit Sub
    seen.Add value, True
    result.Add value
End Sub

Private Function VfmLeftBoundary(ByVal hit As Word.Range) As Boolean
    Dim previous As Word.Range
    If hit.Start <= 0 Then VfmLeftBoundary = True: Exit Function
    Set previous = hit.Duplicate
    previous.SetRange hit.Start - 1, hit.Start
    VfmLeftBoundary = (Len(VefeyatNormalize(previous.Text)) = 0)
End Function

' Native Word Find in plain text (outside fields) for one spelling.
Private Function VfmFindPlain(ByVal area As Word.Range, ByVal value As String, ByVal state As Object, ByVal person As String) As Boolean
    Dim finder As Word.Range, limit As Long, guard As Long
    If area.End <= area.Start Then Exit Function
    limit = area.End
    Set finder = area.Duplicate
    With finder.Find
        .ClearFormatting
        .Text = value
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Do While finder.Find.Execute
        If finder.End > limit Then Exit Do
        guard = guard + 1
        If guard > 200 Then Exit Do
        If finder.Fields.Count = 0 Then
            If VfmLeftBoundary(finder) Then
                If VfmIdentity(state, finder, person) >= 0 Then
                    VfmFindPlain = True
                    Exit Function
                End If
            End If
        End If
        finder.Collapse wdCollapseEnd
    Loop
End Function

Private Function VfmFallbackKey(ByVal doc As Word.Document, ByRef row As Variant, ByVal target As Word.Range) As String
    VfmFallbackKey = CStr(doc.Content.End) & ":" & CStr(doc.Footnotes.Count) & ":" & CStr(target.StoryType) & ":" & _
        CStr(target.Start) & ":" & VfpPersonKey(row)
End Function

' True when Word itself finds this person's name before the target although
' the document memory found none: the memory missed pasted or remote edits.
Private Function VfmStaleEarlier(ByVal doc As Word.Document, ByVal state As Object, ByRef row As Variant, _
                                 ByVal query As String, ByVal target As Word.Range) As Boolean
    Dim note As Word.Footnote, mainLimit As Long, noteLimit As Long, found As Boolean
    Dim values As Collection, value As Variant, area As Word.Range, person As String
    On Error GoTo Unavailable
    If state.Exists("FallbackKey") Then
        If CStr(state("FallbackKey")) = VfmFallbackKey(doc, row, target) Then Exit Function
    End If
    If target.StoryType = wdMainTextStory Then
        mainLimit = target.Start
        For Each note In doc.Footnotes
            If note.Reference.Start >= mainLimit Then Exit For
            noteLimit = note.Range.End
        Next note
    ElseIf target.StoryType = wdFootnotesStory Then
        For Each note In doc.Footnotes
            If target.InRange(note.Range) Then
                mainLimit = note.Reference.Start: noteLimit = target.Start: found = True
                Exit For
            End If
        Next note
        If Not found Then Exit Function
    Else
        Exit Function
    End If
    person = VfpPersonKey(row)
    Set values = VfmLiteralAliases(state, row, query)
    For Each value In values
        If mainLimit > 0 Then
            Set area = doc.Range(0, mainLimit)
            If VfmFindPlain(area, CStr(value), state, person) Then VfmStaleEarlier = True: Exit Function
        End If
        If noteLimit > 0 Then
            Set area = doc.StoryRanges(wdFootnotesStory)
            area.SetRange area.Start, noteLimit
            If VfmFindPlain(area, CStr(value), state, person) Then VfmStaleEarlier = True: Exit Function
        End If
    Next value
Unavailable:
End Function

' Full rebuild of the document memory from inside a running command.
Private Function VfmRebuild(ByVal doc As Word.Document, ByRef state As Object) As Boolean
    Dim wasRunning As Boolean
    wasRunning = mVfbRunning
    state("Ready") = False
    mVfbRunning = False
    Vefeyat0341UiWarmMemory doc
    mVfbRunning = wasRunning
    Set state = VfmState(doc)
    If state Is Nothing Then
        VfaWarn VfText("Belgede bellekte olmayan bir \u00f6nceki ge\u00e7i\u015f bulundu, fakat belge belle\u011fi yenilenemedi. Tarih eklenmedi; Belle\u011fi yenile komutunu kullan\u0131p yeniden deneyin.")
        Exit Function
    End If
    If Not CBool(state("Ready")) Then
        VfaWarn VfText("Belgede bellekte olmayan bir \u00f6nceki ge\u00e7i\u015f bulundu, fakat belge belle\u011fi yenilenemedi. Tarih eklenmedi; Belle\u011fi yenile komutunu kullan\u0131p yeniden deneyin.")
        Exit Function
    End If
    VfmRebuild = True
End Function

' ---- Kisaltma degisince mevcut tarihler ----

' Counts (countOnly) or rewrites Vefeyat date fields to the new prefix. Only a
' field whose visible date is exactly the stored source date in one of the
' known prefixes is changed; a hand-edited date is kept and counted.
Private Function VfpReformatDates(ByVal doc As Word.Document, ByVal prefix As String, ByVal countOnly As Boolean, _
                                  ByRef kept As Long) As Long
    Dim region As Word.Range, field As Word.Field, id As String, raw As String, current As String, wanted As String
    Dim key As Variant, known As Boolean, total As Long
    kept = 0
    For Each region In VfaRegions(doc)
        For Each field In region.Fields
            id = VfaFieldBinding(doc, field)
            If Len(id) > 0 Then
                raw = VfaReadVariable(doc, VfaVariableName(id, "Raw"))
                current = field.Result.Text
                wanted = " " & VfpFormatDate(raw, prefix)
                If Len(raw) = 0 Then
                    kept = kept + 1
                ElseIf StrComp(current, wanted, vbBinaryCompare) <> 0 Then
                    known = False
                    For Each key In Array("original", "o", "ol", "v")
                        If StrComp(current, " " & VfpFormatDate(raw, CStr(key)), vbBinaryCompare) = 0 Then known = True: Exit For
                    Next key
                    If known Then
                        total = total + 1
                        If Not countOnly Then
                            field.Locked = False
                            field.Result.Text = wanted
                            field.Locked = True
                            field.ShowCodes = False
                            If StrComp(field.Result.Text, wanted, vbBinaryCompare) <> 0 Then
                                Err.Raise VfaInvalidData, "Vefeyat", "Word did not retain the reformatted date."
                            End If
                        End If
                    Else
                        kept = kept + 1
                    End If
                End If
            End If
        Next field
    Next region
    VfpReformatDates = total
End Function

Private Function VfpPrefixLabel(ByVal key As String) As String
    Select Case key
        Case "o": VfpPrefixLabel = VfText("\u00f6.")
        Case "ol": VfpPrefixLabel = VfText("\u00f6l.")
        Case "v": VfpPrefixLabel = "v."
        Case Else: VfpPrefixLabel = VfText("kaynaktaki bi\u00e7im")
    End Select
End Function

' Native acceptance helper on the synthetic acceptance document only.
Private Sub VftPrefixNative(ByVal doc As Word.Document, ByRef count As Long)
    Dim row As Variant, named As Word.Range, id As String, changed As Boolean, kept As Long, field As Word.Field, edited As Word.Field
    doc.Content.Text = "Prefix Person one. Prefix Person two." & vbCr
    row = Array("Prefix Person", "Prefix Person Full", "", "(v. 500/1106)", "Synthetic", "")
    VfaEnsureNamespace doc, changed
    Set named = doc.Range(Len("Prefix Person one. "), Len("Prefix Person one. Prefix Person"))
    id = VdiNewBinding(doc, row, "Prefix Person", changed)
    VfaInsertManaged doc, named, VfpFormatDate(CStr(row(3)), "original"), id, changed
    Set named = doc.Range(0, Len("Prefix Person"))
    id = VdiNewBinding(doc, row, "Prefix Person", changed)
    VfaInsertManaged doc, named, VfpFormatDate(CStr(row(3)), "original"), id, changed
    VftAssert VfpReformatDates(doc, "o", True, kept) = 2 And kept = 0, "prefix change counts managed dates", count
    Set edited = doc.Fields(2)
    edited.Locked = False: edited.Result.Text = " (v. 500/1107)": edited.Locked = True
    VftAssert VfpReformatDates(doc, "o", False, kept) = 1 And kept = 1, "prefix change keeps hand-edited date", count
    Set field = doc.Fields(1)
    VftAssert field.Result.Text = " " & VfpFormatDate(CStr(row(3)), "o"), "prefix change rewrites managed date", count
    VftAssert doc.Fields(2).Result.Text = " (v. 500/1107)", "hand-edited date untouched", count
    VftAssert VfpReformatDates(doc, "o", True, kept) = 0, "prefix change is idempotent", count
    Do While doc.Fields.Count > 0
        doc.Fields(1).Locked = False
        doc.Fields(1).Delete
    Loop
End Sub
