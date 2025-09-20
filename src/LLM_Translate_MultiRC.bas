' Add Support for Multiple Row and Column Translations using your model of choice.
' This module provides functions to translate multiple columns and rows of text using an LLM API.
' It is related to the LLM_REVIEW_TRANSLATE module as originally created by ychoi-kr (Yong Choi).
' Intended to run for LM Studio users with OpenAI-compatible API endpoints.

' This module adds the ability to translate multiple rows of a single column in a single API call, which is more efficient.
' It also includes a simple in-memory cache to avoid redundant translations during a session.

' Hardware used for testing:
' Laptop: Lenovo ThinkBook 16 G7 IML (21MS)
' CPU: Intel Core Ultra 7 125H (14 cores, 28 threads)
' RAM: 32GB DDR5 SODIMM
' GPU: NVIDIA GeForce RTX 2060 Super OC (ASUS DUAL-RTX2060S-O8G-EVO, see: https://www.techpowerup.com/gpu-specs/asus-dual-rtx-2060-super-evo-oc.b7178)
' OS: Windows 11 Pro
' GPU Dock: EXP GDC TH3P4G3, Thunderbolt 3 to PCIe 3.0 x16
' Docking Station: Dell WD22TB4 - 180W
' LLM: Local deployment of nvidia_riva-translate-4b-instruct on LM Studio (https://huggingface.co/tensorblock/nvidia_Riva-Translate-4B-Instruct-GGUF)

' You can adjust this value based on your LLM's capabilities and hardware performance.
' Number of rows to process in one batch call. I would not exceed 200 for most GPUs.

'Function to translate multiple columns of text at batch rate.

Option Explicit ' Add the config starting with "Private Const" to "session cache" to the first block in LLM_Functions.bas

' =========================
' DEFAULTS Config
' =========================
Private Const CHUNK_ROWS As Long = 20
Private Const UDF_CHUNK_ROWS As Long = 20
Private Const ROW_DELIM As String = "<<<__ROW_DELIM__>>>"
Private Const DEFAULT_BASE_URL As String   = "http://localhost:1234/v1"   ' LM Studio default
Private Const DEFAULT_MODEL    As String   = "your-lmstudio-model-name"   ' e.g., "qwen2.5-7b-instruct"
Private Const DEFAULT_CHUNK    As Long     = 200                          ' rows per request
Private Const DEFAULT_DEST_COL As String   = "H"                          ' destination column
Private Const DEFAULT_TLANG    As String   = "en"                         ' target language
' =========================
' It will ask you for a destination column (any column), target language, and optional model/base URL.

' Optional: show a live status in Excel's status bar while running
Private Const SHOW_STATUS As Boolean = True

' Session cache for duplicates in one run
Private gTranslateCache As Object   ' Scripting.Dictionary

' =========================
' MANUAL-RUN MACRO
' =========================
' 1) Select a single contiguous column (e.g., E2:E20000)
' 2) Alt+F8 → BatchTranslate_WithDefaults → Run
' 3) Accept defaults (press Enter) or type overrides for this run
Public Sub BatchTranslate_WithDefaults()
    Dim sel As Range, ws As Worksheet, srcCol As Range, destColRange As Range
    Dim destColIndex As Long, rowsCount As Long, startRow As Long, countRows As Long
    Dim prevCalc As XlCalculation, prevSU As Boolean, prevEE As Boolean

    Dim inDestCol As String, inChunk As String, inModel As String, inBase As String
    Dim inTL As String, inSL As String, inPrompt As String, inTemp As String, inMaxTok As String
    Dim targetLang As String, sourceLang As String, customPrompt As String
    Dim temperature As Variant, maxTokens As Variant, model As Variant, baseUrl As Variant
    Dim hasData As Boolean

    On Error GoTo ErrHandler

    ' --- Validate selection ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a single contiguous column (e.g., E2:E20000).", vbExclamation: Exit Sub
    End If
    Set sel = Selection
    If sel.Areas.Count > 1 Or sel.Columns.Count <> 1 Then
        MsgBox "Please select exactly one column block (e.g., E2:E20000).", vbExclamation: Exit Sub
    End If

    Set ws = sel.Worksheet
    Set srcCol = sel

    ' --- Gather overrides with defaults pre-filled (press Enter = use default) ---
    inDestCol = InputBox("Destination column (letter or number)" & vbCrLf & _
                         "Press Enter to use default: " & DEFAULT_DEST_COL, _
                         "Destination Column", DEFAULT_DEST_COL)
    If StrPtr(inDestCol) = 0 Then inDestCol = DEFAULT_DEST_COL ' Cancel => default
    destColIndex = ResolveColumnIndex(inDestCol)
    If destColIndex < 1 Or destColIndex > Columns.Count Then
        MsgBox "Invalid destination column.", vbCritical: Exit Sub
    End If

    inChunk = InputBox("Chunk size = max rows per request" & vbCrLf & _
                       "Press Enter to use default: " & DEFAULT_CHUNK, _
                       "Chunk Size", CStr(DEFAULT_CHUNK))
    If StrPtr(inChunk) = 0 Or Trim$(inChunk) = "" Then
        countRows = DEFAULT_CHUNK
    Else
        countRows = CLng(inChunk)
        If countRows < 1 Then countRows = DEFAULT_CHUNK
    End If

    inTL = InputBox("Target language (e.g., en, ko, ja)" & vbCrLf & _
                    "Press Enter to use default: " & DEFAULT_TLANG, _
                    "Target Language", DEFAULT_TLANG)
    If StrPtr(inTL) = 0 Or Trim$(inTL) = "" Then
        targetLang = DEFAULT_TLANG
    Else
        targetLang = Trim$(inTL)
    End If

    inPrompt = InputBox("Custom prompt (optional). If set, it overrides default translation instruction." & vbCrLf & _
                        "Press Enter to leave empty.", "Custom Prompt", "")
    If StrPtr(inPrompt) = 0 Then inPrompt = ""
    customPrompt = inPrompt

    inSL = InputBox("Source language (optional – blank = auto)", _
                    "Source Language", "")
    If StrPtr(inSL) = 0 Then inSL = ""
    sourceLang = inSL

    inTemp = InputBox("Temperature (optional numeric – blank = provider default)", _
                      "Temperature", "")
    If StrPtr(inTemp) = 0 Or Trim$(inTemp) = "" Then
        temperature = Empty
    Else
        temperature = CDbl(inTemp)
    End If

    inMaxTok = InputBox("Max tokens (optional numeric – blank = default)", _
                        "Max Tokens", "")
    If StrPtr(inMaxTok) = 0 Or Trim$(inMaxTok) = "" Then
        maxTokens = Empty
    Else
        maxTokens = CLng(inMaxTok)
    End If

    inModel = InputBox("Model (optional – blank = default)" & vbCrLf & _
                       "Press Enter to use default: " & DEFAULT_MODEL, _
                       "Model", DEFAULT_MODEL)
    If StrPtr(inModel) = 0 Or Trim$(inModel) = "" Then
        model = DEFAULT_MODEL
    Else
        model = Trim$(inModel)
    End If

    inBase = InputBox("Base URL (optional – blank = default)" & vbCrLf & _
                      "Press Enter to use default: " & DEFAULT_BASE_URL, _
                      "Base URL", DEFAULT_BASE_URL)
    If StrPtr(inBase) = 0 Or Trim$(inBase) = "" Then
        baseUrl = DEFAULT_BASE_URL
    Else
        baseUrl = Trim$(inBase)
    End If

' --- helpers: tolerant numeric parsing (place once at module bottom) ---
' Returns Empty if blank or non-numeric. Accepts either "." or "," as decimal.
Private Function ParseDoubleOpt(ByVal s As String) As Variant
    Dim ds As String: ds = Application.International(xlDecimalSeparator)
    s = Trim$(s)
    If s = "" Then Exit Function
    ' normalize alternate decimal separator
    If ds = "." Then s = Replace(s, ",", ".") Else s = Replace(s, ".", ",")
    If IsNumeric(s) Then ParseDoubleOpt = CDbl(s)
End Function

Private Function ParseLongOpt(ByVal s As String) As Variant
    s = Trim$(s)
    If s = "" Then Exit Function
    If IsNumeric(s) Then ParseLongOpt = CLng(s)
End Function

' --- use the helpers in your prompts ---
inTemp = InputBox("Temperature (optional numeric – blank = provider default)", "Temperature", "")
temperature = ParseDoubleOpt(inTemp)     ' Empty if invalid or blank

inMaxTok = InputBox("Max tokens (optional numeric – blank = default)", "Max Tokens", "")
maxTokens = ParseLongOpt(inMaxTok)       ' Empty if invalid or blank

inChunk = InputBox("Chunk size = max rows per request" & vbCrLf & _
                   "Press Enter to use default: " & DEFAULT_CHUNK, "Chunk Size", CStr(DEFAULT_CHUNK))
If IsEmpty(ParseLongOpt(inChunk)) Then
    countRows = DEFAULT_CHUNK
Else
    countRows = ParseLongOpt(inChunk)
    If countRows < 1 Then countRows = DEFAULT_CHUNK
End If
    ' --- Build destination range, confirm overwrite ---
Set destColRange = ws.Range(ws.Cells(sel.Row, destColIndex), _
    ws.Cells(sel.Row + sel.Rows.Count - 1, destColIndex))

    On Error Resume Next
    hasData = (Application.WorksheetFunction.CountA(destColRange) > 0)
    On Error GoTo ErrHandler
    If hasData Then
        If MsgBox("Destination " & destColRange.Address(0, 0) & " contains data. Overwrite?", _
                  vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        destColRange.ClearContents
    End If

    ' --- Init cache and guard Excel state ---
    EnsureCacheReady
    prevSU = Application.ScreenUpdating
    prevEE = Application.EnableEvents
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual     ' keep Excel from recalcing during the run  (MS docs)

    ' --- Process in chunks ---
    Dim totalRows As Long: totalRows = srcCol.Rows.Count
    Dim chunkSize As Long: chunkSize = countRows
    Dim startRowIdx As Long: startRowIdx = 1

    Do While startRowIdx <= totalRows
        Dim thisCount As Long: thisCount = chunkSize
        If startRowIdx + thisCount - 1 > totalRows Then thisCount = totalRows - startRowIdx + 1

        Dim chunkSrc As Range, chunkDst As Range
        Set chunkSrc = srcCol.Cells(startRowIdx, 1).Resize(thisCount, 1)
        Set chunkDst = destColRange.Cells(startRowIdx, 1).Resize(thisCount, 1)

        TranslateChunk chunkSrc, chunkDst, targetLang, sourceLang, customPrompt, _
                       temperature, maxTokens, model, baseUrl, False, Empty

        If SHOW_STATUS Then
            Application.StatusBar = "Translating → " & ColumnLetter(destColIndex) & _
                                    " Rows " & (sel.Row + startRowIdx - 1) & _
                                    "–" & (sel.Row + startRowIdx + thisCount - 2)
        End If
        ' Optional: add a rare DoEvents if you want ESC cancel responsiveness; otherwise omit for speed.
        ' DoEvents

        startRowIdx = startRowIdx + thisCount
    Loop

SafeExit:
    Application.StatusBar = False
    Application.ScreenUpdating = prevSU
    Application.EnableEvents = prevEE
    Application.Calculation = prevCalc
    Set gTranslateCache = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Translation stopped: " & Err.Description, vbExclamation
    Resume SafeExit
End Sub

' =========================
' CHUNK TRANSLATION + CACHE
' =========================
Private Sub TranslateChunk( _
    ByVal src As Range, ByVal dst As Range, _
    ByVal targetLang As String, ByVal sourceLang As String, ByVal customPrompt As String, _
    ByVal temperature As Variant, ByVal maxTokens As Variant, _
    ByVal model As Variant, ByVal baseUrl As Variant, _
    ByVal showThink As Boolean, ByVal apiKey As Variant)

    Dim rows As Long: rows = src.Rows.Count
    Dim inVals As Variant: inVals = src.Value2                  ' bulk read (fast)  (MS Range docs)
    Dim outVals() As Variant: ReDim outVals(1 To rows, 1 To 1)

    Dim sendIdx() As Long, sendText() As String
    Dim nToSend As Long, r As Long

    For r = 1 To rows
        Dim v As Variant: v = inVals(r, 1)
        If IsError(v) Then
            outVals(r, 1) = v
        ElseIf LenB(v) = 0 Then
            outVals(r, 1) = ""
        Else
            Dim key As String: key = BuildCacheKey(CStr(v), targetLang, sourceLang, customPrompt, model, baseUrl)
            If Not gTranslateCache Is Nothing And gTranslateCache.Exists(key) Then
                outVals(r, 1) = gTranslateCache(key)
            Else
                nToSend = nToSend + 1
                ReDim Preserve sendIdx(1 To nToSend)
                ReDim Preserve sendText(1 To nToSend)
                sendIdx(nToSend) = r
                sendText(nToSend) = CStr(v)
            End If
        End If
    Next r

    If nToSend > 0 Then
        Dim batchRes As Variant, i As Long
        batchRes = LLM_TRANSLATE_BATCH(sendText, targetLang, sourceLang, customPrompt, _
                                       temperature, maxTokens, model, baseUrl, False, apiKey)

        For i = 1 To nToSend
            Dim rr As Long: rr = sendIdx(i)
            outVals(rr, 1) = batchRes(i)
            Dim k As String: k = BuildCacheKey(sendText(i), targetLang, sourceLang, customPrompt, model, baseUrl)
            If Not gTranslateCache.Exists(k) Then gTranslateCache.Add k, outVals(rr, 1)
        Next i
    End If
    dst.Value = outVals      ' bulk write VALUES → no formulas, no spills
End Sub

Private Sub EnsureCacheReady()
    If gTranslateCache Is Nothing Then
        Set gTranslateCache = CreateObject("Scripting.Dictionary")
        gTranslateCache.CompareMode = 1 ' vbTextCompare
    End If
End Sub

Private Function BuildCacheKey(ByVal text As String, ByVal targetLang As String, _
                               ByVal sourceLang As String, ByVal customPrompt As String, _
                               ByVal model As Variant, ByVal baseUrl As Variant) As String
    BuildCacheKey = text & "||" & targetLang & "||" & sourceLang & "||" & customPrompt & _
                    "||" & CStr(model) & "||" & CStr(baseUrl)
End Function

Public Function LLM_TRANSLATE_BATCH( _
    ByVal lines As Variant, _
    Optional ByVal targetLang As String = "", _
    Optional ByVal sourceLang As String = "", _
    Optional ByVal customPrompt As String = "", _
    Optional ByVal temperature As Variant, _
    Optional ByVal maxTokens As Variant, _
    Optional ByVal model As Variant, _
    Optional ByVal baseUrl As Variant, _
    Optional ByVal showThink As Boolean = False, _
    Optional ByVal apiKey As Variant _
) As Variant
    Dim modelName As String, effectiveBaseUrl As String
    Call ResolveModelAndBaseUrl(modelName, effectiveBaseUrl, model, baseUrl)

    Dim n As Long
    n = UBound(lines) - LBound(lines) + 1
    If n <= 0 Then
        Dim emptyOut() As Variant
        ReDim emptyOut(1 To 0)
        LLM_TRANSLATE_BATCH = emptyOut
        Exit Function
    End If

    Dim sep As String: sep = ROW_DELIM
    Dim hdr As String
    If customPrompt <> "" Then
        hdr = customPrompt
    ElseIf sourceLang <> "" Then
        hdr = "Translate each item from " & sourceLang & " to " & targetLang & "."
    Else
        hdr = "Translate each item to " & targetLang & "."
    End If

    Dim rules As String
    rules = "There are exactly " & CStr(n) & " items." & vbCrLf & _
            "Items are separated by the delimiter: " & sep & vbCrLf & _
            "Respond with ONLY the translations in the same order, joined with the same delimiter (" & sep & "). " & _
            "No numbering, no extra text. Empty inputs must produce empty outputs."

    Dim inputBlock As String, i As Long
    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then inputBlock = inputBlock & vbCrLf & sep & vbCrLf
        inputBlock = inputBlock & CStr(lines(i))
    Next i

    Dim finalPrompt As String
    finalPrompt = hdr & vbCrLf & vbCrLf & rules & vbCrLf & vbCrLf & "INPUT:" & vbCrLf & inputBlock

    Dim response As String
    response = LLM_Dispatcher(finalPrompt, "", temperature, maxTokens, modelName, effectiveBaseUrl, apiKey)

    Dim body As String
    body = ProcessLLMResponse(response, False)
    body = CleanLLMText(body)

    Dim parts As Variant
    parts = Split(body, sep)

    Dim out() As Variant
    ReDim out(1 To n)

    If UBound(parts) - LBound(parts) + 1 = n Then
        For i = 1 To n
            out(i) = Trim$(parts(LBound(parts) + (i - 1)))
        Next i
        LLM_TRANSLATE_BATCH = out
        Exit Function
    End If

    ' Fallback: per-line if the delimiter spec wasn't followed
    For i = 1 To n
        out(i) = LLM_TRANSLATE(CStr(lines(LBound(lines) + (i - 1))), _
                               targetLang, sourceLang, customPrompt, _
                               temperature, maxTokens, model, baseUrl, showThink, apiKey)
        'DoEvents    ' optional / sparing use
    Next i

    LLM_TRANSLATE_BATCH = out
End Function

' =========================
' Small helpers
' =========================
Private Function ResolveColumnIndex(ByVal colRef As Variant) As Long
    Dim s As String, i As Long, res As Long
    If IsNumeric(colRef) Then ResolveColumnIndex = CLng(colRef): Exit Function
    s = UCase$(Trim$(CStr(colRef))): If s = "" Then Exit Function
    For i = 1 To Len(s)
        Dim ch As Integer: ch = Asc(Mid$(s, i, 1))
        If ch < 65 Or ch > 90 Then Exit Function
        res = res * 26 + (ch - 64)
    Next i
    ResolveColumnIndex = res
End Function

Private Function ColumnLetter(ByVal colIndex As Long) As String
    Dim q As Long, r As Long, s As String
    q = colIndex
    Do While q > 0
        r = (q - 1) Mod 26
        s = Chr$(65 + r) & s
        q = (q - 1) \ 26
    Loop
    ColumnLetter = s
End Function

Private Function CleanLLMText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, "```", "")
    s = Trim$(s)
    CleanLLMText = s
End Function