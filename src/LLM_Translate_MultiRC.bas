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
' Number of rows to process in one batch call. I would not exceed 200 for most LLMs.

' The user must specify the focus cell range or cell to translate.
' Usage - Select a single-column range (e.g. A2:A234).
' Usage(continued) - Run: Translate_SelectedColumn2Column_200 destColumn:="H", targetLang:="en", _
'   model:="your_model_name", baseURL:="your_base_url".
' Sample call: Translate_SelectedColumn2Column_200 destColumn:="H", targetLang:="en", _
'   model:="nvidia_riva-translate-4b-instruct", baseURL:="http://localhost:1234/v1/"
'Function to translate multiple columns of text at batch rate.

Option Explicit ' Add the config starting with "Private Const" to the first block in LLM_Functions.bas

' =========================
' Config
' =========================
Private Const CHUNK_ROWS As Long = 200
Private Const UDF_CHUNK_ROWS As Long = 200
Private Const ROW_DELIM As String = "<<<__ROW_DELIM__>>>"


Private gTranslateCache As Object ' Scripting.Dictionary (session cache)
' =========================
' Public macro:
'   Translate the selected single column INTO a fixed destination column (e.g., "H"),
'   processing in 200-row chunks per LLM call.
' =========================
'
' Usage:
'   1) Select a single contiguous column range, e.g., E2:E200 or E2:E20000.
'   2) Run: TranslateSelectedColumnToColumn_200 destColumn:="H", targetLang:="en", _
'          model:="your-lmstudio-model", baseUrl:="http://localhost:1234/v1"
'
Public Sub TranslateSelectedColumnToColumn_200( _
    Optional ByVal destColumn As Variant = "H", _
    Optional ByVal targetLang As String = "en", _
    Optional ByVal sourceLang As String = "", _
    Optional ByVal customPrompt As String = "", _
    Optional ByVal temperature As Variant, _
    Optional ByVal maxTokens As Variant, _
    Optional ByVal model As Variant, _
    Optional ByVal baseUrl As Variant, _
    Optional ByVal showThink As Boolean = False, _
    Optional ByVal apiKey As Variant _
)
    Dim sel As Range, ws As Worksheet
    Dim destColIndex As Long
    Dim rowsCount As Long, startRow As Long, countRows As Long
    Dim srcCol As Range, destColRange As Range
    Dim hasData As Boolean

    On Error GoTo ErrHandler

    ' --- Validate selection (must be a single contiguous column) ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a single contiguous column range (e.g., E2:E200).", vbExclamation
        Exit Sub
    End If

    Set sel = Selection
    If sel.Areas.Count > 1 Or sel.Columns.Count <> 1 Then
        MsgBox "Select exactly one continuous column block (e.g., E2:E200).", vbExclamation
        Exit Sub
    End If
    ' Selection.Areas ensures single-area selection behavior.  [2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.areas)

    ' --- Ensure we have either a target language or a custom prompt (matches the UDF logic) ---
    If Trim$(targetLang) = "" And Trim$(customPrompt) = "" Then
        targetLang = InputBox("Target language (e.g., en, ko, Japanese). Leave blank only if using a custom prompt.", _
                              "Translate Column → Specific Column", "en")
        If StrPtr(targetLang) = 0 Then Exit Sub ' Cancelled
        If Trim$(targetLang) = "" And Trim$(customPrompt) = "" Then
            MsgBox "Either targetLang or customPrompt is required.", vbExclamation
            Exit Sub
        End If
    End If

    ' --- Resolve destination column index from "H"/"AA"/8, etc. ---
    destColIndex = ResolveColumnIndex(destColumn)
    If destColIndex < 1 Or destColIndex > Columns.Count Then
        MsgBox "Invalid destination column: " & CStr(destColumn), vbCritical
        Exit Sub
    End If

    Set ws = sel.Worksheet
    Set srcCol = sel

    ' Build destination range: same row span, fixed column
    Set destColRange = ws.Range(ws.Cells(sel.Row, destColIndex), _
                                ws.Cells(sel.Row + sel.Rows.Count - 1, destColIndex))

    ' --- Warn if overwriting existing content in destination ---
    On Error Resume Next
    hasData = (Application.WorksheetFunction.CountA(destColRange) > 0) ' WorksheetFunction.CountA  [5](https://learn.microsoft.com/en-us/office/vba/api/Excel.WorksheetFunction.CountA)
    On Error GoTo ErrHandler

    If hasData Then
        If MsgBox("Destination " & destColRange.Address(0, 0) & " contains data. Overwrite?", _
                  vbQuestion + vbYesNo, "Confirm Overwrite") <> vbYes Then
            Exit Sub
        End If
    End If

    ' --- Init cache ---
    EnsureCacheReady

    ' --- Guard UI / calc state ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual   ' Manual calc during long operations  [3](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation)

    rowsCount = srcCol.Rows.Count
    startRow = 1

    Do While startRow <= rowsCount
        countRows = CHUNK_ROWS
        If startRow + countRows - 1 > rowsCount Then
            countRows = rowsCount - startRow + 1
        End If

        Dim chunkSrc As Range, chunkDst As Range
        Set chunkSrc = srcCol.Cells(startRow, 1).Resize(countRows, 1)      ' Range addressing via Cells/Resize  [1](https://learn.microsoft.com/en-us/office/vba/api/excel.range%28object%29)
        Set chunkDst = destColRange.Cells(startRow, 1).Resize(countRows, 1)

        ' Translate this 200-row chunk (or smaller at end)
        TranslateChunk chunkSrc, chunkDst, targetLang, sourceLang, customPrompt, _
                       temperature, maxTokens, model, baseUrl, showThink, apiKey

        ' Progress in the status bar; restore with False later  [4](https://learn.microsoft.com/en-us/office/vba/api/excel.application.statusbar)
        Application.StatusBar = "Translating → " & ColumnLetter(destColIndex) & _
                                " : rows " & (sel.Row + startRow - 1) & "–" & _
                                (sel.Row + startRow + countRows - 2) & " ..."
        DoEvents  ' Yield to UI (use sparingly)  [9](https://learn.microsoft.com/vi-vn/office/vba/language/reference/user-interface-help/doevents-function)

        startRow = startRow + countRows
    Loop

SafeExit:
    ' --- Restore app state, always ---
    Application.StatusBar = False           ' Restore default status text  [4](https://learn.microsoft.com/en-us/office/vba/api/excel.application.statusbar)
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic   ' restore to automatic  [3](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation)
    Exit Sub

ErrHandler:
    MsgBox "Translation stopped: " & Err.Description, vbExclamation
    Resume SafeExit
End Sub


' =========================
' Chunk translator + cache
' =========================
Private Sub TranslateChunk( _
    ByVal src As Range, _
    ByVal dst As Range, _
    ByVal targetLang As String, _
    ByVal sourceLang As String, _
    ByVal customPrompt As String, _
    ByVal temperature As Variant, _
    ByVal maxTokens As Variant, _
    ByVal model As Variant, _
    ByVal baseUrl As Variant, _
    ByVal showThink As Boolean, _
    ByVal apiKey As Variant _
)
    Dim rows As Long: rows = src.Rows.Count
    Dim inVals As Variant: inVals = src.Value2     ' Efficient bulk read/write of Range values  [1](https://learn.microsoft.com/en-us/office/vba/api/excel.range%28object%29)
    Dim outVals() As Variant: ReDim outVals(1 To rows, 1 To 1)

    Dim r As Long
    Dim sendIdx() As Long, sendText() As String
    Dim nToSend As Long: nToSend = 0

    ' Pass 1: blanks/errors → copy; cache hits → copy; others → collect for batch
    For r = 1 To rows
        Dim v As Variant: v = inVals(r, 1)

        If IsError(v) Then
            outVals(r, 1) = v
        ElseIf LenB(v) = 0 Then
            outVals(r, 1) = ""
        Else
            Dim key As String
            key = BuildCacheKey(CStr(v), targetLang, sourceLang, customPrompt, model, baseUrl)
            If gTranslateCache.Exists(key) Then
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

    ' Pass 2: batch-translate collected rows
    If nToSend > 0 Then
        Dim batchRes As Variant
        batchRes = LLM_TRANSLATE_BATCH(sendText, targetLang, sourceLang, customPrompt, _
                                       temperature, maxTokens, model, baseUrl, False, apiKey)
        Dim i As Long
        For i = 1 To nToSend
            Dim rr As Long: rr = sendIdx(i)
            outVals(rr, 1) = batchRes(i)
            Dim k As String
            k = BuildCacheKey(sendText(i), targetLang, sourceLang, customPrompt, model, baseUrl)
            If Not gTranslateCache.Exists(k) Then gTranslateCache.Add k, outVals(rr, 1)
        Next i
    End If

    dst.Value = outVals
End Sub

Private Function BuildCacheKey( _
    ByVal text As String, _
    ByVal targetLang As String, _
    ByVal sourceLang As String, _
    ByVal customPrompt As String, _
    ByVal model As Variant, _
    ByVal baseUrl As Variant _
) As String
    BuildCacheKey = text & "||" & targetLang & "||" & sourceLang & "||" & customPrompt & _
                    "||" & CStr(model) & "||" & CStr(baseUrl)
End Function

Private Sub EnsureCacheReady()
    If gTranslateCache Is Nothing Then
        On Error Resume Next
        ' Try early binding first (requires Tools→References: Microsoft Scripting Runtime)
        ' Otherwise, fall back to late binding.
        Set gTranslateCache = CreateObject("Scripting.Dictionary")  ' Dictionary object  [7](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object)
        On Error GoTo 0
        If Not gTranslateCache Is Nothing Then
            gTranslateCache.CompareMode = 1 ' vbTextCompare
        End If
    End If
End Sub

' =========================
' Batch call (1 LLM request for many lines)
' =========================
'
' lines: 1-based 1-D array of strings
' Returns: 1-based 1-D array of translations (same length)
'
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
    ' Build batch prompt with a hard delimiter contract
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
            out(i) = Trim$(parts(LBound(parts) + (i - 1)))   ' Trim leading/trailing whitespace  [8](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ltrim-rtrim-and-trim-functions)
        Next i
        LLM_TRANSLATE_BATCH = out
        Exit Function
    End If

    ' Fallback: per-line translation to preserve alignment if the model ignored the delimiter rule
    For i = 1 To n
        out(i) = LLM_TRANSLATE(CStr(lines(LBound(lines) + (i - 1))), _
                               targetLang, sourceLang, customPrompt, _
                               temperature, maxTokens, model, baseUrl, showThink, apiKey)
        DoEvents   ' keep UI responsive  [9](https://learn.microsoft.com/vi-vn/office/vba/language/reference/user-interface-help/doevents-function)
    Next i

    LLM_TRANSLATE_BATCH = out
End Function
Private Function CleanLLMText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, "```", "")
    s = Trim$(s)   ' Trim whitespace  [8](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ltrim-rtrim-and-trim-functions)
    CleanLLMText = s
End Function

' =========================
' Small helpers (Range/columns)
' =========================
Private Function ResolveColumnIndex(ByVal colRef As Variant) As Long
    ' Accepts "H", "AA", 8, etc.
    Dim s As String, i As Long, res As Long
    If IsNumeric(colRef) Then
        ResolveColumnIndex = CLng(colRef)
        Exit Function
    End If
    s = UCase$(Trim$(CStr(colRef)))
    If s = "" Then
        ResolveColumnIndex = 0
        Exit Function
    End If
    res = 0
    For i = 1 To Len(s)
        Dim ch As Integer
        ch = Asc(Mid$(s, i, 1))
        If ch < 65 Or ch > 90 Then
            ResolveColumnIndex = 0
            Exit Function
        End If
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

' =========================
' Batch UDF for translating a single-column range in chunks (≤200 rows per call). 
' Returns a 2-D array (rows x 1) suitable for spilling from the top cell.
Public Function LLM_TRANSLATE_RANGE_BATCH( _
    ByVal rng As Range, _
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
    On Error GoTo FailHard

    If rng Is Nothing Then
        LLM_TRANSLATE_RANGE_BATCH = CVErr(xlErrRef)
        Exit Function
    End If
    If rng.Columns.Count <> 1 Then
        ' Worksheet UDF: return an error if not a single column
        LLM_TRANSLATE_RANGE_BATCH = CVErr(xlErrValue)
        Exit Function
    End If

    Dim rows As Long: rows = rng.Rows.Count
    Dim inVals As Variant: inVals = rng.Value2 ' 2-D [1..rows, 1..1]
    Dim outArr() As Variant: ReDim outArr(1 To rows, 1 To 1)

    ' Flatten into a 1-D 1-based array of strings for batch calls
    Dim flat() As String: ReDim flat(1 To rows)
    Dim isErr() As Boolean: ReDim isErr(1 To rows)
    Dim r As Long, v As Variant
    For r = 1 To rows
        v = inVals(r, 1)
        If IsError(v) Then
            isErr(r) = True
            flat(r) = ""       ' placeholder; we’ll restore the error later
        ElseIf LenB(v) = 0 Then
            flat(r) = ""       ' empty row stays empty
        Else
            flat(r) = CStr(v)
        End If
    Next r

    ' Call the batch helper in 200-row chunks
    Dim startRow As Long, countRows As Long
    Dim i As Long, subOut As Variant
    Dim outFlat() As Variant: ReDim outFlat(1 To rows)

    startRow = 1
    Do While startRow <= rows
        countRows = UDF_CHUNK_ROWS
        If startRow + countRows - 1 > rows Then
            countRows = rows - startRow + 1
        End If

        ' Build subarray: 1..countRows
        Dim subLines() As String
        ReDim subLines(1 To countRows)
        For i = 1 To countRows
            subLines(i) = flat(startRow + i - 1)
        Next i

        ' One LLM request for this chunk
        subOut = LLM_TRANSLATE_BATCH( _
                     subLines, targetLang, sourceLang, customPrompt, _
                     temperature, maxTokens, model, baseUrl, showThink, apiKey)

        ' Copy chunk back
        For i = 1 To countRows
            outFlat(startRow + i - 1) = subOut(i)
        Next i

        startRow = startRow + countRows
    Loop

    ' Shape into a 2-D array (rows x 1); restore any source errors
    For r = 1 To rows
        If isErr(r) Then
            outArr(r, 1) = inVals(r, 1) ' preserve original Excel error
        Else
            outArr(r, 1) = outFlat(r)
        End If
    Next r

    LLM_TRANSLATE_RANGE_BATCH = outArr
    Exit Function

FailHard:
    LLM_TRANSLATE_RANGE_BATCH = CVErr(xlErrValue)
End Function

' Run once (manually) to register help text for the Function Wizard
Public Sub RegisterUDFHelp()
    Application.MacroOptions _
        Macro:="LLM_TRANSLATE_RANGE_BATCH", _
        Description:="Batch-translate a single-column range in chunks of up to 200 rows via your LLM endpoint.", _
        Category:="User Defined", _
        ArgumentDescriptions:=Array( _
            "Range to translate (single column)", _
            "Target language (e.g., ""en"", ""ko"", ""Japanese"")", _
            "Source language (optional)", _
            "Custom prompt (optional; overrides default translate instruction)", _
            "Temperature (optional, numeric)", _
            "maxTokens (optional, numeric)", _
            "Model name (optional; e.g., LM Studio model id)", _
            "Base URL (optional; e.g., http://localhost:1234/v1)", _
            "Show hidden reasoning (Boolean; usually False)", _
            "API key (optional)" _
        )
End Sub
' =========================