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

Option Explicit ' Add the config starting with "Private Const" to "session cache" to the first block in LLM_Functions.bas

' =========================
' Config
' =========================
Private Const CHUNK_ROWS As Long = 20
Private Const UDF_CHUNK_ROWS As Long = 20
Private Const ROW_DELIM As String = "<<<__ROW_DELIM__>>>"
Private gTranslateCache As Object ' Scripting.Dictionary (session cache)


' =========================
' PUBLIC MACRO (manual run only)
' =========================
' Select a single contiguous column (e.g., E2:E200), then run this macro.
' It will ask you for a destination column (any column), target language, and optional model/base URL.
Public Sub BatchTranslate_ToChosenColumn()
    Dim sel As Range, ws As Worksheet
    Dim srcCol As Range, destColRange As Range
    Dim rowsCount As Long, startRow As Long, countRows As Long

    Dim destColIndex As Long
    Dim targetLang As String, sourceLang As String, customPrompt As String
    Dim temperature As Variant, maxTokens As Variant
    Dim model As Variant, baseUrl As Variant
    Dim showThink As Boolean, apiKey As Variant

    Dim hasData As Boolean
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean, prevEnableEvents As Boolean
    Dim userInput As Variant

    On Error GoTo ErrHandler

    ' ---- Validate selection ----
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a single contiguous column range (e.g., E2:E200).", vbExclamation
        Exit Sub
    End If

    Set sel = Selection
    If sel.Areas.Count > 1 Or sel.ws.Columns.Count <> 1 Then
        MsgBox "Select exactly one continuous column block (e.g., E2:E20).", vbExclamation
        Exit Sub
    End If

    Set ws = sel.Worksheet
    Set srcCol = sel

    ' ---- Ask for destination column (click or type) ----
    ' First try: let the user CLICK any cell in the destination column
    userInput = Application.InputBox( _
                    Prompt:="Choose destination column:" & vbCrLf & _
                            "① Click any cell in the desired column, OR" & vbCrLf & _
                            "② Press Cancel, then you'll be asked to type a column letter (e.g., H).", _
                    Title:="Select Destination Column", _
                    Type:=8) ' Range

    If VarType(userInput) = vbBoolean And userInput = False Then
        ' User pressed Cancel → ask for a column letter or number
        Dim sCol As String
        sCol = InputBox("Enter destination column (e.g., H or 8). Leave blank to cancel.", _
                        "Destination Column", "H")
        If StrPtr(sCol) = 0 Or Trim$(sCol) = "" Then Exit Sub
        destColIndex = ResolveColumnIndex(sCol)
    Else
        ' They clicked a cell/range; take the first cell's column
        destColIndex = userInput.Cells(1, 1).Column
    End If

    If destColIndex < 1 Or destColIndex > ws.Columns.Count Then
        MsgBox "Invalid destination column.", vbCritical
        Exit Sub
    End If

    ' ---- Ask for translation options ----
    targetLang = InputBox("Target language (e.g., en, ko, Japanese)." & vbCrLf & _
                          "(Leave blank only if using a custom prompt.)", _
                          "Target Language", "en")
    If StrPtr(targetLang) = 0 Then Exit Sub ' canceled

    customPrompt = InputBox("Custom prompt (optional). If provided, it overrides the default translate instruction." & vbCrLf & _
                            "Leave blank to use the default translate prompt.", _
                            "Custom Prompt (optional)", "")
    If StrPtr(customPrompt) = 0 Then Exit Sub ' canceled

    If Trim$(targetLang) = "" And Trim$(customPrompt) = "" Then
        MsgBox "Either target language or custom prompt is required.", vbExclamation
        Exit Sub
    End If

    ' Optional: source language
    sourceLang = InputBox("Source language (optional). Leave blank for auto.", _
                          "Source Language (optional)", "")
    If StrPtr(sourceLang) = 0 Then Exit Sub

    ' Optional: temperature / max tokens
    Dim t As String, mt As String
    t = InputBox("Temperature (optional numeric). Leave blank to use provider default.", _
                 "Temperature (optional)", "")
    If StrPtr(t) = 0 Then Exit Sub
    temperature = IIf(Trim$(t) = "", Empty, CDbl(t))

    mt = InputBox("Max tokens (optional numeric). Leave blank for provider default.", _
                  "Max Tokens (optional)", "")
    If StrPtr(mt) = 0 Then Exit Sub
    If Trim$(mt) = "" Then
        maxTokens = Empty
    Else
        maxTokens = CLng(mt)
    End If

    ' Optional: model / base URL (use your ResolveModelAndBaseUrl defaults if left empty)
    model = InputBox("Model name (optional). Leave blank to use your default.", _
                     "Model (optional)", "")
    If StrPtr(model) = 0 Then Exit Sub

    baseUrl = InputBox("Base URL (optional). Leave blank to use your default." & vbCrLf & _
                       "LM Studio example: http://localhost:1234/v1", _
                       "Base URL (optional)", "")
    If StrPtr(baseUrl) = 0 Then Exit Sub

    ' Optional: showThink (hidden reasoning) & key (if your dispatcher uses them)
    showThink = False ' keep default; change to True if you want to request it
    apiKey = Empty    ' LM Studio typically ignores API key; set if needed

    ' ---- Build destination range in the same rows as the selection ----
    Set destColRange = ws.Range(ws.Cells(sel.Row, destColIndex), _
                                ws.Cells(sel.Row + sel.Rows.Count - 1, destColIndex))

    ' Warn if overwriting existing content
    On Error Resume Next
    hasData = (Application.WorksheetFunction.CountA(destColRange) > 0)
    On Error GoTo ErrHandler
    If hasData Then
        If MsgBox("Destination " & destColRange.Address(0, 0) & " contains data. Overwrite?", _
                  vbQuestion + vbYesNo, "Confirm Overwrite") <> vbYes Then
            Exit Sub
        End If
    End If

    ' ---- Prepare cache & protect app state ----
    EnsureCacheReady

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ---- Chunk through the selection ----
    rowsCount = srcCol.Rows.Count
    startRow = 1

    Do While startRow <= rowsCount
        countRows = CHUNK_ROWS
        If startRow + countRows - 1 > rowsCount Then
            countRows = rowsCount - startRow + 1
        End If

        Dim chunkSrc As Range, chunkDst As Range
        Set chunkSrc = srcCol.Cells(startRow, 1).Resize(countRows, 1)
        Set chunkDst = destColRange.Cells(startRow, 1).Resize(countRows, 1)

        TranslateChunk chunkSrc, chunkDst, targetLang, sourceLang, customPrompt, _
                       temperature, maxTokens, model, baseUrl, showThink, apiKey

        Application.StatusBar = "Translating → " & ColumnLetter(destColIndex) & _
                                " : rows " & (sel.Row + startRow - 1) & "–" & _
                                (sel.Row + startRow + countRows - 2) & " ..."
        DoEvents

        startRow = startRow + countRows
    Loop
SafeExit:
    Application.StatusBar = False
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    Exit Sub

ErrHandler:
    MsgBox "Translation stopped: " & Err.Description, vbExclamation
    Resume SafeExit
End Sub


' =========================
' CHUNK TRANSLATION + CACHE
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
    Dim inVals As Variant: inVals = src.Value2
    Dim outVals() As Variant: ReDim outVals(1 To rows, 1 To 1)

    Dim r As Long
    Dim sendIdx() As Long, sendText() As String
    Dim nToSend As Long: nToSend = 0

    ' First pass: blanks/errors from src are preserved; cache hits are re-used
    For r = 1 To rows
        Dim v As Variant: v = inVals(r, 1)
        If IsError(v) Then
            outVals(r, 1) = v
        ElseIf LenB(v) = 0 Then
            outVals(r, 1) = ""
        Else
            Dim key As String
            key = BuildCacheKey(CStr(v), targetLang, sourceLang, customPrompt, model, baseUrl)
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
    ' Batch-translate remaining
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
            If Not gTranslateCache Is Nothing Then
                If Not gTranslateCache.Exists(k) Then gTranslateCache.Add k, outVals(rr, 1)
            End If
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
        Set gTranslateCache = CreateObject("Scripting.Dictionary") ' Late binding; no reference required
        On Error GoTo 0
        If Not gTranslateCache Is Nothing Then gTranslateCache.CompareMode = 1 ' vbTextCompare
    End If
End Sub

' =========================
' BATCH CALL (1 LLM REQUEST FOR MANY LINES)
' =========================
' lines: 1-based 1-D array of strings
' Returns: 1-based 1-D array of translations (same length)
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

    ' Fallback: per-line translation if the delimiter protocol wasn't followed
    For i = 1 To n
        out(i) = LLM_TRANSLATE(CStr(lines(LBound(lines) + (i - 1))), _
                               targetLang, sourceLang, customPrompt, _
                               temperature, maxTokens, model, baseUrl, showThink, apiKey)
        DoEvents
    Next i

    LLM_TRANSLATE_BATCH = out
End Function

Private Function CleanLLMText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, "```", "")
    s = Trim$(s)
    CleanLLMText = s
End Function

' =========================
' Helpers for column parsing
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
    ByVal destColumn As Variant, _
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
    If rng.ws.Columns.Count <> 1 Then
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


' New worksheet UDF that adds a Destination Column argument.
' IMPORTANT: UDFs cannot write to other cells—Excel will place the returned array
' where the formula is entered. We validate the placement using Application.Caller. (See MS docs)
' https://learn.microsoft.com/en-us/office/vba/api/excel.application.caller
Public Function LLM_TRANSLATE_RANGE_TO( _
    ByVal rng As Range, _
    ByVal destColumn As Variant, _
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

    ' Basic validation: single column range
    If rng Is Nothing Or rng.ws.Columns.Count <> 1 Then
        LLM_TRANSLATE_RANGE_TO = CVErr(xlErrValue)
        Exit Function
    End If

    ' Validate destination column against where the formula is entered
    Dim destColIndex As Long
    destColIndex = ResolveColumnIndex(destColumn)
    If destColIndex < 1 Or destColIndex > ws.Columns.Count Then
        LLM_TRANSLATE_RANGE_TO = CVErr(xlErrValue)
        Exit Function
    End If

    ' Application.Caller returns the calling Range for UDFs in cells (or arrays).
    ' If the caller is a multi-cell spill, we only need its top-left cell’s column.
    Dim callerRange As Variant
    callerRange = Application.Caller ' could be Range / String / Error (per docs)
    If TypeName(callerRange) = "Range" Then
        Dim callerCol As Long
        callerCol = callerRange.Columns(1).Column
        If callerCol <> destColIndex Then
            ' Placed in the "wrong" column—return an error to signal misplacement.
            ' (UDFs cannot write elsewhere by design.)
            LLM_TRANSLATE_RANGE_TO = CVErr(xlErrValue)
            Exit Function
        End If
    End If

    ' Forward work to the batch translator (chunks of 200 inside it).
    ' NOTE: UDF returns a 2-D array (rows x 1) that will spill from the formula cell.
    LLM_TRANSLATE_RANGE_TO = LLM_TRANSLATE_RANGE_BATCH( _
                                rng, targetLang, sourceLang, customPrompt, _
                                temperature, maxTokens, model, baseUrl, showThink, apiKey)
    Exit Function
FailHard:
    LLM_TRANSLATE_RANGE_TO = CVErr(xlErrValue)
End Function

' =========================
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
    
 '--- LLM_TRANSLATE_RANGE_TO (new; includes Destination Column) ---
    Application.MacroOptions _
        Macro:="LLM_TRANSLATE_RANGE_TO", _
        Description:="Batch-translate a single-column range and *choose* a destination column. NOTE: Excel writes UDF results where the formula is entered; place the formula in the destination column's top cell. The function validates that placement.", _
        Category:="User Defined", _
        ArgumentDescriptions:=Array( _
            "Range to translate (single column)", _
            "Destination column letter or index (e.g., ""H"" or 8). Place formula in the first cell of that column; UDFs cannot write to other cells.", _
            "Target language (e.g., ""en"", ""ko"", ""Japanese"")", _
            "Source language (optional)", _
            "Custom prompt (optional; overrides default translate instruction)", _
            "Temperature (optional, numeric)", _
            "maxTokens (optional, numeric)", _
            "Model name (optional; e.g., LM Studio model id)", _
            "Base URL (optional; e.g., http://localhost:1234/v1)", _
            "Show hidden reasoning (Boolean; usually FALSE)", _
            "API key (optional)" _
        )
End Sub

