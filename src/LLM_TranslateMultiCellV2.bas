Option Explicit

'========================
' User-tunable defaults
'========================
' TODO: Sort in alphabetical order
' These private constants define default settings for the translation macro and assume that you are using LM Studio locally.
' You can modify these defaults as needed, or override them via the user prompt macro.
' Note: LM Studio must be running and the specified model must be loaded on your GPU for this to work
Private Const DEFAULT_BASE_URL As String = "http://localhost:1234/v1/"                  ' LM Studio default server
Private Const DEFAULT_MODEL_NAME As String = "nvidia_riva-translate-4b-instruct@q4_k_m" ' LM Studio default model variant with API quantization compatibility
Private Const DEFAULT_MODEL As String = "nvidia_riva-translate-4b-instruct"             ' LM Studio default model
Private Const DEFAULT_MODEL_LIST As String = "http://localhost:1234/v1/models/"         ' Get the list of available models from LM Studio
Private Const DEFAULT_TARGET_LANG As String = "English"                                 ' e.g., "Korean", "Japanese", "French"
Private Const DEFAULT_SOURCE_LANG As String = "Arabic"                                  ' Optional: set if you want to lock source lang
Private Const DEFAULT_TEMPERATURE As Double = 0.2                                       ' Low temperature for deterministic translations
Private Const DEFAULT_MAX_TOKENS As Long = 512                                          ' Adjust as needed
Private Const DEFAULT_BATCH_SIZE As Long = 20                                           ' How many rows per API call
Private Const DEFAULT_OUTPUT_COL_OFFSET As Long = 2                                     ' Write to the right of the source column by default
Private Const DEFAULT_OVERWRITE As Boolean = False                                      ' Skip non-empty outputs unless True
Private Const DEFAULT_SHOW_THINK As Boolean = False                                     ' Do not request/parse <think> tags
Private Const DEFAULT_PROMPT_TEMPLATE As String = "<s>System: You are an expert at translating text from {DEFAULT_SOURCE_LANG} to {DEFAULT_TARGET_LANG}.</s>" & _
"<s>User: What is the {DEFAULT_TARGET_LANG} translation of the sentence: {srcText}?</s> <s>Assistant\n <br>" ' Use built-in prompt if empty"
'========================
Const TargetLang As String = DEFAULT_TARGET_LANG
Const SourceLang As String = DEFAULT_SOURCE_LANG
Const OutputColOffset As Long = DEFAULT_OUTPUT_COL_OFFSET
Const BatchSize As Long = DEFAULT_BATCH_SIZE
Const Overwrite As Boolean = DEFAULT_OVERWRITE
Const ModelList As String = DEFAULT_MODEL_LIST
Const BaseUrl As String = DEFAULT_BASE_URL
Const MaxTokens As Long = DEFAULT_MAX_TOKENS
Const prompt As String = DEFAULT_PROMPT_TEMPLATE
Const Temperature As Double = DEFAULT_TEMPERATURE
Const Model As String = DEFAULT_MODEL
Const ModelName As String = DEFAULT_MODEL_NAME
Const srcText As String = "" ' Placeholder, this is the text to translate, will be set in the loop
'================================================
' Start of Office VBA macro module
'================================================

' Sanitisation/Cleaning of Unicode character logic
' --- Core translation helpers ---
' Build a simple JSON payload for the LLM API request
Private Function EscapeText(ByVal srcText As String) As String
    Dim result As String
    result = Replace(srcText, "\", "\\")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, "\", "\\" & """")
    result = Replace(result, "\\" & """", """")
    EscapeText = result
End Function

' Extract the "content" field from the JSON response
Private Function UnescapeText(ByVal srcText As String) As String
    Dim result As String
    result = Replace(srcText, "\n", vbLf)
    result = Replace(result, "\\", "\")
    result = Replace(result, """", "\""")
    UnescapeText = result
End Function

' Helper function to properly escape text for JSON
Private Function EscapeForJSON(ByVal srcText As String) As String
    Dim result As String
    result = Replace(srcText, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbTab, "\t")
    result = Replace(result, Chr(8), "\b")
    result = Replace(result, Chr(12), "\f")
    EscapeForJSON = result
End Function

' Safely write text to a cell, ignoring errors
' This function is used to avoid runtime errors when writing to cells that may be protected or invalid
Private Sub SafeWriteCell(ByVal cell As Range, ByVal value As String)
    On Error Resume Next
    If Not cell Is Nothing Then
        cell.Value = value
    End If
    On Error GoTo 0
End Sub

' Build a simple JSON payload for the LLM API request
Private Function BuildJsonPayload_Simple(ByVal ModelName As String, ByVal prompt As String, _
                                          ByVal Temperature As Double, ByVal MaxTokens As Long, _
                                          ByVal SourceLang As String, ByVal TargetLang As String, ByVal srcText As String) As String
    ' Construct the messages array for the chat-based API
    Dim systemMessage As String
    Dim userMessage As String

    systemMessage = "You are an expert at translating text from " & SourceLang & " to " & TargetLang & "."
    userMessage = "What is the " & TargetLang & " translation of the sentence: " & srcText

    Dim jsonPayload As String
    jsonPayload = "{""model"":""" & ModelName & """," & _
                  """messages"": [" & _
                  "{""role"": ""system"", ""content"": """ & EscapeText(systemMessage) & """}," & _
                  "{""role"": ""user"", ""content"": """ & EscapeText(userMessage) & """}" & _
                  "]," & _
                  """temperature"":" & CStr(Temperature) & "," & _
                  """max_tokens"":" & CStr(MaxTokens) & "}"
    Debug.Print "Generated JSON Payload: " & jsonPayload ' Log the payload to Immediate Window
    BuildJsonPayload_Simple = jsonPayload
End Function
'========================
' LLM API interaction and translation logic - Target: LM Studio
' Extract the "content" field from the JSON response
Private Function ExtractContent(ByVal response As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = """content"":\s*""([\s\S]*?)""\s*(?:,|\})"
    regEx.IgnoreCase = True
    regEx.Global = False

    Dim matches As Object
    Set matches = regEx.Execute(response)

    If matches.Count > 0 Then
        ExtractContent = Replace(matches(0).SubMatches(0), Chr(34), "")
    Else
        Debug.Print "Error: Failed to parse response. Response: " & response ' Log the full response for debugging
        ExtractContent = "Error: Failed to parse response"
    End If
End Function

' --- LLM Base function for LM Studio ---
Private Function LLM_Base(srcText As String, prompt As String, Temperature As Double, MaxTokens As Long, _
                            ModelName As String, BaseUrl As String, Model As String) As String
    If IsMissing(srcText) Or IsEmpty(srcText) Then
        LLM_Base = "Error: Source text is missing. Please provide a valid input."
        Exit Function
    End If

    If IsMissing(prompt) Or IsEmpty(prompt) Then
        prompt = DEFAULT_PROMPT_TEMPLATE
    End If

    If IsMissing(Temperature) Or IsEmpty(Temperature) Then
        Temperature = DEFAULT_TEMPERATURE
    End If
    
    If IsMissing(MaxTokens) Or IsEmpty(MaxTokens) Then
        MaxTokens = DEFAULT_MAX_TOKENS
    End If
    
    If IsMissing(ModelName) Or IsEmpty(ModelName) Then
        ModelName = DEFAULT_MODEL_NAME
    End If

    If IsMissing(BaseUrl) Or IsEmpty(BaseUrl) Then
        BaseUrl = DEFAULT_BASE_URL
    End If

    If IsMissing(Model) Or IsEmpty(Model) Then
        Model = DEFAULT_MODEL
    End If

    If Err.Number <> 0 Then
        LLM_Base = Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    ' Call the LLM API as a JSON Payload.
    Dim jsonPayload As String
    jsonPayload = BuildJsonPayload_Simple(ModelName, prompt, Temperature, MaxTokens, SourceLang, TargetLang, srcText)
    
    ' Send the HTTP request and get the response
    Dim response As String
    response = SendLLMRequest(BaseUrl, jsonPayload)

    If Left$(response, 6) = "Error:" Then
        LLM_Base = response ' Propagate error message
        Exit Function
    End If
    
    Debug.Print "Processed Response: " & response ' Log the processed response
    LLM_Base = UnescapeText(ExtractContent(response))
End Function

' --- LLM API dispatcher ---
Private Function LLM_Dispatcher(prompt As String, ModelList As String, ModelName As String, BaseUrl As String) As String
    Dim response As String
    Call Call_Model_LMStudio(ModelName, BaseUrl, Model)
    
    On Error GoTo ErrorHandler

    response = LLM_Base(srcText, prompt, Temperature, MaxTokens, ModelName, BaseUrl, Model)
    
    LLM_Dispatcher = response
    Exit Function

ErrorHandler:
    LLM_Dispatcher = "Error: " & Err.Description
    Err.Clear
End Function

' Placeholder for model call, can be expanded for different models/APIs
Private Function Call_Model_LMStudio(ModelName As String, BaseUrl As String, Model As String)
End Function
' --- LLM API HTTP request ---
' Send LLM API request and return the response from LM Studio over HTTP.
Private Function SendLLMRequest(ByVal BaseUrl As String, ByVal jsonPayload As String) As String
    Dim http As Object
    Dim url As String
    url = BaseUrl & "chat/completions"
    Set http = CreateObject("MSXML2.XMLHTTP")

    On Error GoTo ErrorHandler

    Debug.Print "Sending HTTP Request to URL: " & url ' Log the URL
    Debug.Print "Payload: " & jsonPayload ' Log the payload

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonPayload

    If http.Status = 200 Then
        Debug.Print "Response: " & http.responseText ' Log the response
        SendLLMRequest = http.responseText
    Else
        Dim serverMsg As String
        serverMsg = http.responseText
        If Len(serverMsg) > 0 Then
            SendLLMRequest = "Error: " & http.Status & " " & http.statusText & " - " & serverMsg
        Else
            SendLLMRequest = "Error: " & http.Status & " " & http.statusText
        End If
        Debug.Print "Error Response: " & SendLLMRequest ' Log the error response
        Exit Function
    End If

ErrorHandler:
    If Err.Number <> 0 Then
        Dim errMsg As String
        errMsg = "Error: " & Err.Number & " - " & Err.Description
        Debug.Print errMsg ' Log the error message
        SendLLMRequest = errMsg
    End If
    On Error GoTo 0
    Set http = Nothing
End Function

' --- Output normalization ---
Private Function SafeText(ByVal v As Variant) As String
    On Error GoTo Fallback
    If IsError(v) Or IsEmpty(v) Then
        SafeText = ""
    Else
        SafeText = CStr(v)
    End If
    Exit Function
Fallback:
    SafeText = ""
End Function

' Safely write text to a non-empty cell
Private Function HasNonEmptyText(ByVal cell As Range) As Boolean
    If Not cell Is Nothing Then
        HasNonEmptyText = Len(Trim$(cell.Value)) > 0
    Else
        HasNonEmptyText = False
    End If
End Function

' Normalize/trim the LLM_LIST output to exactly batchCount items
Private Function NormalizeListOutput(ByVal items As Variant, ByVal batchCount As Long) As String()
    Dim result() As String
    Dim n As Long, i As Long
    If IsArray(items) Then
        n = UBound(items) - LBound(items) + 1
        ReDim result(0 To batchCount - 1)
        ' Copy what we have
        For i = 0 To Application.WorksheetFunction.Min(n, batchCount) - 1
            result(i) = SafeText(items(LBound(items) + i))
        Next i
        ' If fewer returned, pad with empty strings
        For i = n To batchCount - 1
            result(i) = ""
        Next i
    Else
        ReDim result(0 To batchCount - 1)
        For i = 0 To batchSize - 1
            result(i) = SafeText(CStr(items))
        Next i
    End If
    NormalizeListOutput = result
End Function

' --- Main macro ---
' Entry point: translate the current Selection to a different column
Public Sub TranslateSelection_WithLMStudio()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim rowsData As Collection
    Dim i As Long, totalRows As Long, startRowIndex As Long
    Dim srcVals() As String, tgtVals() As String
    Dim wroteCount As Long, skippedCount As Long
    Dim calcState As XlCalculation
    Dim eventsState As Boolean, screenState As Boolean, statusSaved As Variant

    Debug.Print "Starting TranslateSelection_WithLMStudio macro..." ' Log start of macro

    On Error GoTo CleanFail

    If Selection Is Nothing Then
        MsgBox "Please select the source cells to translate.", vbExclamation
        Debug.Print "No selection made. Exiting macro." ' Log no selection
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set rng = Intersect(Selection, ws.UsedRange)
    If rng Is Nothing Then
        MsgBox "Selection contains no used cells.", vbExclamation
        Debug.Print "Selection contains no used cells. Exiting macro." ' Log empty selection
        Exit Sub
    End If
    If rng.Columns.Count > 1 Then
        MsgBox "Please select a single column or the first column of a multi-column selection.", vbExclamation
        Debug.Print "Invalid selection: More than one column selected. Exiting macro." ' Log invalid selection
        Exit Sub
    End If

    ' Guard Excel state
    screenState = Application.ScreenUpdating
    eventsState = Application.EnableEvents
    calcState = Application.Calculation
    statusSaved = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True

    ' Collect row-wise non-empty items and their target cells
    Dim srcCells As Collection
    Dim tgtCells As Collection
    Set srcCells = New Collection
    Set tgtCells = New Collection

    Dim firstCol As Long: firstCol = rng.Columns(1).Column
    Dim outCol As Long: outCol = firstCol + OutputColOffset

    Dim r As Range
    For Each r In rng.Cells
        ' Process only single-column or multi-column selections row-wise using the first column
        If r.Column = firstCol Then
            Dim srcText As String
            srcText = SafeText(r.Value)
            If Len(Trim$(srcText)) > 0 Then
                ' Determine target cell in the output column, same row
                Dim tCell As Range
                Set tCell = ws.Cells(r.Row, outCol)
                ' Skip if output already has content and not overwriting
                If Not Overwrite Then
                    If HasNonEmptyText(tCell) Then
                        skippedCount = skippedCount + 1
                        Debug.Print "Skipping row " & r.Row & ": Target cell already has content." ' Log skipped row
                        GoTo NextCell
                    End If
                End If
                srcCells.Add r
                tgtCells.Add tCell
                Debug.Print "Added row " & r.Row & " to translation queue." ' Log added row
            End If
        End If
NextCell:
    Next r

    totalRows = srcCells.Count
    If totalRows = 0 Then
        Application.StatusBar = False
        MsgBox "Nothing to translate (either source empty or outputs already filled).", vbInformation
        Debug.Print "No rows to translate. Exiting macro." ' Log no rows to translate
        GoTo CleanExit
    End If

    ' Process each cell individually using the recommended prompt format
    For i = 1 To totalRows
        srcText = SafeText(srcCells(i).Value)
        Debug.Print "Translating row " & srcCells(i).Row & ": " & srcText ' Log source text

        ' Translate the source text using the prompt template
        Dim translation As Variant
        translation = LLM_Base(srcText, prompt, Temperature, MaxTokens, ModelName, BaseUrl, Model)

        ' Handle errors and write output
        If IsError(translation) Then
            Debug.Print "Error translating row " & srcCells(i).Row & ": " & CStr(translation) ' Log error
            SafeWriteCell tgtCells(i), CStr(translation)
        ElseIf IsArray(translation) Then
            Debug.Print "Translation (array) for row " & srcCells(i).Row & ": " & SafeText(translation(LBound(translation))) ' Log array translation
            SafeWriteCell tgtCells(i), SafeText(translation(LBound(translation)))
        Else
            Debug.Print "Translation for row " & srcCells(i).Row & ": " & SafeText(translation) ' Log translation
            SafeWriteCell tgtCells(i), SafeText(translation)
        End If
        wroteCount = wroteCount + 1
        Application.StatusBar = "Translated " & wroteCount & " of " & totalRows & " (skipped: " & skippedCount & ")"
        DoEvents
    Next i
    Application.StatusBar = "Done. Wrote: " & wroteCount & " | Skipped: " & skippedCount

CleanExit:
    ' Restore Excel state
    Application.ScreenUpdating = screenState
    Application.EnableEvents = eventsState
    Application.Calculation = calcState
    Application.StatusBar = statusSaved
    Debug.Print "Macro completed successfully. Wrote: " & wroteCount & " | Skipped: " & skippedCount ' Log completion
    Exit Sub

CleanFail:
    ' Try to restore state on error
    On Error Resume Next
    Application.ScreenUpdating = screenState
    Application.EnableEvents = eventsState
    Application.Calculation = calcState
    Application.StatusBar = statusSaved
    Debug.Print "Macro failed with error: " & Err.Description ' Log error
    On Error GoTo 0
    MsgBox "Translation failed: " & Err.Description, vbCritical
End Sub

' --- User prompt macro ---
' Public Sub ()
    'Dim tgtLang As String
    'Dim srcLang As String
    'Dim colOffset As Long
    'Dim batchSize As Long
    'Dim overwriteChoice As VbMsgBoxResult
    
    ' Ask for target language
    'tgtLang = InputBox("Enter the target language for translation:", _
                       "Target Language", TargetLang)
    'If Len(Trim$(tgtLang)) = 0 Then Exit Sub

    ' Ask for source language
    'srcLang = InputBox("Enter the source language (optional, leave blank to auto-detect):", _
                       "Source Language", SourceLang)
    
    ' Ask for output column offset
    'colOffset = CLng(InputBox( _
        "Enter the column offset from the first selected column " & vbCrLf & _
        "(e.g., 1 = next column, 2 = two columns to the right):", _
        "Output Column Offset", OutputColOffset))
    
    ' Ask for batch size
    'batchSize = CLng(InputBox( _
        "Enter the batch size (number of rows per API call):", _
        "Batch Size", BatchSize))
    
    ' Ask whether to overwrite existing translations
    'overwriteChoice = MsgBox("Overwrite existing translations in the target column?", _
                              vbYesNo + vbQuestion, "Overwrite?")
    
    ' Call the main macro with the chosen settings
    'TranslateSelection_WithLMStudio
'End Sub"

' --- Utility (commented) ---
' Private Function ProcessLLMResponse(response As String, Optional showThink As Boolean = False) As Variant
    ' ...existing code...
'End Function

'Private Function RemoveLeadingLineBreaks(ByVal text As String) As String
    ' ...existing code...
'End Function