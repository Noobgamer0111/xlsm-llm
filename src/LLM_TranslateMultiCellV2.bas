Option Explicit

'========================
' User-tunable defaults
'========================
Private Const DEFAULT_BASE_URL As String = "http://localhost:1234/v1/"  ' LM Studio default server
Private Const DEFAULT_MODEL_NAME As String = "nvidia_riva-translate-4b-instruct"               ' Change to your local model id if needed
Private Const DEFAULT_TARGET_LANG As String = "English"                  ' e.g., "Korean", "Japanese", "French"
Private Const DEFAULT_SOURCE_LANG As String = "Arabic"                         ' Optional: set if you want to lock source lang
Private Const DEFAULT_TEMPERATURE As Double = 0.2                        ' Low temperature for deterministic translations
Private Const DEFAULT_MAX_TOKENS As Long = 512                           ' Adjust as needed
Private Const DEFAULT_BATCH_SIZE As Long = 20                            ' How many rows per API call
Private Const DEFAULT_OUTPUT_COL_OFFSET As Long = 1                      ' Write to the next column by default
Private Const DEFAULT_OVERWRITE As Boolean = False                       ' Skip non-empty outputs unless True
Private Const DEFAULT_SHOW_THINK As Boolean = False                      ' Do not request/parse <think> tags

' Entry point: translate the current Selection to a different column
Public Sub TranslateSelection_WithLMStudio( _
    Optional ByVal TargetLang As String = DEFAULT_TARGET_LANG, _
    Optional ByVal OriginalLang As String = DEFAULT_SOURCE_LANG, _
    Optional ByVal OutputColOffset As Long = DEFAULT_OUTPUT_COL_OFFSET, _
    Optional ByVal BatchSize As Long = DEFAULT_BATCH_SIZE, _
    Optional ByVal Overwrite As Boolean = DEFAULT_OVERWRITE, _
    Optional ByVal ModelName As String = DEFAULT_MODEL_NAME, _
    Optional ByVal BaseUrl As String = DEFAULT_BASE_URL, _
    Optional ByVal Temperature As Double = DEFAULT_TEMPERATURE, _
    Optional ByVal MaxTokens As Long = DEFAULT_MAX_TOKENS)

    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim rowsData As Collection
    Dim i As Long, totalRows As Long, startRowIndex As Long
    Dim srcVals() As String, tgtVals() As String
    Dim wroteCount As Long, skippedCount As Long
    Dim calcState As XlCalculation
    Dim eventsState As Boolean, screenState As Boolean, statusSaved As Variant
    
    On Error GoTo CleanFail

    If Selection Is Nothing Then
        MsgBox "Please select the source cells to translate.", vbExclamation
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    Set rng = Intersect(Selection, ws.UsedRange)
    If rng Is Nothing Then
        MsgBox "Selection contains no used cells.", vbExclamation
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
            srcText = SafeCellText(r)
            If Len(Trim$(srcText)) > 0 Then
                ' Determine target cell in the output column, same row
                Dim tCell As Range
                Set tCell = ws.Cells(r.Row, outCol)
                ' Skip if output already has content and not overwriting
                If Not Overwrite Then
                    If HasNonEmptyText(tCell) Then
                        skippedCount = skippedCount + 1
                        GoTo NextCell
                    End If
                End If
                srcCells.Add r
                tgtCells.Add tCell
            End If
        End If
NextCell:
    Next r
    
    totalRows = srcCells.Count
    If totalRows = 0 Then
        Application.StatusBar = False
        MsgBox "Nothing to translate (either source empty or outputs already filled).", vbInformation
        GoTo CleanExit
    End If

    ' Process in batches using LLM_LIST from LLM_Functions.bas
    ' We instruct the model to return exactly N <item> entries, one per source line.
    startRowIndex = 1
    Do While startRowIndex <= totalRows
        Dim endRowIndex As Long
        endRowIndex = Application.WorksheetFunction.Min(startRowIndex + BatchSize - 1, totalRows)
        
        ' Build batch prompt
        Dim prompt As String
        prompt = "<s>System" & vbCrLf & _
                 "You are an expert at translating text from " & OriginalLang & " to " & targetLang & ".</s>" & vbCrLf & _
                 "<s>User" & vbCrLf & _
                 "What is the " & targetLang & " translation of the sentence: " & srcText & "?</s>" & vbCrLf & _
                 "<s>Assistant" & vbCrLf & _
                 "<br>"
        
        ' Call LLM_LIST (returns Variant array of strings)
        Dim items As Variant
        items = LLM_LIST(prompt, ModelName, BaseUrl, DEFAULT_SHOW_THINK)
        ' LLM_LIST returns either a 0-based array of strings or an error text; handle both
        If IsErrorLike(items) Then
            ' Write errors into targets for traceability, but keep moving
            Dim ei As Long
            For ei = startRowIndex To endRowIndex
                Dim errCell As Range
                Set errCell = tgtCells(ei)
                SafeWriteCell errCell, CStr(items)
            Next ei
            wroteCount = wroteCount + (endRowIndex - startRowIndex + 1)
        ElseIf IsArray(items) Then
            ' Ensure counts match; if not, attempt simple alignment
            Dim batchCount As Long
            batchCount = endRowIndex - startRowIndex + 1
            
            Dim outArr() As String
            outArr = NormalizeListOutput(items, batchCount)
            
            ' Write to sheet
            Dim k As Long, idx As Long
            idx = 0
            For k = startRowIndex To endRowIndex
                Dim wCell As Range
                Set wCell = tgtCells(k)
                SafeWriteCell wCell, outArr(idx)
                idx = idx + 1
            Next k
            wroteCount = wroteCount + batchCount
        Else
            ' Unexpected type: write the stringified result
            Dim ui As Long
            For ui = startRowIndex To endRowIndex
                SafeWriteCell tgtCells(ui), CStr(items)
            Next ui
            wroteCount = wroteCount + (endRowIndex - startRowIndex + 1)
        End If
        
        Application.StatusBar = "Translated " & wroteCount & " of " & totalRows & " (skipped: " & skippedCount & ")"
        DoEvents ' Yield UI
        startRowIndex = endRowIndex + 1
    Loop

    Application.StatusBar = "Done. Wrote: " & wroteCount & " | Skipped: " & skippedCount

CleanExit:
    ' Restore Excel state
    Application.ScreenUpdating = screenState
    Application.EnableEvents = eventsState
    Application.Calculation = calcState
    Application.StatusBar = statusSaved
    Exit Sub

CleanFail:
    ' Try to restore state on error
    On Error Resume Next
    Application.ScreenUpdating = screenState
    Application.EnableEvents = eventsState
    Application.Calculation = calcState
    Application.StatusBar = statusSaved
    On Error GoTo 0
    MsgBox "Translation failed: " & Err.Description, vbCritical
End Sub

' Build a strict batch prompt that maps one source line to one target <item>
Private Function BuildBatchPrompt(ByVal srcCells As Collection, ByVal sIdx As Long, ByVal eIdx As Long, ByVal targetLang As String, ByVal OriginalLang As String) As String
    Dim sb As String
    Dim i As Long
    
    sb = ""
    sb = sb & "Translate each of the following items from " & OriginalLang & " to " & targetLang & "." & vbCrLf
    sb = sb & "Output ONLY a <list> containing one <item> for each input, in order, with NO commentary, NO instructions, and NO repetition of the original text." & vbCrLf
    sb = sb & "For example, if the input is:" & vbCrLf
    sb = sb & "<item>Hello</item><item>World</item>" & vbCrLf
    sb = sb & "and the target language is French, output:" & vbCrLf
    sb = sb & "<list><item>Bonjour</item><item>Monde</item></list>" & vbCrLf
    sb = sb & "Now translate the following:" & vbCrLf
    sb = sb & "<list>"
    
    For i = sIdx To eIdx
        Dim srcText As String
        srcText = SafeCellText(srcCells(i))
        ' Escape any XML-looking sequences lightly to avoid breaking tags
        srcText = Replace(srcText, "<", "＜")
        srcText = Replace(srcText, ">", "＞")
        sb = sb & "<item>" & srcText & "</item>"
    Next i
    
    sb = sb & "</list>"
    BuildBatchPrompt = sb
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
        For i = 0 To batchCount - 1
            result(i) = SafeText(CStr(items))
        Next i
    End If
    NormalizeListOutput = result
End Function

' Safe cell value to string
Private Function SafeCellText(ByVal c As Range) As String
    On Error GoTo Fallback
    If c Is Nothing Then
        SafeCellText = ""
        Exit Function
    End If
    If IsError(c.Value2) Then
        SafeCellText = ""
    ElseIf IsEmpty(c.Value2) Then
        SafeCellText = ""
    Else
        SafeCellText = CStr(c.Value2)
    End If
    Exit Function
Fallback:
    SafeCellText = ""
End Function

' Test if cell has any non-empty text
Private Function HasNonEmptyText(ByVal c As Range) As Boolean
    Dim s As String
    s = SafeCellText(c)
    HasNonEmptyText = (Len(Trim$(s)) > 0)
End Function

' Safe write to cell (avoid type mismatch)
Private Sub SafeWriteCell(ByVal c As Range, ByVal text As String)
    On Error Resume Next
    c.Value2 = text
    On Error GoTo 0
End Sub

' Is a Variant "error-like" text from LLM functions
Private Function IsErrorLike(ByVal v As Variant) As Boolean
    If IsArray(v) Then
        IsErrorLike = False
    Else
        Dim s As String
        On Error Resume Next
        s = CStr(v)
        On Error GoTo 0
        If Len(s) = 0 Then
            IsErrorLike = False
        Else
            IsErrorLike = (Left$(s, 6) = "Error:")
        End If
    End If
End Function

' Ensure Variant value becomes a clean string
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

Public Sub Run_TranslateSelection_Prompt()
    Dim tgtLang As String
    Dim srcLang As String
    Dim colOffset As Long
    Dim batchSize As Long
    Dim overwriteChoice As VbMsgBoxResult
    
    ' Ask for target language
    tgtLang = InputBox("Enter the target language for translation:", _
                       "Target Language", DEFAULT_TARGET_LANG)
    If Len(Trim$(tgtLang)) = 0 Then Exit Sub

    ' Ask for source language
    srcLang = InputBox("Enter the source language (optional, leave blank to auto-detect):", _
                       "Source Language", DEFAULT_SOURCE_LANG)
    
    ' Ask for output column offset
    colOffset = CLng(InputBox( _
        "Enter the column offset from the first selected column " & vbCrLf & _
        "(e.g., 1 = next column, 2 = two columns to the right):", _
        "Output Column Offset", DEFAULT_OUTPUT_COL_OFFSET))
    
    ' Ask for batch size
    batchSize = CLng(InputBox( _
        "Enter the batch size (number of rows per API call):", _
        "Batch Size", DEFAULT_BATCH_SIZE))
    
    ' Ask whether to overwrite existing translations
    overwriteChoice = MsgBox("Overwrite existing translations in the target column?", _
                              vbYesNo + vbQuestion, "Overwrite?")
    
    ' Call the main macro with the chosen settings
    Call TranslateSelection_WithLMStudio( _
        TargetLang:=tgtLang, _
        OutputColOffset:=colOffset, _
        BatchSize:=batchSize, _
        Overwrite:=(overwriteChoice = vbYes))
End Sub
