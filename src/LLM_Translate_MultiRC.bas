' Add Support for Multiple Row and Column Translations using your model of choice.
' This module provides functions to translate multiple columns and rows of text using an LLM API.
' It is related to the LLM_REVIEW_TRANSLATE module as originally created by ychoi-kr (Yong Choi).

Option Explicit
Const originalText as String = "Range or Cell to Translate with Review"
' The user must specify the focus cell range or cell to translate.


'Function to translate multiple columns of text
Function LLM_REVIEW_TRANSLATION(originalText As String, translatedText As String, _
                                Optional Range As String = "", _
                                Optional temperature As Variant, Optional maxTokens As Variant, _
                                Optional model As Variant, Optional baseUrl As Variant, _
                                Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant 

    'Define the model and base URL.
    Dim modelName As String, effectiveBaseUrl As String
    Call ResolveModelAndBaseUrl(modelName, effectiveBaseUrl, model, baseUrl)
    
    Dim fullPrompt As String

    If focus = "" Then
        fullPrompt = "Review the following translation for accuracy, grammar, fluency, and overall quality. " & _
                     "Provide feedback and suggest improvements if necessary." & vbCrLf & _
                     "Original text: " & originalText & vbCrLf & _
                     "Translated text: " & translatedText
    Else
'Provide the model with a specific focus area for review.
        fullPrompt = "Review the following translation with a focus on " & focus & ". " & _
                     "Provide feedback and suggest improvements if necessary." & vbCrLf & _
                     "Original text: " & originalText & vbCrLf & _
                     "Translated text: " & translatedText
    End If
    
    'Give the full prompt to the LLM for processing and show the model's 
    Dim response As String
    response = LLM_Dispatcher(fullPrompt, "", temperature, maxTokens, modelName, effectiveBaseUrl, apiKey)
    
    ' Process and return the LLM's response.
    LLM_REVIEW_TRANSLATION = ProcessLLMResponse(response, showThink)
End Function