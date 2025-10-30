Attribute VB_Name = "LearnerInitialsMacro"
' Word VBA Macro: Auto-add Jared M. Learner Initials
' This macro automatically triggers when any Word document opens
' Works with: new documents, existing documents (.docx, .doc), and templates (.dotm, .dotx)
' Searches for "Jared Learner" and "JL" (case-sensitive, standalone words)
' Prompts user to add middle initial if found

Private Sub Document_Open()
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim foundJaredLearner As Boolean
    Dim foundJL As Boolean
    Dim response As VbMsgBoxResult
    Dim replacementCount As Integer
    
    ' Initialize flags
    foundJaredLearner = False
    foundJL = False
    replacementCount = 0
    
    ' Check if document exists and is accessible
    If ActiveDocument Is Nothing Then
        Exit Sub
    End If
    
    ' Check if document is read-only (skip if read-only to prevent errors)
    ' Note: Templates opened for editing will not be read-only
    If ActiveDocument.ReadOnly Then
        Exit Sub
    End If
    
    ' Check if this is a template - templates are handled the same as regular documents
    ' ActiveDocument.Type will be wdTypeTemplate for .dotm/.dotx files
    ' This is fine - we can still search and replace in templates
    
    ' Check if document has content
    ' Works for all document types: new documents, existing documents, and templates
    If ActiveDocument.Content.Text = "" Or Len(Trim(ActiveDocument.Content.Text)) = 0 Then
        Exit Sub
    End If
    
    ' Search for "Jared Learner" (case-sensitive, matches standalone phrase even at boundaries)
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Text = "Jared Learner"
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        If .Execute Then
            foundJaredLearner = True
        End If
    End With
    
    ' Search for "JL" (case-sensitive, matches standalone word even at boundaries)
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Text = "JL"
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        If .Execute Then
            foundJL = True
        End If
    End With
    
    ' If either phrase is found, prompt the user
    If foundJaredLearner Or foundJL Then
        response = MsgBox("Include Mr. Learner's initial?", vbYesNo + vbQuestion, "Confirmation")
        
        If response = vbYes Then
            ' Replace "Jared Learner" with "Jared M. Learner" (preserves surrounding punctuation and spacing)
            If foundJaredLearner Then
                On Error Resume Next
                Set rng = ActiveDocument.Content
                With rng.Find
                    .ClearFormatting
                    .Text = "Jared Learner"
                    .Replacement.Text = "Jared M. Learner"
                    .MatchCase = True
                    .MatchWholeWord = True
                    .MatchWildcards = False
                    If .Execute(Replace:=wdReplaceAll) Then
                        replacementCount = replacementCount + 1
                    End If
                End With
                On Error GoTo ErrorHandler
            End If
            
            ' Replace "JL" with "JML" (preserves surrounding punctuation and spacing)
            If foundJL Then
                On Error Resume Next
                Set rng = ActiveDocument.Content
                With rng.Find
                    .ClearFormatting
                    .Text = "JL"
                    .Replacement.Text = "JML"
                    .MatchCase = True
                    .MatchWholeWord = True
                    .MatchWildcards = False
                    If .Execute(Replace:=wdReplaceAll) Then
                        replacementCount = replacementCount + 1
                    End If
                End With
                On Error GoTo ErrorHandler
            End If
            
            ' Confirmation message (only if replacements were attempted)
            If replacementCount > 0 Then
                MsgBox "Done!", vbInformation, "Confirmation"
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Suppress error messages for common non-critical errors
    If Err.Number <> 0 Then
        ' Only show error for unexpected issues (not for read-only, empty docs, etc.)
        If Err.Number <> 4198 And Err.Number <> 5941 Then
            MsgBox "An error occurred: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Error"
        End If
    End If
    Exit Sub
End Sub

