Sub PromptInstruction(Message As String)
    'Create Prompt
    ThisDrawing.Utility.Prompt (Message)
    
End Sub
Sub DeleteSelectionSets()
    'Delete All SelectionSets
    SelectionSetCount = ThisDrawing.SelectionSets.Count
    If SelectionSetCount > 0 Then
        For n = 1 To SelectionSetCount
            ThisDrawing.SelectionSets(0).Delete
        Next
    End If
    
End Sub
Function SelectElements(Message, SelectionSetName As String)
    
    'Returns SelectionSets
    DeleteSelectionSets
    PromptInstruction (Message)
    Set thisSelectionSet = ThisDrawing.SelectionSets.Add(SelectionSetName)
    With thisSelectionSet
        .SelectOnScreen
    End With
    
    Set SelectElements = thisSelectionSet
End Function
