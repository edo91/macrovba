Attribute VB_Name = "Modulo2"
Sub MergeConcatCells()
Attribute MergeConcatCells.VB_Description = "Merge and Concatenate"
Attribute MergeConcatCells.VB_ProcData.VB_Invoke_Func = "m\n14"

    'Execute Macro on selection
    Set myRange = Application.Selection
    
    'Set up manual concatenation
    temp = ""
    
    For Each x In myRange
    
        'Convert all formulas in values
        If x.HasFormula Then

            x.Formula = x.Value

        End If
    
        
        'Concatenate only non missing vale
        If x.Value <> "" Then
        
            'Ifelse needed to avoid initial blanks
            If temp = "" Then
                temp = CStr(x.Value)
            Else
                temp = temp & " - " & CStr(x.Value)
            End If
        
        'Force value to missing to avoid merging message
        x.Value = ""
        
        End If
        
    Next
    
    'Merge and assign new value
    myRange.Merge
    myRange.Value = temp
        
End Sub
