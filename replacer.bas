Attribute VB_Name = "Module1"
Private Function dictionary()
    Dim defs(1000) As String
    
    '============================================================================================
    '   Only edit the part between these lines.
    '   Add as many as you need here, but make sure they have consecutive numbers.
    
    defs(0) = "USA us america"
    defs(1) = "CAN cn"
    defs(2) = "UTP utopia utop ut"
    
    
    '============================================================================================
    
    dictionary = defs
End Function

Sub excel_replacer()
Attribute excel_replacer.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False

    '   remember active cell
    Dim a As Range
    Set a = activeCell
    
    Dim vals
    vals = dictionary()
    
    Dim v() As String
    
    '   select the column
    a.EntireColumn.Select
    
    For i = LBound(vals) To UBound(vals)
        If vals(i) <> "" Then
                
            v = Split(vals(i), " ")
            
            For j = 1 To UBound(v)
            
                '   replace
                Selection.Replace What:=v(j), Replacement:=v(0), LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False
            
            Next j
            
            '   final replacement for case insensitivity
            Selection.Replace What:=v(0), Replacement:=v(0), LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
        End If
    Next i
    
    '   restore original selection
    a.Select
    
    Application.ScreenUpdating = True
End Sub





