Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Flowchart: Process 1")).Select
    With Selection.ShapeRange.TextFrame2.TextRange.Font
        .NameComplexScript = "ÇlÇr ÉSÉVÉbÉN"
        .NameFarEast = "ÇlÇr ÉSÉVÉbÉN"
        .Name = "ÇlÇr ÉSÉVÉbÉN"
    End With
End Sub
