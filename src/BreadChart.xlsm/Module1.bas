Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Flowchart: Process 1")).Select
    With Selection.ShapeRange.TextFrame2.TextRange.Font
        .NameComplexScript = "�l�r �S�V�b�N"
        .NameFarEast = "�l�r �S�V�b�N"
        .Name = "�l�r �S�V�b�N"
    End With
End Sub
