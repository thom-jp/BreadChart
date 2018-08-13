Attribute VB_Name = "MainModule"
Option Explicit
Public Enum Mode
    mDefault = 0
    mDeletion = 2
    mNormalProcessInput = 3
    mProcessConnection = 5
    mJudgementProcessInput = 4
End Enum

Public Enum Direction
    North = 1
    West = 2
    South = 3
    East = 4
End Enum
Public FormerClickedShape As Shape

Public Function OppositeDirection(D As Direction) As Direction
    Dim ret As Direction
    If D < 3 Then
        ret = D + 2
    Else
        ret = D - 2
    End If
    OppositeDirection = ret
End Function

Sub Click()
    Dim ClickedShape As Shape: Set ClickedShape _
        = ChartSheet.Shapes(Application.Caller)

    If ChartSheet.CurrentMode <> mProcessConnection Then Set FormerClickedShape = Nothing
    
    Dim ProcessText As String
    
    Select Case ChartSheet.CurrentMode
    
    Case Mode.mDeletion
        Call DeactivateProcess(ClickedShape)
        
    Case Mode.mNormalProcessInput
        ProcessText = InputBox("入力してください", , ClickedShape.TextFrame2.TextRange.Text)
        If ProcessText = vbNullString Then Exit Sub
        Call ActivateProcess(ClickedShape)
        Call ChangeProcessType(ClickedShape, msoShapeFlowchartProcess)
        ClickedShape.TextFrame2.TextRange.Text = ProcessText
        
    Case Mode.mJudgementProcessInput
        ProcessText = InputBox("入力してください", , ClickedShape.TextFrame2.TextRange.Text)
        If ProcessText = vbNullString Then Exit Sub
        Call ActivateProcess(ClickedShape)
        Call ChangeProcessType(ClickedShape, msoShapeFlowchartDecision)
        ClickedShape.Line.ForeColor.RGB = rgbWhiteSmoke
        ClickedShape.Fill.ForeColor.RGB = rgbWhiteSmoke
        ClickedShape.TextFrame2.TextRange.Text = ProcessText
        
    Case Mode.mProcessConnection
        If Not FormerClickedShape Is Nothing Then
        
            Dim flowConnector As Shape
            Set flowConnector = ChartSheet.Shapes.AddConnector(msoConnectorElbow, 10, 10, 30, 30)
            flowConnector.Line.EndArrowheadStyle = msoArrowheadOpen
            flowConnector.Line.Weight = 1.5
            flowConnector.Line.ForeColor.RGB = vbBlack
            
            Dim ConnectDirection As Direction
            ConnectDirection = DetectDirection(FormerClickedShape, ClickedShape)
            
            flowConnector.ConnectorFormat.BeginConnect _
                ConnectedShape:=FormerClickedShape, _
                ConnectionSite:=ConnectDirection
                
            flowConnector.ConnectorFormat.EndConnect _
                ConnectedShape:=ClickedShape, _
                ConnectionSite:=OppositeDirection(ConnectDirection)
        End If
        Set FormerClickedShape = ClickedShape
        
    Case Else
        MsgBox "モードを選択してください。", vbInformation
        
    End Select
End Sub

Sub ActivateProcess(process As Shape)
    With process
        .Line.ForeColor.RGB = vbBlack
        .Line.Weight = 2
        .Fill.Transparency = 0
        .Line.DashStyle = msoLineSolid
        .Fill.ForeColor.RGB = vbWhite
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbBlack
    End With
End Sub

Sub DeactivateProcess(process As Shape)
    With process
        .AutoShapeType = msoShapeFlowchartProcess
        .Line.Weight = 0.25
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Fill.Transparency = 1
        .Line.DashStyle = msoLineDash
        With .TextFrame2
            .TextRange.Delete
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        With .TextFrame
            .HorizontalOverflow = xlOartHorizontalOverflowOverflow
            .VerticalOverflow = xlOartVerticalOverflowOverflow
        End With
    End With
End Sub

Sub ResetToTemplate()
    If vbOK <> MsgBox( _
        "現在の内容をクリアして、ひな形に戻しますか?", _
        vbOKCancel + vbExclamation, _
        "確認") Then Exit Sub

    Set FormerClickedShape = Nothing
    Call DeleteAllChart
    Const ProcessWidth As Double = 100
    Const ProcessHeight As Double = 40
    
    Dim r As Range
    For Each r In ChartSheet.Range("BreadRange")
        Dim sh As Shape
        Set sh = ChartSheet.Shapes.AddShape( _
            Type:=msoShapeFlowchartProcess, _
            Left:=r.Left + (r.Width - ProcessWidth) / 2, _
            Top:=r.Top + (r.Height - ProcessHeight) / 2, _
            Width:=ProcessWidth, _
            Height:=ProcessHeight)
        Call DeactivateProcess(sh)
        sh.OnAction = "Click"
    Next
    Call ChartSheet.ClearButtonState
    ChartSheet.CurrentMode = mDefault
End Sub

Sub DeleteAllChart()
    Dim s As Shape
    For Each s In ActiveSheet.Shapes
        If s.Type <> msoFormControl Then    'For exclude buttons
            s.Delete
        End If
    Next
End Sub

Function DetectDirection(s1 As Shape, s2 As Shape) As Direction
    Dim s1HCenter: s1HCenter = s1.Left + (s1.Width / 2)
    Dim s2HCenter: s2HCenter = s2.Left + (s2.Width / 2)
    Dim HDistance: HDistance = s1HCenter - s2HCenter
    
    Dim s1VCenter: s1VCenter = s1.Top + (s1.Height / 2)
    Dim s2VCenter: s2VCenter = s2.Top + (s2.Height / 2)
    Dim VDistance: VDistance = s1VCenter - s2VCenter

    If Abs(VDistance) < (s1.Height + s2.Height) / 2 Then
        If HDistance > 0 Then
            DetectDirection = West
        Else
            DetectDirection = East
        End If
    Else
        If VDistance > 0 Then
            DetectDirection = North
        Else
            DetectDirection = South
        End If
    End If
End Function

Sub ChangeProcessType(TargetShape As Shape, T As MsoAutoShapeType)
    Dim processConnectors As New Collection
    Dim s As Shape
    
    '現在TargetShapeに接続されたコネクタを一覧化しておく
    For Each s In TargetShape.Parent.Shapes
        If s.Connector Then
            If s.ConnectorFormat.BeginConnected Then
                If s.ConnectorFormat.BeginConnectedShape Is TargetShape Then
                    processConnectors.Add _
                        Array(s, s.ConnectorFormat.BeginConnectionSite, True) 'True=Begin
                End If
            End If
            If s.ConnectorFormat.EndConnected Then
                If s.ConnectorFormat.EndConnectedShape Is TargetShape Then
                    processConnectors.Add _
                        Array(s, s.ConnectorFormat.EndConnectionSite, False) 'False=End
                End If
            End If
        End If
    Next
    
    'シェイプタイプを切り替える。
    TargetShape.AutoShapeType = T
    'このとき、コネクタの接続が外れる
    
    '一覧に登録されたコネクタをTargetShapeに再接続する
    Dim c As Variant
    For Each c In processConnectors
        Dim processConnector As Shape: Set processConnector = c(0)
        If c(2) Then 'True=Begin, False=End
            processConnector.ConnectorFormat.BeginConnect TargetShape, c(1)
        Else
            processConnector.ConnectorFormat.EndConnect TargetShape, c(1)
        End If
    Next
End Sub

Sub CompleteChart()
    If vbOK <> MsgBox( _
        "一度完成させると編集モードに戻せないのでご注意ください。" & vbNewLine & "続行しますか?", _
        vbOKCancel + vbExclamation, _
        "確認") Then Exit Sub
    Dim s As Shape
    Dim Arr() As String
    ReDim Arr(0)
    For Each s In ChartSheet.Shapes
        If s.Type <> msoFormControl Then
            s.OnAction = vbNullString
            If s.Connector Then
                If s.Height < 2 Or s.Width < 2 Then
                    s.ConnectorFormat.Type = msoConnectorStraight
                End If
            End If
            If s.Fill.Transparency = 1 And s.AutoShapeType = msoShapeFlowchartProcess Then
                s.Delete
            End If
        End If
    Next
    Call ChartSheet.ClearButtonState
End Sub




