Attribute VB_Name = "MainModule"
'DEBUG_MODEがTrueのとき、参照設定した型を使用するように
'ディレクティブで制御しておきます。Falseの時はObject型を使用します。
#Const DEBUG_MODE = False
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
        
        '指定しないとテーマフォント扱いになり、別ブック化したときにフォントが変わってはみ出す為。
        With ClickedShape.TextFrame2.TextRange.Font
            .NameComplexScript = "ＭＳ ゴシック"
            .NameFarEast = "ＭＳ ゴシック"
            .Name = "ＭＳ ゴシック"
        End With
        
    Case Mode.mJudgementProcessInput
        ProcessText = InputBox("入力してください", , ClickedShape.TextFrame2.TextRange.Text)
        If ProcessText = vbNullString Then Exit Sub
        Call ActivateProcess(ClickedShape)
        Call ChangeProcessType(ClickedShape, msoShapeFlowchartDecision)
        ClickedShape.Line.ForeColor.RGB = ConfigSheet.GetValue("JudgeLineColor")
        ClickedShape.Fill.ForeColor.RGB = ConfigSheet.GetValue("JudgeFillColor")
        ClickedShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ConfigSheet.GetValue("JudgeFontColor")
        ClickedShape.TextFrame2.TextRange.Text = ProcessText
        
        '指定しないとテーマフォント扱いになり、別ブック化したときにフォントが変わってはみ出す為。
        With ClickedShape.TextFrame2.TextRange.Font
            .NameComplexScript = "ＭＳ ゴシック"
            .NameFarEast = "ＭＳ ゴシック"
            .Name = "ＭＳ ゴシック"
        End With
        
    Case Mode.mProcessConnection
        If Not (FormerClickedShape Is Nothing Or IsShiftKeyPressed) Then
            Dim flowConnector As Shape
            Set flowConnector = ChartSheet.Shapes.AddConnector(msoConnectorElbow, 10, 10, 30, 30)
            flowConnector.Line.EndArrowheadStyle = msoArrowheadOpen
            flowConnector.Line.Weight = 1.5
            flowConnector.Line.ForeColor.RGB = ConfigSheet.GetValue("ConnectorColor")
            
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
        .Line.ForeColor.RGB = ConfigSheet.GetValue("ProcessLineColor")
        .Line.Weight = 2
        .Fill.Transparency = 0
        .Line.DashStyle = msoLineSolid
        .Fill.ForeColor.RGB = ConfigSheet.GetValue("ProcessFillColor")
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ConfigSheet.GetValue("ProcessFontColor")
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
    Call ribbon.ResetMode
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
    Dim chartName As String: chartName _
        = InputBox("チャート名を入力してください。ブランクの場合はキャンセルします。")
    If chartName = "" Then
        MsgBox "キャンセルしました。", vbInformation
        Exit Sub
    End If
    
    Dim fileName As Variant
    fileName = _
        Application.GetSaveAsFilename( _
             InitialFileName:=chartName _
           , FileFilter:="Excel ブック(*.xlsx),*.xlsx" _
           , FilterIndex:=1 _
           , Title:="保存先の指定" _
           )
#If DEBUG_MODE Then
    Dim fso As FileSystemObject
#Else
    Dim fso As Object
#End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(fileName) Then
        If vbYes <> MsgBox("ファイルは既に存在します。上書きしますか？", vbYesNo + vbDefaultButton2 + vbExclamation) Then
            MsgBox "キャンセルしました。", vbInformation
            Exit Sub
        End If
    End If
    
    'GetSaveAsFilenameは、正常な場合はStringで戻ってくるので、
    'ここはNot fileNameとはできず、Falseと比較する必要がある。
    If fileName = False Then
        MsgBox "キャンセルしました。", vbInformation
        Exit Sub
    End If
    
    '以下、なぜ一度セーブして開き直してるかというと、
    'ChartSheetをコピーした後生成されたActiveWorkbookに対し、
    'Sheets(1)等でシートにアクセスできないエラーが発生した為である。
    '他のブックでこんなことは無かったので、原因不明。
    '■エラー発生を確認した環境
    '   OS 名　Microsoft Windows 10 Home
    '   OSバージョン　10.0.17134 ビルド 17134
    '   Excel 2013(15.0.5101.1000) MSO(15.0.5101.1000) 32 ビット
    'ひょっとしてマクロが記載されたシートだと不具合が出るのかと思って
    '一度コピーしたブックをxlsx形式で保存し、閉じてから開き直したらシートに対して正常に操作できた。
    ChartSheet.Copy after:=ChartSheet
    Dim sh As Worksheet: Set sh = ActiveSheet
    sh.Name = "Sheet1"
    sh.Range("B1").Value = chartName
    Dim s As Shape
    For Each s In sh.Shapes
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
    ChartSheet.Select
    
    sh.Move
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName
    Application.DisplayAlerts = True
    ActiveWorkbook.Close
    Call ribbon.ResetMode
    
    'いくつかの環境で、ribbon.ResetModeを実行するとエラーが発生することが判明した。
    'EnableEventsをFalseにしたり、DoEventsを挟んでみたり、実行の位置を変えてみたが、
    '別名保存したファイルを閉じると改善することが判明。ただ同じプロシージャ内で開き直すとまたエラーになることが判明し、
    '現在のプロシージャとは切り離すための苦肉の策としてOnTime実行呼び出しとしている。これは成功する。
    '注意) 殆どの環境ではこんなことをしなくてもうまくいく為、本件についてこうしたらうまくいったというアドバイスは特に募集しない。
    Call Application.OnTime(Now + TimeValue("00:00:01"), "'OpenSavedFile " & """" & fileName & """'")
    
End Sub

Sub OpenSavedFile(f)
    Workbooks.Open f
End Sub
