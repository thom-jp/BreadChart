Attribute VB_Name = "MainModule"
'DEBUG_MODE��True�̂Ƃ��A�Q�Ɛݒ肵���^���g�p����悤��
'�f�B���N�e�B�u�Ő��䂵�Ă����܂��BFalse�̎���Object�^���g�p���܂��B
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
        ProcessText = InputBox("���͂��Ă�������", , ClickedShape.TextFrame2.TextRange.Text)
        If ProcessText = vbNullString Then Exit Sub
        Call ActivateProcess(ClickedShape)
        Call ChangeProcessType(ClickedShape, msoShapeFlowchartProcess)
        ClickedShape.TextFrame2.TextRange.Text = ProcessText
        
        '�w�肵�Ȃ��ƃe�[�}�t�H���g�����ɂȂ�A�ʃu�b�N�������Ƃ��Ƀt�H���g���ς���Ă͂ݏo���ׁB
        With ClickedShape.TextFrame2.TextRange.Font
            .NameComplexScript = "�l�r �S�V�b�N"
            .NameFarEast = "�l�r �S�V�b�N"
            .Name = "�l�r �S�V�b�N"
        End With
        
    Case Mode.mJudgementProcessInput
        ProcessText = InputBox("���͂��Ă�������", , ClickedShape.TextFrame2.TextRange.Text)
        If ProcessText = vbNullString Then Exit Sub
        Call ActivateProcess(ClickedShape)
        Call ChangeProcessType(ClickedShape, msoShapeFlowchartDecision)
        ClickedShape.Line.ForeColor.RGB = ConfigSheet.GetValue("JudgeLineColor")
        ClickedShape.Fill.ForeColor.RGB = ConfigSheet.GetValue("JudgeFillColor")
        ClickedShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ConfigSheet.GetValue("JudgeFontColor")
        ClickedShape.TextFrame2.TextRange.Text = ProcessText
        
        '�w�肵�Ȃ��ƃe�[�}�t�H���g�����ɂȂ�A�ʃu�b�N�������Ƃ��Ƀt�H���g���ς���Ă͂ݏo���ׁB
        With ClickedShape.TextFrame2.TextRange.Font
            .NameComplexScript = "�l�r �S�V�b�N"
            .NameFarEast = "�l�r �S�V�b�N"
            .Name = "�l�r �S�V�b�N"
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
        MsgBox "���[�h��I�����Ă��������B", vbInformation
        
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
        "���݂̓��e���N���A���āA�ЂȌ`�ɖ߂��܂���?", _
        vbOKCancel + vbExclamation, _
        "�m�F") Then Exit Sub

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
    
    '����TargetShape�ɐڑ����ꂽ�R�l�N�^���ꗗ�����Ă���
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
    
    '�V�F�C�v�^�C�v��؂�ւ���B
    TargetShape.AutoShapeType = T
    '���̂Ƃ��A�R�l�N�^�̐ڑ����O���
    
    '�ꗗ�ɓo�^���ꂽ�R�l�N�^��TargetShape�ɍĐڑ�����
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
        = InputBox("�`���[�g������͂��Ă��������B�u�����N�̏ꍇ�̓L�����Z�����܂��B")
    If chartName = "" Then
        MsgBox "�L�����Z�����܂����B", vbInformation
        Exit Sub
    End If
    
    Dim fileName As Variant
    fileName = _
        Application.GetSaveAsFilename( _
             InitialFileName:=chartName _
           , FileFilter:="Excel �u�b�N(*.xlsx),*.xlsx" _
           , FilterIndex:=1 _
           , Title:="�ۑ���̎w��" _
           )
#If DEBUG_MODE Then
    Dim fso As FileSystemObject
#Else
    Dim fso As Object
#End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(fileName) Then
        If vbYes <> MsgBox("�t�@�C���͊��ɑ��݂��܂��B�㏑�����܂����H", vbYesNo + vbDefaultButton2 + vbExclamation) Then
            MsgBox "�L�����Z�����܂����B", vbInformation
            Exit Sub
        End If
    End If
    
    'GetSaveAsFilename�́A����ȏꍇ��String�Ŗ߂��Ă���̂ŁA
    '������Not fileName�Ƃ͂ł����AFalse�Ɣ�r����K�v������B
    If fileName = False Then
        MsgBox "�L�����Z�����܂����B", vbInformation
        Exit Sub
    End If
    
    '�ȉ��A�Ȃ���x�Z�[�u���ĊJ�������Ă邩�Ƃ����ƁA
    'ChartSheet���R�s�[�����㐶�����ꂽActiveWorkbook�ɑ΂��A
    'Sheets(1)���ŃV�[�g�ɃA�N�Z�X�ł��Ȃ��G���[�����������ׂł���B
    '���̃u�b�N�ł���Ȃ��Ƃ͖��������̂ŁA�����s���B
    '���G���[�������m�F������
    '   OS ���@Microsoft Windows 10 Home
    '   OS�o�[�W�����@10.0.17134 �r���h 17134
    '   Excel 2013(15.0.5101.1000) MSO(15.0.5101.1000) 32 �r�b�g
    '�Ђ���Ƃ��ă}�N�����L�ڂ��ꂽ�V�[�g���ƕs����o��̂��Ǝv����
    '��x�R�s�[�����u�b�N��xlsx�`���ŕۑ����A���Ă���J����������V�[�g�ɑ΂��Đ���ɑ���ł����B
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
    
    '�������̊��ŁAribbon.ResetMode�����s����ƃG���[���������邱�Ƃ����������B
    'EnableEvents��False�ɂ�����ADoEvents������ł݂���A���s�̈ʒu��ς��Ă݂����A
    '�ʖ��ۑ������t�@�C�������Ɖ��P���邱�Ƃ������B���������v���V�[�W�����ŊJ�������Ƃ܂��G���[�ɂȂ邱�Ƃ��������A
    '���݂̃v���V�[�W���Ƃ͐؂藣�����߂̋���̍�Ƃ���OnTime���s�Ăяo���Ƃ��Ă���B����͐�������B
    '����) �w�ǂ̊��ł͂���Ȃ��Ƃ����Ȃ��Ă����܂������ׁA�{���ɂ��Ă��������炤�܂��������Ƃ����A�h�o�C�X�͓��ɕ�W���Ȃ��B
    Call Application.OnTime(Now + TimeValue("00:00:01"), "'OpenSavedFile " & """" & fileName & """'")
    
End Sub

Sub OpenSavedFile(f)
    Workbooks.Open f
End Sub
