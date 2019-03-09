Attribute VB_Name = "Ribbon"
'参考サイト
'http://fnya.cocolog-nifty.com/blog/2014/02/vba-1e63.html
'http://www.ka-net.org/ribbon/ri64.html
'https://thom.hateblo.jp/entry/2018/06/13/043244 (手前味噌)

#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As LongPtr)
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As Long)
#End If

Private Const RIBBON_TABNAME = "MyOriginalRibbon"
Dim toggleProcess As Boolean
Dim toggleJudge As Boolean
Dim toggleConnector As Boolean
Dim toggleDeletion As Boolean
Dim ribbonUI As IRibbonUI
Private flg As Boolean

#If VBA7 And Win64 Then
Private Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
  Dim p As LongPtr
#Else
Private Function GetRibbon(ByVal lRibbonPointer As Long) As Object
  Dim p As Long
#End If
  Dim ribbonObj As Object
  
  MoveMemory ribbonObj, lRibbonPointer, LenB(lRibbonPointer)
  Set GetRibbon = ribbonObj
  p = 0: MoveMemory ribbonObj, p, LenB(p) '後始末
End Function
Sub Ribbon_onLoad(ribbon As IRibbonUI)
    ribbon.ActivateTab RIBBON_TABNAME
    Set ribbonUI = ribbon
    
    ConfigSheet.LetValue "RibbonPointer", CStr(ObjPtr(ribbon))
    flg = True
End Sub
Sub toggleButton_getPressed(control As IRibbonControl, ByRef returnValue)
  Select Case control.ID
    Case "tgl1"
      returnValue = toggleProcess
    Case "tgl2"
      returnValue = toggleJudge
    Case "tgl3"
      returnValue = toggleConnector
    Case "tgl4"
      returnValue = toggleDeletion
  End Select
End Sub
Sub toggleButton_onAction(control As IRibbonControl, pressed As Boolean)
    Call ResetMode
    If pressed Then
        Select Case control.ID
        Case "tgl1"
            toggleProcess = True
            ChartSheet.CurrentMode = mNormalProcessInput
        Case "tgl2"
            toggleJudge = True
            ChartSheet.CurrentMode = mJudgementProcessInput
        Case "tgl3"
            toggleConnector = True
            ChartSheet.CurrentMode = mProcessConnection
            Set MainModule.FormerClickedShape = Nothing
        Case "tgl4"
            toggleDeletion = True
            ChartSheet.CurrentMode = mDeletion
        End Select
    End If
    
    If ribbonUI Is Nothing Then Set ribbonUI = GetRibbon(ConfigSheet.GetValue("RibbonPointer"))
    ribbonUI.Invalidate
End Sub

Sub ResetMode()
    ChartSheet.CurrentMode = mDefault
    toggleProcess = False
    toggleJudge = False
    toggleConnector = False
    toggleDeletion = False
    If ribbonUI Is Nothing Then Set ribbonUI = GetRibbon(ConfigSheet.GetValue("RibbonPointer"))
    ribbonUI.Invalidate
End Sub

Sub RibbonMacros(control As IRibbonControl)
    Application.Run control.Tag, control
End Sub

Private Sub Dummy(control As IRibbonControl)
    Application.Run control.ID
End Sub

Private Sub R_ResetToTemplate()
    MainModule.ResetToTemplate
End Sub

Private Sub R_CompleteChart()
    MainModule.CompleteChart
End Sub
