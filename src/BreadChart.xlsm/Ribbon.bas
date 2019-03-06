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
Dim toggleButton1 As Boolean
Dim toggleButton2 As Boolean
Dim toggleButton3 As Boolean
Dim toggleButton4 As Boolean
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
      returnValue = toggleButton1
    Case "tgl2"
      returnValue = toggleButton2
    Case "tgl3"
      returnValue = toggleButton3
    Case "tgl4"
      returnValue = toggleButton4
  End Select
End Sub
Sub toggleButton_onAction(control As IRibbonControl, pressed As Boolean)
    Select Case control.ID
    Case "tgl1"
        toggleButton1 = pressed
        If toggleButton1 = True Then
            toggleButton2 = False
            toggleButton3 = False
            toggleButton4 = False
            ChartSheet.CurrentMode = mNormalProcessInput
        Else
            ChartSheet.CurrentMode = mDefault
        End If
    Case "tgl2"
        toggleButton2 = pressed
        If toggleButton2 = True Then
            toggleButton1 = False
            toggleButton3 = False
            toggleButton4 = False
            ChartSheet.CurrentMode = mJudgementProcessInput
        Else
            ChartSheet.CurrentMode = mDefault
        End If
    Case "tgl3"
        Set MainModule.FormerClickedShape = Nothing
        toggleButton3 = pressed
        If toggleButton3 = True Then
            toggleButton1 = False
            toggleButton2 = False
            toggleButton4 = False
            ChartSheet.CurrentMode = mProcessConnection
        Else
            ChartSheet.CurrentMode = mDefault
        End If
    Case "tgl4"
        toggleButton4 = pressed
        If toggleButton4 = True Then
            toggleButton1 = False
            toggleButton2 = False
            toggleButton3 = False
            ChartSheet.CurrentMode = mDeletion
        Else
            ChartSheet.CurrentMode = mDefault
        End If
    End Select
    
    If ribbonUI Is Nothing Then Set ribbonUI = GetRibbon(ConfigSheet.GetValue("RibbonPointer"))
    ribbonUI.Invalidate
End Sub

Sub ResetToggle()
    toggleButton1 = False
    toggleButton2 = False
    toggleButton3 = False
    toggleButton4 = False
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
