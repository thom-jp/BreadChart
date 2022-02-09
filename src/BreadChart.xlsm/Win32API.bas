Attribute VB_Name = "Win32API"
#If Win64 Then
    Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As LongLong) As Integer
#Else
    #If VBA6 Or VBA5 Then
        Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer
    #Else
        Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer
    #End If
#End If
'GetAsyncKeyStateは、-32768, 1, -32767, 0 のうちいずれかの16ビット整数を返す。
'これは二進数に直したときのビットに意味がある。
'-32768: 1000 0000 0000 0000  最上位ビットが1なら、現在そのキーが押されていることを示す。
'1         : 0000 0000 0000 0001 最下位ビットが1なら、最後のGetAsyncKeyState呼び出しの後にそのキーが押されたことを示す。
'-32767: 1000 0000 0000 0001 従って、これは両方に該当することを示す。
'0         :0000 0000 0000 0000 これは、どちらでもないことを示す。
'つまり現在キーが押されているかどうかを知るには、GetAsyncKeyStateの結果を-32768のAndマスクに掛け、
'-32768になればOKということ。

Function IsShiftKeyPressed() As Boolean
    Const KEY_PRESSED = -32768      '1000 0000 0000 0000 最上位ビットが1であることを示す。
    IsShiftKeyPressed = (GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) = KEY_PRESSED
End Function
