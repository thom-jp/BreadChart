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
'GetAsyncKeyState�́A-32768, 1, -32767, 0 �̂��������ꂩ��16�r�b�g������Ԃ��B
'����͓�i���ɒ������Ƃ��̃r�b�g�ɈӖ�������B
'-32768: 1000 0000 0000 0000  �ŏ�ʃr�b�g��1�Ȃ�A���݂��̃L�[��������Ă��邱�Ƃ������B
'1         : 0000 0000 0000 0001 �ŉ��ʃr�b�g��1�Ȃ�A�Ō��GetAsyncKeyState�Ăяo���̌�ɂ��̃L�[�������ꂽ���Ƃ������B
'-32767: 1000 0000 0000 0001 �]���āA����͗����ɊY�����邱�Ƃ������B
'0         :0000 0000 0000 0000 ����́A�ǂ���ł��Ȃ����Ƃ������B
'�܂茻�݃L�[��������Ă��邩�ǂ�����m��ɂ́AGetAsyncKeyState�̌��ʂ�-32768��And�}�X�N�Ɋ|���A
'-32768�ɂȂ��OK�Ƃ������ƁB

Function IsShiftKeyPressed() As Boolean
    Const KEY_PRESSED = -32768      '1000 0000 0000 0000 �ŏ�ʃr�b�g��1�ł��邱�Ƃ������B
    IsShiftKeyPressed = (GetAsyncKeyState(vbKeyShift) And KEY_PRESSED) = KEY_PRESSED
End Function
