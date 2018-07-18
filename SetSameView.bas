Attribute VB_Name = "SetSameView"
Sub SetSameView()

    '�ϐ��錾
    Dim s As Object
    Dim defaultSheet As Object
    Dim dispMag As Integer
    Dim focus As String
    Dim cursor As String
    Dim focusSht As String
    
    SetSameViewFormMod.Show
    
    '�t�H�[����Ԋm�F
    If SetSameViewFormMod.status <> vbOK Then
        Exit Sub
    
    ElseIf Not (IsNumeric(SetSameViewFormMod.TextBoxMag)) Then
        MsgBox "�w��\���{���͐��l�Ƃ��Ė����ł�"
        Exit Sub
    
    End If
                                      
    '�\���{���̎擾
    On Error GoTo whenOverFlowOccurred
    dispMag = CInt(SetSameViewFormMod.TextBoxMag)
    
    '�t�H�[�J�X�ʒu�擾
    focus = SetSameViewFormMod.TextBoxFocus
    cursor = SetSameViewFormMod.TextBoxCursor
    focusSht = SetSameViewFormMod.ComboBoxFocusShtNames.Text
    Application.ScreenUpdating = False
    
    
    
    '�J�[�\���ʒu�C���E�\���{���ύX
    On Error GoTo whenZoomFailed
    Set defaultSheet = ActiveSheet
    For Each s In ActiveWorkbook.Sheets
        s.Activate
        ActiveWindow.ScrollRow = Range(focus).Row
        ActiveWindow.ScrollColumn = Range(focus).Column
        Range(focus).Select
        ActiveWindow.Zoom = dispMag
        Range(cursor).Select
    Next s
    Worksheets(focusSht).Activate
    
    Application.ScreenUpdating = True
    MsgBox "Done!"
    
    Exit Sub
    
whenOverFlowOccurred:
    MsgBox "Application.InputBox���\�b�h�ŗ�O" & vbLf & _
           "�I�[�o�[�t���[�̉\��������܂�"
    
    Exit Sub
    
whenZoomFailed:
    Application.ScreenUpdating = True
    MsgBox "Window.Zoom�v���p�e�B�ŗ�O" & vbLf & _
           "�w��\���{�����J�[�\���������s���ȉ\��������܂�"
    
    Exit Sub
    
End Sub

