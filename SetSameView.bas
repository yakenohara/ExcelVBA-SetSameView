Attribute VB_Name = "�S�V�[�g�\���{��and�J�[�\���C��"
Sub �S�V�[�g�\���{��and�J�[�\���C��()

    '�ϐ��錾
    Dim s As Object
    Dim defaultSheet As Object
    Dim dispMag As Integer
    Dim focus As String
    Dim cursor As String
    Dim focusSht As String
    
    �\���{��and�J�[�\���ݒ�.Show
    
    '�t�H�[����Ԋm�F
    If �\���{��and�J�[�\���ݒ�.status <> vbOK Then
        Exit Sub
    
    ElseIf Not (IsNumeric(�\���{��and�J�[�\���ݒ�.TextBoxMag)) Then
        MsgBox "�w��\���{���͐��l�Ƃ��Ė����ł�"
        Exit Sub
    
    End If
                                      
    '�\���{���̎擾
    On Error GoTo whenOverFlowOccurred
    dispMag = CInt(�\���{��and�J�[�\���ݒ�.TextBoxMag)
    
    '�t�H�[�J�X�ʒu�擾
    focus = �\���{��and�J�[�\���ݒ�.TextBoxFocus
    cursor = �\���{��and�J�[�\���ݒ�.TextBoxCursor
    focusSht = �\���{��and�J�[�\���ݒ�.ComboBoxFocusShtNames.Text
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

