VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetSameViewFormMod 
   Caption         =   "SetSameView"
   ClientHeight    =   8265.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5610
   OleObjectBlob   =   "SetSameViewFormMod.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "SetSameViewFormMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Public status As Integer

Const defaultMag As Integer = 100
Const defaultFocus As String = "A1"

'Dim isFirstBoot As Boolean: isFirstBoot = True


Private Sub btnCANCEL_Click()
    status = vbCancel
    Me.Hide
    
End Sub

Private Sub BtnDefault_Click()
    
    Call setDefault
    
End Sub

Private Sub BtnNowView_Click()

    Dim cautionMessage As String: cautionMessage = "�Z�����I��(�J�[�\��)����Ă��܂���B" & vbLf & vbLf & _
                                                   "�Z�����I��(�J�[�\��)����Ă��Ȃ��ꍇ�̃J�[�\���ʒu�́A" & vbLf & _
                                                   "���݂̃t�H�[�J�X�ʒu����ɂȂ�܂��B"
    
    Me.TextBoxMag = ActiveWindow.Zoom
    
    Me.CheckBoxCloserToA1.Value = False

    Me.TextBoxFocus = ActiveWindow.VisibleRange(1).Address(False, False)
    
    If (Selection Is Nothing) Or (Not (TypeName(Selection) = "Range")) Then
        
        retVal = MsgBox(Prompt:=cautionMessage, Buttons:=vbExclamation)
        
        Me.TextBoxCursor = Me.TextBoxFocus
        
    Else
        
        Me.TextBoxCursor = Selection.Address(False, False)
        
    End If
    
    topLeftAddress = defaultFocus
    If ActiveWindow.FreezePanes Then 'freeze pain �L���̏ꍇ
    
        'unfreezed pain �͈͂̂̍���Z���̃A�h���X���Z�o
        
        Dim px_topLeftCell As Range
        
        If ActiveWindow.Panes.Count = 4 Then '���4�����̏ꍇ
            Set p1 = ActiveWindow.Panes(1)
            Set p1_bottomRightCell = p1.VisibleRange.Item(p1.VisibleRange.Count)
            Set px_topLeftCell = Cells(p1_bottomRightCell.Row + 1, p1_bottomRightCell.Column + 1)
            
        Else '2�����̏ꍇ
        
            If ActiveWindow.SplitRow = 0 Then '���E2�����̏ꍇ(ActiveWindow.SplitColumn = 0 �̏ꍇ)
                Set px_topLeftCell = Cells(1, ActiveWindow.Column + 1)
            
            Else '�㉺2�����̏ꍇ
                Set px_topLeftCell = Cells(ActiveWindow.SplitRow + 1, 1)
                
            End If
        
        End If
        
        topLeftAddress = px_topLeftCell.Address(False, False)
        
    End If
    
    If topLeftAddress = Me.TextBoxCursor And _
        topLeftAddress = Me.TextBoxFocus Then '�\���͈͍��オ�I������Ă���ꍇ
        
        Me.CheckBoxCloserToA1.Value = True
        
        Me.CheckBoxSameAsFocus.Value = True
        Me.CheckBoxSameAsFocus.Enabled = False
                
        Me.TextBoxFocus = defaultFocus
        Me.TextBoxFocus.Enabled = False
        
        Me.TextBoxCursor = defaultFocus
        Me.TextBoxCursor.Enabled = False
    
    ElseIf Me.TextBoxCursor = Me.TextBoxFocus Then
        Me.CheckBoxSameAsFocus.Value = True
        Me.TextBoxCursor.Enabled = False
        
    Else
        Me.CheckBoxSameAsFocus.Value = False
        Me.TextBoxCursor.Enabled = True
    
    End If
    
    Call BtnSelectNowSht_Click

End Sub


Private Sub btnOK_Click()
    status = vbOK
    Me.Hide
    
End Sub

Private Sub BtnSelectNowSht_Click()
    Dim counter As Long: counter = 0
    
    For Each sht In Sheets
        If ActiveSheet.Name = sht.Name Then
            
            Me.ComboBoxFocusShtNames.ListIndex = counter
            
        End If
        
        counter = counter + 1
        
    Next
    
End Sub

Private Sub BtnSelect1stSht_Click()
    
    Me.ComboBoxFocusShtNames.ListIndex = 0
    
End Sub

Private Sub CheckBoxCloserToA1_Change()
    
    If Me.CheckBoxCloserToA1.Value Then
        
        Me.CheckBoxSameAsFocus.Enabled = False
        
        Me.TextBoxFocus = defaultFocus
        Me.TextBoxFocus.Enabled = False
        
        Me.TextBoxCursor = defaultFocus
        Me.TextBoxCursor.Enabled = False
        
        
    Else
    
        Me.CheckBoxSameAsFocus.Enabled = True
        Me.TextBoxFocus.Enabled = True
        
        If Not (Me.CheckBoxSameAsFocus) Then
            Me.TextBoxCursor.Enabled = True
        End If
        
    End If
    
End Sub

Private Sub CheckBoxSameAsFocus_Change()
    
    If Me.CheckBoxSameAsFocus.Value Then
        
        Me.TextBoxCursor = Me.TextBoxFocus
        Me.TextBoxCursor.Enabled = False
        
    Else
    
        Me.TextBoxCursor.Enabled = True
    
    End If
    
End Sub

Private Sub TextBoxFocus_Change()
    
    If Me.CheckBoxSameAsFocus.Value Then
        
        Me.TextBoxCursor = Me.TextBoxFocus
        
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Call setDefault
    
End Sub

Private Sub UserForm_Activate()
    '�ϐ�
    Dim didFound As Boolean
    
    '�V�[�g���Đݒ�
    tmpStr = ComboBoxFocusShtNames.Text
    Me.ComboBoxFocusShtNames.Clear
    
    counter = 0
    For Each sht In Sheets
        
        Me.ComboBoxFocusShtNames.AddItem sht.Name
        
        If tmpStr = sht.Name Then
            didFound = True
            Me.ComboBoxFocusShtNames.ListIndex = counter
            
        End If
        
        counter = counter + 1
        
    Next
    
    If Not (didFound) Then
        Me.ComboBoxFocusShtNames.ListIndex = 0
    
    End If
    
End Sub

Private Sub setDefault()
    
    '������
    status = vbCancel
    
    Me.TextBoxMag = defaultMag
    Me.CheckBoxCloserToA1.Value = True
    Me.TextBoxFocus = defaultFocus
    Me.TextBoxFocus.Enabled = False
    Me.CheckBoxSameAsFocus.Value = True
    Me.CheckBoxSameAsFocus.Enabled = False
    Me.TextBoxCursor = defaultFocus
    Me.TextBoxCursor.Enabled = False
    Me.TextBoxMag.SetFocus
    Me.TextBoxMag.SelStart = 0
    Me.TextBoxMag.SelLength = Len(Me.TextBoxMag)
    
    Me.ComboBoxFocusShtNames.Clear
    
    For Each sht In Sheets
        
        Me.ComboBoxFocusShtNames.AddItem sht.Name
        
    Next
    
    Me.ComboBoxFocusShtNames.ListIndex = 0
    
    Me.CheckBoxEveryBook.Value = False
    
End Sub
