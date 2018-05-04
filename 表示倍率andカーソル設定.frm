VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �\���{��and�J�[�\���ݒ� 
   Caption         =   "�\���{��and�J�[�\���ݒ�"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   OleObjectBlob   =   "�\���{��and�J�[�\���ݒ�.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�\���{��and�J�[�\���ݒ�"
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
    
    didFound = False
    
    '�V�[�g������
    For Each sht In Sheets
        
        If Me.ComboBoxFocusShtNames.Text = sht.Name Then
            didFound = True
            Exit For
        End If
        
    Next
    
    '�V�[�g��������Ȃ������ꍇ
    If Not (didFound) Then
        
        '�R���{�{�b�N�X�̏�����
        
        Me.ComboBoxFocusShtNames.Clear
        
        For Each sht In Sheets
            
            Me.ComboBoxFocusShtNames.AddItem sht.Name
            
        Next
        
        Me.ComboBoxFocusShtNames.ListIndex = 0
        
    End If
    
End Sub

Private Sub setDefault()
    
    '������
    status = vbCancel
    
    Me.TextBoxMag = defaultMag
    Me.TextBoxFocus = defaultFocus
    Me.CheckBoxSameAsFocus.Value = True
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
    
End Sub
