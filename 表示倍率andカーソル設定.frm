VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �\���{��and�J�[�\���ݒ� 
   Caption         =   "�\���{��and�J�[�\���ݒ�"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3270
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

Private Sub btnCANCEL_Click()
    status = vbCancel
    Me.Hide
    
End Sub

Private Sub btnOK_Click()
    status = vbOK
    Me.Hide
    
End Sub

Private Sub UserForm_Activate()
    '������
    status = vbCancel
    Me.TextBoxMag = defaultMag
    Me.TextBoxFocus = defaultFocus
    Me.TextBoxMag.SetFocus
    Me.TextBoxMag.SelStart = 0
    Me.TextBoxMag.SelLength = Len(Me.TextBoxMag)
    
End Sub

