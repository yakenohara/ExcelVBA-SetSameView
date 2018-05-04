VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 表示倍率andカーソル設定 
   Caption         =   "表示倍率andカーソル設定"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3270
   OleObjectBlob   =   "表示倍率andカーソル設定.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "表示倍率andカーソル設定"
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
    '初期化
    status = vbCancel
    Me.TextBoxMag = defaultMag
    Me.TextBoxFocus = defaultFocus
    Me.TextBoxMag.SetFocus
    Me.TextBoxMag.SelStart = 0
    Me.TextBoxMag.SelLength = Len(Me.TextBoxMag)
    
End Sub

