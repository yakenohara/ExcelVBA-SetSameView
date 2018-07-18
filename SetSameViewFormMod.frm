VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetSameViewFormMod 
   Caption         =   "表示倍率andカーソル設定"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   OleObjectBlob   =   "SetSameViewFormMod.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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

    Dim cautionMessage As String: cautionMessage = "セルが選択(カーソル)されていません。" & vbLf & vbLf & _
                                                   "セルが選択(カーソル)されていない場合のカーソル位置は、" & vbLf & _
                                                   "現在のフォーカス位置左上になります。"
    
    Me.TextBoxMag = ActiveWindow.Zoom

    Me.TextBoxFocus = ActiveWindow.VisibleRange(1).Address(False, False)
    
    If (Selection Is Nothing) Or (Not (TypeName(Selection) = "Range")) Then
        
        retVal = MsgBox(Prompt:=cautionMessage, Buttons:=vbExclamation)
        
        Me.TextBoxCursor = Me.TextBoxFocus
        
    Else
        
        Me.TextBoxCursor = Selection.Address(False, False)
        
    End If
    
    If Me.TextBoxCursor = Me.TextBoxFocus Then
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
    '変数
    Dim didFound As Boolean
    
    'シート名再設定
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
    
    '初期化
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
