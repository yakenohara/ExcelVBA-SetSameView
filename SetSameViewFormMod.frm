VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetSameViewFormMod 
   Caption         =   "SetSameView"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
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


Private Sub buttn_cancel_Click()
    status = vbCancel
    Me.Hide
    
End Sub

Private Sub buttn_set_all_as_default_Click()
    
    Call setDefault
    
End Sub

Private Sub buttn_set_all_as_current_Click()

    Dim cautionMessage As String: cautionMessage = "セルが選択(カーソル)されていません。" & vbLf & vbLf & _
                                                   "セルが選択(カーソル)されていない場合のカーソル位置は、" & vbLf & _
                                                   "現在のフォーカス位置左上になります。"
    
    Me.txtbx_zoom_level = ActiveWindow.Zoom
    
    Me.chkbx_top_left.Value = False

    Me.txtbx_top_left_address_of_view = ActiveWindow.VisibleRange(1).Address(False, False)
    
    If (Selection Is Nothing) Or (Not (TypeName(Selection) = "Range")) Then
        
        retVal = MsgBox(Prompt:=cautionMessage, Buttons:=vbExclamation)
        
        Me.txtbx_range_address_to_select = Me.txtbx_top_left_address_of_view
        
    Else
        
        Me.txtbx_range_address_to_select = Selection.Address(False, False)
        
    End If
    
    topLeftAddress = defaultFocus
    If ActiveWindow.FreezePanes Then 'freeze pain 有効の場合
    
        'unfreezed pain 範囲のの左上セルのアドレスを算出
        
        Dim px_topLeftCell As Range
        
        If ActiveWindow.Panes.Count = 4 Then '画面4分割の場合
            Set p1 = ActiveWindow.Panes(1)
            Set p1_bottomRightCell = p1.VisibleRange.Item(p1.VisibleRange.Count)
            Set px_topLeftCell = Cells(p1_bottomRightCell.Row + 1, p1_bottomRightCell.Column + 1)
            
        Else '2分割の場合
        
            If ActiveWindow.SplitRow = 0 Then '左右2分割の場合(ActiveWindow.SplitColumn = 0 の場合)
                Set px_topLeftCell = Cells(1, ActiveWindow.Column + 1)
            
            Else '上下2分割の場合
                Set px_topLeftCell = Cells(ActiveWindow.SplitRow + 1, 1)
                
            End If
        
        End If
        
        topLeftAddress = px_topLeftCell.Address(False, False)
        
    End If
    
    If topLeftAddress = Me.txtbx_range_address_to_select And _
        topLeftAddress = Me.txtbx_top_left_address_of_view Then '表示範囲左上が選択されている場合
        
        Me.chkbx_top_left.Value = True
        
        Me.chkbx_same_as_top_left_address_of_view.Value = True
        Me.chkbx_same_as_top_left_address_of_view.Enabled = False
                
        Me.txtbx_top_left_address_of_view = defaultFocus
        Me.txtbx_top_left_address_of_view.Enabled = False
        
        Me.txtbx_range_address_to_select = defaultFocus
        Me.txtbx_range_address_to_select.Enabled = False
    
    ElseIf Me.txtbx_range_address_to_select = Me.txtbx_top_left_address_of_view Then
        Me.chkbx_same_as_top_left_address_of_view.Value = True
        Me.txtbx_range_address_to_select.Enabled = False
        
    Else
        Me.chkbx_same_as_top_left_address_of_view.Value = False
        Me.txtbx_range_address_to_select.Enabled = True
    
    End If
    
    Call buttn_set_sht_current_Click

End Sub


Private Sub buttn_ok_Click()
    status = vbOK
    Me.Hide
    
End Sub

Private Sub buttn_set_sht_current_Click()
    Dim counter As Long: counter = 0
    
    For Each sht In Sheets
        If ActiveSheet.Name = sht.Name Then
            
            Me.cmbbx_sheet_name_to_activate.ListIndex = counter
            
        End If
        
        counter = counter + 1
        
    Next
    
End Sub

Private Sub buttn_set_sht_1st_Click()
    
    Me.cmbbx_sheet_name_to_activate.ListIndex = 0
    
End Sub

Private Sub chkbx_top_left_Change()
    
    If Me.chkbx_top_left.Value Then
        
        Me.chkbx_same_as_top_left_address_of_view.Enabled = False
        
        Me.txtbx_top_left_address_of_view = defaultFocus
        Me.txtbx_top_left_address_of_view.Enabled = False
        
        Me.txtbx_range_address_to_select = defaultFocus
        Me.txtbx_range_address_to_select.Enabled = False
        
        
    Else
    
        Me.chkbx_same_as_top_left_address_of_view.Enabled = True
        Me.txtbx_top_left_address_of_view.Enabled = True
        
        If Not (Me.chkbx_same_as_top_left_address_of_view) Then
            Me.txtbx_range_address_to_select.Enabled = True
        End If
        
    End If
    
End Sub

Private Sub chkbx_same_as_top_left_address_of_view_Change()
    
    If Me.chkbx_same_as_top_left_address_of_view.Value Then
        
        Me.txtbx_range_address_to_select = Me.txtbx_top_left_address_of_view
        Me.txtbx_range_address_to_select.Enabled = False
        
    Else
    
        Me.txtbx_range_address_to_select.Enabled = True
    
    End If
    
End Sub

Private Sub txtbx_top_left_address_of_view_Change()
    
    If Me.chkbx_same_as_top_left_address_of_view.Value Then
        
        Me.txtbx_range_address_to_select = Me.txtbx_top_left_address_of_view
        
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Call setDefault
    
End Sub

Private Sub UserForm_Activate()
    '変数
    Dim didFound As Boolean
    
    'シート名再設定
    tmpStr = cmbbx_sheet_name_to_activate.Text
    Me.cmbbx_sheet_name_to_activate.Clear
    
    counter = 0
    For Each sht In Sheets
        
        Me.cmbbx_sheet_name_to_activate.AddItem sht.Name
        
        If tmpStr = sht.Name Then
            didFound = True
            Me.cmbbx_sheet_name_to_activate.ListIndex = counter
            
        End If
        
        counter = counter + 1
        
    Next
    
    If Not (didFound) Then
        Me.cmbbx_sheet_name_to_activate.ListIndex = 0
    
    End If
    
End Sub

Private Sub setDefault()
    
    '初期化
    status = vbCancel
    
    Me.txtbx_zoom_level = defaultMag
    Me.chkbx_top_left.Value = True
    Me.txtbx_top_left_address_of_view = defaultFocus
    Me.txtbx_top_left_address_of_view.Enabled = False
    Me.chkbx_same_as_top_left_address_of_view.Value = True
    Me.chkbx_same_as_top_left_address_of_view.Enabled = False
    Me.txtbx_range_address_to_select = defaultFocus
    Me.txtbx_range_address_to_select.Enabled = False
    Me.txtbx_zoom_level.SetFocus
    Me.txtbx_zoom_level.SelStart = 0
    Me.txtbx_zoom_level.SelLength = Len(Me.txtbx_zoom_level)
    
    Me.cmbbx_sheet_name_to_activate.Clear
    
    For Each sht In Sheets
        
        Me.cmbbx_sheet_name_to_activate.AddItem sht.Name
        
    Next
    
    Me.cmbbx_sheet_name_to_activate.ListIndex = 0
    
    Me.chkbx_all_books.Value = False
    
End Sub



