VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetSameViewFormMod 
   Caption         =   "SetSameView"
   ClientHeight    =   8250.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   OleObjectBlob   =   "SetSameViewFormMod.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "SetSameViewFormMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<License>------------------------------------------------------------
'
' Copyright (c) 2019 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'<Settings>----------------------------------------------------------------

'
' Me.show の直前で各 Input 要素に設定する値の種類
'
' * 0 -> Default
' * 1 -> Current
' * Other -> Do nothing
'
Const INITIALIZE_BY As Integer = 1

Const DEFAULT_ZOOM_LEVEL As Integer = 100
Const DEFAULT_ADDRESS_TO_SELECT As String = "A1"
Const DEFAULT_MINIMIZE_RIBBON As Boolean = True

'---------------------------------------------------------------</Settings>

Private ended_in As Integer
Private bool_change_event_enabled As Boolean

'<Controller>-----------------------------------------------------------------------------

'
' フォームを表示してユーザの選択内容を返却する
'
' 返却値型は MsgBox 関数と同じ(※) 型・意味とし、以下3種類のみを使用する
'
' | Constant | Value | Description                                                     |
' | -------- | ----- | --------------------------------------------------------------- |
' | vbOK     | 1     | OK押下                                                          |
' | vbCancel | 2     | Cancel押下                                                      |
' | vbAbort  | 3     | ウィンドウ右上 `x` クリックもしくは Alt + F4 でウィドウクローズ |
'
' ※
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
'
'
Public Function showForm()
    
    ended_in = vbAbort 'ウィンドウ右上 `×` クリックもしくは
        'Alt + F4 でウィドウクローズの場合はこの数値を返す

    '表示前 設定チェック
    Select Case INITIALIZE_BY
        Case 0 'Default 化指定
            Call setDefault
        Case 1 'Current 指定
            Call setCurrent
        Case Else '`何もしない` が指定
            'nothing to do
    End Select

    Me.Show
    ' ↑
    ' ｜ この間に GUI 操作
    ' ↓
    showForm = ended_in 'ユーザー選択内容の返却
    
End Function

'----------------------------------------------------------------------------</Controller>

'<Life cycle of Form>-----------------------------------------------------------------------------

'
' FormObject の load 時
' (呼び出し側モジュールで`SetSameViewFormMod`にアクセスした時に、load済みでなかった場合のみ)
' に実行される
'
Private Sub UserForm_Initialize()
    
    Call setDefault
    
End Sub

Private Sub UserForm_Activate()

    'note
    ' UserForm がメモリに展開された状態で シート名編集を行うと、
    ' cmbbx_sheet_name_to_activate の アイテムコレクション と不一致となるので、
    ' セットしなおす
    ' (UserForm を .show させる SetSameView() 側の処理終了時に、Unload しないようにした場合を想定した実装)

    Dim didFound As Boolean
    
    'シート名再設定
    tmpStr = cmbbx_sheet_name_to_activate.Text '選択済みのシート名を保存
    Me.cmbbx_sheet_name_to_activate.Clear
    
    counter = 0
    For Each sht In Sheets
        
        Me.cmbbx_sheet_name_to_activate.AddItem sht.Name
        
        If tmpStr = sht.Name Then '保存していたシート名が見つかった場合
            didFound = True
            Me.cmbbx_sheet_name_to_activate.ListIndex = counter 'シート名を選択
        End If
        
        counter = counter + 1
    Next
    
    If Not (didFound) Then
        Me.cmbbx_sheet_name_to_activate.ListIndex = 0
    
    End If
    
End Sub

'----------------------------------------------------------------------------</Life cycle of Form>


'<GUI Events>-----------------------------------------------------------------------------

'
' `Top left` check box
'
Private Sub chkbx_top_left_Change()

    If bool_change_event_enabled Then
    
        bool_change_event_enabled = False

        If Me.chkbx_top_left.Value Then 'checked
            
            Me.txtbx_top_left_address_of_view.Enabled = False
            
            Me.txtbx_range_address_to_select.Enabled = False
            
            Me.chkbx_same_as_top_left_address_of_view.Enabled = False
            
        Else 'unchecked
        
            Me.txtbx_top_left_address_of_view.Enabled = True
            
            If Not (Me.chkbx_same_as_top_left_address_of_view.Value) Then
                Me.txtbx_range_address_to_select.Enabled = True
            End If

            Me.chkbx_same_as_top_left_address_of_view.Enabled = True
            
        End If

        bool_change_event_enabled = True

    End If
End Sub

'
' `Top left address of view` text box
'
Private Sub txtbx_top_left_address_of_view_Change()

    If bool_change_event_enabled Then
    
        If Me.chkbx_same_as_top_left_address_of_view.Value Then 'checked
            Me.txtbx_range_address_to_select.Value = Me.txtbx_top_left_address_of_view.Value
        End If
    End If
End Sub

'
' `Same as top left address of view` check box
'
Private Sub chkbx_same_as_top_left_address_of_view_Change()

    If bool_change_event_enabled Then
    
        If Me.chkbx_same_as_top_left_address_of_view.Value Then 'checked
            
            Me.txtbx_range_address_to_select.Value = Me.txtbx_top_left_address_of_view.Value
            Me.txtbx_range_address_to_select.Enabled = False
            
        Else 'unchecked
        
            Me.txtbx_range_address_to_select.Enabled = True
        
        End If
    End If
End Sub

'
' `Current` button (in Sheet to activate area)
'
Private Sub buttn_set_sht_current_Click()
    
    Dim counter As Long: counter = 0

    For Each sht In Sheets
        If ActiveSheet.Name = sht.Name Then
            Me.cmbbx_sheet_name_to_activate.ListIndex = counter
            Exit For
        End If
        counter = counter + 1
    Next
    
End Sub

'
' `1st` button (in Sheet to activate area)
'
Private Sub buttn_set_sht_1st_Click()
    
    Me.cmbbx_sheet_name_to_activate.ListIndex = 0
    
End Sub

'
' `Set as default` Button
'
Private Sub buttn_set_all_as_default_Click()
    
    Call setDefault
    
End Sub

'
' `Set as current` Button
'
Private Sub buttn_set_all_as_current_Click()

    Call setCurrent

End Sub

'
' `OK` button
' この関数をバインドした Object (CommandButton Object) には `Default` プロパティに True を設定している
' その為、他のCommandButton Objectにフォーカスがない場合に enter キーを押してもこの関数は発火する

'
Private Sub buttn_ok_Click()
    ended_in = vbOK
    Me.Hide
    
End Sub

'
' `CANCEL` button
'
' NOTE
' この関数をバインドした Object (CommandButton Object) には `Cancel` プロパティに True を設定している
' その為、Esc キーを押すか、ボタンにフォーカスがあるときに enter キーを押してもこの関数は発火する
'
Private Sub buttn_cancel_Click()
    ended_in = vbCancel
    Me.Hide
    
End Sub

'----------------------------------------------------------------------------</GUI Events>

'<Common>-----------------------------------------------------------------------------

'
' デフォルト設定を反映させる
'
Private Sub setDefault()

    bool_change_event_enabled = False
    
    Me.txtbx_zoom_level.Value = DEFAULT_ZOOM_LEVEL

    Me.chkbx_top_left.Value = True

    Me.txtbx_top_left_address_of_view.Value = DEFAULT_ADDRESS_TO_SELECT
    Me.txtbx_top_left_address_of_view.Enabled = False

    Me.txtbx_range_address_to_select.Value = DEFAULT_ADDRESS_TO_SELECT
    Me.txtbx_range_address_to_select.Enabled = False

    Me.chkbx_same_as_top_left_address_of_view.Value = True
    Me.chkbx_same_as_top_left_address_of_view.Enabled = False
    
    Me.cmbbx_sheet_name_to_activate.Clear
    For Each sht In Sheets
        Me.cmbbx_sheet_name_to_activate.AddItem sht.Name
    Next
    Me.cmbbx_sheet_name_to_activate.ListIndex = 0 '最初のシート名を選択
    
    Me.chkbx_minimize_ribbon.Value = DEFAULT_MINIMIZE_RIBBON

    Me.chkbx_all_books.Value = False
    
    bool_change_event_enabled = True
    
End Sub

Private Sub setCurrent()

    Dim range_top_left_cell_of_unfreezed_pane As Range
    Dim range_imaginary_top_left_cell As Range
    Dim range_address_to_select As Range

    '<画面表示エリアの TopLeftCell 算出>---------------------------
    
    Dim long_freezed_panes_row_count As Long
    Dim long_freezed_panes_col_count As Long
    
    If (ActiveWindow.FreezePanes) Then

        If ActiveWindow.Panes.Count = 4 Then '画面4分割の場合

            '画面左上 Pain の画面占有範囲サイズを取得
            long_freezed_panes_row_count = ActiveWindow.Panes(1).VisibleRange.Rows.Count
            long_freezed_panes_col_count = ActiveWindow.Panes(1).VisibleRange.Columns.Count
            
        Else '2分割の場合

            If ActiveWindow.SplitRow = 0 Then '左右2分割の場合

                '画面上部 Pain の画面占有範囲サイズ(行数のみ)を取得
                long_freezed_panes_row_count = 0
                long_freezed_panes_col_count = ActiveWindow.Panes(1).VisibleRange.Columns.Count
            
            Else '上下2分割の場合 (activewindow.SplitColumn = 0 の場合)
                
                '画面左部 Pain の画面占有範囲サイズ(列数のみ)を取得
                long_freezed_panes_row_count = ActiveWindow.Panes(1).VisibleRange.Rows.Count
                long_freezed_panes_col_count = 0
                
            End If
        End If
    
    Else 'Activewindow は freeze されていない場合

        long_freezed_panes_row_count = 0
        long_freezed_panes_col_count = 0

    End If

    Set range_top_left_cell_of_unfreezed_pane = getTopLeftCellOfUnfreezedPane(ActiveWindow)

    Set range_imaginary_top_left_cell = ActiveWindow.ActiveSheet.Cells( _
        ActiveWindow.ScrollRow - long_freezed_panes_row_count, _
        ActiveWindow.ScrollColumn - long_freezed_panes_col_count _
    )

    '--------------------------</画面表示エリアの TopLeftCell 算出>

    '<Selection が示す Range 算出>---------------------------------

    If (TypeName(Selection) = "Range") Then '選択中の Object が Range の場合
    
        Set range_address_to_select = Selection '現在の選択範囲を取得
        
    Else '選択中のオブジェクトが Range Object でない場合
        
        Set selectionRange = getRangeFromSelectionObj(Selection)
        
        If (selectionRange Is Nothing) Then 'Selection の cell 占有領域の算出ができなかった場合
            
            retVal = MsgBox( _
                Prompt:= _
                    "Any cell or range is not selected. " & vbCrLf & _
                    "top left cell address of active window `" & range_top_left_cell_of_unfreezed_pane.Address(False, False) & "` will be set.", _
                Buttons:=vbExclamation _
            )

            Set range_address_to_select = range_top_left_cell_of_unfreezed_pane
            
        Else 'Selection の cell 占有領域の算出ができた場合
            
            retVal = MsgBox( _
                Prompt:= _
                    "Object type `" & TypeName(Selection) & "` selected. " & vbCrLf & _
                    "Ooccupied range address by that selection `" & selectionRange.Address(False, False) & "` will be set.", _
                Buttons:=vbExclamation _
            )

            Set range_address_to_select = selectionRange
            
        End If
        
    End If

    '--------------------------------</Selection が示す Range 算出>
    
    bool_change_event_enabled = False

    Me.txtbx_zoom_level.Value = ActiveWindow.Zoom

    If _
    ( _
        (range_top_left_cell_of_unfreezed_pane.Address = range_address_to_select.Address) And _
        (range_top_left_cell_of_unfreezed_pane.Item(1).Row = ActiveWindow.ScrollRow) And _
        (range_top_left_cell_of_unfreezed_pane.Item(1).Column = ActiveWindow.ScrollColumn) _
    ) Then 'ウィンドウ表示範囲も、選択セル(単一セル選択状態)も、左上のセルになっている

        Me.chkbx_top_left.Value = True

    Else
        Me.chkbx_top_left.Value = False

    End If

    Me.txtbx_top_left_address_of_view.Value = range_imaginary_top_left_cell.Address(False, False)
    Me.txtbx_top_left_address_of_view.Enabled = Not (Me.chkbx_top_left.Value)

    Me.txtbx_range_address_to_select.Value = range_address_to_select.Address(False, False)
    If (range_imaginary_top_left_cell.Address = range_address_to_select.Address) Then '画面右上セルと選択選択が同じ場合
        Me.txtbx_range_address_to_select.Enabled = False
    Else
        Me.txtbx_range_address_to_select.Enabled = Not (Me.chkbx_top_left.Value)
    End If

    If (range_imaginary_top_left_cell.Address = range_address_to_select.Address) Then '画面右上セルと選択選択が同じ場合
        Me.chkbx_same_as_top_left_address_of_view.Value = True
    Else
        Me.chkbx_same_as_top_left_address_of_view.Value = False
    End If
    Me.chkbx_same_as_top_left_address_of_view.Enabled = Not (Me.chkbx_top_left.Value)


    Dim counter As Long: counter = 0
    For Each sht In Sheets
        If ActiveSheet.Name = sht.Name Then
            Me.cmbbx_sheet_name_to_activate.ListIndex = counter
            Exit For
        End If
        counter = counter + 1
    Next

    Me.chkbx_minimize_ribbon.Value = Application.CommandBars.GetPressedMso("MinimizeRibbon")
    
    bool_change_event_enabled = True

End Sub

'
' オブジェクトのCell選択範囲を Range object にして返す
' 取得できなかった場合は Nothing を返す
'
Private Function getRangeFromSelectionObj(ByVal selectionObj As Object) As Variant

    Dim ret As Variant
    
    If selectionObj Is Nothing Then
        Set ret = Nothing ' Nothing を返す
    
    ElseIf (TypeName(selectionObj)) = "Range" Then ' Range オブジェyクトの場合
        Set ret = selectionObj 'そのまま返す
    
    Else ' Range オブジェyクトでない場合
        On Error GoTo TOP_LEFT_CELL_IS_NOT_DEFINED
        'TopLeftCell, BottomRightCell property を使って範囲を取得する
        Set ret = Range(selectionObj.TopLeftCell, selectionObj.BottomRightCell)
    
    End If
    
    Set getRangeFromSelectionObj = ret
    Exit Function
    
TOP_LEFT_CELL_IS_NOT_DEFINED:
    Set ret = Nothing
    Set getRangeFromSelectionObj = ret
    Exit Function
    
End Function

'
'unfreezed な pain 範囲のの左上セルのアドレスを取得する
'
Private Function getTopLeftCellOfUnfreezedPane(ByVal obj_window As Window) As Range

    Dim px_topLeftCell As Range

    If obj_window.FreezePanes Then 'freeze pain 有効の場合
    
        If obj_window.Panes.Count = 4 Then '画面4分割の場合
            Set p1 = obj_window.Panes(1)
            Set p1_bottomRightCell = getEdgeCellFromRange( _
                rangeObj:=p1.VisibleRange, _
                bottom:=True, _
                right:=True _
            ) 'pane(1)の範囲の右下のセルを取得
            Set px_topLeftCell = obj_window.ActiveSheet.Cells(p1_bottomRightCell.Row + 1, p1_bottomRightCell.Column + 1) 'pane(1)の範囲の1つ右下を設定
            
        Else '2分割の場合
        
            If obj_window.SplitRow = 0 Then '左右2分割の場合
                Set p1 = obj_window.Panes(1)
                Set p1_topRightCell = getEdgeCellFromRange( _
                    rangeObj:=p1.VisibleRange, _
                    bottom:=False, _
                    right:=True _
                ) 'pane(1)の範囲の右上のセルを取得
                Set px_topLeftCell = obj_window.ActiveSheet.Cells(1, p1_topRightCell.Column + 1) 'pane(1)の範囲の1つ右列の一番上を設定
            
            Else '上下2分割の場合 (obj_window.SplitColumn = 0 の場合)
                Set p1 = obj_window.Panes(1)
                Set p1_bottomLeftCell = getEdgeCellFromRange( _
                    rangeObj:=p1.VisibleRange, _
                    bottom:=True, _
                    right:=False _
                ) 'pane(1)の範囲の左下のセルを取得
                Set px_topLeftCell = obj_window.ActiveSheet.Cells(p1_bottomLeftCell.Row + 1, 1) 'pane(1)の範囲の1つ下行の一番左を設定
                
            End If
        
        End If

    Else
        Set px_topLeftCell = obj_window.ActiveSheet.Cells(1, 1) 'A1 セルを返す
        
    End If

    Set getTopLeftCellOfUnfreezedPane = px_topLeftCell

End Function

'
' Rangeオブジェクトの左上/右上/左下/右下のセルを返す
'
Private Function getEdgeCellFromRange(ByVal rangeObj As Range, ByVal bottom As Boolean, ByVal right As Boolean) As Range
    
    '変数
    Dim ret As Range
    Dim rowOffset As Long
    Dim colOffset As Long
    
    'Range 左上からの Row 相対位置の算出
    If bottom Then '最下部取得指定の場合
        rowOffset = rangeObj.Rows.Count - 1
    Else '最上部取得指定の場合
        rowOffset = 0
    End If
    
    'Range 左上からの Column 相対位置の算出
    If right Then '最右部取得指定の場合
        colOffset = rangeObj.Columns.Count - 1
    Else '最左部取得指定の場合
        colOffset = 0
    End If
    
    '返却値設定
    Set ret = rangeObj.Parent.Cells( _
        rangeObj.Item(1).Row + rowOffset, _
        rangeObj.Item(1).Column + colOffset _
    )
    
    Set getEdgeCellFromRange = ret '返却

End Function

'----------------------------------------------------------------------------</Common>


