Attribute VB_Name = "SetSameView"
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

Sub SetSameView()

    'フォームの表示 & ユーザー選択状態の取得
    formEndsIn = SetSameViewFormMod.showForm()

    'フォームは ウィンドウ右上 `×` クリックもしくは Alt + F4 でウィドウクローズ された
    If formEndsIn = vbAbort Then
        Exit Sub '何もせず終了
    End If
    
    'フォーム状態確認
    If formEndsIn = vbOK Then
    
        '<フォームの設定内容の型チェック>-----------------------------------------------------------------
        Dim bool_type_ok As Boolean: bool_type_ok = True  'OKを格納する(NG になったときだけ、 False にする)
        
        'Zoom level text box のチェック
        Dim int_tmp_val As Integer
        bool_is_int = cIntSafely(SetSameViewFormMod.txtbx_zoom_level.Value, int_tmp_val) 'Zoom level を取得 & Integer 変換可能かどうかチェック
        If Not (bool_is_int) Then 'Zoom level の設定値は Integer に変換不可能
            retOfMsgBox = MsgBox("Invalid Zoom level :`" & int_tmp_val & "` specified", vbCritical) 'エラーをメッセージで表示
            bool_type_ok = False
        End If
        '----------------------------------------------------------------</フォームの設定内容の型チェック>
        
        If (bool_type_ok) Then 'フォーム設定内容の型チェック OK の場合
            
            'フォームの設定内容の取得
            Dim dict_view_setting As Object
            Set dict_view_setting = CreateObject("Scripting.Dictionary")
            
            With dict_view_setting
                .Add "prop_int_zoom_level", int_tmp_val
                .Add "prop_bool_top_left_option_enabled", SetSameViewFormMod.chkbx_top_left.Value
                .Add "prop_str_top_left_address_of_view", SetSameViewFormMod.txtbx_top_left_address_of_view.Value
                .Add "prop_str_range_address_to_select", SetSameViewFormMod.txtbx_range_address_to_select.Value
                .Add "prop_str_sheet_name_to_activate", SetSameViewFormMod.cmbbx_sheet_name_to_activate.Value
                .Add "prop_bool_minimize_ribbon_option_enabled", SetSameViewFormMod.chkbx_minimize_ribbon.Value
                .Add "prop_bool_all_books_option_enabled", SetSameViewFormMod.chkbx_all_books.Value
            End With
            
            'View反映
            Application.ScreenUpdating = False
            ret = satSameViewIterator(dict_view_setting)
            Application.ScreenUpdating = True

            If ret Then
                MsgBox "Done!"
            End If
            
        'Note
        '  フォーム設定内容の型チェック NG の場合のメッセージ処理は、
        '  <フォームの設定内容の型チェック></フォームの設定内容の型チェック>
        '  の内部で行う
            
        End If

    End If

    'フォーム開放
    Unload SetSameViewFormMod
    
End Sub

Private Function satSameViewIterator(ByVal dict_view_setting As Object) As Boolean

    '<処理対象 WorkBook を collection 化>------------------------------------------------------
    
    Set collection_books_to_operate = New Collection

    If dict_view_setting.Item("prop_bool_all_books_option_enabled") Then 'すべてのブック処理の場合
       
        For Each wbk In Workbooks
            If Windows(wbk.Name).Visible Then 'Visible == ture なWorkBookのみ処理する
                collection_books_to_operate.Add wbk
            End If
        Next
   
    Else 'AcriveWorkBookのみの場合
        collection_books_to_operate.Add ActiveWorkbook
       
    End If

    '-----------------------------------------------------</処理対象 WorkBook を collection 化>
    
    Set obj_book_to_activate = ActiveWorkbook '処理終了時にアクティブにするブックを記録

    Dim str_top_left_address_of_view As String
    Dim str_range_address_to_select As String
    
    'View設定ループ
    For Each bk In collection_books_to_operate
       
        bk.Activate
       
        'ribbon
        bool_ribbon_is_minimized = Application.CommandBars.GetPressedMso("MinimizeRibbon")
        If (bool_ribbon_is_minimized <> dict_view_setting.Item("prop_bool_minimize_ribbon_option_enabled")) Then 'リボンの 表示 / 非表示状態が 設定値と異なる場合
            Application.CommandBars.ExecuteMso "MinimizeRibbon" 'リボン表示 / 非表示の切り替え
        End If
       
        bool_found_sheet_to_activate = False 'アクティブ化 対象シートの存在

        For Each obj_sheet In bk.Sheets
           
            obj_sheet.Activate
            Set range_top_left_of_unfreezed_pain = getTopLeftCellOfUnfreezedPane(ActiveWindow)
            Dim range_top_left_of_specified As Range
            
            On Error GoTo EXCEPTION_VIEW_SET_FAILED
            
            If dict_view_setting.Item("prop_bool_top_left_option_enabled") Then '左上セルにあわせた View 設定指定の場合

                str_top_left_address_of_view = range_top_left_of_unfreezed_pain.Address
                str_range_address_to_select = range_top_left_of_unfreezed_pain.Address
                Set range_top_left_of_specified = Range(str_range_address_to_select)

            Else '左上セルにあわせた View 設定指定が無効(=form の text box で指定した Cell Address を使用する指定)の場合
                
                str_top_left_address_of_view = dict_view_setting.Item("prop_str_top_left_address_of_view")
                str_range_address_to_select = dict_view_setting.Item("prop_str_range_address_to_select")
                
                Set range_imaginary_top_left_of_specified = Range(str_top_left_address_of_view)
                arr_imaginary_p1_range_count = getImaginaryPane1_sRangeCount(ActiveWindow)
                long_lbound = LBound(arr_imaginary_p1_range_count)
                
                '
                Set range_top_left_of_specified = Cells( _
                    range_imaginary_top_left_of_specified.Row + arr_imaginary_p1_range_count(long_lbound), _
                    range_imaginary_top_left_of_specified.Column + arr_imaginary_p1_range_count(long_lbound + 1) _
                )

            End If
            
            ActiveWindow.Zoom = dict_view_setting.Item("prop_int_zoom_level")
            ActiveWindow.ScrollRow = range_top_left_of_specified.Row
            ActiveWindow.ScrollColumn = range_top_left_of_specified.Column
            Range(str_range_address_to_select).Select
            
            On Error GoTo 0

            'アクティブ化 対象シートかどうかチェック
            If obj_sheet.Name = dict_view_setting.Item("prop_str_sheet_name_to_activate") Then
                bool_found_sheet_to_activate = True
            End If
           
        Next obj_sheet
       
        'フォーカスシートの設定
        If bool_found_sheet_to_activate Then 'フォーカスすべきシートが存在する
            bk.Worksheets(dict_view_setting.Item("prop_str_sheet_name_to_activate")).Activate
           
        Else 'フォーカスすべきシートが存在しない
            bk.Worksheets(1).Activate '先頭のシートを選択
           
        End If
   
    Next bk
   
    obj_book_to_activate.Activate '処理開始時のブックを Active に戻す
    satSameViewIterator = True
    Exit Function

EXCEPTION_VIEW_SET_FAILED:
    
    retOfMsg = MsgBox( _
        "Exception occurred. As a cause, Specified zoom level or cursor format may be invalid", _
        vbCritical _
    )
    
    satSameViewIterator = False
    Exit Function

End Function

'
' String を Integer に変換する
' 成功したら TRUE, 失敗したら FALSE を返す
'
Private Function cIntSafely(ByVal fromThisString As String, ByRef toThisInt As Integer) As Boolean

    Dim ret As Boolean
    
    If Not (IsNumeric(fromThisString)) Then '数値に変換不可能な場合
        
        ret = False '失敗を格納
    
    Else '数値に変換可能な場合
        
        On Error GoTo EXCEPTION_OVERFLOWED 'CInt() でオーバーフローの場合は EXCEPTION_OVERFLOWED に Go
        toThisInt = CInt(fromThisString) '指定変数に格納
        ret = True
        
    End If
    
    cIntSafely = ret '成功 / 失敗状態を返却
    Exit Function
    
EXCEPTION_OVERFLOWED:
    ret = False '失敗を格納
    cIntSafely = ret
    Exit Function
        
    
End Function

'
'unfreezed な pain 範囲の左上セルのアドレスを取得する
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
                Set px_topLeftCell = obj_window.ActiveSheet.Cells(1, p1_topRightCell.Column + 1) 'pane(1)の範囲の1つ右を設定
            
            Else '上下2分割の場合 (obj_window.SplitColumn = 0 の場合)
                Set p1 = obj_window.Panes(1)
                Set p1_bottomLeftCell = getEdgeCellFromRange( _
                    rangeObj:=p1.VisibleRange, _
                    bottom:=True, _
                    right:=False _
                ) 'pane(1)の範囲の左下のセルを取得
                Set px_topLeftCell = obj_window.ActiveSheet.Cells(p1_bottomLeftCell.Row + 1, 1) 'pane(1)の範囲の1つ下を設定
                
            End If
        
        End If

    Else
        Set px_topLeftCell = obj_window.ActiveSheet.Cells(1, 1) 'A1 セルを返す
        
    End If

    Set getTopLeftCellOfUnfreezedPane = px_topLeftCell

End Function

'
' 左上にFreezed Pane が存在すると仮定した場合の、
' その Pane のセル専有範囲(rows.count, columns.count)配列を算出して返す
'
Private Function getImaginaryPane1_sRangeCount(ByVal obj_window As Window) As Variant
    
    Dim arr_ret As Variant '返却値
    
    If obj_window.FreezePanes Then 'freeze pain 有効の場合
    
        
        If obj_window.Panes.Count = 4 Then '画面4分割の場合
            Set p1 = obj_window.Panes(1)
            arr_ret = Array( _
                p1.VisibleRange.Rows.Count, _
                p1.VisibleRange.Columns.Count _
            )
            
        Else '2分割の場合
        
            If obj_window.SplitRow = 0 Then '左右2分割の場合
                Set p1 = obj_window.Panes(1)
                arr_ret = Array( _
                    0, _
                    p1.VisibleRange.Columns.Count _
                )
            
            Else '上下2分割の場合 (obj_window.SplitColumn = 0 の場合)
                Set p1 = obj_window.Panes(1)
                arr_ret = Array( _
                    p1.VisibleRange.Rows.Count, _
                    0 _
                )
                
            End If
        
        End If

    Else
        arr_ret = Array(0, 0) '範囲0 で返す
        
    End If

    getImaginaryPane1_sRangeCount = arr_ret

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



