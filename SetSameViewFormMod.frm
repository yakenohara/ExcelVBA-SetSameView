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

Public hided_in As Integer

Const DEFAULT_ZOOM_LEVEL As Integer = 100
Const DEFAULT_ADDRESS_TO_SELECT As String = "A1"
Const DEFAULT_MINIMIZE_RIBBON As Boolean = True

'<Life cycle of Form>-----------------------------------------------------------------------------

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

'----------------------------------------------------------------------------</Life cycle of Form>


'<GUI Events>-----------------------------------------------------------------------------

'
' `Top left` check box
'
Private Sub chkbx_top_left_Change()
    
    If Me.chkbx_top_left.Value Then
        
        Me.chkbx_same_as_top_left_address_of_view.Enabled = False
        
        Me.txtbx_top_left_address_of_view = DEFAULT_ADDRESS_TO_SELECT
        Me.txtbx_top_left_address_of_view.Enabled = False
        
        Me.txtbx_range_address_to_select = DEFAULT_ADDRESS_TO_SELECT
        Me.txtbx_range_address_to_select.Enabled = False
        
        
    Else
    
        Me.chkbx_same_as_top_left_address_of_view.Enabled = True
        Me.txtbx_top_left_address_of_view.Enabled = True
        
        If Not (Me.chkbx_same_as_top_left_address_of_view) Then
            Me.txtbx_range_address_to_select.Enabled = True
        End If
        
    End If
    
End Sub

'
' `Top left address of view` text box
'
Private Sub txtbx_top_left_address_of_view_Change()
    
    If Me.chkbx_same_as_top_left_address_of_view.Value Then
        
        Me.txtbx_range_address_to_select = Me.txtbx_top_left_address_of_view
        
    End If
    
End Sub

'
' `Same as top left address of view` check box
'
Private Sub chkbx_same_as_top_left_address_of_view_Change()
    
    If Me.chkbx_same_as_top_left_address_of_view.Value Then
        
        Me.txtbx_range_address_to_select = Me.txtbx_top_left_address_of_view
        Me.txtbx_range_address_to_select.Enabled = False
        
    Else
    
        Me.txtbx_range_address_to_select.Enabled = True
    
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

    Me.txtbx_zoom_level = ActiveWindow.Zoom
    
    Me.chkbx_top_left.Value = False

    'todo グラフだけを表示しているシートを表示中の場合にコケる
    Me.txtbx_top_left_address_of_view = ActiveWindow.VisibleRange(1).Address(False, False)
    
    If (Selection Is Nothing) Or (Not (TypeName(Selection) = "Range")) Then
    
        str_range_address_to_select = ""
    
        Set selectionRange = getRangeFromSelectionObj(Selection)
        
        If (selectionRange Is Nothing) Then 'Selection の cell 占有領域の算出ができなかった場合
            retVal = MsgBox( _
                Prompt:= _
                    "Any cell or range is not selected. " & vbCrLf & _
                    "top left cell address of active window `" & Me.txtbx_top_left_address_of_view & "` will be set.", _
                Buttons:=vbExclamation _
            )
            
            str_range_address_to_select = Me.txtbx_top_left_address_of_view.Value
            
        Else 'Selection の cell 占有領域の算出ができた場合
            
            str_range_address_to_select = selectionRange.Address(False, False)
                
            retVal = MsgBox( _
                Prompt:= _
                    "Object type `" & TypeName(Selection) & "` selected. " & vbCrLf & _
                    "Coccupied range address by that selection `" & str_range_address_to_select & "` will be set.", _
                Buttons:=vbExclamation _
            )
            
        End If
        
        Me.txtbx_range_address_to_select = str_range_address_to_select
        
    Else
        
        Me.txtbx_range_address_to_select = Selection.Address(False, False)
        
    End If
    
    topLeftAddress = DEFAULT_ADDRESS_TO_SELECT
    If ActiveWindow.FreezePanes Then 'freeze pain 有効の場合
    
        'unfreezed pain 範囲のの左上セルのアドレスを算出
        
        Dim px_topLeftCell As Range
        
        If ActiveWindow.Panes.Count = 4 Then '画面4分割の場合
            Set p1 = ActiveWindow.Panes(1)
            Set p1_bottomRightCell = p1.VisibleRange.Item(p1.VisibleRange.Count)
            Set px_topLeftCell = Cells(p1_bottomRightCell.Row + 1, p1_bottomRightCell.Column + 1)
            
        Else '2分割の場合
        
            If ActiveWindow.SplitRow = 0 Then '左右2分割の場合
                'todo ActiveWindow.Panes(1).VisibleRange の右上のセルの一つ右のセルにする
'                Set p1 = ActiveWindow.Panes(1)
'                Set xx = p1.VisibleRange
'                Set p1_bottomRightCell = p1.VisibleRange.Item(p1.VisibleRange.Count)
                Set px_topLeftCell = Cells(1, ActiveWindow.SplitColumn + 1)
            
            Else '上下2分割の場合 (ActiveWindow.SplitColumn = 0 の場合)
                'todo ActiveWindow.Panes(1).VisibleRange の左下のセルの一つ下のセルにする
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
                
        Me.txtbx_top_left_address_of_view = DEFAULT_ADDRESS_TO_SELECT
        Me.txtbx_top_left_address_of_view.Enabled = False
        
        Me.txtbx_range_address_to_select = DEFAULT_ADDRESS_TO_SELECT
        Me.txtbx_range_address_to_select.Enabled = False
    
    ElseIf Me.txtbx_range_address_to_select = Me.txtbx_top_left_address_of_view Then
        Me.chkbx_same_as_top_left_address_of_view.Value = True
        Me.txtbx_range_address_to_select.Enabled = False
        
    Else
        Me.chkbx_same_as_top_left_address_of_view.Value = False
        Me.txtbx_range_address_to_select.Enabled = True
    
    End If
    
    Call buttn_set_sht_current_Click '現在選択中のシート名を選択
    
    'リボンの表示 / 非表示状態の反映
    bool_ribbon_is_minimized = Application.CommandBars.GetPressedMso("MinimizeRibbon")
    Me.chkbx_minimize_ribbon.Value = bool_ribbon_is_minimized

End Sub

'
' `OK` button
'
Private Sub buttn_ok_Click()
    hided_in = vbOK
    Me.Hide
    
End Sub

'
' `CANCEL` button
'
Private Sub buttn_cancel_Click()
    hided_in = vbCancel
    Me.Hide
    
End Sub

'----------------------------------------------------------------------------</GUI Events>

'<Common>-----------------------------------------------------------------------------

Private Sub setDefault()
    
    '初期化
    hided_in = vbCancel
    
    Me.txtbx_zoom_level = DEFAULT_ZOOM_LEVEL
    Me.chkbx_top_left.Value = True
    Me.txtbx_top_left_address_of_view = DEFAULT_ADDRESS_TO_SELECT
    Me.txtbx_top_left_address_of_view.Enabled = False
    Me.chkbx_same_as_top_left_address_of_view.Value = True
    Me.chkbx_same_as_top_left_address_of_view.Enabled = False
    Me.txtbx_range_address_to_select = DEFAULT_ADDRESS_TO_SELECT
    Me.txtbx_range_address_to_select.Enabled = False
    Me.txtbx_zoom_level.SetFocus
    Me.txtbx_zoom_level.SelStart = 0
    Me.txtbx_zoom_level.SelLength = Len(Me.txtbx_zoom_level)
    
    Me.cmbbx_sheet_name_to_activate.Clear
    
    For Each sht In Sheets
        
        Me.cmbbx_sheet_name_to_activate.AddItem sht.Name
        
    Next
    
    Me.cmbbx_sheet_name_to_activate.ListIndex = 0
    
    Me.chkbx_minimize_ribbon.Value = DEFAULT_MINIMIZE_RIBBON
    Me.chkbx_all_books.Value = False
    
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

'----------------------------------------------------------------------------</Common>


