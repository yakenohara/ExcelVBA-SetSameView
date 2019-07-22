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

    '�t�H�[���̕\�� & ���[�U�[�I����Ԃ̎擾
    formEndsIn = SetSameViewFormMod.showForm()

    '�t�H�[���� �E�B���h�E�E�� `�~` �N���b�N�������� Alt + F4 �ŃE�B�h�E�N���[�Y ���ꂽ
    If formEndsIn = vbAbort Then
        Exit Sub '���������I��
    End If
    
    '�t�H�[����Ԋm�F
    If formEndsIn = vbOK Then
    
        '<�t�H�[���̐ݒ���e�̌^�`�F�b�N>-----------------------------------------------------------------
        Dim bool_type_ok As Boolean: bool_type_ok = True  'OK���i�[����(NG �ɂȂ����Ƃ������A False �ɂ���)
        
        'Zoom level text box �̃`�F�b�N
        Dim int_tmp_val As Integer
        bool_is_int = cIntSafely(SetSameViewFormMod.txtbx_zoom_level.Value, int_tmp_val) 'Zoom level ���擾 & Integer �ϊ��\���ǂ����`�F�b�N
        If Not (bool_is_int) Then 'Zoom level �̐ݒ�l�� Integer �ɕϊ��s�\
            retOfMsgBox = MsgBox("Invalid Zoom level :`" & int_tmp_val & "` specified", vbCritical) '�G���[�����b�Z�[�W�ŕ\��
            bool_type_ok = False
        End If
        '----------------------------------------------------------------</�t�H�[���̐ݒ���e�̌^�`�F�b�N>
        
        If (bool_type_ok) Then '�t�H�[���ݒ���e�̌^�`�F�b�N OK �̏ꍇ
            
            '�t�H�[���̐ݒ���e�̎擾
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
            
            'View���f
            Application.ScreenUpdating = False
            ret = satSameViewIterator(dict_view_setting)
            Application.ScreenUpdating = True

            If ret Then
                MsgBox "Done!"
            End If
            
        'Note
        '  �t�H�[���ݒ���e�̌^�`�F�b�N NG �̏ꍇ�̃��b�Z�[�W�����́A
        '  <�t�H�[���̐ݒ���e�̌^�`�F�b�N></�t�H�[���̐ݒ���e�̌^�`�F�b�N>
        '  �̓����ōs��
            
        End If

    End If

    '�t�H�[���J��
    Unload SetSameViewFormMod
    
End Sub

Private Function satSameViewIterator(ByVal dict_view_setting As Object) As Boolean

    '<�����Ώ� WorkBook �� collection ��>------------------------------------------------------
    
    Set collection_books_to_operate = New Collection

    If dict_view_setting.Item("prop_bool_all_books_option_enabled") Then '���ׂẴu�b�N�����̏ꍇ
       
        For Each wbk In Workbooks
            If Windows(wbk.Name).Visible Then 'Visible == ture ��WorkBook�̂ݏ�������
                collection_books_to_operate.Add wbk
            End If
        Next
   
    Else 'AcriveWorkBook�݂̂̏ꍇ
        collection_books_to_operate.Add ActiveWorkbook
       
    End If

    '-----------------------------------------------------</�����Ώ� WorkBook �� collection ��>
    
    Set obj_book_to_activate = ActiveWorkbook '�����I�����ɃA�N�e�B�u�ɂ���u�b�N���L�^

    Dim str_top_left_address_of_view As String
    Dim str_range_address_to_select As String
    
    'View�ݒ胋�[�v
    For Each bk In collection_books_to_operate
       
        bk.Activate
       
        'ribbon
        bool_ribbon_is_minimized = Application.CommandBars.GetPressedMso("MinimizeRibbon")
        If (bool_ribbon_is_minimized <> dict_view_setting.Item("prop_bool_minimize_ribbon_option_enabled")) Then '���{���� �\�� / ��\����Ԃ� �ݒ�l�ƈقȂ�ꍇ
            Application.CommandBars.ExecuteMso "MinimizeRibbon" '���{���\�� / ��\���̐؂�ւ�
        End If
       
        bool_found_sheet_to_activate = False '�A�N�e�B�u�� �ΏۃV�[�g�̑���

        For Each obj_sheet In bk.Sheets
           
            obj_sheet.Activate
            Set range_top_left_of_unfreezed_pain = getTopLeftCellOfUnfreezedPane(ActiveWindow)
            Dim range_top_left_of_specified As Range
            
            On Error GoTo EXCEPTION_VIEW_SET_FAILED
            
            If dict_view_setting.Item("prop_bool_top_left_option_enabled") Then '����Z���ɂ��킹�� View �ݒ�w��̏ꍇ

                str_top_left_address_of_view = range_top_left_of_unfreezed_pain.Address
                str_range_address_to_select = range_top_left_of_unfreezed_pain.Address
                Set range_top_left_of_specified = Range(str_range_address_to_select)

            Else '����Z���ɂ��킹�� View �ݒ�w�肪����(=form �� text box �Ŏw�肵�� Cell Address ���g�p����w��)�̏ꍇ
                
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

            '�A�N�e�B�u�� �ΏۃV�[�g���ǂ����`�F�b�N
            If obj_sheet.Name = dict_view_setting.Item("prop_str_sheet_name_to_activate") Then
                bool_found_sheet_to_activate = True
            End If
           
        Next obj_sheet
       
        '�t�H�[�J�X�V�[�g�̐ݒ�
        If bool_found_sheet_to_activate Then '�t�H�[�J�X���ׂ��V�[�g�����݂���
            bk.Worksheets(dict_view_setting.Item("prop_str_sheet_name_to_activate")).Activate
           
        Else '�t�H�[�J�X���ׂ��V�[�g�����݂��Ȃ�
            bk.Worksheets(1).Activate '�擪�̃V�[�g��I��
           
        End If
   
    Next bk
   
    obj_book_to_activate.Activate '�����J�n���̃u�b�N�� Active �ɖ߂�
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
' String �� Integer �ɕϊ�����
' ���������� TRUE, ���s������ FALSE ��Ԃ�
'
Private Function cIntSafely(ByVal fromThisString As String, ByRef toThisInt As Integer) As Boolean

    Dim ret As Boolean
    
    If Not (IsNumeric(fromThisString)) Then '���l�ɕϊ��s�\�ȏꍇ
        
        ret = False '���s���i�[
    
    Else '���l�ɕϊ��\�ȏꍇ
        
        On Error GoTo EXCEPTION_OVERFLOWED 'CInt() �ŃI�[�o�[�t���[�̏ꍇ�� EXCEPTION_OVERFLOWED �� Go
        toThisInt = CInt(fromThisString) '�w��ϐ��Ɋi�[
        ret = True
        
    End If
    
    cIntSafely = ret '���� / ���s��Ԃ�ԋp
    Exit Function
    
EXCEPTION_OVERFLOWED:
    ret = False '���s���i�[
    cIntSafely = ret
    Exit Function
        
    
End Function

'
'unfreezed �� pain �͈͂̍���Z���̃A�h���X���擾����
'
Private Function getTopLeftCellOfUnfreezedPane(ByVal obj_window As Window) As Range

    Dim px_topLeftCell As Range

    If obj_window.FreezePanes Then 'freeze pain �L���̏ꍇ
    
        
        If obj_window.Panes.Count = 4 Then '���4�����̏ꍇ
            Set p1 = obj_window.Panes(1)
            Set p1_bottomRightCell = getEdgeCellFromRange( _
                rangeObj:=p1.VisibleRange, _
                bottom:=True, _
                right:=True _
            ) 'pane(1)�͈̔͂̉E���̃Z�����擾
            Set px_topLeftCell = obj_window.ActiveSheet.Cells(p1_bottomRightCell.Row + 1, p1_bottomRightCell.Column + 1) 'pane(1)�͈̔͂�1�E����ݒ�
            
        Else '2�����̏ꍇ
        
            If obj_window.SplitRow = 0 Then '���E2�����̏ꍇ
                Set p1 = obj_window.Panes(1)
                Set p1_topRightCell = getEdgeCellFromRange( _
                    rangeObj:=p1.VisibleRange, _
                    bottom:=False, _
                    right:=True _
                ) 'pane(1)�͈̔͂̉E��̃Z�����擾
                Set px_topLeftCell = obj_window.ActiveSheet.Cells(1, p1_topRightCell.Column + 1) 'pane(1)�͈̔͂�1�E��ݒ�
            
            Else '�㉺2�����̏ꍇ (obj_window.SplitColumn = 0 �̏ꍇ)
                Set p1 = obj_window.Panes(1)
                Set p1_bottomLeftCell = getEdgeCellFromRange( _
                    rangeObj:=p1.VisibleRange, _
                    bottom:=True, _
                    right:=False _
                ) 'pane(1)�͈̔͂̍����̃Z�����擾
                Set px_topLeftCell = obj_window.ActiveSheet.Cells(p1_bottomLeftCell.Row + 1, 1) 'pane(1)�͈̔͂�1����ݒ�
                
            End If
        
        End If

    Else
        Set px_topLeftCell = obj_window.ActiveSheet.Cells(1, 1) 'A1 �Z����Ԃ�
        
    End If

    Set getTopLeftCellOfUnfreezedPane = px_topLeftCell

End Function

'
' �����Freezed Pane �����݂���Ɖ��肵���ꍇ�́A
' ���� Pane �̃Z����L�͈�(rows.count, columns.count)�z����Z�o���ĕԂ�
'
Private Function getImaginaryPane1_sRangeCount(ByVal obj_window As Window) As Variant
    
    Dim arr_ret As Variant '�ԋp�l
    
    If obj_window.FreezePanes Then 'freeze pain �L���̏ꍇ
    
        
        If obj_window.Panes.Count = 4 Then '���4�����̏ꍇ
            Set p1 = obj_window.Panes(1)
            arr_ret = Array( _
                p1.VisibleRange.Rows.Count, _
                p1.VisibleRange.Columns.Count _
            )
            
        Else '2�����̏ꍇ
        
            If obj_window.SplitRow = 0 Then '���E2�����̏ꍇ
                Set p1 = obj_window.Panes(1)
                arr_ret = Array( _
                    0, _
                    p1.VisibleRange.Columns.Count _
                )
            
            Else '�㉺2�����̏ꍇ (obj_window.SplitColumn = 0 �̏ꍇ)
                Set p1 = obj_window.Panes(1)
                arr_ret = Array( _
                    p1.VisibleRange.Rows.Count, _
                    0 _
                )
                
            End If
        
        End If

    Else
        arr_ret = Array(0, 0) '�͈�0 �ŕԂ�
        
    End If

    getImaginaryPane1_sRangeCount = arr_ret

End Function
'
' Range�I�u�W�F�N�g�̍���/�E��/����/�E���̃Z����Ԃ�
'
Private Function getEdgeCellFromRange(ByVal rangeObj As Range, ByVal bottom As Boolean, ByVal right As Boolean) As Range
    
    '�ϐ�
    Dim ret As Range
    Dim rowOffset As Long
    Dim colOffset As Long
    
    'Range ���ォ��� Row ���Έʒu�̎Z�o
    If bottom Then '�ŉ����擾�w��̏ꍇ
        rowOffset = rangeObj.Rows.Count - 1
    Else '�ŏ㕔�擾�w��̏ꍇ
        rowOffset = 0
    End If
    
    'Range ���ォ��� Column ���Έʒu�̎Z�o
    If right Then '�ŉE���擾�w��̏ꍇ
        colOffset = rangeObj.Columns.Count - 1
    Else '�ō����擾�w��̏ꍇ
        colOffset = 0
    End If
    
    '�ԋp�l�ݒ�
    Set ret = rangeObj.Parent.Cells( _
        rangeObj.Item(1).Row + rowOffset, _
        rangeObj.Item(1).Column + colOffset _
    )
    
    Set getEdgeCellFromRange = ret '�ԋp

End Function



