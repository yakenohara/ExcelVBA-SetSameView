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

    '�ϐ��錾
    Dim obj_sheet As Object
    Dim int_zoom_level As Integer
    Dim bool_top_left_option_enabled As Boolean
    Dim str_top_left_address_of_view As String
    Dim str_range_address_to_select As String
    Dim str_sheet_name_to_activate As String
    Dim collection_opened_books As Collection
    Dim obj_book_to_activate As Workbook
    
    SetSameViewFormMod.Show
    
    '�t�H�[����Ԋm�F
    If SetSameViewFormMod.hided_in <> vbOK Then
        Exit Sub
    
    ElseIf Not (IsNumeric(SetSameViewFormMod.txtbx_zoom_level.Value)) Then 'zoom level �����l�^�łȂ��ꍇ
        retOfMsgBox = MsgBox("Invalid Zoom level type:`" & TypeName(SetSameViewFormMod.txtbx_zoom_level.Value) & "` specified", vbCritical)
        Exit Sub '�I��
    
    End If
                                      
    '�\���{���̎擾
    On Error GoTo C_INT_FUNC_OVERFLOWED
    int_zoom_level = CInt(SetSameViewFormMod.txtbx_zoom_level.Value)
    
    bool_top_left_option_enabled = SetSameViewFormMod.chkbx_top_left
    
    '�t�H�[�J�X�ʒu�擾
    str_top_left_address_of_view = SetSameViewFormMod.txtbx_top_left_address_of_view
    str_range_address_to_select = SetSameViewFormMod.txtbx_range_address_to_select
    str_sheet_name_to_activate = SetSameViewFormMod.cmbbx_sheet_name_to_activate.Text
    Application.ScreenUpdating = False
    
    Set collection_opened_books = New Collection
    
    If SetSameViewFormMod.chkbx_all_books.Value Then '���ׂẴu�b�N�����̏ꍇ
        
        For Each wbk In Workbooks
            If Windows(wbk.Name).Visible Then 'Visible == ture ��WorkBook�̂ݏ�������
                collection_opened_books.Add wbk
            End If
        Next
    
    Else 'AcriveWorkBook�݂̂̏ꍇ
        collection_opened_books.Add ActiveWorkbook
        
    End If
    
    
    '�J�[�\���ʒu�C���E�\���{���ύX
    On Error GoTo ZOOM_FAILED
    Set obj_book_to_activate = ActiveWorkbook
    
    For Each bk In collection_opened_books
        
        bk.Activate
        
        shtFound = False
        
        For Each obj_sheet In bk.Sheets
            
            obj_sheet.Activate
            
            ActiveWindow.Zoom = int_zoom_level
            
            If bool_top_left_option_enabled Then
            
                If ActiveWindow.FreezePanes Then '�E�B���h�E�g�Œ肪�L���̏ꍇ
                    
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
                
                    str_top_left_address_of_view = px_topLeftCell.Address
                    str_range_address_to_select = px_topLeftCell.Address
                    
                Else
                    str_top_left_address_of_view = "A1"
                    str_range_address_to_select = "A1"
                
                End If
                
            End If
            
            ActiveWindow.ScrollRow = Range(str_top_left_address_of_view).Row
            ActiveWindow.ScrollColumn = Range(str_top_left_address_of_view).Column
            Range(str_range_address_to_select).Select
            
            
            If obj_sheet.Name = str_sheet_name_to_activate Then
                shtFound = True
            End If
            
        Next obj_sheet
        
        '�t�H�[�J�X�V�[�g�̐ݒ�
        If shtFound Then '�t�H�[�J�X���ׂ��V�[�g�����݂���
            bk.Worksheets(str_sheet_name_to_activate).Activate
            
        Else '�t�H�[�J�X���ׂ��V�[�g�����݂��Ȃ�
            bk.Worksheets(1).Activate '�擪�̃V�[�g��I��
            
        End If
    
    Next bk
    
    obj_book_to_activate.Activate
    
    Application.ScreenUpdating = True
    MsgBox "Done!"
    
    Exit Sub
    
C_INT_FUNC_OVERFLOWED:
    retOfMsg = MsgBox( _
        "Cannot cast into Integer specified zoom level:`" & str(SetSameViewFormMod.txtbx_zoom_level.Value) & "`", _
        vbCritical _
    )
    
    Exit Sub
    
ZOOM_FAILED:
    Application.ScreenUpdating = True
    retOfMsg = MsgBox( _
        "Exception occurred. As a cause, Specified display magnification or cursor format may be invalid", _
        vbCritical _
    )
    
    Exit Sub
    
End Sub

