VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetSameViewFormMod 
   Caption         =   "SetSameView"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   OleObjectBlob   =   "SetSameViewFormMod.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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

Const DEFAULT_ZOOM_LEVEL As Integer = 100
Const DEFAULT_ADDRESS_TO_SELECT As String = "A1"
Const DEFAULT_MINIMIZE_RIBBON As Boolean = True

Private ended_in As Integer

'<Controller>-----------------------------------------------------------------------------

'
' �t�H�[����\�����ă��[�U�̑I����e��ԋp����
'
' �ԋp�l�^�� MsgBox �֐��Ɠ���(��) �^�E�Ӗ��Ƃ��A�ȉ�3��ނ݂̂��g�p����
'
' | Constant | Value | Description                                                      |
' | -------- | ----- | ---------------------------------------------------------------  |
' | vbOK     | 1     | OK����                                                           |
' | vbCancel | 2     | Cancel����                                                       |
' | vbAbort  | 3     | �E�B���h�E�E�� `�~` �N���b�N�������� Alt + F4 �ŃE�B�h�E�N���[�Y |
'
' ��
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
'
'
Public Function showForm()
    
    ended_in = vbAbort '�E�B���h�E�E�� `�~` �N���b�N��������
        'Alt + F4 �ŃE�B�h�E�N���[�Y�̏ꍇ�͂��̐��l��Ԃ�
        
    Me.Show
    showForm = ended_in '���[�U�[�I����e�̕ԋp
    
End Function

'----------------------------------------------------------------------------</Controller>

'<Life cycle of Form>-----------------------------------------------------------------------------

'
' FormObject �� load ��
' (�Ăяo�������W���[����`SetSameViewFormMod`�ɃA�N�Z�X�������ɁAload�ς݂łȂ������ꍇ�̂�)
' �Ɏ��s�����
'
Private Sub UserForm_Initialize()
    
    Call setDefault
    
End Sub

Private Sub UserForm_Activate()
    '�ϐ�
    Dim didFound As Boolean
    
    '�V�[�g���Đݒ�
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

    'todo �O���t������\�����Ă���V�[�g��\�����̏ꍇ�ɃR�P��
    Me.txtbx_top_left_address_of_view = ActiveWindow.VisibleRange(1).Address(False, False)
    
    If (Selection Is Nothing) Or (Not (TypeName(Selection) = "Range")) Then
    
        str_range_address_to_select = ""
    
        Set selectionRange = getRangeFromSelectionObj(Selection)
        
        If (selectionRange Is Nothing) Then 'Selection �� cell ��L�̈�̎Z�o���ł��Ȃ������ꍇ
            retVal = MsgBox( _
                Prompt:= _
                    "Any cell or range is not selected. " & vbCrLf & _
                    "top left cell address of active window `" & Me.txtbx_top_left_address_of_view & "` will be set.", _
                Buttons:=vbExclamation _
            )
            
            str_range_address_to_select = Me.txtbx_top_left_address_of_view.Value
            
        Else 'Selection �� cell ��L�̈�̎Z�o���ł����ꍇ
            
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
    If ActiveWindow.FreezePanes Then 'freeze pain �L���̏ꍇ
    
        'unfreezed pain �͈͂̂̍���Z���̃A�h���X���Z�o
        
        Dim px_topLeftCell As Range
        
        If ActiveWindow.Panes.Count = 4 Then '���4�����̏ꍇ
            Set p1 = ActiveWindow.Panes(1)
            Set p1_bottomRightCell = getEdgeCellFromRange( _
                rangeObj:=p1.VisibleRange, _
                bottom:=True, _
                right:=True _
            ) 'pane(1)�͈̔͂̉E���̃Z�����擾
            Set px_topLeftCell = Cells(p1_bottomRightCell.Row + 1, p1_bottomRightCell.Column + 1) 'pane(1)�͈̔͂�1�E����ݒ�
            
        Else '2�����̏ꍇ
        
            If ActiveWindow.SplitRow = 0 Then '���E2�����̏ꍇ
                Set p1 = ActiveWindow.Panes(1)
                Set p1_topRightCell = getEdgeCellFromRange( _
                    rangeObj:=p1.VisibleRange, _
                    bottom:=False, _
                    right:=True _
                ) 'pane(1)�͈̔͂̉E��̃Z�����擾
                Set px_topLeftCell = Cells(1, p1_topRightCell.Column + 1) 'pane(1)�͈̔͂�1�E��ݒ�
            
            Else '�㉺2�����̏ꍇ (ActiveWindow.SplitColumn = 0 �̏ꍇ)
                Set p1 = ActiveWindow.Panes(1)
                Set p1_bottomLeftCell = getEdgeCellFromRange( _
                    rangeObj:=p1.VisibleRange, _
                    bottom:=True, _
                    right:=False _
                ) 'pane(1)�͈̔͂̍����̃Z�����擾
                Set px_topLeftCell = Cells(p1_bottomLeftCell.Row + 1, 1) 'pane(1)�͈̔͂�1����ݒ�
                
            End If
        
        End If
        
        topLeftAddress = px_topLeftCell.Address(False, False)
        
    End If
    
    If topLeftAddress = Me.txtbx_range_address_to_select And _
        topLeftAddress = Me.txtbx_top_left_address_of_view Then '�\���͈͍��オ�I������Ă���ꍇ
        
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
    
    Call buttn_set_sht_current_Click '���ݑI�𒆂̃V�[�g����I��
    
    '���{���̕\�� / ��\����Ԃ̔��f
    bool_ribbon_is_minimized = Application.CommandBars.GetPressedMso("MinimizeRibbon")
    Me.chkbx_minimize_ribbon.Value = bool_ribbon_is_minimized

End Sub

'
' `OK` button
' ���̊֐����o�C���h���� Object (CommandButton Object) �ɂ� `Default` �v���p�e�B�� True ��ݒ肵�Ă���
' ���ׁ̈A����CommandButton Object�Ƀt�H�[�J�X���Ȃ��ꍇ�� enter �L�[�������Ă����̊֐��͔��΂���

'
Private Sub buttn_ok_Click()
    ended_in = vbOK
    Me.Hide
    
End Sub

'
' `CANCEL` button
'
' NOTE
' ���̊֐����o�C���h���� Object (CommandButton Object) �ɂ� `Cancel` �v���p�e�B�� True ��ݒ肵�Ă���
' ���ׁ̈AEsc �L�[���������A�{�^���Ƀt�H�[�J�X������Ƃ��� enter �L�[�������Ă����̊֐��͔��΂���
'
Private Sub buttn_cancel_Click()
    ended_in = vbCancel
    Me.Hide
    
End Sub

'----------------------------------------------------------------------------</GUI Events>

'<Common>-----------------------------------------------------------------------------

'
' �f�t�H���g�ݒ�𔽉f������
'
Private Sub setDefault()
    
    '������
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
' �I�u�W�F�N�g��Cell�I��͈͂� Range object �ɂ��ĕԂ�
' �擾�ł��Ȃ������ꍇ�� Nothing ��Ԃ�
'
Private Function getRangeFromSelectionObj(ByVal selectionObj As Object) As Variant

    Dim ret As Variant
    
    If selectionObj Is Nothing Then
        Set ret = Nothing ' Nothing ��Ԃ�
    
    ElseIf (TypeName(selectionObj)) = "Range" Then ' Range �I�u�W�Fy�N�g�̏ꍇ
        Set ret = selectionObj '���̂܂ܕԂ�
    
    Else ' Range �I�u�W�Fy�N�g�łȂ��ꍇ
        On Error GoTo TOP_LEFT_CELL_IS_NOT_DEFINED
        'TopLeftCell, BottomRightCell property ���g���Ĕ͈͂��擾����
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

'----------------------------------------------------------------------------</Common>


