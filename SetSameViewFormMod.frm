VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetSameViewFormMod 
   Caption         =   "SetSameView"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
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

Public hided_in As Integer

Const DEFAULT_ZOOM_LEVEL As Integer = 100
Const DEFAULT_ADDRESS_TO_SELECT As String = "A1"

'<Life cycle of Form>-----------------------------------------------------------------------------

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

    Dim cautionMessage As String: cautionMessage = "�Z�����I��(�J�[�\��)����Ă��܂���B" & vbLf & vbLf & _
                                                   "�Z�����I��(�J�[�\��)����Ă��Ȃ��ꍇ�̃J�[�\���ʒu�́A" & vbLf & _
                                                   "���݂̃t�H�[�J�X�ʒu����ɂȂ�܂��B"
    
    Me.txtbx_zoom_level = ActiveWindow.Zoom
    
    Me.chkbx_top_left.Value = False

    Me.txtbx_top_left_address_of_view = ActiveWindow.VisibleRange(1).Address(False, False)
    
    If (Selection Is Nothing) Or (Not (TypeName(Selection) = "Range")) Then
        
        retVal = MsgBox(Prompt:=cautionMessage, Buttons:=vbExclamation)
        
        Me.txtbx_range_address_to_select = Me.txtbx_top_left_address_of_view
        
    Else
        
        Me.txtbx_range_address_to_select = Selection.Address(False, False)
        
    End If
    
    topLeftAddress = DEFAULT_ADDRESS_TO_SELECT
    If ActiveWindow.FreezePanes Then 'freeze pain �L���̏ꍇ
    
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
    
    Call buttn_set_sht_current_Click

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
    
    '������
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
    
    Me.chkbx_all_books.Value = False
    
End Sub

'----------------------------------------------------------------------------</Common>
