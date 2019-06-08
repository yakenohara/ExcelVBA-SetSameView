Attribute VB_Name = "SetSameView"
'<License>------------------------------------------------------------
'
' Copyright (c) 2018 Shinnosuke Yakenohara
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

    '変数宣言
    Dim s As Object
    Dim dispMag As Integer
    Dim closerToA1 As Boolean
    Dim focus As String
    Dim cursor As String
    Dim focusSht As String
    Dim books As Collection
    Dim toActiveBk As Workbook
    
    SetSameViewFormMod.Show
    
    'フォーム状態確認
    If SetSameViewFormMod.status <> vbOK Then
        Exit Sub
    
    ElseIf Not (IsNumeric(SetSameViewFormMod.TextBoxMag)) Then
        MsgBox "指定表示倍率は数値として無効です"
        Exit Sub
    
    End If
                                      
    '表示倍率の取得
    On Error GoTo whenOverFlowOccurred
    dispMag = CInt(SetSameViewFormMod.TextBoxMag)
    
    closerToA1 = SetSameViewFormMod.CheckBoxCloserToA1
    
    'フォーカス位置取得
    focus = SetSameViewFormMod.TextBoxFocus
    cursor = SetSameViewFormMod.TextBoxCursor
    focusSht = SetSameViewFormMod.ComboBoxFocusShtNames.Text
    Application.ScreenUpdating = False
    
    Set books = New Collection
    
    If SetSameViewFormMod.CheckBoxEveryBook.Value Then 'すべてのブック処理の場合
        
        For Each wbk In Workbooks
            If Windows(wbk.Name).Visible Then 'Visible == ture なWorkBookのみ処理する
                books.Add wbk
            End If
        Next
    
    Else 'AcriveWorkBookのみの場合
        books.Add ActiveWorkbook
        
    End If
    
    
    'カーソル位置修正・表示倍率変更
    On Error GoTo whenZoomFailed
    Set toActiveBk = ActiveWorkbook
    
    For Each bk In books
        
        bk.Activate
        
        shtFound = False
        
        For Each s In bk.Sheets
            
            s.Activate
            
            ActiveWindow.Zoom = dispMag
            
            If closerToA1 Then
            
                If ActiveWindow.FreezePanes Then 'ウィンドウ枠固定が有効の場合
                    Set p1 = ActiveWindow.Panes(1)
                    Set p1_bottomRightCell = p1.VisibleRange.Item(p1.VisibleRange.Count)
                    Set p4_topLeftCell = Cells(p1_bottomRightCell.Row + 1, p1_bottomRightCell.Column + 1)
                    
                    focus = p4_topLeftCell.Address
                    cursor = p4_topLeftCell.Address
                
                Else
                    focus = "A1"
                    cursor = "A1"
                
                End If
                
            End If
            
            ActiveWindow.ScrollRow = Range(focus).Row
            ActiveWindow.ScrollColumn = Range(focus).Column
            Range(cursor).Select
            
            
            If s.Name = focusSht Then
                shtFound = True
            End If
            
        Next s
        
        'フォーカスシートの設定
        If shtFound Then 'フォーカスすべきシートが存在する
            bk.Worksheets(focusSht).Activate
            
        Else 'フォーカスすべきシートが存在しない
            bk.Worksheets(1).Activate '先頭のシートを選択
            
        End If
    
    Next bk
    
    toActiveBk.Activate
    
    Application.ScreenUpdating = True
    MsgBox "Done!"
    
    Exit Sub
    
whenOverFlowOccurred:
    MsgBox "Application.InputBoxメソッドで例外" & vbLf & _
           "オーバーフローの可能性があります"
    
    Exit Sub
    
whenZoomFailed:
    Application.ScreenUpdating = True
    MsgBox "Window.Zoomプロパティで例外" & vbLf & _
           "指定表示倍率かカーソル書式が不正な可能性があります"
    
    Exit Sub
    
End Sub







