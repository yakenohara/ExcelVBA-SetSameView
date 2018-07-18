Attribute VB_Name = "SetSameView"
Sub SetSameView()

    '変数宣言
    Dim s As Object
    Dim dispMag As Integer
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
            ActiveWindow.ScrollRow = Range(focus).Row
            ActiveWindow.ScrollColumn = Range(focus).Column
            Range(focus).Select
            ActiveWindow.Zoom = dispMag
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

