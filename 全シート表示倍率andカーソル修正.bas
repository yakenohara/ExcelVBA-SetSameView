Attribute VB_Name = "全シート表示倍率andカーソル修正"
Sub 全シート表示倍率andカーソル修正()

    '変数宣言
    Dim s As Object
    Dim defaultSheet As Object
    Dim dispMag As Integer
    Dim focus As String
    
    表示倍率andカーソル設定.Show
    
    'フォーム状態確認
    If 表示倍率andカーソル設定.status <> vbOK Then
        Exit Sub
    
    ElseIf Not (IsNumeric(表示倍率andカーソル設定.TextBoxMag)) Then
        MsgBox "指定表示倍率は数値として無効です"
        Exit Sub
    
    End If
                                      
    '表示倍率の取得
    On Error GoTo whenOverFlowOccurred
    dispMag = CInt(表示倍率andカーソル設定.TextBoxMag)
    
    'フォーカス位置取得
    focus = 表示倍率andカーソル設定.TextBoxFocus
    
    Application.ScreenUpdating = False
    
    
    
    'カーソル位置修正・表示倍率変更
    On Error GoTo whenZoomFailed
    Set defaultSheet = ActiveSheet
    For Each s In ActiveWorkbook.Sheets
        s.Activate
        ActiveWindow.ScrollRow = Range(focus).Row
        ActiveWindow.ScrollColumn = Range(focus).Column
        Range(focus).Select
        ActiveWindow.Zoom = dispMag
    Next s
    Worksheets(1).Activate
    
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

