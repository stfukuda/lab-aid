Attribute VB_Name = "Module_Display"
Option Explicit

' Module_Common 依存
Public EventEnabled As Boolean
Public Const BTN_INSTANT As String = "即時反映ボタン"
Public Const VISIBLE_TOKEN As String = "1:表示" ' E列の入力規則: 空 または 1:表示

'=========================================================================
' メソッド名: InArray
' 概要    : 配列内に指定文字列が含まれるか判定する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - value (String): 判定する値
' - arr (Variant): 1次元配列
' 戻り値 : Boolean - 含まれる場合 True
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Private Function InArray(value As String, arr As Variant) As Boolean
    Dim v

    For Each v In arr
        If StrComp(CStr(v), value, vbTextCompare) = 0 Then
            InArray = True
            Exit Function
        End If
    Next v
End Function

'=========================================================================
' メソッド名: ToggleColumnVisibility
' 概要    : マスタコードとアイテムコードに合致する左右ペア列の表示/非表示を切り替える
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - mastCode (String): マスタコード
' - colCode (String): アイテムコード
' - visibleFlag (Boolean): True=表示, False=非表示
' - firstDataCol (Long): 左側データ開始列
' - lastColLeft (Long): 左側列数
' - mirrorFirstCol (Long): 右側開始列
' - wsMas (Worksheet): マスタシート
' 戻り値 : なし
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Public Sub ToggleColumnVisibility(mastCode As String, colCode As String, visibleFlag As Boolean, _
                                  firstDataCol As Long, lastColLeft As Long, mirrorFirstCol As Long, wsMas As Worksheet)
    Dim colNum As Long, mirrorColIndex As Long

    For colNum = firstDataCol To firstDataCol + lastColLeft - 1
        If StrComp(Trim$(CStr(wsMas.Cells(Row_MasterCode, colNum).Value)), mastCode, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(wsMas.Cells(Row_ColumnCode, colNum).Value)), colCode, vbTextCompare) = 0 Then
            wsMas.Columns(colNum).Hidden = Not visibleFlag
            mirrorColIndex = mirrorFirstCol + (colNum - firstDataCol)
            wsMas.Columns(mirrorColIndex).Hidden = Not visibleFlag
            Exit For
        End If
    Next colNum
End Sub

'=========================================================================
' メソッド名: ToggleColumnGroup
' 概要    : 指定列の表示切替
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - mastCode (String): マスタコード
' - colCode (String): アイテムコード
' - visibleFlag (Boolean): True=表示, False=非表示
' - firstDataCol (Long): 左側データ開始列
' - lastColLeft (Long): 左側列数
' - mirrorFirstCol (Long): 右側開始列
' - wsMas (Worksheet): マスタシート
' 戻り値 : なし
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Private Sub ToggleColumnGroup(mastCode As String, colCode As String, visibleFlag As Boolean, _
                              firstDataCol As Long, lastColLeft As Long, mirrorFirstCol As Long, wsMas As Worksheet)
    ToggleColumnVisibility mastCode, colCode, visibleFlag, firstDataCol, lastColLeft, mirrorFirstCol, wsMas
End Sub

'=========================================================================
' メソッド名: UpdateInstantUpdateButton
' 概要    : 即時反映ボタンの表示/色をセルの値に合わせて更新する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Private Sub UpdateInstantUpdateButton()
    Dim btn As Button
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(SHEET_SET)
    Set btn = ws.Buttons(BTN_INSTANT)

    EventEnabled = CBool(ws.Cells(1, 1).Value)

    If EventEnabled Then
        btn.Caption = "即時反映(ON)"
        btn.Font.Color = vbGreen
    Else
        btn.Caption = "即時反映(OFF)"
        btn.Font.Color = vbRed
    End If
End Sub

'=========================================================================
' メソッド名: SheetChangeHandler
' 概要    : 設定シートの表示フラグ変更に応じて列の表示/非表示を切り替える
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - Target (Range): 変更されたセル範囲
' 戻り値 : なし
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Public Sub SheetChangeHandler(ByVal Target As Range)
    Dim wsSet As Worksheet
    Dim wsMas As Worksheet
    Dim rngChanged As Range
    Dim mastCode As String
    Dim colCode As String
    Dim firstDataCol As Long
    Dim lastCol As Long
    Dim lastColLeft As Long
    Dim mirrorFirstCol As Long
    Dim cell As Range
    Dim visibleFlag As Boolean
    Dim disp As String
    Dim prevEvents As Boolean

    EventEnabled = CBool(Sheets(SHEET_SET).Cells(1, 1).Value)
    If Not EventEnabled Then Exit Sub

    Set wsSet = Sheets(SHEET_SET)
    Set wsMas = Sheets(SHEET_MASTER)
    Set rngChanged = Intersect(Target, wsSet.Range("E4:E" & wsSet.Rows.Count))
    If rngChanged Is Nothing Then Exit Sub

    firstDataCol = 2
    lastCol = wsMas.Cells(Row_MasterCode, wsMas.Columns.Count).End(xlToLeft).Column
    lastColLeft = (lastCol - firstDataCol + 1) \ 2
    mirrorFirstCol = firstDataCol + lastColLeft

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo CleanFail

    For Each cell In rngChanged
        mastCode = Trim$(CStr(wsSet.Cells(cell.Row, "A").Value))
        colCode = Trim$(CStr(wsSet.Cells(cell.Row, "C").Value))
        If mastCode <> "" And colCode <> "" Then
            disp = Replace(Trim$(CStr(wsSet.Cells(cell.Row, "E").Value)), " ", "")
            visibleFlag = (disp = VISIBLE_TOKEN)
            ToggleColumnGroup mastCode, colCode, visibleFlag, _
                               firstDataCol, lastColLeft, mirrorFirstCol, wsMas
        End If
    Next cell

CleanExit:
    Application.EnableEvents = prevEvents
    Exit Sub
CleanFail:
    Application.EnableEvents = prevEvents
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'=========================================================================
' メソッド名: ApplyFlagToMaster
' 概要    : 設定シートの表示フラグをマスタ全体に一括反映する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Public Sub ApplyFlagToMaster()
    Dim wsSet As Worksheet
    Dim wsMas As Worksheet
    Dim mastCode As String
    Dim colCode As String
    Dim lastRowSet As Long
    Dim firstDataCol As Long
    Dim lastCol As Long
    Dim lastColLeft As Long
    Dim mirrorFirstCol As Long
    Dim r As Long
    Dim visibleFlag As Boolean
    Dim disp As String

    Set wsSet = Sheets(SHEET_SET)
    Set wsMas = Sheets(SHEET_MASTER)
    firstDataCol = 2
    lastCol = wsMas.Cells(Row_MasterCode, wsMas.Columns.Count).End(xlToLeft).Column
    lastColLeft = (lastCol - firstDataCol + 1) \ 2
    mirrorFirstCol = firstDataCol + lastColLeft
    lastRowSet = wsSet.Cells(wsSet.Rows.Count, "A").End(xlUp).Row

    For r = 4 To lastRowSet
        mastCode = Trim$(CStr(wsSet.Cells(r, "A").Value))
        colCode = Trim$(CStr(wsSet.Cells(r, "C").Value))
        If mastCode <> "" And colCode <> "" Then
            disp = Replace(Trim$(CStr(wsSet.Cells(r, "E").Value)), " ", "")
            visibleFlag = (disp = VISIBLE_TOKEN)
            ToggleColumnGroup mastCode, colCode, visibleFlag, _
                               firstDataCol, lastColLeft, mirrorFirstCol, wsMas
        End If
    Next r
    MsgBox "マスタ列表示設定を一括反映しました。", vbInformation
End Sub

'=========================================================================
' メソッド名: ToggleInstantUpdate
' 概要    : 即時反映モードのON/OFFを切り替え、ボタンを更新する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Public Sub ToggleInstantUpdate()
    Dim prevEvents As Boolean

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    EventEnabled = Not CBool(Sheets(SHEET_SET).Cells(1, 1).Value)
    Sheets(SHEET_SET).Cells(1, 1).Value = EventEnabled

    Application.EnableEvents = prevEvents

    UpdateInstantUpdateButton
    MsgBox IIf(EventEnabled, "即時反映機能を有効にしました。", "即時反映機能を無効にしました。"), _
           IIf(EventEnabled, vbInformation, vbExclamation)
End Sub

'=========================================================================
' メソッド名: ShowAllColumns
' 概要    : マスタシートの全列を表示する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Public Sub ShowAllColumns()
    Dim wsMas As Worksheet

    Set wsMas = Sheets(SHEET_MASTER)

    wsMas.Columns.Hidden = False
    MsgBox "マスタシートのすべての列を表示しました。", vbInformation
End Sub

'=========================================================================
' メソッド名: ToggleReferenceTables
' 概要    : 引用テーブル系シートの表示/非表示をまとめて切り替える
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Public Sub ToggleReferenceTables()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim anyVisible As Boolean
    Dim i As Long
    Dim newVisible As XlSheetVisibility
    Dim btn As Button
    Dim setWs As Worksheet
    Dim prevActiveName As String
    Dim prevScreenUpdating As Boolean
    Dim msg As String

    sheetNames = Array(SHEET_RELATION, SHEET_QA_INDEX, SHEET_PROD_CONV, SHEET_COM_CONV, _
                       SHEET_RENAISSANCE, SHEET_RANK, SHEET_PROD_COMMON)

    prevActiveName = ActiveSheet.Name
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetNames(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If ws.Visible = xlSheetVisible Then
                anyVisible = True
                Exit For
            End If
        End If
        Set ws = Nothing
    Next i

    newVisible = IIf(anyVisible, xlSheetHidden, xlSheetVisible)

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetNames(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then ws.Visible = newVisible
        Set ws = Nothing
    Next i

    Set setWs = ThisWorkbook.Sheets(SHEET_SET)
    On Error Resume Next
    Set btn = setWs.Buttons("引用テーブル表示ボタン")
    On Error GoTo 0
    If Not btn Is Nothing Then
        If newVisible = xlSheetVisible Then
            btn.Caption = "引用テーブル(表示)"
            btn.Font.Color = vbGreen
        Else
            btn.Caption = "引用テーブル(非表示)"
            btn.Font.Color = vbRed
        End If
    End If

    If newVisible = xlSheetVisible Then
        msg = "引用テーブルを表示しました。"
    Else
        msg = "引用テーブルを非表示にしました。"
    End If
    On Error Resume Next
    ThisWorkbook.Sheets(prevActiveName).Activate
    On Error GoTo 0
    Application.ScreenUpdating = prevScreenUpdating
    MsgBox msg, vbInformation
End Sub
'=========================================================================
' メソッド名: ToggleMasterTables
' 概要    : 各マスタ系シートの表示/非表示をまとめて切り替える
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'=========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================
Public Sub ToggleMasterTables()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim anyVisible As Boolean
    Dim i As Long
    Dim newVisible As XlSheetVisibility
    Dim btn As Button
    Dim setWs As Worksheet
    Dim prevActiveName As String
    Dim prevScreenUpdating As Boolean
    Dim msg As String

    sheetNames = Array("計算式マスタ", "単位マスタ", "試験項目マスタ", "ホルダマスタ", _
                       "文字データマスタ", "プロトコルマスタ2", "プロトコルマスタ3", "規格マスタ", _
                       "サンプルマスタ", "サンプルグループマスタ", "サンプル別試験項目マスタ", "試薬マスタ")

    prevActiveName = ActiveSheet.Name
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetNames(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If ws.Visible = xlSheetVisible Then
                anyVisible = True
                Exit For
            End If
        End If
        Set ws = Nothing
    Next i

    newVisible = IIf(anyVisible, xlSheetHidden, xlSheetVisible)

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetNames(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then ws.Visible = newVisible
        Set ws = Nothing
    Next i

    Set setWs = ThisWorkbook.Sheets(SHEET_SET)
    On Error Resume Next
    Set btn = setWs.Buttons("各マスタ表示ボタン")
    On Error GoTo 0
    If Not btn Is Nothing Then
        If newVisible = xlSheetVisible Then
            btn.Caption = "各マスタ(表示)"
            btn.Font.Color = vbGreen
        Else
            btn.Caption = "各マスタ(非表示)"
            btn.Font.Color = vbRed
        End If
    End If

    If newVisible = xlSheetVisible Then
        msg = "各マスタを表示しました。"
    Else
        msg = "各マスタを非表示にしました。"
    End If

    On Error Resume Next
    ThisWorkbook.Sheets(prevActiveName).Activate
    On Error GoTo 0
    Application.ScreenUpdating = prevScreenUpdating
    MsgBox msg, vbInformation
End Sub