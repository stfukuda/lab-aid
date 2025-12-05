Attribute VB_Name = "Module_MasterManager"
Option Explicit

' Module_Common の定数・行列定義に依存

'=========================================================================='
' メソッド名: MasterInfo (Type)
' 概要    : マスタシートのメタとデータを一時保持する構造体
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - なし
' 戻り値 : Type - MasterInfo
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Public Type MasterInfo
    firstDataCol As Long
    lastColLeft As Long
    lastCol As Long
    mirrorFirstCol As Long
    maxRow As Long
    dataArr As Variant
    mastCodeArr As Variant
    colCodeArr As Variant
End Type

'=========================================================================='
' メソッド名: GetColumnIndex
' 概要    : マスタコードと列コードに一致する左側列の列番号を返す
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - mastCode (String): マスタコード
' - colCode (String): 列コード
' - firstCol (Long): 左側データ開始列
' - totalLeftCols (Long): 左側列数
' - ws (Worksheet): 参照先シート（省略時はマスタシート）
' 戻り値 : Long - 一致列の列番号（見つからない場合は0）
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Public Function GetColumnIndex(mastCode As String, colCode As String, _
                               firstCol As Long, totalLeftCols As Long, Optional ws As Worksheet) As Long
    Dim c As Long

    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets(SHEET_MASTER)
    For c = firstCol To firstCol + totalLeftCols - 1
        If ws.Cells(Row_MasterCode, c).Value = mastCode _
           And ws.Cells(Row_ColumnCode, c).Value = colCode Then
            GetColumnIndex = c
            Exit Function
        End If
    Next
    GetColumnIndex = 0
End Function

'=========================================================================='
' メソッド名: ConvertRestrictionToRegex
' 概要    : 入力制限文字列をVBScript.RegExpのパターンに変換する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - restriction (String): 入力制限の説明文字列
' 戻り値 : String - 正規表現パターン（制限なしの場合は空文字）
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Public Function ConvertRestrictionToRegex(restriction As String) As String
    Dim pattern As String, startPos As Long, endPos As Long
    Dim tmp As String, i As Long, ch As String, safeTmp As String

    If restriction = "" Or restriction = "チェックなし" Then Exit Function
    If InStr(restriction, "英大文字") > 0 Then pattern = pattern & "A-Z"
    If InStr(restriction, "英小文字") > 0 Then pattern = pattern & "a-z"
    If InStr(restriction, "数字") > 0 Then pattern = pattern & "0-9"
    If InStr(restriction, "記号リスト") > 0 Then
        startPos = InStr(restriction, "(")
        endPos = InStr(restriction, ")")
        If startPos > 0 And endPos > startPos Then
            tmp = Replace(Mid(restriction, startPos + 1, endPos - startPos - 1), " ", "")
            For i = 1 To Len(tmp)
                ch = Mid(tmp, i, 1)
                Select Case ch
                    Case "\\", "-", "]", "^", "["
                        safeTmp = safeTmp & "\" & ch
                    Case Else
                        safeTmp = safeTmp & ch
                End Select
            Next i
            pattern = pattern & safeTmp
        End If
    End If
    If pattern <> "" Then ConvertRestrictionToRegex = "^[" & pattern & "]+$"
End Function

'=========================================================================='
' メソッド名: ConditionMatches
' 概要    : 条件タイプに応じて値が条件を満たすか判定する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - value (String): 判定対象値
' - condType (String): 条件種別
' 戻り値 : Boolean - 条件を満たす場合 True
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Public Function ConditionMatches(value As String, condType As String) As Boolean
    Dim firstChar As String

    Select Case condType
        Case "FirstCharNumGE1"
            If value <> "" Then
                firstChar = Left(value, 1)
                If IsNumeric(firstChar) And Val(firstChar) >= 1 Then ConditionMatches = True
            End If
        Case "FirstTwoCharsIs02"
            If value <> "" Then
                If Left(value, 2) = "02" Then ConditionMatches = True
            End If
        Case "FirstCharIsEFX"
            If value <> "" Then
                firstChar = UCase(Left(value, 1))
                If firstChar = "E" Or firstChar = "F" Or firstChar = "X" Then ConditionMatches = True
            End If
        Case "FirstCharIsZ"
            If value <> "" Then
                If UCase(Left(value, 1)) = "Z" Then ConditionMatches = True
            End If
        Case "FirstCharIsE"
            If value <> "" Then
                If UCase(Left(value, 1)) = "E" Then ConditionMatches = True
            End If
        Case "FirstCharIsAT"
            If value <> "" Then
                firstChar = UCase(Left(value, 1))
                If firstChar = "A" Or firstChar = "T" Then ConditionMatches = True
            End If
        Case "FirstCharIsR"
            If value <> "" Then
                If UCase(Left(value, 1)) = "R" Then ConditionMatches = True
            End If
        Case "FirstCharIs1"
            If value <> "" Then
                If Left(value, 1) = "1" Then ConditionMatches = True
            End If
        Case "ContainsC"
            If value <> "" Then
                If InStr(1, value, "C", vbTextCompare) > 0 Then ConditionMatches = True
            End If
        Case "FirstCharNotH"
            If value <> "" Then
                If UCase(Left(value, 1)) <> "H" Then ConditionMatches = True
            End If
    End Select
End Function

'=========================================================================='
' メソッド名: AppendError
' 概要    : 右側のエラー列にメッセージを書き込み、件数を加算する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - ws (Worksheet): マスタシート
' - colNum (Long): 左側の対象列番号
' - mirrorFirstCol (Long): 右側の先頭列番号
' - firstDataCol (Long): 左側の先頭列番号
' - r (Long): 行番号
' - errMsg (String): エラーメッセージ
' - errCount() (Long): 列ごとのエラー件数配列
' - errArr (Variant): 右側エラー出力用配列
' 戻り値 : なし
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Public Sub AppendError(ws As Worksheet, ByVal colNum As Long, mirrorFirstCol As Long, _
                       firstDataCol As Long, r As Long, errMsg As String, errCount() As Long, errArr As Variant)
    Dim mirrorColIndex As Long

    mirrorColIndex = colNum - firstDataCol + 1
    If errArr(r, mirrorColIndex) = "" Then
        errArr(r, mirrorColIndex) = errMsg
        errCount(colNum - firstDataCol + 1) = errCount(colNum - firstDataCol + 1) + 1
    Else
        errArr(r, mirrorColIndex) = errArr(r, mirrorColIndex) & "," & errMsg
    End If
End Sub

'=========================================================================='
' メソッド名: CheckRequiredFields
' 概要    : 必須列（見出し先頭が*）の未入力をチェックし、エラーを記録する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - wsMas (Worksheet): マスタシート
' - info (MasterInfo): 事前読み込み済み情報
' - errCount() (Long): 列ごとのエラー件数配列
' - errArr (Variant): 右側エラー出力用配列
' 戻り値 : なし
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Private Sub CheckRequiredFields(wsMas As Worksheet, info As MasterInfo, errCount() As Long, errArr As Variant)
    Dim colNum As Long
    Dim r As Long

    For colNum = 1 To info.lastColLeft
        If Left(info.dataArr(Row_ColumnName_Top, colNum), 1) = "*" Then
            For r = Row_DataStart To info.maxRow
                If Trim(info.dataArr(r, colNum)) = "" Then
                    AppendError wsMas, info.firstDataCol + colNum - 1, info.mirrorFirstCol, info.firstDataCol, _
                                r, "未入力", errCount, errArr
                End If
            Next r
        End If
    Next colNum
End Sub

'=========================================================================='
' メソッド名: CheckLengthAndRestriction
' 概要    : 文字数制限と入力制限に基づき各セルを検証する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - wsMas (Worksheet): マスタシート
' - info (MasterInfo): 事前読み込み済み情報
' - errCount() (Long): 列ごとのエラー件数配列
' - errArr (Variant): 右側エラー出力用配列
' 戻り値 : なし
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Private Sub CheckLengthAndRestriction(wsMas As Worksheet, info As MasterInfo, errCount() As Long, errArr As Variant)
    Dim colNum As Long
    Dim r As Long
    Dim maxLen As Long
    Dim lengthText As String
    Dim restriction As String
    Dim pattern As String
    Dim colType As String
    Dim regex As Object
    Dim valTmp As String

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True

    For colNum = 1 To info.lastColLeft
        colType = Trim(info.dataArr(Row_ColumnType, colNum))
        If colType = "フラグ" Then GoTo NextCol

        maxLen = 0
        If info.dataArr(Row_Length, colNum) <> "" Then
            lengthText = Replace(info.dataArr(Row_Length, colNum), "文字", "")
            lengthText = Replace(lengthText, " ", "")
            If IsNumeric(lengthText) Then maxLen = CLng(lengthText)
        End If

        restriction = Trim(info.dataArr(Row_InputRestriction, colNum))
        If restriction = "" Or restriction = "チェックなし" Then
            pattern = ""
        ElseIf restriction = "数字" Then
            pattern = "^-?\d+$"
        Else
            pattern = ConvertRestrictionToRegex(restriction)
        End If

        For r = Row_DataStart To info.maxRow
            valTmp = Trim(info.dataArr(r, colNum))
            If maxLen > 0 And Len(valTmp) > maxLen Then
                AppendError wsMas, info.firstDataCol + colNum - 1, info.mirrorFirstCol, info.firstDataCol, _
                            r, "文字数超過 (" & Len(valTmp) & "/" & maxLen & ")", errCount, errArr
            End If
            If pattern <> "" And valTmp <> "" Then
                regex.Pattern = pattern
                If Not regex.Test(valTmp) Then
                    AppendError wsMas, info.firstDataCol + colNum - 1, info.mirrorFirstCol, info.firstDataCol, _
                                r, "入力制限違反 (" & valTmp & ")", errCount, errArr
                End If
            End If
        Next r
NextCol:
    Next colNum
End Sub

'=========================================================================='
' メソッド名: CheckConditionalRequired
' 概要    : 条件付き必須ルールに従い未入力をチェックする
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - wsMas (Worksheet): マスタシート
' - info (MasterInfo): 事前読み込み済み情報
' - errCount() (Long): 列ごとのエラー件数配列
' - errArr (Variant): 右側エラー出力用配列
' 戻り値 : なし
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Private Sub CheckConditionalRequired(wsMas As Worksheet, info As MasterInfo, errCount() As Long, errArr As Variant)
    Dim rules As Variant
    Dim rule As Variant
    Dim mastCode As String
    Dim condColCode As String
    Dim reqCodes As Variant
    Dim colCond As Long
    Dim reqColNums() As Long
    Dim i As Long
    Dim j As Long
    Dim r As Long
    Dim valCond As String

    rules = GetConditionalRules()

    For Each rule In rules
        mastCode = CStr(rule(0))
        condColCode = CStr(rule(1))

        colCond = 0
        For i = 1 To info.lastColLeft
            If info.mastCodeArr(1, i) = mastCode And info.colCodeArr(1, i) = condColCode Then
                colCond = i: Exit For
            End If
        Next i
        If colCond = 0 Then GoTo NextRule

        If Not IsArray(rule(3)) Then GoTo NextRule
        reqCodes = rule(3)
        ReDim reqColNums(LBound(reqCodes) To UBound(reqCodes))
        For i = LBound(reqCodes) To UBound(reqCodes)
            reqColNums(i) = 0
            For j = 1 To info.lastColLeft
                If info.mastCodeArr(1, j) = mastCode And info.colCodeArr(1, j) = CStr(reqCodes(i)) Then
                    reqColNums(i) = j: Exit For
                End If
            Next j
        Next i

        For r = Row_DataStart To info.maxRow
            valCond = Trim(info.dataArr(r, colCond))
            If ConditionMatches(valCond, CStr(rule(2))) Then
                For i = LBound(reqColNums) To UBound(reqColNums)
                    If reqColNums(i) > 0 And Trim(info.dataArr(r, reqColNums(i))) = "" Then
                        AppendError wsMas, info.firstDataCol + reqColNums(i) - 1, info.mirrorFirstCol, info.firstDataCol, _
                                    r, "未入力", errCount, errArr
                    End If
                Next i
            End If
        Next r
NextRule:
    Next rule
End Sub

'=========================================================================='
' メソッド名: CheckDATLENRange
' 概要    : EDTF種別に応じてDATLENの範囲・必須をチェックする
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - wsMas (Worksheet): マスタシート
' - info (MasterInfo): 事前読み込み済み情報
' - errCount() (Long): 列ごとのエラー件数配列
' - errArr (Variant): 右側エラー出力用配列
' 戻り値 : なし
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Private Sub CheckDATLENRange(wsMas As Worksheet, info As MasterInfo, errCount() As Long, errArr As Variant)
    Dim colEDTF As Long
    Dim colDATLEN As Long
    Dim i As Long
    Dim r As Long
    Dim edtVal As String
    Dim lenVal As String

    For i = 1 To info.lastColLeft
        If info.mastCodeArr(1, i) = "LDBMTST" And info.colCodeArr(1, i) = "EDTF" Then colEDTF = i
        If info.mastCodeArr(1, i) = "LDBMTST" And info.colCodeArr(1, i) = "DATLEN" Then colDATLEN = i
    Next i
    If colEDTF = 0 Or colDATLEN = 0 Then Exit Sub

    For r = Row_DataStart To info.maxRow
        edtVal = UCase(Trim(info.dataArr(r, colEDTF)))
        lenVal = Trim(info.dataArr(r, colDATLEN))
        Select Case Left(edtVal, 1)
            Case "E", "F"
                If lenVal = "" Then
                    AppendError wsMas, info.firstDataCol + colDATLEN - 1, info.mirrorFirstCol, info.firstDataCol, _
                                r, "未入力", errCount, errArr
                ElseIf Not IsNumeric(lenVal) Or Int(Val(lenVal)) <> Val(lenVal) _
                   Or Val(lenVal) < 1 Or Val(lenVal) > 10 Then
                    AppendError wsMas, info.firstDataCol + colDATLEN - 1, info.mirrorFirstCol, info.firstDataCol, _
                                r, "1～10の整数が必要", errCount, errArr
                End If
            Case "X"
                If lenVal = "" Then
                    AppendError wsMas, info.firstDataCol + colDATLEN - 1, info.mirrorFirstCol, info.firstDataCol, _
                                r, "未入力", errCount, errArr
                ElseIf Not IsNumeric(lenVal) Or Int(Val(lenVal)) <> Val(lenVal) _
                   Or Val(lenVal) < -9 Or Val(lenVal) > 10 Then
                    AppendError wsMas, info.firstDataCol + colDATLEN - 1, info.mirrorFirstCol, info.firstDataCol, _
                                r, "-9～10の整数が必要", errCount, errArr
                End If
        End Select
    Next r
End Sub

'=========================================================================='
' メソッド名: CheckKeyConsistency
' 概要    : 太字列をキーとして、同一キー内で他列の値差異を検出する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - wsMas (Worksheet): マスタシート
' - info (MasterInfo): 事前読み込み済み情報
' - errCount() (Long): 列ごとのエラー件数配列
' - errArr (Variant): 右側エラー出力用配列
' 戻り値 : なし
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Private Sub CheckKeyConsistency(wsMas As Worksheet, info As MasterInfo, errCount() As Long, errArr As Variant)
    Dim isBold() As Boolean
    Dim dict As Object
    Dim mc As Variant
    Dim keyStr As String
    Dim c As Variant
    Dim r As Long
    Dim kv As Variant
    Dim rowIdx As Variant
    Dim colNum As Long
    Dim keyRowsMap As Object
    Dim baseRow As Long
    Dim keyVal As String
    Dim colsWithDiff As Collection
    Dim baseVal As String
    Dim hasDiff As Boolean
    Dim curVal As String

    ReDim isBold(1 To info.lastColLeft)
    For colNum = info.firstDataCol To info.firstDataCol + info.lastColLeft - 1
        isBold(colNum - info.firstDataCol + 1) = wsMas.Cells(Row_ColumnName_Top, colNum).Font.Bold
    Next colNum

    Set dict = CreateObject("Scripting.Dictionary")
    For colNum = 1 To info.lastColLeft
        keyStr = CStr(info.mastCodeArr(1, colNum))
        If Not dict.Exists(keyStr) Then
            dict(keyStr) = Array(New Collection, New Collection)
        End If
        If isBold(colNum) Then
            dict(keyStr)(0).Add colNum
        Else
            dict(keyStr)(1).Add colNum
        End If
    Next colNum

    For Each mc In dict.Keys
        Set keyRowsMap = CreateObject("Scripting.Dictionary")
        For r = Row_DataStart To info.maxRow
            keyVal = ""
            For Each c In dict(mc)(0)
                keyVal = keyVal & "|" & Trim(info.dataArr(r, c))
            Next c
            If Replace(keyVal, "|", "") <> "" Then
                If Not keyRowsMap.Exists(keyVal) Then Set keyRowsMap(keyVal) = New Collection
                keyRowsMap(keyVal).Add r
            End If
        Next r

        For Each kv In keyRowsMap.Keys
            Set colsWithDiff = New Collection
            baseRow = CLng(keyRowsMap(kv).Item(1))
            For Each c In dict(mc)(1)
                baseVal = Trim(info.dataArr(baseRow, c))
                hasDiff = False
                For Each rowIdx In keyRowsMap(kv)
                    curVal = Trim(info.dataArr(CLng(rowIdx), c))
                    If Not (baseVal = "" And curVal = "") Then
                        If baseVal <> curVal Then hasDiff = True: Exit For
                    End If
                Next rowIdx
                If hasDiff Then colsWithDiff.Add c
            Next c

            If colsWithDiff.Count > 0 Then
                For Each rowIdx In keyRowsMap(kv)
                    For Each c In colsWithDiff
                        AppendError wsMas, info.firstDataCol + c - 1, info.mirrorFirstCol, info.firstDataCol, _
                                    CLng(rowIdx), "キー不整合", errCount, errArr
                    Next c
                Next rowIdx
            End If
        Next kv
    Next mc
End Sub

'=========================================================================='
' メソッド名: ValidateMasterData
' 概要    : マスタシートのデータを読み込み、各種バリデーションを実行する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - なし
' 戻り値 : なし
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Public Sub ValidateMasterData()
    Dim wsMas As Worksheet
    Dim info As MasterInfo
    Dim errCount() As Long
    Dim errArr As Variant
    Dim colNum As Long

    Set wsMas = Sheets(SHEET_MASTER)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    info = LoadMasterData(wsMas)
    If info.lastColLeft <= 0 Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "マスタシートに検証対象列がありません。", vbExclamation
        Exit Sub
    End If
    ReDim errCount(1 To info.lastColLeft)
    ReDim errArr(Row_DataStart To info.maxRow, 1 To info.lastColLeft)

    wsMas.Range(wsMas.Cells(Row_ErrorCount, info.mirrorFirstCol), wsMas.Cells(Row_ErrorCount, info.lastCol)).ClearContents
    wsMas.Range(wsMas.Cells(Row_DataStart, info.mirrorFirstCol), wsMas.Cells(wsMas.Rows.Count, info.lastCol)).ClearContents

    CheckRequiredFields wsMas, info, errCount, errArr
    CheckLengthAndRestriction wsMas, info, errCount, errArr
    CheckConditionalRequired wsMas, info, errCount, errArr
    CheckDATLENRange wsMas, info, errCount, errArr
    CheckKeyConsistency wsMas, info, errCount, errArr

    wsMas.Range(wsMas.Cells(Row_DataStart, info.mirrorFirstCol), wsMas.Cells(info.maxRow, info.lastCol)).Value = errArr

    For colNum = info.firstDataCol To info.firstDataCol + info.lastColLeft - 1
        wsMas.Cells(Row_ErrorCount, info.mirrorFirstCol + (colNum - info.firstDataCol)).Value = _
            IIf(errCount(colNum - info.firstDataCol + 1) > 0, errCount(colNum - info.firstDataCol + 1) & "件", "")
    Next colNum

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'=========================================================================='
' メソッド名: LoadMasterData
' 概要    : マスタシートのメタ情報とデータをまとめて読み込む
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - ws (Worksheet): 読み込み対象シート（省略時はマスタ）
' 戻り値 : MasterInfo - 先頭列/列数/最大行/各種配列を含む構造体
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Public Function LoadMasterData(Optional ws As Worksheet) As MasterInfo
    Dim info As MasterInfo

    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets(SHEET_MASTER)

    info.firstDataCol = 2
    info.lastCol = ws.Cells(Row_MasterCode, ws.Columns.Count).End(xlToLeft).Column
    info.lastColLeft = (info.lastCol - info.firstDataCol + 1) \ 2
    info.mirrorFirstCol = info.firstDataCol + info.lastColLeft
    info.maxRow = ws.Cells(ws.Rows.Count, info.firstDataCol).End(xlUp).Row
    If info.maxRow < Row_DataStart Then info.maxRow = Row_DataStart

    info.mastCodeArr = ws.Range(ws.Cells(Row_MasterCode, info.firstDataCol), _
                                ws.Cells(Row_MasterCode, info.firstDataCol + info.lastColLeft - 1)).Value
    info.colCodeArr = ws.Range(ws.Cells(Row_ColumnCode, info.firstDataCol), _
                               ws.Cells(Row_ColumnCode, info.firstDataCol + info.lastColLeft - 1)).Value
    info.dataArr = ws.Range(ws.Cells(1, info.firstDataCol), _
                            ws.Cells(info.maxRow, info.firstDataCol + info.lastColLeft - 1)).Value
    LoadMasterData = info
End Function

'=========================================================================='
' メソッド名: GetColumnFromInfo
' 概要    : MasterInfoからマスタコード/列コードに一致する列番号を取得する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - info (MasterInfo): 取得済みマスタ情報
' - mastCode (String): マスタコード
' - colCode (String): 列コード
' 戻り値 : Long - 列番号（見つからない場合は0）
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Public Function GetColumnFromInfo(info As MasterInfo, mastCode As String, colCode As String) As Long
    Dim i As Long
    For i = 1 To info.lastColLeft
        If info.mastCodeArr(1, i) = mastCode And info.colCodeArr(1, i) = colCode Then
            GetColumnFromInfo = info.firstDataCol + i - 1
            Exit Function
        End If
    Next i
End Function

'=========================================================================='
' メソッド名: WriteBackLeftSide
' 概要    : 左側データ配列をシートに一括書き戻す
' 作成日  : 2025/12/03
' 作成者  : dc11449
'=========================================================================='
' パラメーター:
' - ws (Worksheet): 書き込み対象シート
' - info (MasterInfo): 取得済みマスタ情報
' - dataArr (Variant): 左側データ配列
' 戻り値 : なし
'=========================================================================='
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'=========================================================================='
Public Sub WriteBackLeftSide(ws As Worksheet, info As MasterInfo, dataArr As Variant)
    Dim lastRow As Long

    If IsEmpty(dataArr) Then Exit Sub
    lastRow = UBound(dataArr, 1)
    ws.Range(ws.Cells(1, info.firstDataCol), ws.Cells(lastRow, info.firstDataCol + info.lastColLeft - 1)).Value = dataArr
End Sub



