Attribute VB_Name = "Module_PreProcess"
Option Explicit

' Module_Common と Module_MasterManager のヘルパーに依存

'==========================================================================
' メソッド名: GetNewCommonNo
' 概要    : 共通項目変換シートから新しい共通項目番号を取得する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'==========================================================================
' パラメーター:
' - wsConv (Worksheet): 共通項目変換シート
' - prodCode (String): 製品コード
' - mgNo (String): 管理基準番号
' - oldCommonNo (String): 旧共通項目番号
' 戻り値 : String - 新しい共通項目番号（見つからない場合は空文字）
'==========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'==========================================================================
Private Function GetNewCommonNo(wsConv As Worksheet, prodCode As String, mgNo As String, oldCommonNo As String) As String
    Dim rr As Long, lastRow As Long
    lastRow = wsConv.Cells(wsConv.Rows.Count, "A").End(xlUp).Row
    For rr = 2 To lastRow
        If Trim(wsConv.Cells(rr, "A").Value) = prodCode _
           And Trim(wsConv.Cells(rr, "B").Value) = mgNo _
           And Trim(wsConv.Cells(rr, "C").Value) = oldCommonNo Then
            GetNewCommonNo = wsConv.Cells(rr, "E").Value
            Exit Function
        End If
    Next rr
    GetNewCommonNo = ""
End Function

'==========================================================================
' メソッド名: GetItemCodeFromProd
' 概要    : 製品コードから品目コードを取得する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'==========================================================================
' パラメーター:
' - wsConv (Worksheet): 製品コード変換シート
' - prodCode (String): 製品コード
' 戻り値 : String - 品目コード（見つからない場合は空文字）
'==========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'==========================================================================
Private Function GetItemCodeFromProd(wsConv As Worksheet, prodCode As String) As String
    Dim rr As Long, lastRow As Long
    lastRow = wsConv.Cells(wsConv.Rows.Count, "A").End(xlUp).Row

    For rr = 2 To lastRow
        If Trim(wsConv.Cells(rr, "A").Value) = prodCode Then
            GetItemCodeFromProd = wsConv.Cells(rr, "B").Value
            Exit Function
        End If
    Next rr

    GetItemCodeFromProd = ""
End Function

'==========================================================================
' メソッド名: ApplyConstantsFromSetting
' 概要    : 設定シートの定数値をマスタシートの該当列に一括反映する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'==========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'==========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'==========================================================================
Public Sub ApplyConstantsFromSetting()
    Dim wsSet As Worksheet, wsMas As Worksheet
    Dim firstDataCol As Long, lastCol As Long, lastColLeft As Long
    Dim mirrorFirstCol As Long, lastRowSet As Long
    Dim mastCode As String, colCode As String, constVal As Variant
    Dim colIndex As Long
    Dim lastRowMas As Long
    Dim r As Long

    Set wsSet = Sheets(SHEET_SET)
    Set wsMas = Sheets(SHEET_MASTER)

    firstDataCol = 2
    lastCol = wsMas.Cells(Row_MasterCode, wsMas.Columns.Count).End(xlToLeft).Column
    lastColLeft = (lastCol - firstDataCol + 1) \ 2
    mirrorFirstCol = firstDataCol + lastColLeft
    lastRowSet = wsSet.Cells(wsSet.Rows.Count, "A").End(xlUp).Row
    lastRowMas = wsMas.Cells(wsMas.Rows.Count, firstDataCol).End(xlUp).Row

    For r = 2 To lastRowSet
        mastCode = wsSet.Cells(r, "A").Value
        colCode = wsSet.Cells(r, "C").Value
        constVal = wsSet.Cells(r, "F").Value

        If mastCode <> "" And colCode <> "" And constVal <> "" Then
            colIndex = Module_MasterManager.GetColumnIndex(mastCode, colCode, firstDataCol, lastColLeft, wsMas)
            If colIndex > 0 Then
                wsMas.Range(wsMas.Cells(Row_DataStart, colIndex), wsMas.Cells(lastRowMas, colIndex)).Value = constVal
            End If
        End If
    Next r
End Sub

'==========================================================================
' メソッド名: ApplyRelations
' 概要    : リレーションシートの定義に従い引用元列の値を引用先列へコピーする
' 作成日  : 2025/12/03
' 作成者  : dc11449
'==========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'==========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'==========================================================================
Public Sub ApplyRelations()
    Dim wsRel As Worksheet, wsMas As Worksheet
    Dim lastRowRel As Long, r As Long
    Dim destMast As String, destCol As String
    Dim srcMast As String, srcCol As String
    Dim destIndex As Long, srcIndex As Long
    Dim firstDataCol As Long, lastCol As Long, lastColLeft As Long, mirrorFirstCol As Long
    Dim lastRowMas As Long

    Set wsRel = Sheets(SHEET_RELATION)
    Set wsMas = Sheets(SHEET_MASTER)

    firstDataCol = 2
    lastCol = wsMas.Cells(Row_MasterCode, wsMas.Columns.Count).End(xlToLeft).Column
    lastColLeft = (lastCol - firstDataCol + 1) \ 2
    mirrorFirstCol = firstDataCol + lastColLeft

    lastRowRel = wsRel.Cells(wsRel.Rows.Count, "A").End(xlUp).Row
    lastRowMas = wsMas.Cells(wsMas.Rows.Count, firstDataCol).End(xlUp).Row

    Application.ScreenUpdating = False

    For r = 2 To lastRowRel
        destMast = wsRel.Cells(r, "A").Value
        destCol = wsRel.Cells(r, "B").Value
        srcMast = wsRel.Cells(r, "C").Value
        srcCol = wsRel.Cells(r, "D").Value

        destIndex = Module_MasterManager.GetColumnIndex(destMast, destCol, firstDataCol, lastColLeft, wsMas)
        srcIndex = Module_MasterManager.GetColumnIndex(srcMast, srcCol, firstDataCol, lastColLeft, wsMas)

        If destIndex > 0 And srcIndex > 0 Then
            wsMas.Range(wsMas.Cells(Row_DataStart, destIndex), wsMas.Cells(lastRowMas, destIndex)).Value = _
                wsMas.Range(wsMas.Cells(Row_DataStart, srcIndex), wsMas.Cells(lastRowMas, srcIndex)).Value
        End If
    Next r

    Application.ScreenUpdating = True
End Sub

'==========================================================================
' メソッド名: ApplyRenaissanceQuote
' 概要    : ルネサンス品目シートから情報をマスタシートへコピーする
' 作成日  : 2025/12/03
' 作成者  : dc11449
'==========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'==========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'==========================================================================
Public Sub ApplyRenaissanceQuote()
    Dim wsRen As Worksheet, wsMas As Worksheet
    Dim dictRen As Object
    Dim lastRowRen As Long, r As Long
    Dim firstDataCol As Long, lastColLeft As Long
    Dim colItemCode As Long, colName1 As Long, colName2 As Long, colCompany As Long
    Dim lastRowMas As Long
    Dim i As Long

    Set wsRen = Sheets(SHEET_RENAISSANCE)
    Set wsMas = Sheets(SHEET_MASTER)

    firstDataCol = 2
    lastColLeft = (wsMas.Cells(Row_MasterCode, wsMas.Columns.Count).End(xlToLeft).Column - firstDataCol + 1) \ 2

    colItemCode = Module_MasterManager.GetColumnIndex("LDBMSMP", "SMPINFX31", firstDataCol, lastColLeft, wsMas)
    colName2 = Module_MasterManager.GetColumnIndex("LDBMSMP", "SMPINF21", firstDataCol, lastColLeft, wsMas)
    colName1 = Module_MasterManager.GetColumnIndex("LDBMSMP", "SMPINF22", firstDataCol, lastColLeft, wsMas)
    colCompany = Module_MasterManager.GetColumnIndex("LDBMSMP", "SMPINF24", firstDataCol, lastColLeft, wsMas)

    Set dictRen = CreateObject("Scripting.Dictionary")
    lastRowRen = wsRen.Cells(wsRen.Rows.Count, "B").End(xlUp).Row

    For r = 2 To lastRowRen
        Dim renItemCode As String
        renItemCode = Trim(wsRen.Cells(r, "B").Value)
        If renItemCode <> "" Then
            dictRen(renItemCode) = Array(wsRen.Cells(r, "D").Value, _
                                         wsRen.Cells(r, "E").Value, _
                                         wsRen.Cells(r, "A").Value)
        End If
    Next r

    lastRowMas = wsMas.Cells(wsMas.Rows.Count, colItemCode).End(xlUp).Row

    For i = Row_DataStart To lastRowMas
        Dim mItemCode As String
        mItemCode = Trim(wsMas.Cells(i, colItemCode).Value)

        If dictRen.Exists(mItemCode) Then
            Dim arrVal As Variant
            arrVal = dictRen(mItemCode)
            wsMas.Cells(i, colName1).Value = arrVal(0)
            wsMas.Cells(i, colName2).Value = arrVal(1)
            wsMas.Cells(i, colCompany).Value = arrVal(2)
        End If
    Next i
End Sub

'==========================================================================
' メソッド名: ApplyStandard
' 概要    : 管理基準ランクシートを元に規格値・上下限をマスタへ反映する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'==========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'==========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'==========================================================================
Public Sub ApplyStandard()
    Dim wsRank As Worksheet, wsProdConv As Worksheet, wsComConv As Worksheet, wsMas As Worksheet
    Dim dictStd As Object
    Dim info As MasterInfo
    Dim arrData As Variant
    Dim idxItemCode As Long, idxTSTCD As Long, idxStdNo As Long, idxCommonNo As Long
    Dim idxLL1REF As Long, idxUL1REF As Long, idxLL1 As Long, idxUL1 As Long, idxTSTINF03 As Long
    Dim lastRowMas As Long
    Dim key As String
    Dim i As Long

    Set wsRank = Sheets(SHEET_RANK)
    Set wsComConv = Sheets(SHEET_COM_CONV)
    Set wsProdConv = Sheets(SHEET_PROD_CONV)
    Set wsMas = Sheets(SHEET_MASTER)

    info = Module_MasterManager.LoadMasterData(wsMas)
    arrData = info.dataArr

    Set dictStd = CreateObject("Scripting.Dictionary")

    ' 管理基準ランクを辞書化（キー: 品目コード|管理基準番号|共通項目番号）
    Dim lastRowRank As Long, r As Long
    lastRowRank = wsRank.Cells(wsRank.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRowRank
        Dim prodCode As String, mgNo As String, commonNo As String, useCommonNo As String
        Dim itemCode As String, stdFlag As Long, llVal As Variant, ulVal As Variant, specVal As Variant

        prodCode = Trim(wsRank.Cells(r, "B").Value)
        mgNo = Trim(wsRank.Cells(r, "C").Value)
        commonNo = Trim(wsRank.Cells(r, "D").Value)

        useCommonNo = GetNewCommonNo(wsComConv, prodCode, mgNo, commonNo)
        If useCommonNo = "" Then useCommonNo = commonNo

        itemCode = GetItemCodeFromProd(wsProdConv, prodCode)
        If itemCode = "" Then GoTo SkipRank

        stdFlag = wsRank.Cells(r, "F").Value
        llVal = wsRank.Cells(r, "G").Value
        ulVal = wsRank.Cells(r, "H").Value
        specVal = wsRank.Cells(r, "I").Value

        key = itemCode & "|" & mgNo & "|" & useCommonNo
        dictStd(key) = Array(mgNo, stdFlag, llVal, ulVal, specVal)
SkipRank:
    Next r

    idxItemCode = Module_MasterManager.GetColumnFromInfo(info, "LDBMSMP", "SMPINFX31") - info.firstDataCol + 1
    idxTSTCD = Module_MasterManager.GetColumnFromInfo(info, "LDBMTST", "TSTCD") - info.firstDataCol + 1
    idxStdNo = Module_MasterManager.GetColumnFromInfo(info, "LDBMSMP", "管理基準番号") - info.firstDataCol + 1
    idxCommonNo = Module_MasterManager.GetColumnFromInfo(info, "LDBMSMP", "共通項目番号") - info.firstDataCol + 1
    idxLL1REF = Module_MasterManager.GetColumnFromInfo(info, "LDBMSPC", "LL1REF") - info.firstDataCol + 1
    idxUL1REF = Module_MasterManager.GetColumnFromInfo(info, "LDBMSPC", "UL1REF") - info.firstDataCol + 1
    idxLL1 = Module_MasterManager.GetColumnFromInfo(info, "LDBMSPC", "LL1") - info.firstDataCol + 1
    idxUL1 = Module_MasterManager.GetColumnFromInfo(info, "LDBMSPC", "UL1") - info.firstDataCol + 1
    idxTSTINF03 = Module_MasterManager.GetColumnFromInfo(info, "LDBMTST", "TSTINF03") - info.firstDataCol + 1

    If idxItemCode <= 0 Or idxTSTCD <= 0 Or idxStdNo <= 0 Or idxCommonNo <= 0 Then Exit Sub

    lastRowMas = info.maxRow

    For i = Row_DataStart To lastRowMas
        Dim mItemCode As String, tstCode As String, mCommonNo As String
        Dim mgNoM As String, arrVal As Variant

        mItemCode = Trim(arrData(i, idxItemCode))

        tstCode = Trim(arrData(i, idxTSTCD))
        If Len(tstCode) >= 6 Then
            mCommonNo = Mid(tstCode, 3, 4)
        Else
            mCommonNo = ""
        End If

        If Right(tstCode, 4) <> "0001" Then GoTo NextRowStd

        mgNoM = Trim(arrData(i, idxStdNo))
        key = mItemCode & "|" & mgNoM & "|" & mCommonNo

        If dictStd.Exists(key) Then
            arrVal = dictStd(key)

            arrData(i, idxStdNo) = arrVal(0)
            arrData(i, idxCommonNo) = mCommonNo

            Select Case arrVal(1)
                Case 1
                    If idxLL1REF > 0 Then arrData(i, idxLL1REF) = arrVal(2)
                    If idxLL1 > 0 And Trim(arrData(i, idxLL1)) = "" Then arrData(i, idxLL1) = arrVal(2)
                Case 2
                    If idxUL1REF > 0 Then arrData(i, idxUL1REF) = arrVal(3)
                    If idxUL1 > 0 And Trim(arrData(i, idxUL1)) = "" Then arrData(i, idxUL1) = arrVal(3)
                Case 3
                    If idxLL1REF > 0 Then arrData(i, idxLL1REF) = arrVal(2)
                    If idxUL1REF > 0 Then arrData(i, idxUL1REF) = arrVal(3)
                    If idxLL1 > 0 And Trim(arrData(i, idxLL1)) = "" Then arrData(i, idxLL1) = arrVal(2)
                    If idxUL1 > 0 And Trim(arrData(i, idxUL1)) = "" Then arrData(i, idxUL1) = arrVal(3)
            End Select

            If idxTSTINF03 > 0 Then arrData(i, idxTSTINF03) = arrVal(4)
        End If
NextRowStd:
    Next i

    Module_MasterManager.WriteBackLeftSide wsMas, info, arrData
End Sub

'==========================================================================
' メソッド名: ApplyEnglishTestInfo
' 概要    : 製品共通項目シートの英字検査内容をマスタへ反映する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'==========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'==========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'==========================================================================
Public Sub ApplyEnglishTestInfo()
    Dim wsRank As Worksheet, wsProd As Worksheet, wsProdConv As Worksheet, wsComConv As Worksheet, wsMas As Worksheet
    Dim dictEngSource As Object, dictEngApply As Object
    Dim info As MasterInfo
    Dim arrData As Variant
    Dim idxItemCode As Long, idxTSTCD As Long, idxStdNo As Long
    Dim idxTSTINF01 As Long, idxTSTINF02 As Long
    Dim lastRowMas As Long
    Dim i As Long

    Set wsRank = Sheets(SHEET_RANK)
    Set wsProd = Sheets(SHEET_PROD_COMMON)
    Set wsComConv = Sheets(SHEET_COM_CONV)
    Set wsProdConv = Sheets(SHEET_PROD_CONV)
    Set wsMas = Sheets(SHEET_MASTER)

    info = Module_MasterManager.LoadMasterData(wsMas)
    arrData = info.dataArr

    Set dictEngSource = CreateObject("Scripting.Dictionary")
    Set dictEngApply = CreateObject("Scripting.Dictionary")

    ' 製品共通項目: 英字検査内容辞書（キー: 事業部|共通項目番号）
    Dim lastRowProd As Long, r As Long
    lastRowProd = wsProd.Cells(wsProd.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRowProd
        Dim deptEng As String, commonEng As String
        deptEng = Trim(wsProd.Cells(r, "A").Value)
        commonEng = Trim(wsProd.Cells(r, "B").Value)
        If commonEng <> "" Then
            dictEngSource(deptEng & "|" & commonEng) = Array(wsProd.Cells(r, "F").Value, wsProd.Cells(r, "G").Value)
        End If
    Next r

    ' 管理基準ランクを元にマスタキーへ英字内容をマッピング
    Dim lastRowRank As Long
    lastRowRank = wsRank.Cells(wsRank.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRowRank
        Dim dept As String, prodCode As String, mgNo As String, commonNo As String, useCommonNo As String
        Dim itemCode As String, engKey As String, key As String
        Dim eng1 As String, eng2 As String

        dept = Trim(wsRank.Cells(r, "A").Value)
        prodCode = Trim(wsRank.Cells(r, "B").Value)
        mgNo = Trim(wsRank.Cells(r, "C").Value)
        commonNo = Trim(wsRank.Cells(r, "D").Value)

        useCommonNo = GetNewCommonNo(wsComConv, prodCode, mgNo, commonNo)
        If useCommonNo = "" Then useCommonNo = commonNo

        itemCode = GetItemCodeFromProd(wsProdConv, prodCode)
        If itemCode = "" Then GoTo SkipRankEng

        eng1 = ""
        eng2 = ""
        engKey = dept & "|" & commonNo
        If dictEngSource.Exists(engKey) Then
            eng1 = dictEngSource(engKey)(0)
            eng2 = dictEngSource(engKey)(1)
        End If

        key = itemCode & "|" & mgNo & "|" & useCommonNo
        dictEngApply(key) = Array(eng1, eng2)
SkipRankEng:
    Next r

    idxItemCode = Module_MasterManager.GetColumnFromInfo(info, "LDBMSMP", "SMPINFX31") - info.firstDataCol + 1
    idxTSTCD = Module_MasterManager.GetColumnFromInfo(info, "LDBMTST", "TSTCD") - info.firstDataCol + 1
    idxStdNo = Module_MasterManager.GetColumnFromInfo(info, "LDBMSMP", "管理基準番号") - info.firstDataCol + 1
    idxTSTINF01 = Module_MasterManager.GetColumnFromInfo(info, "LDBMTST", "TSTINF01") - info.firstDataCol + 1
    idxTSTINF02 = Module_MasterManager.GetColumnFromInfo(info, "LDBMTST", "TSTINF02") - info.firstDataCol + 1

    If idxItemCode <= 0 Or idxTSTCD <= 0 Or idxStdNo <= 0 Then Exit Sub

    lastRowMas = info.maxRow

    For i = Row_DataStart To lastRowMas
        Dim mItemCode As String, tstCode As String, mCommonNo As String
        Dim mgNoM As String, arrEng As Variant, keyEng As String

        mItemCode = Trim(arrData(i, idxItemCode))

        tstCode = Trim(arrData(i, idxTSTCD))
        If Len(tstCode) >= 6 Then
            mCommonNo = Mid(tstCode, 3, 4)
        Else
            mCommonNo = ""
        End If

        If Right(tstCode, 4) <> "0001" Then GoTo NextRowEng

        mgNoM = Trim(arrData(i, idxStdNo))
        keyEng = mItemCode & "|" & mgNoM & "|" & mCommonNo

        If dictEngApply.Exists(keyEng) Then
            arrEng = dictEngApply(keyEng)
            If idxTSTINF01 > 0 Then arrData(i, idxTSTINF01) = arrEng(0)
            If idxTSTINF02 > 0 Then arrData(i, idxTSTINF02) = arrEng(1)
        End If
NextRowEng:
    Next i

    Module_MasterManager.WriteBackLeftSide wsMas, info, arrData
End Sub

'==========================================================================
' メソッド名: ApplyTestItemNumbers
' 概要    : マスタシート(LDBMTST)の試験項目に上から順に試験項目番号(TSTNO)を付与する
' 作成日  : 2025/12/03
' 作成者  : dc11449
'==========================================================================
' パラメーター:
' - なし
' 戻り値 : なし
'==========================================================================
' 修正履歴 :
' - 2025/12/03: dc11449 - 新規作成
'==========================================================================
Public Sub ApplyTestItemNumbers()
    Dim wsMas As Worksheet
    Dim info As MasterInfo
    Dim arrData As Variant
    Dim idxTSTCD As Long, idxTSTNO As Long
    Dim lastRowMas As Long
    Dim i As Long, seq As Long

    Set wsMas = Sheets(SHEET_MASTER)

    info = Module_MasterManager.LoadMasterData(wsMas)
    arrData = info.dataArr

    idxTSTCD = Module_MasterManager.GetColumnFromInfo(info, "LDBMTST", "TSTCD") - info.firstDataCol + 1
    idxTSTNO = Module_MasterManager.GetColumnFromInfo(info, "LDBMTST", "TSTNO") - info.firstDataCol + 1
    If idxTSTCD <= 0 Or idxTSTNO <= 0 Then Exit Sub

    lastRowMas = info.maxRow

    seq = 1
    For i = Row_DataStart To lastRowMas
        If Trim(arrData(i, idxTSTCD)) <> "" Then
            arrData(i, idxTSTNO) = seq
            seq = seq + 1
        Else
            arrData(i, idxTSTNO) = ""
        End If
    Next i

    Module_MasterManager.WriteBackLeftSide wsMas, info, arrData
End Sub


