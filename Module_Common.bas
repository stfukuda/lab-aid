Attribute VB_Name = "Module_Common"
Option Explicit

' 共通定義
Public Const SHEET_SET As String = "設定"
Public Const SHEET_MASTER As String = "マスタ"
Public Const SHEET_RELATION As String = "リレーション"
Public Const SHEET_RENAISSANCE As String = "ルネサンス品目"
Public Const SHEET_RANK As String = "管理基準ランク"
Public Const SHEET_PROD_COMMON As String = "製品共通項目"
Public Const SHEET_COM_CONV As String = "共通項目変換"
Public Const SHEET_PROD_CONV As String = "製品コード変換"
Public Const SHEET_QA_INDEX As String = "品証事業部索引"

Public Enum MasterRows
    Row_Header1 = 1
    Row_ErrorCount = 2
    Row_MasterCode = 3
    Row_ColumnCode = 4
    Row_ColumnName_Top = 5
    Row_ColumnName_Under = 6
    Row_Length = 7
    Row_ColumnType = 8
    Row_InputRestriction = 9
    Row_DataStart = 11
End Enum

' 条件付き必須チェック専用ルール
' 要素: {マスタコード, 条件列コード, 条件タイプ, 必須列コード配列}
Public Function GetConditionalRules() As Variant
    Dim col As Collection
    Set col = New Collection

    col.Add Array("LDBMSMP", "SMPINFX15", "FirstCharNumGE1", Array("SMPINFX12", "SMPINFX14"))
    col.Add Array("LDBMSMP", "SMPINFX20", "FirstCharNumGE1", Array("SMPINFX17", "SMPINFX19"))
    col.Add Array("LDBMSMP", "SMPINFX27", "FirstCharIs1", Array("SMPINFX28"))
    col.Add Array("LDBMSMP", "SMPINFX43", "FirstCharNumGE1", Array("SMPINFX40", "SMPINFX41"))
    col.Add Array("LDBMHLD", "HLDINFX39", "FirstTwoCharsIs02", Array("HLDINFX38"))
    col.Add Array("LDBMHLD", "DATSRC0", "FirstCharIsE", Array("EXP", "RECKEY1"))
    col.Add Array("LDBMHLD", "DATSRC0", "FirstCharIsAT", Array("RECKEY1"))
    col.Add Array("LDBMHLD", "DATSRC0", "FirstCharIsR", Array("RECKEY1", "RECKEY2"))
    col.Add Array("LDBMTST", "EDTF", "FirstCharIsEFX", Array("RNDF"))
    col.Add Array("LDBMTST", "EDTF", "FirstCharIsZ", Array("EXP", "EXPCDR"))
    col.Add Array("LDBMDAT", "SPCCHKF", "ContainsC", Array("DAT1"))
    col.Add Array("LDVMREAG", "REAGINFX01", "FirstCharIs1", Array("REAGINFX02", "REAGINFX03"))
    col.Add Array("LDBMPRT", "DATSRC0", "FirstCharNotH", Array("DATCNT", "DATCHK"))

    GetConditionalRules = CollectionToArray(col)
End Function

Private Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long

    ReDim arr(0 To col.Count - 1)

    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i

    CollectionToArray = arr
End Function
