Attribute VB_Name = "M_SheetCtl"
Option Explicit

Public Const C_LAST_ROW As Long = 65536
Public Const C_LAST_COL As Long = 256

' ************************************************
' 名称：列番号変換
' 引数：列番号
' 備考：列番号をアルファベットに変換 colConv(2)→B
' ************************************************
Function colConv(ByVal nArgColNo As Long) As String
    Const C_AscA As Long = 65 ' Asc("A")の結果
    Const C_AlfabetCnt As Long = 26 ' アルファベット件数
    If nArgColNo Then
        nArgColNo = nArgColNo - 1
        colConv = colConv(nArgColNo \ C_AlfabetCnt) & Chr(nArgColNo Mod C_AlfabetCnt + C_AscA)
    End If
End Function


' ************************************************
' 名称：最終行取得
' 引数：列番号
' 備考：なし
' ************************************************
Function getLastRowNo(ByRef oArgSheet As Worksheet, ByVal nArgColNo As Long) As Long
    getLastRowNo = oArgSheet.Cells(C_LAST_ROW, nArgColNo).End(xlUp).Row
End Function
Function getLastColNo(ByRef oArgSheet As Worksheet, ByVal nArgRowNo As Long) As Long
    getLastColNo = oArgSheet.Cells(nArgRowNo, C_LAST_COL).End(xlToLeft).Column
End Function

' ************************************************
' 名称：ダブルクォーテーションで文字を囲む
' 引数：処理対象文字列
' 備考：なし
' ************************************************
Function packDblQuote(ByVal sTargetStr As String) As String
    packDblQuote = C_DBL_QUOTE & sTargetStr & C_DBL_QUOTE
End Function
