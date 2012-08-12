VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sheetTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 列定義
Enum COL_Test
    ダミー
    空欄
    関数名
    備考1
    備考2
    備考3
    想定結果
    実行結果
    判定
    引数開始
End Enum

' 行定義
Enum ROW_Test
    ダミー
    表題
    空欄
    ヘッダ
    値開始
End Enum


' ************************************************
' 名称：テスト実施
' 引数：なし
' ************************************************
Private Sub cb_TestRun_Click()
    Application.ScreenUpdating = False
    Call testRun
    Application.ScreenUpdating = True
End Sub
Private Sub testRun()
    Dim nRow As Long
    Dim nLastRow As Long
    Dim nLastCol As Long
    
On Error GoTo ERR
    nLastRow = getLastRowNo(sheetTest, COL_Test.関数名)
    With sheetTest
        For nRow = ROW_Test.値開始 To nLastRow
            ' 最終列位置取得
            nLastCol = getLastColNo(sheetTest, nRow)
        
            ' 着色初期化
            With .Range(.Cells(nRow, COL_Test.引数開始), .Cells(nRow, C_LAST_COL))
                .Font.Underline = False
                .Interior.ColorIndex = xlNone
            End With
        
            ' 式組立
            Dim sFormula As String
            sFormula = "=" & .Cells(nRow, COL_Test.関数名).Value & "("
            sFormula = sFormula & ""
            
            ' 引数組立
            Dim nCol As Long
            Dim sDelim As String
            nCol = COL_Test.引数開始
            sDelim = ""
            
            For nCol = COL_Test.引数開始 To nLastCol
                Dim sArgument As String
                sArgument = .Cells(nRow, nCol)
                If sArgument = ")" Then Exit For
                Call editCell(nRow, nCol)
                sFormula = sFormula & sDelim & packDblQuote(sArgument)
                sDelim = ","
            Next
            sFormula = sFormula & ")"

            ' 式反映
            .Cells(nRow, COL_Test.実行結果).Formula = sFormula
            
            ' 下線を引いてスペースを認識可能にする
            .Cells(nRow, COL_Test.想定結果).Font.Underline = True
            .Cells(nRow, COL_Test.実行結果).Font.Underline = True
            
            ' 判定設定
            Dim sFormulaEqual As String
            Dim sCStr As String
            sCStr = C_DBL_QUOTE & C_DBL_QUOTE & " & "
            sFormulaEqual = "(" & sCStr & colConv(COL_Test.実行結果) & nRow _
                & "=" & sCStr & colConv(COL_Test.想定結果) & nRow & ")"
            .Cells(nRow, COL_Test.判定).Formula _
                = "=IF(ISERR(" & colConv(COL_Test.実行結果) & nRow & ")," _
                & "FALSE," & sFormulaEqual & ")"
        Next
    End With
    
OK:
    Exit Sub
ERR:
    Dim sMsg As String
    sMsg = "cb_TestRun_Click:" & vbNewLine & ERR.Number & ":" & ERR.Description & vbNewLine & "nRow>" & nRow
    Debug.Print (sMsg)
    Call MsgBox(sMsg)
End Sub
Private Sub editCell(nArgRow As Long, nArgCol As Long)
    Dim oCell As Range
    Set oCell = sheetTest.Cells(nArgRow, nArgCol)
    With oCell
        .Font.Underline = True
        .Interior.ColorIndex = 34
    End With
End Sub
