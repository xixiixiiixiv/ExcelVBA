VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Enum COL_FileList
    ダミー
    全文
    フルパス
    フォルダ
    ファイル名
    拡張子
    行
    位置
    文字コード
    内容
    列数
    
    ' 任意指定
    読込パス = 2
End Enum

Enum ROW_FileList
    ダミー
    設定
    備考
    ヘッダ
    結果開始
End Enum
        
Enum E_LineKind
    info
    Folder
    file
End Enum
        
Private Const C_SPLIT_CHR As String = "@"
Private moFSO As FileSystemObject
Private mnCount As Long ' 検出件数
Private mnMaxDepth As Long ' サブフォルダ最大の深さ
Private msStartPath As String ' 検索開始パス
        
        
' ヘッダ設定
Private Sub setHeader()
    Call sheetReader.Cells(ROW_FileList.ヘッダ, 1).Activate
    Call setVal(全文, "全文")
    Call setVal(フルパス, "フルパス")
    Call setVal(フォルダ, "フォルダ")
    Call setVal(ファイル名, "ファイル名")
    Call setVal(拡張子, "拡張子")
    Call setVal(行, "行")
    Call setVal(位置, "位置")
    Call setVal(文字コード, "文字コード")
    Call setVal(内容, "内容")
End Sub


' 「開始」ボタン押下時処理
Private Sub cb_Run_Click()
    ' クリップボードの内容を検査
    If Application.ClipboardFormats(1) <> xlClipboardFormatText Then
        MsgBox ("テキストがコピーされていません。" & vbNewLine & "sakuraのgrep結果を全文コピーして「貼付け」ボタンを押してください。")
        Exit Sub
    End If
    
    Application.ScreenUpdating = False ' 画面更新停止し高速化
    Call Initialize ' 初期化
    
    ' クリップボードより貼付け
    sheetReader.Cells(ROW_FileList.結果開始, COL_FileList.全文).PasteSpecial
    
    Call analyzeText ' 解析実施
    
    ' 体裁調整
    Call setHeader ' ヘッダ設定
    Call setAutoFilter ' オートフィルタ
    Call setPageSetup ' 印刷設定
    
    Application.ScreenUpdating = True ' 画面更新停止を解除
    MsgBox ("完了")
End Sub
    
    
' 値書込簡素化
Private Sub setVal(ByVal nArgCol As COL_FileList, ByVal sValue As String)
    ActiveCell.Offset(0, nArgCol - 1).Value = sValue
End Sub
Private Sub setLink(ByVal nArgCol As COL_FileList, ByVal sPath As String)
    ActiveCell.Offset(0, nArgCol - 1).Formula = "=HYPERLINK(" & C_DBL_QUOTE & sPath & C_DBL_QUOTE & ")"
End Sub
Private Sub delVal(ByVal nArgCol As COL_FileList, ByVal sValue As String)
    ActiveCell.Offset(0, nArgCol - 1).Clear
End Sub
    
    
' 初期化
Private Sub Initialize()
    With sheetReader
        ' 既存結果消去
        With .Range(.Cells(ROW_FileList.ヘッダ, 1), .Cells(C_LAST_ROW, C_LAST_COL))
            .Clear ' 消去
            .NumberFormatLocal = "@"
            .Font.Size = 9
            .Font.Name = "ＭＳ ゴシック"
            .VerticalAlignment = xlTop
            .WrapText = False
        End With
        
        ' 既存結果消去
        Call setFormat(全文)
        Call setFormat(フルパス)
        Call setFormat(フォルダ)
        Call setFormat(内容)
        Call setFormat(行, "#")
        Call setFormat(位置, "#")
        
        .Columns(COL_FileList.内容).WrapText = True
        
        ' ヘッダ行書式
        With .Rows(ROW_FileList.ヘッダ)
            .Font.Bold = True
        End With
    End With
End Sub
        
        
' 書式設定
Private Sub setFormat(ByVal nArgCol As COL_FileList, Optional ByVal sArgFormat As String = "G/標準")
    With sheetReader
        With .Range(.Cells(ROW_FileList.結果開始, nArgCol), .Cells(C_LAST_ROW, nArgCol))
            .NumberFormatLocal = sArgFormat
        End With
    End With
End Sub
        
' オートフィルタ設定
Private Sub setAutoFilter()
    With sheetReader
        With .Range(.Cells(ROW_FileList.ヘッダ, 1), .Cells(C_LAST_ROW, COL_FileList.列数))
            .AutoFilter
        End With
    End With
End Sub
        
    
' 印刷設定
Private Sub setPageSetup()
    With sheetReader.PageSetup
        .PrintTitleRows = "$1:$" & ROW_FileList.ヘッダ ' 固定行
        
        .CenterHeader = "&F &A" ' ヘッダ
        .CenterFooter = "&P/&N" ' フッタ
        
        ' 横１ページで印刷
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
End Sub
    
    
' ファイルを全行読込み「全行」列に書込
Private Sub readFile(ByRef sArgPath As String)
    Dim nRow As Long
    nRow = ROW_FileList.結果開始

    Dim nFileNo As Integer
    Dim sBuf As String
    nFileNo = FreeFile
    Open sArgPath For Input As #nFileNo
On Error GoTo ERR
    Do Until EOF(nFileNo)
        Line Input #nFileNo, sBuf
        sheetReader.Cells(nRow, COL_FileList.全文).Value = sBuf
        nRow = nRow + 1
    Loop
OK:
    GoTo FINALLY
ERR:
    Debug.Print "readFile:" & ERR.Number & ":" & ERR.Description
    GoTo FINALLY
FINALLY:
    Close nFileNo
    Exit Sub
End Sub


' 読込結果解析
Private Sub analyzeText()
    Dim nRow As Long
    Dim nLastRow As Long
    nLastRow = getLastRowNo(sheetReader, COL_FileList.全文)

On Error GoTo ERR
    For nRow = ROW_FileList.結果開始 To nLastRow
        Call sheetReader.Cells(nRow, 1).Activate
        If analyzeLine(sheetReader.Cells(nRow, COL_FileList.全文).Value) <> C_OK Then GoTo ERR
    Next
OK:
    GoTo FINALLY
ERR:
    Debug.Print "analyzeText:" & nRow & "行目>" & ERR.Number & ":" & ERR.Description
    GoTo FINALLY
FINALLY:
    Exit Sub
End Sub


' １行を解析
Private Function analyzeLine(ByRef sLine As String) As Integer
On Error GoTo ERR
    
    ' 文節で分解
    Dim sTmp As String
    Dim nKind As E_LineKind
    Dim sWords() As String
    sTmp = sLine
    sTmp = getRegExpReplace(sTmp, "\(([0-9]+),([0-9]+)\)  \[(.+)\]\: ", getReplaceStr(3))
    sWords = Split(sTmp, C_SPLIT_CHR)
    
    If UBound(sWords) > 0 Then
        ' ファイル名
        sTmp = getFromArray(sWords, 0)
        Call setLink(フルパス, sTmp)
        Call setLink(フォルダ, getParentDir(sTmp))
        Call setVal(ファイル名, getFileName(sTmp))
        Call setVal(拡張子, getExt(sTmp))
            
        Call setVal(行, getFromArray(sWords, 1))
        Call setVal(位置, getFromArray(sWords, 2))
        Call setVal(文字コード, getFromArray(sWords, 3))
        Call setVal(内容, getFromArray(sWords, 4))
    End If
    
OK:
    analyzeLine = C_OK
    GoTo FINALLY
ERR:
    analyzeLine = C_ERR
    Debug.Print "analyzeLine:" & ERR.Number & ":" & ERR.Description
    GoTo FINALLY
FINALLY:
    Exit Function
End Function

' 置換後文字列を生成
Private Function getReplaceStr(ByVal nWordCount As Integer) As String
    Dim i As Integer
    getReplaceStr = C_SPLIT_CHR & "$1" & C_SPLIT_CHR
    For i = 2 To nWordCount
        getReplaceStr = getReplaceStr & "$" & i & C_SPLIT_CHR
    Next
End Function

' 配列から要素を抽出する際のインデックス検査を簡略化
Private Function getFromArray(ByRef sArray() As String, ByVal nIndex As Integer) As String
    If UBound(sArray) < nIndex Then
        getFromArray = ""
    Else
        getFromArray = sArray(nIndex)
    End If
End Function
