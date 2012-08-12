Attribute VB_Name = "M_Funcs"
Option Explicit

Private Const C_NOT_FOUND As Integer = -1

' ************************************************
' 名称：正規表現検索
' 引数：処理対象文字列、パターン
' 備考：複数マッチした場合は先頭のみ返します
' 参考：http://officetanaka.net/excel/vba/tips/tips38.htm
' ************************************************
Public Function getRegExpMatchStr(ByVal saArgTargetStr As String _
    , ByVal sArgPattern As String _
)
On Error GoTo ERR
    Dim oRegExp As Object
    Dim strPattern As String
    Dim reMatch As Object
    
    Set oRegExp = CreateObject("VBScript.RegExp")
    With oRegExp
        .Pattern = sArgPattern      '検索パターンを設定
        .IgnoreCase = True          '大文字と小文字を区別しない
        .Global = True              '文字列全体を検索
        
        ' 一致確認
        Set reMatch = .Execute(saArgTargetStr)
        If reMatch.Count > 0 Then
            getRegExpMatchStr = reMatch(0)
        Else
            getRegExpMatchStr = ""
        End If
    End With
OK:
    GoTo FINALLY
ERR:
    Debug.Print "getRegExpMatchStr:" & ERR.Number & ":" & ERR.Description
    getRegExpMatchStr = "*Err*" ' 暫定 エラー時は当文言を返す
    GoTo FINALLY
FINALLY:
    Set oRegExp = Nothing
End Function


' ************************************************
' 名称：正規表現置換
' 引数：処理対象文字列、パターン、置換後文字列
' 備考：正規表現の条件にマッチする文字列を置換します
' 参考：http://officetanaka.net/excel/vba/tips/tips38.htm
' ************************************************
Public Function getRegExpReplace(ByVal saArgTargetStr As String _
    , ByVal sArgPattern As String _
    , ByVal sArgReplace As String _
)
On Error GoTo ERR
    Dim oRegExp As Object
    Dim strPattern As String
    Dim reMatch As Object
    
    Set oRegExp = CreateObject("VBScript.RegExp")
    With oRegExp
        .Pattern = sArgPattern      '検索パターンを設定
        .IgnoreCase = True          '大文字と小文字を区別しない
        .Global = True              '文字列全体を検索
        
        ' 置換
        getRegExpReplace = .Replace(saArgTargetStr, sArgReplace)
    End With
OK:
    GoTo FINALLY
ERR:
    Debug.Print "getRegExpReplace:" & ERR.Number & ":" & ERR.Description
    getRegExpReplace = "*Err*" ' 暫定 エラー時は当文言を返す
    GoTo FINALLY
FINALLY:
    Set oRegExp = Nothing
End Function


' ************************************************
' 名称：末尾から文字位置検索
' 引数：検索対象、検索文字
' 備考：検出されなかった場合は-1を返す
' ************************************************
Public Function LastIndexOf(ByRef sTarget As String, ByRef sSearchChr As String) As Integer
    Dim nIndex As Integer
    For nIndex = Len(sTarget) To 1 Step -1
        If Mid(sTarget, nIndex, 1) = sSearchChr Then
            LastIndexOf = nIndex
            Exit Function
        End If
    Next
    LastIndexOf = C_NOT_FOUND
    Exit Function
End Function


' ************************************************
' 名称：所属フォルダ取得
' 引数：フルパス
' 備考：フルパスからファイルが存在するフォルダのパスを抽出
' ************************************************
Public Function getParentDir(ByRef sArgPath As String) As String
    Dim nIndex As Integer
    nIndex = LastIndexOf(sArgPath, "\")
    If nIndex > 0 Then
        getParentDir = Left(sArgPath, nIndex - 1)
    Else
        getParentDir = ""
    End If
End Function


' ************************************************
' 名称：拡張子取得
' 引数：ファイル名
' 備考：ファイル名から拡張子を抽出
' ************************************************
Public Function getExt(ByRef sArgPath As String) As String
    Dim nIndex As Integer
    nIndex = LastIndexOf(sArgPath, ".")
    If nIndex > 0 Then
        getExt = Mid(sArgPath, nIndex + 1)
    Else
        getExt = ""
    End If
End Function


' ************************************************
' 名称：拡張子取得
' 引数：ファイル名
' 備考：ファイル名から拡張子を抽出
' ************************************************
Public Function getFileName(ByRef sArgPath As String) As String
    Dim nIndex As Integer
    nIndex = LastIndexOf(sArgPath, "\")
    If nIndex > 0 Then
        getFileName = Mid(sArgPath, nIndex + 1)
    Else
        getFileName = sArgPath
    End If
End Function

