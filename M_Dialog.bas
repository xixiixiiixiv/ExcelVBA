Attribute VB_Name = "M_Dialog"
Option Explicit

' ************************************************
' 名称：「ファイルを開く」ダイアログを開く
' 引数：パス初期位置
' 備考：検出されなかった場合は-1を返す
' ************************************************
Function OpenFileDialog(ByVal sDefaultPath As String) As String
    Dim FileDlg As Office.FileDialog
    Dim filePath As String
    Dim Result As Integer 'ファイルを選択したかどうか（0=選択していない, -1=選択した）
    
On Error GoTo ERR
    filePath = ""
    
    ' 初期表示フォルダを指定
    If sDefaultPath = "" Or Dir(sDefaultPath) = "" Then
        sDefaultPath = "C:\"
    Else
        sDefaultPath = getParentDir(sDefaultPath)
    End If
    
    'ファイルダイアログの種類を指定
    Set FileDlg = Application.FileDialog(msoFileDialogFilePicker)
    With FileDlg
      .InitialFileName = sDefaultPath 'ダイアログで最初に表示するファイルパス
      .Filters.Add "テキストファイル", "*.txt; *.csv"  'ダイアログに表示するファイル種別
      .Filters.Add "すべてのファイル", "*.*" 'ダイアログに表示するファイル種別
      .FilterIndex = 1 'ダイアログ表示時に有効にするファイル種別のインデックス
      .AllowMultiSelect = False '複数ファイルの選択を無効にする
    End With
    Result = FileDlg.Show()
    
    'ファイルが選択された場合、選択したファイル名を返す
    If Result = -1 Then
        filePath = FileDlg.SelectedItems(1)
    End If
OK:
    OpenFileDialog = filePath
    GoTo FINALLY
ERR:
    OpenFileDialog = ""
    GoTo FINALLY
FINALLY:
  Set FileDlg = Nothing
End Function
