Attribute VB_Name = "module_DdlToDataBook"
Option Explicit

Const CONTROL_SHEET_NAME = "CONTROL"
Const BASE_SHEET_NAME = "template"

Const DDL_RANGE = "B38"


' CREATE文をもとに、空のデータ表を作成する
' 設定値をあつめてメインに渡す係
Sub createSkeletonFromDdl()
    
'    Application.ScreenUpdating = False
    
    ' コントロールブック、シートを取得
    Dim wbControl As Workbook: Set wbControl = ActiveWorkbook
    Dim wsControl As Worksheet: Set wsControl = wbControl.Worksheets(CONTROL_SHEET_NAME)
    
    ' 設定値変数
    Dim wrDdl As Range: Set wrDdl = wsControl.Range(DDL_RANGE)
    
    ' シート定義 コピー元のシート
    Dim wsBase As Worksheet: Set wsBase = wbControl.Worksheets(BASE_SHEET_NAME)
    
    Call createSkeletonMain(wsBase, wrDdl)
    
'    Application.DisplayAlerts = True
    MsgBox "done"
    
End Sub

' CREATE文をもとに、空のデータ表を作成する
' メイン処理
' @param wsBase データ表作成時のコピー元
' @param wrDdl 1つ目のCREATE文のセルのrange
Private Function createSkeletonMain(wsBase As Worksheet, wrDdl As Range)
    
    ' 生成するデータ表のブックを新規作成
    Workbooks.Add
    Dim wbData As Workbook: Set wbData = ActiveWorkbook
    
    ' 新規作成したブックのシートを、シート数が1になるまで削る
    Application.DisplayAlerts = False
    Do While wbData.Worksheets.Count <> 1
        wbData.Worksheets(1).delete
    Loop
    Application.DisplayAlerts = True

    Do While wrDdl.value <> ""
        ' コピー元からコピーして空のデータ表を作成
        wsBase.copy after:=wbData.Worksheets(wbData.Worksheets.Count)
        Dim wsData As Worksheet
        Set wsData = wbData.Worksheets(wbData.Worksheets.Count)
        ' DDL→データ表への定義転記
        Dim ddl As String: ddl = wrDdl.value
        Call createSkeletonOne(wsData, ddl)
        
        Set wrDdl = wrDdl.Offset(1, 0)
    Loop
    
    
    ' ブック作成時に残した1シートを削除する
    Application.DisplayAlerts = False
    wbData.Worksheets(1).delete
    Application.DisplayAlerts = True
    
End Function


' CREATE文をもとに、空のデータ表を作成する
' @param wsOut データ表
' @param ddl CREATE文
Private Function createSkeletonOne(wsOut As Worksheet, ddl As String)
    

    ' セル位置設定 出力側 1項目目の列の物理名の行、C3
    Dim wrOut As Range: Set wrOut = wsOut.Range("C3")
    
    ' ループ内で使う変数
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    Dim matches As Variant
    Dim match As Variant
    
    ' 1行ずつ処理
    Dim lines() As String: lines = Split(ddl, vbLf)
    Dim line As Variant
    For Each line In lines
            
        ' 行がCREATE TABLEで始まる場合、そこからテーブル名を取得
        If InStr(line, "CREATE TABLE") = 1 Then
            regex.Pattern = "`([^`]+)`"
            Set matches = regex.Execute(line)
            Dim tableName: tableName = matches.Item(0).SubMatches.Item(0)
            wsOut.Name = tableName
        End If
        
        ' カラム定義列の場合
        ' MEMO: COMMENTを論理名として使いたいが、うまく取れていない。
        regex.Pattern = "^ +`([^`]+)` +([^ ]+).+?(COMMENT '([^']+)')?"
        Set matches = regex.Execute(line)
        If matches.Count <> 0 Then
            Dim physicalName As String: physicalName = matches.Item(0).SubMatches.Item(0)
            Dim dataType As String: dataType = matches.Item(0).SubMatches.Item(1)
            
            ' うまく取れていない
            ' logicalName = matches.Item(0).SubMatches.Item(3)
            
            ' 一つの正規表現でまとめて取得したかったがうまくいかないので、COMMENTを単独で取得
            Dim logicalName As String: logicalName = ""
            regex.Pattern = "COMMENT '([^']+)'"
            Set matches = regex.Execute(line)
            If matches.Count <> 0 Then
                logicalName = matches.Item(0).SubMatches.Item(0)
            End If
        
            wrOut.value = physicalName
            wrOut.Offset(1, 0).value = logicalName
            wrOut.Offset(2, 0).value = convertDataType(dataType)

            ' 次の項目にセル位置を移動
            Set wrOut = wrOut.Offset(0, 1)
        
        End If
        
            
    Next line
    
    ' 書式をコピー
    wsOut.Range("C3:C8").copy
    wsOut.Range(wsOut.Range("D3"), wrOut.Offset(5, -1)).PasteSpecial xlPasteFormats
    wsOut.Range("A1").Select
    
End Function


' データ型を変換する
' 必要なのは型名だけなので、それ以外を除去する。
Private Function convertDataType(dataType As String) As String
    
    Dim s As String: s = dataType
    
    ' まず先頭末尾の空白を除去
    s = Trim(s)
    
    ' ( があればそれ以降を除去
    Dim pos As Long
    pos = InStr(s, "(")
    If pos <> 0 Then
        s = Left(s, pos - 1)
    End If
    
    ' 空白があればそれ以降を除去
    pos = InStr(s, " ")
    If pos <> 0 Then
        s = Left(s, pos - 1)
    End If
    
    convertDataType = s
End Function

