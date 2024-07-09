VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConvTableDataToSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INDENT2 = "        "
Const INDENT3 = "            "


' ---------------------------------------------
' DELETE文を生成
' ---------------------------------------------
Public Function generateDelete(data As TableData)
    Dim s As String
    s = ""
    
    Dim tableName As String
    tableName = data.tableName
    
    Dim i As Long
    For i = 1 To data.columnNames.Count
        Dim columnName As String, deleteCondition As String
        columnName = data.columnNames.Item(i)
        deleteCondition = data.deleteConditions.Item(i)
        If deleteCondition <> "" Then
            Select Case Len(s)
                Case 0
                    s = "DELETE FROM " & tableName & " WHERE " & columnName & " " & deleteCondition
                Case Else
                    s = s & " AND " & columnName & " " & deleteCondition
            End Select
        End If
    Next i
    If Len(s) > 0 Then
        s = s & ";"
    Else
'        ' 削除条件がひとつもなかった場合は、処理が速いのでTRUNCATEにする
'        s = "TRUNCATE TABLE " & data.tableName & ";"
        ' TRUNCATEだと外部キー参照されているテーブルには使えないようなのでDELETEにした
        s = "DELETE FROM " & data.tableName & ";"
    End If
    
    generateDelete = s
End Function

' -----------------------------------------------
' テーブルデータをもとにINSERT文を生成する
' @param dbms 対象のDBMS。対応しているのはmysql,oracle,laravel
' @param data テーブルデータのオブジェクト
' @param multipleInsertCount mysqlのbulk insertをする際の、INSERT1本での登録最大数。
' @param nullString フィールド値がこれに一致するものは、nullであるとみなす。
' -----------------------------------------------
Public Function generateInsert(dbms As String, data As TableData, multipleInsertCount As Long, nullString As String)
    
    ' laravelの場合はsqlではないので別処理にする
    If dbms = "laravel" Then
        Set generateInsert = generateInsertLaravel(data, nullString)
        Exit Function
    End If
    
    Dim c As New Collection
    Set generateInsert = c
    
    ' INSERT文のINSERT INTO～部分
    Dim insertPhrase As String
    insertPhrase = "INSERT INTO " & data.tableName
    
    ' INSERT文のカラム名部分 ex. (id, field1, field2)
    Dim columnPhrase As String
    columnPhrase = ""
    Dim v As Variant
    For Each v In data.columnNames
        If Len(columnPhrase) = 0 Then
            columnPhrase = "(" & v
        Else
            columnPhrase = columnPhrase & ", " & v
        End If
    Next v
    columnPhrase = columnPhrase & ")"
    
    ' value句に含むレコード数のカウンタ
    Dim recCount As Long
    recCount = 0
    ' value句を溜め込むワーク
    Dim valuePhrases As String
    valuePhrases = ""
    
    ' mysql以外の場合、マルチプルインサートを無効化させるために、引数を改変する
    If dbms <> "mysql" Then
        multipleInsertCount = 1
    End If
    
    ' データ1件ごとループ
    Dim record As Collection
    For Each v In data.records
        Set record = v
        
        ' VALUES句生成
        Dim valuePhrase As String
        valuePhrase = generateValuePhrase(dbms, record, data.columnTypes, nullString)
        
        If Len(valuePhrases) = 0 Then
            valuePhrases = "VALUES" & valuePhrase
        Else
            ' mysqlのマルチプルインサートの2件目以降
            valuePhrases = valuePhrases & "," & valuePhrase
        End If
        
        
        recCount = recCount + 1
        ' 「DBMSがmysqlかつマルチプルインサートを有効化している場合」以外は、常に真
        ' 「DBMSがmysqlかつマルチプルインサートを有効化している場合」は、所定件数に達した場合のみ出力
        If recCount >= multipleInsertCount Then
            c.Add insertPhrase & columnPhrase & valuePhrases & ";"
            recCount = 0
            valuePhrases = ""
        End If
    Next v
    
    ' mysqlのマルチプルインサートの残り分の出力
    If Len(valuePhrases) > 0 Then
        c.Add insertPhrase & columnPhrase & valuePhrases & ";"
    End If

    Set generateInsert = c
End Function

' ----------------------------------------------------------
' VALUES句を生成する
' @param dbms 対象のDBMS。対応しているのはmysql,oracle
' @param record レコードのフィールド値のコレクション
' @param columnTypes レコードのフィールド型のコレクション
' @param nullString フィールド値がこれに一致するものは、nullであるとみなす。
' ----------------------------------------------------------
Private Function generateValuePhrase(dbms As String, record As Collection, columnTypes As Collection, nullString As String)
    Dim str As String, columnString As String
    str = ""
    Dim i As Long
    ' 1レコード内の項目ごとにループ
    For i = 1 To record.Count
        
        ' 項目の値の文字列を取得
        columnString = editColumnString(dbms, record.Item(i), columnTypes.Item(i), nullString)
        
        ' VALUES句を生成
        If Len(str) = 0 Then
            str = "(" & columnString
        Else
            str = str & ", " & columnString
        End If
    Next i
    str = str & ")"
    
    generateValuePhrase = str
End Function

' -----------------------------------------------------------------
' VALUES内の各項目の文字列を生成する
' @param dbms 対象のDBMS。対応しているのはmysql,oracle
' @param value データ値
' @param dataType データ型
' @param nullString フィールド値がこれに一致するものは、nullであるとみなす。
' -----------------------------------------------------------------
Private Function editColumnString(dbms As String, value As String, dataType As String, nullString As String)
    
    If value = nullString Then
        editColumnString = "null"
    Else
        Select Case dbms
        Case "mysql"
            Select Case UCase(dataType)
            Case "CHAR", "VARCHAR", "TEXT", "ENUM", "DATE", "DATETIME", "TIMESTAMP"
                ' クォートで括る。対象文字列に含まれるクォートはエスケープする。
                editColumnString = "'" & Replace(value, "'", "''") & "'"
            Case Else
                editColumnString = value
            End Select
        Case "oracle"
            Select Case UCase(dataType)
            Case "DATE"
                ' TO_DATE関数
                editColumnString = "TO_DATE('" & value & "', 'YYYY/MM/DD HH24:MI:SS')"
            Case "CHAR", "VARCHAR", "NVARCHAR", "VARCHAR2", "NVARCHAR2"
                ' クォートで括る。対象文字列に含まれるクォートはエスケープする。
                editColumnString = "'" & Replace(value, "'", "''") & "'"
            Case Else
                editColumnString = value
            End Select
        Case Else
            MsgBox "dbms is unknown: [" & dbms & "]"
            End
        End Select
    End If
End Function

' -----------------------------------------------
' テーブルデータをもとにLaravel用のシーダーを生成する
' @param dbms 対象のDBMS。対応しているのはmysql,oracle
' @param data テーブルデータのオブジェクト
' @param multipleInsertCount mysqlのbulk insertをする際の、INSERT1本での登録最大数。
' @param nullString フィールド値がこれに一致するものは、nullであるとみなす。
' -----------------------------------------------
Private Function generateInsertLaravel(data As TableData, nullString As String)
    Dim c As New Collection
    
    ' レコードの配列開始
    c.Add INDENT2 & "$recs = ["
    
    Dim v As Variant
    For Each v In data.records
        Dim record As New Collection
        Set record = v
        Dim rec As String
        rec = generateInsertLaravelOne(data.columnNames, data.columnTypes, record, nullString)
        c.Add rec
    Next v
    
    ' レコードの配列終了
    c.Add INDENT2 & "];"
    ' レコード配列を用いてDBにcreateをかける
    c.Add INDENT2 & "foreach ($recs as $rec) {"
    c.Add INDENT3 & "\App\Models\" & data.tableName & "::create($rec);"
    c.Add INDENT2 & "}"
    
    Set generateInsertLaravel = c
End Function

' -----------------------------------------------
' Laravelのレコード1件分の連想配列を作成する
' @param names カラム名のCollection
' @param columnTypes カラムのデータ型のCollection
' @param values カラム値のCollection
' @param nullString フィールド値がこれに一致するものは、nullであるとみなす。
' -----------------------------------------------
Private Function generateInsertLaravelOne(names As Collection, columnTypes As Collection, values As Collection, nullString As String)
    Dim s As String
    s = INDENT3 & "["
    
    Dim i As Long
    For i = 1 To names.Count
        ' 2カラム目以降はカンマをつける
        If i > 1 Then
            s = s & ", "
        End If
        s = s & "'" & names.Item(i) & "' => " & editColumnStringForLaravel(values.Item(i), columnTypes.Item(i), nullString)
    Next i
    s = s & "],"

    generateInsertLaravelOne = s
End Function

' -----------------------------------------------
' Laravelのカラム1個の値を作成する
' @param value カラム値
' @param columnType カラムのデータ型
' @param nullString フィールド値がこれに一致するものは、nullであるとみなす。
' -----------------------------------------------
Private Function editColumnStringForLaravel(value As String, columnType As String, nullString As String)
    Dim s As String
    If value = nullString Then
        ' null
        s = "null"
    Else
        If columnType = "string" Then
            ' 文字列（ダブルクォートで括る）
            s = """" & value & """"
        Else
            ' 文字列以外（ダブルクォートで括らない）
            s = value
        End If
    End If

    editColumnStringForLaravel = s
End Function



