VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tables As New Collection
Private sql As New Collection

Private folderPath As String
Private dataFileName As String
Private outputDmlToText As Boolean
Private textFileNamePrefix As String
Private textFileNameFromExcel As Boolean
Private textFileNameAddDateTime As Boolean
Private addDeleteStatement As Boolean
Private nullString As String
Private sheetToIgnore As Variant
Private dbms As String
Private multipleInsertCount As Long
Private oracleEndWithExit As Boolean
Private dataBook As Workbook

' -----------------------------
' 初期処理
' 設定値を取り込む。
' -----------------------------
Public Sub init(aFolderPath As String, aDataFileName As String, aOutputDmlToText As Boolean, aTextFileNamePrefix As String, aTextFileNameFromExcel As Boolean, aTextFileNameAddDateTime As Boolean, aAddDeleteStatement As Boolean, aNullString As String, aSheetToIgnore As Variant, aDbms As String, aMultipleInsertCount As Long, aOracleEndWithExit As Boolean, aDataBook As Workbook)
    folderPath = aFolderPath
    dataFileName = aDataFileName
    outputDmlToText = aOutputDmlToText
    textFileNamePrefix = aTextFileNamePrefix
    textFileNameFromExcel = aTextFileNameFromExcel
    textFileNameAddDateTime = aTextFileNameAddDateTime
    addDeleteStatement = aAddDeleteStatement
    nullString = aNullString
    dbms = aDbms
    multipleInsertCount = aMultipleInsertCount
    oracleEndWithExit = aOracleEndWithExit
    sheetToIgnore = aSheetToIgnore
    
    Set dataBook = aDataBook
End Sub

' ------------------------------
' データ読み込み
' ------------------------------
Public Sub read()
    Dim ws As Worksheet
    For Each ws In dataBook.Worksheets
        Dim found As Variant
        found = Filter(sheetToIgnore, ws.Name)
        If UBound(found) <> -1 Then
            ' シート名が無視対象のシート名配列に一致した場合、処理しない
        Else
            'シート内容をテーブルオブジェクトに読み込む
            Dim table As New TableData
            Dim iConvExcelToTableData As New ConvExcelToTableData
            Set table = iConvExcelToTableData.exec(ws)
            Call tables.Add(table)
        End If
    Next ws
End Sub


' ------------------------------------------
' 読み込んだテーブルデータからSQLを生成
' ------------------------------------------
Public Sub generate()
    Dim conv As New ConvTableDataToSql
    Dim table As TableData
    Dim v As Variant
    
    For Each v In tables
        Set table = v
        
        ' 設定で「DMLにDELETEを追加」にしている場合、条件に応じたDELETE文を追加する
        If addDeleteStatement Then
            Dim delete As String
            delete = conv.generateDelete(table)
            sql.Add delete
        End If
        
        ' INSERT文を生成、追加
        Dim insert As Collection
        Set insert = conv.generateInsert(dbms, table, multipleInsertCount, nullString)
        Dim v2 As Variant
        For Each v2 In insert
            sql.Add v2
        Next v2
    Next v
    
    ' DBMSがoracleの場合、最後にexitをつける
    If dbms = "oracle" And oracleEndWithExit Then
        sql.Add "exit"
    End If
    
End Sub

' ---------------------------------------
' SQLファイルを出力する
' ---------------------------------------
Public Sub writeSqlFile()
    
    ' 出力ファイルパスを生成
    Dim filePath As String
    filePath = folderPath
    If Right(filePath, 1) <> "\" Then
        filePath = filePath & "\"
    End If
    filePath = filePath & textFileNamePrefix
    If textFileNameFromExcel Then
        Dim extentionDeleted As String
        filePath = filePath & Mid(dataFileName, 1, InStrRev(dataFileName, ".") - 1)
    End If
    If textFileNameAddDateTime Then
        filePath = filePath & "_" & Format(Now, "yyyymmdd_HHMM")
    End If
    filePath = filePath & ".txt"
    
    ' 出力ファイルを開く
    Dim fileNumber As Long
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    
    ' SQLを書き出す
    Dim line As Variant
    For Each line In sql
        Print #fileNumber, line
    Next line
    
    ' 閉じる
    Close #fileNumber
    
    
End Sub

