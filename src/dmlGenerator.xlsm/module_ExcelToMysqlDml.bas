Attribute VB_Name = "module_ExcelToMysqlDml"
Option Explicit


Const CONTROL_SHEET_NAME = "CONTROL"
Const COL_SETTING = 3
Const ROW_FOLDER_PATH = 4
Const ROW_DATA_FILE = 5
Const ROW_OUTPUT_DML_TO_TEXT = 6
Const ROW_TEXT_FILENAME_PREFIX = 7
Const ROW_TEXT_FILENAME_FROM_EXCEL = 8
Const ROW_TEXT_FILENAME_YMDHMS = 9
Const ROW_ADD_DELETE_STATEMENT = 10
Const ROW_NULL_STRING = 11
Const ROW_SHEET_TO_IGNORE = 12
Const ROW_DBMS = 13
Const ROW_MULTIPLE_INSERT_COUNT = 14
Const ROW_ORACLE_END_WITH_EXIT = 15


' ***********************************************
' * DML生成
' ***********************************************
Sub generateDmlFromDataSheet()
    
    Application.ScreenUpdating = False
    
    ' コントロールブック、シートを取得
    Dim wbControl As Workbook: Set wbControl = ActiveWorkbook
    Dim wsControl As Worksheet: Set wsControl = wbControl.Worksheets(CONTROL_SHEET_NAME)
    
    ' 設定値変数
    Dim folderPath As String
    Dim dataFileName As String
    Dim outputDmlToText As Boolean
    Dim textFileNamePrefix As String
    Dim textFileNameFromExcel As Boolean
    Dim textFileNameAddDateTime As Boolean
    Dim addDeleteStatement As Boolean
    Dim nullString As String
    Dim sheetToIgnore As Variant
    Dim dbms As String
    Dim multipleInsertCount As Long
    Dim oracleEndWithExit As Boolean
    
    ' 設定値取得
    folderPath = wsControl.Cells(ROW_FOLDER_PATH, COL_SETTING)
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    dataFileName = wsControl.Cells(ROW_DATA_FILE, COL_SETTING)
    If wsControl.Cells(ROW_OUTPUT_DML_TO_TEXT, COL_SETTING).value = "Y" Then
        outputDmlToText = True
    Else
        outputDmlToText = False
    End If
    textFileNamePrefix = wsControl.Cells(ROW_TEXT_FILENAME_PREFIX, COL_SETTING)
    If wsControl.Cells(ROW_TEXT_FILENAME_FROM_EXCEL, COL_SETTING).value = "Y" Then
        textFileNameFromExcel = True
    Else
        textFileNameFromExcel = False
    End If
    If wsControl.Cells(ROW_TEXT_FILENAME_YMDHMS, COL_SETTING).value = "Y" Then
        textFileNameAddDateTime = True
    Else
        textFileNameAddDateTime = False
    End If
    If wsControl.Cells(ROW_ADD_DELETE_STATEMENT, COL_SETTING).value = "Y" Then
        addDeleteStatement = True
    Else
        addDeleteStatement = False
    End If
    nullString = wsControl.Cells(ROW_NULL_STRING, COL_SETTING)
    sheetToIgnore = Split(wsControl.Cells(ROW_SHEET_TO_IGNORE, COL_SETTING).value, ",")
    dbms = wsControl.Cells(ROW_DBMS, COL_SETTING)
    multipleInsertCount = wsControl.Cells(ROW_MULTIPLE_INSERT_COUNT, COL_SETTING)
    If wsControl.Cells(ROW_ORACLE_END_WITH_EXIT, COL_SETTING).value = "Y" Then
        oracleEndWithExit = True
    Else
        oracleEndWithExit = False
    End If
    
    ' DB定義のワークブック取得。既に開いてればそれを使う。
    Dim wbData As Workbook
    Dim wb As Workbook
    Dim wbOpenedOnInit As Boolean
    wbOpenedOnInit = False
    Dim wsInit As Worksheet, wrInit As Range
    For Each wb In Workbooks
        If wb.Name = dataFileName Then
            wbOpenedOnInit = True
            Set wbData = wb
            wb.Activate
            Set wsInit = ActiveSheet
            Set wrInit = ActiveCell
        End If
    Next wb
    If Not wbOpenedOnInit Then
        Workbooks.Open folderPath & dataFileName
        Set wbData = ActiveWorkbook
    End If
    
    ' コントローラ実行
    Dim ctrl As New aController
    Call ctrl.init(folderPath, dataFileName, outputDmlToText, textFileNamePrefix, textFileNameFromExcel, textFileNameAddDateTime, addDeleteStatement, nullString, sheetToIgnore, dbms, multipleInsertCount, oracleEndWithExit, wbData)
    Call ctrl.read
    Call ctrl.generate
    If outputDmlToText Then
        Call ctrl.writeSqlFile
    End If
    
    Application.DisplayAlerts = False
    wbData.Save
    ' 最初から開いてた場合は元のカーソル位置を復元する。そうでない場合は閉じる。
    If wbOpenedOnInit Then
        wbData.Activate
        wsInit.Activate
        wrInit.Activate
    Else
        wbData.Close
    End If
    Application.DisplayAlerts = True
    
    
    MsgBox "done"
    
End Sub



