VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConvExcelToTableData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const ROW_COLUMN_NAME = 3
Const ROW_DATA_TYPE = 5
Const ROW_DELETE_CONDITION = 6
Const ROW_DATA_START = 7

Const COL_INDEX = 1
Const COL_DATA_START = 3





' -------------------------------------------------------
' データシートから、テーブルデータオブジェクトに変換する
' -------------------------------------------------------
Public Function exec(ws As Worksheet)
    Dim table As New TableData
    
    ' 最終列番号を取得
    Dim colLast As Long
    colLast = ws.Cells(ROW_COLUMN_NAME, COL_DATA_START).End(xlToRight).Column
    
    ' テーブル属性取得
    Call table.setTableName(ws.Name)
    Call table.setNames(Range(ws.Cells(ROW_COLUMN_NAME, COL_DATA_START), ws.Cells(ROW_COLUMN_NAME, colLast)))
    Call table.setTypes(Range(ws.Cells(ROW_DATA_TYPE, COL_DATA_START), ws.Cells(ROW_DATA_TYPE, colLast)))
    Call table.setDeleteConditions(Range(ws.Cells(ROW_DELETE_CONDITION, COL_DATA_START), ws.Cells(ROW_DELETE_CONDITION, colLast)))
    
    ' データ取得
    Dim wr As Range
    Set wr = ws.Cells(ROW_DATA_START, COL_INDEX)
    Do Until wr.value = ""
        Call table.addRecord(Range(ws.Cells(wr.row, COL_DATA_START), ws.Cells(wr.row, colLast)))
        Set wr = wr.Offset(1, 0)
    Loop
    
    
    Set exec = table
End Function
