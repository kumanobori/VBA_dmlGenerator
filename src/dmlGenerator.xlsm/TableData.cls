VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tableName As String
Public columnNames As New Collection
Public columnTypes As New Collection
Public deleteConditions As New Collection

Public records As New Collection ' collection of Record

' -------------------------------------
' テーブル名をセットする
' @param aTableName テーブル名
' -------------------------------------
Public Function setTableName(aTableName As String)
    tableName = aTableName
End Function


' -------------------------------------
' フィールド名をセットする
' @param values フィールド名が入力されたセルのRange
' -------------------------------------
Public Function setNames(values As Range)
    Dim v() As Variant
    v = values
    Dim i As Long
    For i = LBound(v, 2) To UBound(v, 2)
        columnNames.Add (v(1, i))
    Next i
End Function

' -------------------------------------
' データ型をセットする
' @param values データ型が入力されたセルのRange
' -------------------------------------
Public Function setTypes(values As Range)
    Dim v As Variant
    v = values
    Dim i As Long
    For i = LBound(v, 2) To UBound(v, 2)
        columnTypes.Add (v(1, i))
    Next i
End Function

' -------------------------------------
' 削除条件をセットする
' @param values 削除条件が入力されたセルのRange
' -------------------------------------
Public Function setDeleteConditions(values As Range)
    Dim v As Variant
    v = values
    Dim i As Long
    For i = LBound(v, 2) To UBound(v, 2)
        deleteConditions.Add (v(1, i))
    Next i
End Function

' -------------------------------------
' レコードを追加する
' @param values レコード内容が入力されたセルのRange
' -------------------------------------
Public Function addRecord(values As Range)
    Dim v As Variant
    v = values
    Dim i As Long
    Dim columns As New Collection
    For i = LBound(v, 2) To UBound(v, 2)
        columns.Add (v(1, i))
    Next i
    records.Add columns
End Function
