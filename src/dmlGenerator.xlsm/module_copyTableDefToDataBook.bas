Attribute VB_Name = "module_copyTableDefToDataBook"
Option Explicit

Const BOOK_MACRO = "Mysqlレコード追加DML生成ツール.xlsm"
Const BOOK_DDL = "テーブル定義(wip).xlsx"
Const BOOK_DATA = "マスタデータ表(wip).xlsx"
Const SHEET_TPL = "base"
' [wip] テーブル定義表をもとに、空のマスタデータ表を生成する
Function copy()

    Dim wbMacro As Workbook, wbDdl As Workbook, wbData As Workbook
    Set wbMacro = Workbooks(BOOK_MACRO)
    Set wbDdl = Workbooks(BOOK_DDL)
    Set wbData = Workbooks(BOOK_DATA)
    
    Dim wsTemplate As Worksheet
    Set wsTemplate = wbMacro.Worksheets(SHEET_TPL)
    
    Dim wsDdl As Worksheet
    For Each wsDdl In wbDdl.Worksheets
        If Mid(wsDdl.Name, 1, 2) = "c_" Then
            Call convertDdlSheetToDmlSheet(wsDdl, wsTemplate, wbData)
        End If
    Next wsDdl
    MsgBox "done"
End Function

' [wip]
Function convertDdlSheetToDmlSheet(wsDdl As Worksheet, wsTemplate As Worksheet, wbData As Workbook)
    wsTemplate.copy after:=wbData.Worksheets(wbData.Worksheets.Count)
    Dim wsNew As Worksheet
    Set wsNew = ActiveSheet
    wsNew.Name = wsDdl.Name
    
    ' カラム
    Dim wrDdl As Range
    Set wrDdl = wsDdl.Range("A8")
    Dim wrNew As Range
    Set wrNew = wsNew.Range("C3")
    Do Until wrDdl.value = ""
    
        ' 書式設定
        wsNew.Range(wrNew, wrNew.Offset(5, 0)).Borders.LineStyle = xlContinuous
        wrNew.Interior.Color = wrNew.Offset(0, -1).Interior.Color
        wrNew.Offset(1, 0).Interior.Color = wrNew.Offset(1, -1).Interior.Color
        wrNew.Offset(2, 0).Interior.Color = wrNew.Offset(2, -1).Interior.Color
        wrNew.Offset(3, 0).Interior.Color = wrNew.Offset(3, -1).Interior.Color
        
        ' 物理名
        wrNew.value = wrDdl.Offset(0, 1).value
        ' 論理名
        wrNew.Offset(1, 0).value = wrDdl.Offset(0, 14).value
        ' 型
        wrNew.Offset(2, 0).value = wrDdl.Offset(0, 2).value
    
        Set wrDdl = wrDdl.Offset(1, 0)
        Set wrNew = wrNew.Offset(0, 1)
    Loop
    
    wsNew.Range("C1", wrNew).EntireColumn.AutoFit

    ' テーブル物理名
    wsNew.Range("C1").value = wsDdl.Name
    ' テーブル論理名
    wsNew.Range("C2").value = wsDdl.Range("C2").value & ":" & wsDdl.Range("C3").value

End Function
