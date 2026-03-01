Attribute VB_Name = "module_copyTableDefToDataBook"
Option Explicit

Const CONTROL_SHEET_NAME = "CONTROL"
Const COL_SETTING = 3
Const ROW_DEF_START = 23

Const BASE_SHEET_NAME = "template"


' 現在開いている他のブックのデータ定義書をもとに、空のデータ表を作成する
' 設定値をあつめてメインに渡す係
Sub createSkeleton()
    
'    Application.ScreenUpdating = False
    
    ' コントロールブック、シートを取得
    Dim wbControl As Workbook: Set wbControl = ActiveWorkbook
    Dim wsControl As Worksheet: Set wsControl = wbControl.Worksheets(CONTROL_SHEET_NAME)
    
    ' 設定値変数
    Dim defFileName As String
    Dim whiteListStr As String
    Dim blackListStr As String
    Dim typeDictionaryStr As String
    
    
    ' 設定値取得
    Dim wkRow As Long: wkRow = ROW_DEF_START - 1
    If wsControl.Cells(wkRow, COL_SETTING - 1).value <> "今開いているA5MK2形式のDB定義書をもとにデータ表を生成する" Then
        MsgBox "設定値のセル位置が想定通りでないです。終了します。"
        End
    End If
    wkRow = wkRow + 1: defFileName = wsControl.Cells(wkRow, COL_SETTING)
    wkRow = wkRow + 1: whiteListStr = wsControl.Cells(wkRow, COL_SETTING)
    wkRow = wkRow + 1: blackListStr = wsControl.Cells(wkRow, COL_SETTING)
    wkRow = wkRow + 1: typeDictionaryStr = wsControl.Cells(wkRow, COL_SETTING)
    
    ' 設定値の加工 作成対象シートを配列化
    Dim whiteList() As String: whiteList = csvToArray(whiteListStr)
    ' 設定値の加工 作成対象外シートを配列化
    Dim blackList() As String: blackList = csvToArray(blackListStr)
    ' 設定値の加工 データ型辞書をdictionary化
    Dim typeDictionary As Object: Set typeDictionary = createDictionary(typeDictionaryStr)
    
    ' ブック定義 定義ブック
    Dim wbDef As Workbook:  Set wbDef = Workbooks(defFileName)
    
    ' シート定義 コピー元のシート
    Dim wsBase As Worksheet: Set wsBase = wbControl.Worksheets(BASE_SHEET_NAME)
    
    Call createSkeletonMain(wbDef, whiteList, blackList, typeDictionary, wsBase)
    
'    Application.DisplayAlerts = True
    MsgBox "done"
    
End Sub

' 現在開いている他のブックのデータ定義書をもとに、空のデータ表を作成する
' メイン処理
' @param wbDef DB定義書
' @param whiteList データ表作成対象の、DB定義書のワークシート名
' @param blackList データ表作成対象外の、DB定義書のワークシート名（未実装）
' @param typeDictionary DB定義書に型が記載されていない場合の、
' @param wsBase データ表作成時のコピー元
Private Function createSkeletonMain(wbDef As Workbook, whiteList() As String, blackList() As String, typeDictionary As Object, wsBase As Worksheet)
    
    ' 生成するデータ表のブックを新規作成
    Workbooks.Add
    Dim wbData As Workbook: Set wbData = ActiveWorkbook
    
    ' 新規作成したブックのシートを、シート数が1になるまで削る
    Application.DisplayAlerts = False
    Do While wbData.Worksheets.Count <> 1
        wbData.Worksheets(1).delete
    Loop
    Application.DisplayAlerts = True

    Dim sheetName As Variant
    Dim wsDef As Worksheet
    Dim wsData As Worksheet
    ' whiteList（作成対象シート名）でループ
    For Each sheetName In whiteList
        
        ' ワークシート取得。なければエラーメッセージを出して終了。
        Set wsDef = Nothing
        On Error Resume Next
        Set wsDef = wbDef.Worksheets(sheetName)
        On Error GoTo 0
        If wsDef Is Nothing Then
            MsgBox "作成対象として指定されているワークシートが存在しません：" & sheetName
            End
        End If
        
        ' コピー元からコピーして空のデータ表を作成
        wsBase.copy after:=wbData.Worksheets(wbData.Worksheets.Count)
        Set wsData = wbData.Worksheets(wbData.Worksheets.Count)
        
        ' DB定義→データ表への定義転記
        Call createSkeletonOne(wsDef, wsData, typeDictionary)
    Next sheetName
    
    
    ' ブック作成時に残した1シートを削除する
    Application.DisplayAlerts = False
    wbData.Worksheets(1).delete
    Application.DisplayAlerts = True
    
End Function


' カンマ区切り文字列を配列にする
Private Function csvToArray(csv As String, Optional separator = ",") As String()
    
    ' カンマ区切りを配列にする
    Dim splitted() As String
    splitted = Split(csv, separator)
    
    ' 前後の空白を除去
    Dim i As Long
    For i = LBound(splitted) To UBound(splitted)
        splitted(i) = Trim(splitted(i))
    Next i
    
    csvToArray = splitted
End Function

' 物理名:データ型 をカンマで区切った文字列を、Dictionaryに変換する
Private Function createDictionary(csv As String) As Object

    ' カンマ区切り文字列を、コロン区切り文字列の配列にする
    Dim colonSeparateds() As String
    colonSeparateds = csvToArray(csv)
    
    ' Dictionary初期化
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    
    
    Dim i As Long
    For i = LBound(colonSeparateds) To UBound(colonSeparateds)
        ' コロン区切り文字列を配列化
        Dim definition() As String: definition = csvToArray(colonSeparateds(i), ":")
        ' Dictionaryに追加
        d.Add definition(0), definition(1)
    Next i
    
    Set createDictionary = d
End Function


' DB定義シートをもとに空のデータ表を作成する
' @param wsDef DB定義書
' @param wsOut データ表
' @param typeDictionary DB定義書に型が書いていない場合に参照する、項目名と型の関連
Private Function createSkeletonOne(wsDef As Worksheet, wsOut As Worksheet, typeDictionary As Object)
    
    ' シート名はテーブル物理名
    wsOut.Name = wsDef.Range("C6").value
    ' シート内のテーブル物理名と論理名セット
    wsOut.Range("C1").value = wsOut.Name
    wsOut.Range("C2").value = wsDef.Range("C5").value
    
    
    ' セル位置設定 DB定義側 A列で値が「No.」の行の次、DB定義1行目
    Dim wrDef As Range: Set wrDef = wsDef.Range("A2")
    Do While wrDef.Offset(-1, 0).value <> "No."
        Set wrDef = wrDef.Offset(1, 0)
    Loop
    
    ' セル位置設定 出力側 1項目目の列の物理名の行、C3
    Dim wrOut As Range: Set wrOut = wsOut.Range("C3")
    
    ' DB定義表のA列が空になるまで1行ずつ処理
    Do While wrDef.value <> ""
    
        ' 次の行をつくる
'        Range(wrOut, wrOut.Offset(5, 0)).copy wrOut.Offset(0, 1)
        
        ' 物理名、論理名、型
        Dim physicalName As String: physicalName = wrDef.Offset(0, 2)
        Dim logicalName As String: logicalName = wrDef.Offset(0, 1)
        Dim dataType As String: dataType = getDataType(physicalName, wrDef.Offset(0, 3), typeDictionary)
        wrOut.value = physicalName
        wrOut.Offset(1, 0).value = logicalName
        wrOut.Offset(2, 0).value = dataType
        
        
        ' 次の項目にセル位置を移動
        Set wrDef = wrDef.Offset(1, 0)
        Set wrOut = wrOut.Offset(0, 1)
    Loop
    
    ' 書式をコピー
    wsOut.Range("C3:C8").copy
    wsOut.Range(wsOut.Range("D3"), wrOut.Offset(5, -1)).PasteSpecial xlPasteFormats
    wsOut.Range("A1").Select
    
End Function

' データ型を取得する
' @param physicalName 項目の物理名
' @param typeOnDef DB定義書上のデータ型
' @param typeDictionary DB定義書に型が書いていない場合に参照する、項目名と型の関連
Private Function getDataType(physicalName As String, typeOnDef As String, typeDictionary As Object)
    
    ' DB定義書にあればそれを使う
    If typeOnDef <> "" Then
        getDataType = typeOnDef
        Exit Function
    End If
    
    ' マクロ設定のデータ型辞書にあればそれを使う
    If typeDictionary.Exists(physicalName) Then
        getDataType = typeDictionary.Item(physicalName)
        Exit Function
    End If
    
    ' なければ undefined とする
    getDataType = "undefined"
    
End Function

