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
' ��������
' �ݒ�l����荞�ށB
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
' �f�[�^�ǂݍ���
' ------------------------------
Public Sub read()
    Dim ws As Worksheet
    For Each ws In dataBook.Worksheets
        Dim found As Variant
        found = Filter(sheetToIgnore, ws.Name)
        If UBound(found) <> -1 Then
            ' �V�[�g���������Ώۂ̃V�[�g���z��Ɉ�v�����ꍇ�A�������Ȃ�
        Else
            '�V�[�g���e���e�[�u���I�u�W�F�N�g�ɓǂݍ���
            Dim table As New TableData
            Dim iConvExcelToTableData As New ConvExcelToTableData
            Set table = iConvExcelToTableData.exec(ws)
            Call tables.Add(table)
        End If
    Next ws
End Sub


' ------------------------------------------
' �ǂݍ��񂾃e�[�u���f�[�^����SQL�𐶐�
' ------------------------------------------
Public Sub generate()
    Dim conv As New ConvTableDataToSql
    Dim table As TableData
    Dim v As Variant
    
    For Each v In tables
        Set table = v
        
        ' �ݒ�ŁuDML��DELETE��ǉ��v�ɂ��Ă���ꍇ�A�����ɉ�����DELETE����ǉ�����
        If addDeleteStatement Then
            Dim delete As String
            delete = conv.generateDelete(table)
            sql.Add delete
        End If
        
        ' INSERT���𐶐��A�ǉ�
        Dim insert As Collection
        Set insert = conv.generateInsert(dbms, table, multipleInsertCount, nullString)
        Dim v2 As Variant
        For Each v2 In insert
            sql.Add v2
        Next v2
    Next v
    
    ' DBMS��oracle�̏ꍇ�A�Ō��exit������
    If dbms = "oracle" And oracleEndWithExit Then
        sql.Add "exit"
    End If
    
End Sub

' ---------------------------------------
' SQL�t�@�C�����o�͂���
' ---------------------------------------
Public Sub writeSqlFile()
    
    ' �o�̓t�@�C���p�X�𐶐�
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
    
    ' �o�̓t�@�C�����J��
    Dim fileNumber As Long
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    
    ' SQL�������o��
    Dim line As Variant
    For Each line In sql
        Print #fileNumber, line
    Next line
    
    ' ����
    Close #fileNumber
    
    
End Sub
