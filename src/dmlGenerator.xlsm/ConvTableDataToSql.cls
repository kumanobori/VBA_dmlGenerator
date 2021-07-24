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

' ---------------------------------------------
' DELETE���𐶐�
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
'        ' �폜�������ЂƂ��Ȃ������ꍇ�́A�����������̂�TRUNCATE�ɂ���
'        s = "TRUNCATE TABLE " & data.tableName & ";"
        ' TRUNCATE���ƊO���L�[�Q�Ƃ���Ă���e�[�u���ɂ͎g���Ȃ��悤�Ȃ̂�DELETE�ɂ���
        s = "DELETE FROM " & data.tableName & ";"
    End If
    
    generateDelete = s
End Function

' -----------------------------------------------
' �e�[�u���f�[�^�����Ƃ�INSERT���𐶐�����
' @param dbms �Ώۂ�DBMS�B�Ή����Ă���̂�mysql,oracle
' @param data �e�[�u���f�[�^�̃I�u�W�F�N�g
' @param multipleInsertCount mysql��bulk insert������ۂ́AINSERT1�{�ł̓o�^�ő吔�B
' @param nullString �t�B�[���h�l������Ɉ�v������̂́Anull�ł���Ƃ݂Ȃ��B
' -----------------------------------------------
Public Function generateInsert(dbms As String, data As TableData, multipleInsertCount As Long, nullString As String)
    Dim c As New Collection
    Set generateInsert = c
    
    ' INSERT����INSERT INTO�`����
    Dim insertPhrase As String
    insertPhrase = "INSERT INTO " & data.tableName
    
    ' INSERT���̃J���������� ex. (id, field1, field2)
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
    
    ' value��Ɋ܂ރ��R�[�h���̃J�E���^
    Dim recCount As Long
    recCount = 0
    ' value��𗭂ߍ��ރ��[�N
    Dim valuePhrases As String
    valuePhrases = ""
    
    ' mysql�ȊO�̏ꍇ�A�}���`�v���C���T�[�g�𖳌��������邽�߂ɁA���������ς���
    If dbms <> "mysql" Then
        multipleInsertCount = 1
    End If
    
    ' �f�[�^1�����ƃ��[�v
    Dim record As Collection
    For Each v In data.records
        Set record = v
        
        ' VALUES�吶��
        Dim valuePhrase As String
        valuePhrase = generateValuePhrase(dbms, record, data.columnTypes, nullString)
        
        If Len(valuePhrases) = 0 Then
            valuePhrases = "VALUES" & valuePhrase
        Else
            ' mysql�̃}���`�v���C���T�[�g��2���ڈȍ~
            valuePhrases = valuePhrases & "," & valuePhrase
        End If
        
        
        recCount = recCount + 1
        ' �uDBMS��mysql���}���`�v���C���T�[�g��L�������Ă���ꍇ�v�ȊO�́A��ɐ^
        ' �uDBMS��mysql���}���`�v���C���T�[�g��L�������Ă���ꍇ�v�́A���茏���ɒB�����ꍇ�̂ݏo��
        If recCount >= multipleInsertCount Then
            c.Add insertPhrase & columnPhrase & valuePhrases & ";"
            recCount = 0
            valuePhrases = ""
        End If
    Next v
    
    ' mysql�̃}���`�v���C���T�[�g�̎c�蕪�̏o��
    If Len(valuePhrases) > 0 Then
        c.Add insertPhrase & columnPhrase & valuePhrases & ";"
    End If

    Set generateInsert = c
End Function

' ----------------------------------------------------------
' VALUES��𐶐�����
' @param dbms �Ώۂ�DBMS�B�Ή����Ă���̂�mysql,oracle
' @param record ���R�[�h�̃t�B�[���h�l�̃R���N�V����
' @param columnTypes ���R�[�h�̃t�B�[���h�^�̃R���N�V����
' @param nullString �t�B�[���h�l������Ɉ�v������̂́Anull�ł���Ƃ݂Ȃ��B
' ----------------------------------------------------------
Private Function generateValuePhrase(dbms As String, record As Collection, columnTypes As Collection, nullString As String)
    Dim str As String, columnString As String
    str = ""
    Dim i As Long
    ' 1���R�[�h���̍��ڂ��ƂɃ��[�v
    For i = 1 To record.Count
        
        ' ���ڂ̒l�̕�������擾
        columnString = editColumnString(dbms, record.Item(i), columnTypes.Item(i), nullString)
        
        ' VALUES��𐶐�
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
' VALUES���̊e���ڂ̕�����𐶐�����
' @param dbms �Ώۂ�DBMS�B�Ή����Ă���̂�mysql,oracle
' @param value �f�[�^�l
' @param dataType �f�[�^�^
' @param nullString �t�B�[���h�l������Ɉ�v������̂́Anull�ł���Ƃ݂Ȃ��B
' -----------------------------------------------------------------
Private Function editColumnString(dbms As String, value As String, dataType As String, nullString As String)
    
    If value = nullString Then
        editColumnString = "null"
    Else
        Select Case dbms
        Case "mysql"
            Select Case UCase(dataType)
            Case "CHAR", "VARCHAR", "TEXT", "ENUM", "DATE", "DATETIME", "TIMESTAMP"
                ' �N�H�[�g�Ŋ���B�Ώە�����Ɋ܂܂��N�H�[�g�̓G�X�P�[�v����B
                editColumnString = "'" & Replace(value, "'", "''") & "'"
            Case Else
                editColumnString = value
            End Select
        Case "oracle"
            Select Case UCase(dataType)
            Case "DATE"
                ' TO_DATE�֐�
                editColumnString = "TO_DATE('" & value & "', 'YYYY/MM/DD HH24:MI:SS')"
            Case "CHAR", "VARCHAR", "NVARCHAR", "VARCHAR2", "NVARCHAR2"
                ' �N�H�[�g�Ŋ���B�Ώە�����Ɋ܂܂��N�H�[�g�̓G�X�P�[�v����B
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

