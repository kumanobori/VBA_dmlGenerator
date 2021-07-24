Attribute VB_Name = "module_copyTableDefToDataBook"
Option Explicit

Const BOOK_MACRO = "Mysql���R�[�h�ǉ�DML�����c�[��.xlsm"
Const BOOK_DDL = "�e�[�u����`(wip).xlsx"
Const BOOK_DATA = "�}�X�^�f�[�^�\(wip).xlsx"
Const SHEET_TPL = "base"
' [wip] �e�[�u����`�\�����ƂɁA��̃}�X�^�f�[�^�\�𐶐�����
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
    
    ' �J����
    Dim wrDdl As Range
    Set wrDdl = wsDdl.Range("A8")
    Dim wrNew As Range
    Set wrNew = wsNew.Range("C3")
    Do Until wrDdl.value = ""
    
        ' �����ݒ�
        wsNew.Range(wrNew, wrNew.Offset(5, 0)).Borders.LineStyle = xlContinuous
        wrNew.Interior.Color = wrNew.Offset(0, -1).Interior.Color
        wrNew.Offset(1, 0).Interior.Color = wrNew.Offset(1, -1).Interior.Color
        wrNew.Offset(2, 0).Interior.Color = wrNew.Offset(2, -1).Interior.Color
        wrNew.Offset(3, 0).Interior.Color = wrNew.Offset(3, -1).Interior.Color
        
        ' ������
        wrNew.value = wrDdl.Offset(0, 1).value
        ' �_����
        wrNew.Offset(1, 0).value = wrDdl.Offset(0, 14).value
        ' �^
        wrNew.Offset(2, 0).value = wrDdl.Offset(0, 2).value
    
        Set wrDdl = wrDdl.Offset(1, 0)
        Set wrNew = wrNew.Offset(0, 1)
    Loop
    
    wsNew.Range("C1", wrNew).EntireColumn.AutoFit

    ' �e�[�u��������
    wsNew.Range("C1").value = wsDdl.Name
    ' �e�[�u���_����
    wsNew.Range("C2").value = wsDdl.Range("C2").value & ":" & wsDdl.Range("C3").value

End Function
