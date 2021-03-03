Attribute VB_Name = "SQLUtility_Insert"
' @summary INSERT���̃e�[�u���E����w�肷�镔�����쐬���쐬���܂�
' @param table �e�[�u����
' @param ranges �񖼂��Z�b�g���ꂽ�Z���͈̔�
' @return �e�[�u���E����w�肷�镔����INSERT����Ԃ��܂�
Public Function usrMakeInsertSQLHead(table As String, ParamArray ranges() As Variant)
    Dim columns As String
    For Each r In ranges
        For Each c In r
            If c.Value = vbNullString Then
                usrSqlInsertColumns = "#�J����������!"
                Exit Function
            End If
            columns = columns & IIf(columns <> vbNullString, ",", vbNullString)
            columns = columns & c.Value
        Next c
    Next r
    usrSqlInsertColumns = "insert into " & table & "(" & columns & ")"
End Function

' @summary INSERT���̓o�^���e���w�肷�镔�����쐬���܂�
' @param ranges �o�^���e���Z�b�g���ꂽ�Z���͈̔�
' @return �o�^���e�iVALUES�j���w�肷�镔����INSERT����Ԃ��܂��B
Public Function usrMakeInsertSQLBody(ParamArray ranges() As Variant)
    Dim values As String
    For Each r In ranges
        For Each c In r
            values = values & IIf(values <> vbNullString, ",", vbNullString)
            ' "null"�A"Null"�A"NULL"�݂����ȕ������ null�Ƃ��ēo�^
            values = values & IIf(LCase(c.Value) = "null", "null", "'" & c.Value & "'")
        Next c
    Next r
    usrSqlInsertValues = "values(" & values & ");"
End Function