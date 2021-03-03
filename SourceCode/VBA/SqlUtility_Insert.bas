Attribute VB_Name = "SQLUtility_Insert"
' @summary INSERT文のテーブル・列を指定する部分を作成を作成します
' @param table テーブル名
' @param ranges 列名がセットされたセルの範囲
' @return テーブル・列を指定する部分のINSERT文を返します
Public Function usrMakeInsertSQLHead(table As String, ParamArray ranges() As Variant)
    Dim columns As String
    For Each r In ranges
        For Each c In r
            If c.Value = vbNullString Then
                usrSqlInsertColumns = "#カラム名が空白!"
                Exit Function
            End If
            columns = columns & IIf(columns <> vbNullString, ",", vbNullString)
            columns = columns & c.Value
        Next c
    Next r
    usrSqlInsertColumns = "insert into " & table & "(" & columns & ")"
End Function

' @summary INSERT文の登録内容を指定する部分を作成します
' @param ranges 登録内容がセットされたセルの範囲
' @return 登録内容（VALUES）を指定する部分のINSERT文を返します。
Public Function usrMakeInsertSQLBody(ParamArray ranges() As Variant)
    Dim values As String
    For Each r In ranges
        For Each c In r
            values = values & IIf(values <> vbNullString, ",", vbNullString)
            ' "null"、"Null"、"NULL"みたいな文字列は nullとして登録
            values = values & IIf(LCase(c.Value) = "null", "null", "'" & c.Value & "'")
        Next c
    Next r
    usrSqlInsertValues = "values(" & values & ");"
End Function