Imports System.ComponentModel
Imports System.Globalization
Imports System.Text

Class MainWindow
    Implements INotifyPropertyChanged


    Private Const Max_Length As Integer = 90
    Private _OrginScript As String
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Property OrginScript As String
        Get
            Return _OrginScript
        End Get
        Set
            _OrginScript = CleanSqlScript(Value)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("OrginScript"))
        End Set
    End Property

    Private Function GenerateSelectVbCode(tbl As SqlTableInfo) As String
        Dim helper As New StringBuilder
        Dim sqlHelper As New StringBuilder

        helper.AppendLine("Dim cmd As SqlCommand = Me.Command")
        helper.AppendLine("cmd.Parameters.Clear()")
        helper.Append("cmd.CommandText =")

        Dim line As Integer = 1
        sqlHelper.Append("""SELECT ")
        For Each f As String In tbl.Feilds
            sqlHelper.Append($"{f},")
            If sqlHelper.Length > Max_Length * Line Then
                sqlHelper.AppendLine(" "" _")
                sqlHelper.Append("& """)
                Line += 1
            End If
        Next
        sqlHelper.Remove(sqlHelper.Length - 1, 1)

        sqlHelper.Append($" FROM {tbl.TableName}")
        sqlHelper.AppendLine("""")

        helper.Append(sqlHelper.ToString)
        helper.AppendLine()

        Return helper.ToString
    End Function

    Private Function GenerateUpdateVbCode(tbl As SqlTableInfo) As String
        Dim helper As New StringBuilder
        Dim sqlHelper As New StringBuilder
        helper.AppendLine("Dim cmd As SqlCommand = Me.Command")
        helper.AppendLine("cmd.Parameters.Clear()")
        helper.Append("cmd.CommandText =")
        sqlHelper.Append($"""UPDATE {tbl.TableName} SET ")
        Dim formatter As TextInfo = Threading.Thread.CurrentThread.CurrentCulture.TextInfo
        Dim line As Integer = 1
        For Each f As String In tbl.Feilds
            sqlHelper.Append($"{f}=@{formatter.ToTitleCase(f)},")
            If sqlHelper.Length > Max_Length * line Then
                sqlHelper.AppendLine(" "" _")
                sqlHelper.Append("& """)
                line += 1
            End If
        Next
        sqlHelper.Remove(sqlHelper.Length - 1, 1)
        sqlHelper.AppendLine("""")

        helper.Append(sqlHelper.ToString)
        helper.AppendLine()

        For Each f As String In tbl.Feilds
            helper.AppendLine($"cmd.Parameters.AddWithValue(""{formatter.ToTitleCase(f)}"", DBNull)")
        Next

        Return helper.ToString
    End Function

    Private Function GenerateInsertVbCode(tbl As SqlTableInfo) As String
        Dim helper As New StringBuilder
        Dim sqlHelper As New StringBuilder
        helper.AppendLine("Dim cmd As SqlCommand = Me.Command")
        helper.AppendLine("cmd.Parameters.Clear()")
        helper.Append("cmd.CommandText =")
        sqlHelper.Append($"""INSERT {tbl.TableName} (")
        Dim line As Integer = 1
        For Each f As String In tbl.Feilds
            sqlHelper.Append($"{f},")
            If sqlHelper.Length > Max_Length * line Then
                sqlHelper.AppendLine(" "" _")
                sqlHelper.Append("& """)
                line += 1
            End If
        Next
        sqlHelper.Remove(sqlHelper.Length - 1, 1)
        sqlHelper.Append(") VALUES (")
        Dim formatter As TextInfo = Threading.Thread.CurrentThread.CurrentCulture.TextInfo
        For Each f As String In tbl.Feilds
            sqlHelper.Append($"@{formatter.ToTitleCase(f)},")
            If sqlHelper.Length > Max_Length * line Then
                sqlHelper.AppendLine(" "" _")
                sqlHelper.Append("& """)
                line += 1
            End If
        Next
        sqlHelper.Remove(sqlHelper.Length - 1, 1)
        sqlHelper.AppendLine(")""")

        helper.Append(sqlHelper.ToString)
        helper.AppendLine()

        For Each f As String In tbl.Feilds
            helper.AppendLine($"cmd.Parameters.AddWithValue(""{formatter.ToTitleCase(f)}"", DBNull)")
        Next
        Return helper.ToString
    End Function

    Private Function FormatSqlScript(sqlScript As String) As String
        Dim scripts As String() = sqlScript.Split(","c)
        Dim helper As New StringBuilder
        Dim pos As Integer = 1
        helper.Append("""")
        For Each w As String In scripts
            helper.Append(w)
            If helper.Length > Max_Length * pos Then
                helper.AppendLine(" "" _")
                helper.Append("& """)
                pos += 1
            End If
        Next
        Return helper.ToString
    End Function

    Private Function GetTableInfoFromInsertSql(script As String) As SqlTableInfo
        Dim tbl As New SqlTableInfo
        Dim wholeParts As String() = (From w In script.Split(New Char() {" "c, CChar(vbCrLf), ","c, CChar(String.Empty), "("c, ")"c, CChar（vbLf）, "["c, "]"c, CChar(vbTab)})
                                      Where w <> String.Empty).ToArray
        Dim hasTableName As Boolean = False
        For Each w In wholeParts
            Select Case w.ToUpper
                Case "INSERT", "INTO", "DBO", "."
                Case "VALUES"
                    Exit For
                Case Else
                    If hasTableName Then
                        tbl.Feilds.Add(w)
                    Else
                        tbl.TableName = w
                        hasTableName = True
                    End If
            End Select
        Next

        Return tbl
    End Function

    Private Function GetTableInfoFromSelectSql(script As String) As SqlTableInfo
        Dim tbl As New SqlTableInfo
        Dim wholeParts As String() = (From w In script.Split(New Char() {" "c, CChar(vbCrLf), ","c, CChar(String.Empty), "("c, ")"c, CChar（vbLf）, "["c, "]"c, CChar(vbTab)})
                                      Where w <> String.Empty).ToArray
        Dim hasTableName As Boolean = False
        For Each w In wholeParts
            Select Case w.ToUpper
                Case "SELECT"
                Case "FROM"
                    hasTableName = True
                Case Else
                    If hasTableName Then
                        tbl.TableName = w
                        Exit For
                    Else
                        tbl.Feilds.Add(w)
                    End If
            End Select
        Next

        Return tbl
    End Function

    Private Function GetVbCodeFromSql(script As String) As String
        Dim paras As String() = FindParameters(script)
        Dim helper As New StringBuilder
        helper.AppendLine("Dim cmd As SqlCommand = Me.Command")
        helper.AppendLine("cmd.Parameters.Clear()")
        helper.Append("cmd.CommandText =")
        Dim segs As String() = (From s In script.Split(New Char() {CChar(vbCrLf), CChar(vbLf), CChar(vbTab)}) Where s <> String.Empty).ToArray
        If segs.Count > 0 Then
            helper.Append($"""{segs(0)} """)
        End If
        If segs.Count > 1 Then
            helper.AppendLine(" _")
            For i As Integer = 1 To segs.Count - 2
                helper.AppendLine($"& ""{segs(i)} "" _")
            Next
            helper.AppendLine($"& ""{segs.Last} """)
        End If

        helper.AppendLine()
        Dim formatter As TextInfo = Threading.Thread.CurrentThread.CurrentCulture.TextInfo
        For Each f As String In paras
            helper.AppendLine($"cmd.Parameters.AddWithValue(""{formatter.ToTitleCase(f)}"", DBNull)")
        Next

        Return helper.ToString
    End Function

    Private Function FindParameters(script As String) As String()
        Dim ps As New Dictionary(Of String, String)
        Dim segs As String() = (From s In script.Split(New Char() {"="c, ","c, " "c, ")"c, ";"c}) Where s <> String.Empty).ToArray
        For Each seg As String In segs
            seg = seg.Trim
            If seg.Length > 0 AndAlso seg.First = "@"c Then
                If seg.Last = ";" Then
                    seg = seg.Substring(0, seg.Length - 1)
                End If
                If Not seg.Contains("@@") Then
                    seg = seg.Substring(1)
                    If Not ps.ContainsKey(seg) Then
                        ps.Add(seg, seg)
                    End If
                End If
            End If
        Next
        Return ps.Keys.ToArray
    End Function

    Private Function CleanSqlScript(script As String) As String
        Dim helper As New StringBuilder
        Dim lines As String() = script.Split(New Char() {"&"c})
        For Each l As String In lines
            l = l.Trim
            If l.Count > 0 Then
                If l.Last = "_" AndAlso l.First = """"c Then
                    helper.AppendLine(l.Substring(1, l.Length - 4))
                ElseIf l.Last = """"c AndAlso l.First = """"c Then
                    helper.AppendLine(l.Substring(1, l.Length - 2))
                Else
                    helper.AppendLine(l)
                End If
            End If
        Next
        Return helper.ToString
    End Function

    Private Function GetTableInfoFromUpdateSql(script As String) As SqlTableInfo
        Dim tbl As New SqlTableInfo
        Dim wholeParts As String() = (From w In script.Split(New Char() {" "c, CChar(vbCrLf), ","c, CChar(String.Empty), "("c, ")"c, CChar（vbLf）, "["c, "]"c, CChar(vbTab)})
                                      Where w <> String.Empty).ToArray
        Dim hasTableName As Boolean = False
        For Each w In wholeParts
            Select Case w.ToUpper
                Case "UPDATE", ".", "DBO"
                Case "WHERE"
                    Exit For
                Case Else
                    If hasTableName Then
                        If w.First = "<"c Then
                            tbl.Feilds.Add(w.Substring(1))
                        End If
                    Else
                        tbl.TableName = w
                        hasTableName = True
                    End If
            End Select
        Next

        Return tbl
    End Function

    Private Sub GenerateCode(sender As Object, e As RoutedEventArgs)
        txtOrg.GetBindingExpression(TextBox.TextProperty).UpdateSource()
        Dim tbl As SqlTableInfo = Nothing
        Dim script As String = OrginScript
        If script IsNot Nothing Then
            If script.Contains("<") Then
                If script.Contains("INSERT") OrElse script.Contains("insert") Then
                    tbl = GetTableInfoFromInsertSql(script)
                    txtResult.Text = GenerateInsertVbCode(tbl)
                ElseIf script.Contains("UPDATE") OrElse script.Contains("update") Then
                    tbl = GetTableInfoFromUpdateSql(script)
                    txtResult.Text = GenerateUpdateVbCode(tbl)
                End If
            Else
                txtResult.Text = GetVbCodeFromSql(OrginScript)
            End If
        End If
    End Sub

    Private Sub GenerateCodeFromSSMS(sender As Object, e As RoutedEventArgs)
        txtOrg.GetBindingExpression(TextBox.TextProperty).UpdateSource()
        Dim tbl As SqlTableInfo = Nothing
        Dim script As String = OrginScript
        If script IsNot Nothing Then
            If script.Contains("INSERT") OrElse script.Contains("insert") Then
                tbl = GetTableInfoFromInsertSql(script)
                txtResult.Text = GenerateInsertVbCode(tbl)
            ElseIf script.Contains("UPDATE") OrElse script.Contains("update") Then
                tbl = GetTableInfoFromUpdateSql(script)
                txtResult.Text = GenerateUpdateVbCode(tbl)
            ElseIf script.Contains("SELECT") OrElse script.Contains("select") Then
                tbl = GetTableInfoFromSelectSql(script)
                txtResult.Text = GenerateSelectVbCode(tbl)
            End If
        End If
    End Sub

    Private Sub GenerateCodeFromSelf(sender As Object, e As RoutedEventArgs)
        txtOrg.GetBindingExpression(TextBox.TextProperty).UpdateSource()
        Dim script As String = OrginScript
        If script IsNot Nothing Then
            txtResult.Text = GetVbCodeFromSql(OrginScript)
        End If
    End Sub
End Class
