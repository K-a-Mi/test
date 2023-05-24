Imports Npgsql

Public Class WebForm2
    Inherits System.Web.UI.Page
    Public conn As New NpgsqlConnection("Server=192.168.0.111; Port=5432; User Id=postgres; Password=brains; Database=brains")
    Public reader As NpgsqlDataReader
    Public list() As String = {"<未選択>"}
    'javascriptに渡す変数を宣言
    Protected data As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '前ページからデータ受け取り
        Dim Login_U As String = Request.QueryString("name")
        Label2.Text = HttpUtility.UrlDecode(Login_U)
        Dim Tenpo_M As String = Request.QueryString("tenpo")
        data = Request.QueryString("tokryak")
        Try
            conn.Open()

            Dim sql As String = "select * from " & Tenpo_M & ";"
            Dim Shift_Chk As NpgsqlCommand = New NpgsqlCommand(sql, conn)
            reader = Shift_Chk.ExecuteReader()

            If DropDownList1.Items.Count < 1 Then

                DropDownList1.Items.Add("<未選択>")

                If DropDownList1.Items.Count = 1 Then

                    Dim i As Integer = 1
                    While reader.Read() = True

                        DropDownList1.Items.Add(reader("pattern"))
                        ReDim Preserve list(i)
                        list(i) = reader("pattern")
                        i = i + 1

                    End While

                End If
            End If

        Catch ex As Exception

        End Try

        'シフト確認メッセージ表示
        'Button2.Attributes("onclick") = "var ret=confirm('シフトは間違いないですか？');if (ret == true) { return true; } else { return false; }"

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs)

        Response.Redirect("WebForm1.aspx")

    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs)

        Try

            If DropDownList1.SelectedItem.Text <> "<未選択>" Then

                'Response.Redirect("WebForm3.aspx?name=" & Label2.Text & "&tenpo=" & Request.QueryString("tokryak"))

            Else

                'シフト選択メッセージ
                'Dim c_script As String = "alert('シフトを選択してください。')"
                'ClientScript.RegisterStartupScript(Me.GetType(), "shift", c_script, True)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
End Class