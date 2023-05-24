Imports Npgsql
Public Class WebForm1
    Inherits System.Web.UI.Page

    'DB接続パス
    Public conn As New NpgsqlConnection("Server=192.168.0.111; Port=5432; User Id=postgres; Password=brains; Database=brains")
    Public reader As NpgsqlDataReader
    Public dt As Date = DateTime.Now


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs)

        Try
            conn.Open()

            Dim sql As String = "select * from tokumt where tokumei = '" + TextBox1.Text + "' and tokucd = '" + TextBox2.Text + "';"
            Dim ID_Chk As NpgsqlCommand = New NpgsqlCommand(sql, conn)

            reader = ID_Chk.ExecuteReader()

            If reader.Read() = True Then

                If Trim(reader("tokumei")) = "master" And Trim(reader("tokucd")) = "0" Then

                    Response.Redirect("WebForm4.aspx")

                End If

                Dim tenpo As String = reader("tenpo")
                    Dim tokryak As String = reader("tokuryak")

                    reader.Close()

                Dim S_chk As String = "select * from time_stamp where tokumei = '" & TextBox1.Text & "' and s_dt = '" & dt.ToString("yyyy/MM/dd") & "';"
                Dim Chk As NpgsqlCommand = New NpgsqlCommand(S_chk, conn)
                    reader = Chk.ExecuteReader()

                    If reader.Read() = False Then

                        'ページ間でデータ引き渡しのため設定(下記ページ移動時、[?value= & 変数名] で移動先に渡す) 
                        Dim Login_U As String = HttpUtility.UrlEncode(TextBox1.Text)
                        Dim Tenpo_M As String = HttpUtility.UrlEncode(tenpo)

                        'CatchでResponse.Write(ex.Message)とセットすることで例外をキャッチできる(第2引数なしの場合)
                        '第2引数でFalseを指定すると例外(Respons.End)を抑制できる
                        Response.Redirect("WebForm2.aspx?name=" & Login_U & "&tenpo=" & Tenpo_M & "&tokryak=" & tokryak)

                    Else

                        Response.Redirect("WebForm3.aspx?name=" & TextBox1.Text & "&tenpo=" & tokryak)

                    End If



                Else

                    Dim c_script As String = "alert('登録されていません');"
                '描画後に実行
                'ClientScript.RegisterClientScriptBlock(Me.GetType(), "key", c_script, True)
                '描画が開始される前に実行
                ClientScript.RegisterStartupScript(Me.GetType(), "key", c_script, True)

            End If
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

        conn.Close()

    End Sub
End Class