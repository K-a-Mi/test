Imports Npgsql

Public Class WebForm3
    Inherits System.Web.UI.Page
    Public conn As New NpgsqlConnection("Server=192.168.0.111; Port=5432; User Id=postgres; Password=brains; Database=brains")
    Public reader As NpgsqlDataReader
    Public tenpo As String
    Dim dt As Date = DateTime.Now
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Label1.Text = Request.QueryString("name")
        Label3.Text = Request.QueryString("tenpo")

        Label2.Text = dt.ToString("yyyy/MM/dd　HH:mm")

        Try
            conn.Open()

            Dim ST_chk As String = "select * from time_stamp where tokumei = '" & Label1.Text & "' and s_dt = '" & dt.ToString("yyyy/MM/dd") & "';"
            Dim ST_cm As NpgsqlCommand = New NpgsqlCommand(ST_chk, conn)
            reader = ST_cm.ExecuteReader()
            'ボタン制御
            If reader.Read() = True Then
                If reader("s_time") <> "" Then

                    Button1.Enabled = False

                    If reader("T_time") <> "" Then

                        Button2.Enabled = False
                    End If

                Else

                    Button2.Enabled = False

                End If

            End If
            reader.Close()

            'テーブルの存在を確認するSQL
            'Dim T_chk As String = "select * from information_schema.tables where table_name = '" & Label1.Text & "';"
            '得意先マスタ取得
            Dim T_ms As String = "select * from tokumt where tokumei = '" & Label1.Text & "' and tokuryak = '" & Label3.Text & "';"
            Dim G_ten As NpgsqlCommand = New NpgsqlCommand(T_ms, conn)
            reader = G_ten.ExecuteReader()

            reader.Read()
            '店舗を取得
            tenpo = reader("tenpo")
            reader.Close()

            '本日のデータがあるかチェック
            Dim S_chk As String = "select * from Time_stamp where S_dt = '" & dt.ToString("yyyy/MM/dd") & "' and tokumei = '" & Label1.Text & "';"
            Dim C_time As NpgsqlCommand = New NpgsqlCommand(S_chk, conn)
            reader = C_time.ExecuteReader()

            If reader.Read() = False Then
                'データなければ名前、日付、時間、シフトをインサートする
                reader.Close()
                Dim S_insert As String = "insert into Time_stamp(tokumei,S_dt,K_time,Shift) values('" & Label1.Text & "','" & dt.ToString("yyyy/MM/dd") & "', (select k_kei from " & tenpo & " where pattern = '" & Request.QueryString("shift") & "'),'" & Request.QueryString("shift") & "');"
                Dim in_com As NpgsqlCommand = New NpgsqlCommand(S_insert, conn)
                reader = in_com.ExecuteReader()

            End If
        Catch ex As Exception

        End Try
        reader.Close()
        conn.Close()
    End Sub

    Protected Sub Button3_Click(sender As Object, e As EventArgs)

        Response.Redirect("WebForm1.aspx")

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs)

        Try
            conn.Open()
            '出勤時間打刻(update)
            Dim S_time As String = "update time_stamp set s_time = '" & dt.ToString("HH:mm") & "' where tokumei = '" & Label1.Text & "' and s_dt = '" & dt.ToString("yyyy/MM/dd") & "';"
            Dim S_tsp As NpgsqlCommand = New NpgsqlCommand(S_time, conn)
            reader = S_tsp.ExecuteReader()

            Dim hello As String = "alert('おはようございます。')"
            ClientScript.RegisterClientScriptBlock(GetType(String), "hello", hello, True)

            Button1.Enabled = False

        Catch ex As Exception

        End Try
        conn.Close()
    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs)

        Try
            conn.Open()
            '退勤時間打刻(update)
            Dim T_time As String = "update time_stamp set t_time = '" & dt.ToString("HH:mm") & "' where tokumei = '" & Label1.Text & "' and s_dt = '" & dt.ToString("yyyy/MM/dd") & "';"
            Dim T_tsp As NpgsqlCommand = New NpgsqlCommand(T_time, conn)
            reader = T_tsp.ExecuteReader()

            Dim bye As String = "alert('お疲れさまでした。')"
            ClientScript.RegisterClientScriptBlock(GetType(String), "bye", bye, True)

            Button2.Enabled = False

        Catch ex As Exception

        End Try
        conn.Close()
    End Sub
End Class