Imports Npgsql
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form1

    'DB接続パス
    Public conn As New NpgsqlConnection("Server=192.168.0.111; Port=5432; User Id=postgres; Password=brains; Database=brains")
    Public conn2 As New NpgsqlConnection("Server=192.168.0.111; Port=5432; User Id=postgres; Password=brains; Database=postgres")

    Public reader As NpgsqlDataReader
    Public reader1 As NpgsqlDataReader
    Public reader2 As NpgsqlDataReader
    Public reader3 As NpgsqlDataReader

    Public mei As String

    'Form2のインスタンス化
    Public f2 As Form2 = New Form2

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'DB接続
        conn.Open()

        '実行するSQL
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("select * from tokumt;", conn)

        Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter(cmd)
        Dim dt As DataTable()

        'da.Fill(dt)

        'SQL実行
        reader = cmd.ExecuteReader()

        'DataGridView1にデータ表示
        DataGridView1.DataSource = dt

        'Readでレコードデータを取得する(次のレコードデータを取得する)
        While reader.Read()

            mei = reader("tokumei")

            'コンボボックスに"tokuryaku"フィールドの取得したデータをセット
            Me.ComboBox1.Items.Add(reader("tokuryak"))

        End While

        Console()

        Ctl()
        conn.Close()

        'End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        MsgBox(mei)

    End Sub

    Private Sub Ctl()

        Dim filename = "C:\Users\User\test1.xlsx"
        'Dim filename = "C:\Users\User\source\repos\sample1\sample1\test1.xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        '保護ファイルを開ける(ない場合はPasswordは省略)
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="test1")

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(1)
        Dim cs As Excel.Range = ws.Cells

        'セルに入力（セル指定）
        cs.Range("A1").Value = "hello"

        'セルに入力（座標指定）
        cs.Cells(2, 1) = "world"

        '確認メッセージ非表示
        ea.DisplayAlerts = False

        '上書き保存
        wb.Save()

        '名前を付けて保存
        'wb.SaveAs2("C:\Users\User\test2.xlsx")

        '保存しないで閉じる
        'wb.Close(False)

        '保存して閉じる（上書き）
        'wb.Close(True)

        '解放
        System.Runtime.InteropServices.Marshal.ReleaseComObject(cs)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ss)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wbs)

        '終了
        ea.Quit()

        '解放
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ea)

    End Sub

    Private Sub Console()

        MsgBox("OK")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Call show_form()

    End Sub

    Public Sub show_form()

        Try
            f2.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()

    End Sub

    Private Sub ex()

        Dim i = 1

        '本日の日時取得
        Dim dt1 As DateTime = DateTime.Now

        Dim filename = "C:\Users\User\test1.xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="test1")

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(1)
        Dim cs As Excel.Range = ws.Cells

        Try
            conn.Open()

            Dim cmd As NpgsqlCommand = New NpgsqlCommand("select * from tokumt;", conn)

            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter(cmd)

            reader = cmd.ExecuteReader()

            While reader.Read()

                i += 1
                cs.Range("B" + i.ToString).Value = Replace(reader("tokucd"), Space(1), String.Empty)
                cs.Range("C" + i.ToString).Value = Replace(reader("tokumei"), Space(1), String.Empty)
                cs.Range("D" + i.ToString).Value = reader("tokuryak").ToString.Trim()
                cs.Range("E" + i.ToString).Value = reader("mukouflg")
                cs.Range("F" + i.ToString).Value = reader("kousinniji")

            End While

            '日時入力(形式指定)
            cs.Range("B1").Value = dt1.ToString("yyyy/MM/dd")
            cs.Range("C1").Value = dt1.ToString("HH:mm:ss")

            'シートの後ろにシートをコピー
            ws.Copy(After:=ws)
            '作業シートをコピーしたシートを選択
            ws = ss("copy1 (2)")
            'シートの名前を変更
            ws.Name = "test1"

            ea.DisplayAlerts = False

            wb.Save()

            'System.Runtime.InteropServices.Marshal.ReleaseComObject(cs)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(ws)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(ss)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(wbs)

            '終了
            ea.Quit()

            '解放
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(ea)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Call ex()

        MsgBox("完了")

        Try

        Catch ex As System.IO.FileNotFoundException

        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        'テーブル有無確認SQL(WHERE table_name = (検索テーブル)
        Dim tb_name = "SELECT * FROM information_schema.tables WHERE table_name = 'test1';"

        Try
            conn2.Open()

            Dim cmd As NpgsqlCommand = New NpgsqlCommand(tb_name, conn2)
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter(cmd)

            reader = cmd.ExecuteReader()

            If reader.Read() = True Then

                MsgBox("存在します。")

            Else

                reader.Close()

                Dim creat_tb2 = "create table public.test1(id character varying(100),name character varying(100),kousinniji timestamp(6) without time zone);"
                Dim creat_tb3 = "comment on table public.test1 is 'テスト1テーブル'; comment on column test1.id is 'ID'; comment on column test1.name is '名称'; comment on column test1.kousinniji is '更新日時';"

                Dim creat_cmd2 As NpgsqlCommand = New NpgsqlCommand(creat_tb2, conn2)
                Dim creat_cmd3 As NpgsqlCommand = New NpgsqlCommand(creat_tb3, conn2)

                'Dim da2 As NpgsqlDataAdapter = New NpgsqlDataAdapter(creat_cmd2)
                reader2 = creat_cmd2.ExecuteReader()
                reader2.Close()

                'Dim da3 As NpgsqlDataAdapter = New NpgsqlDataAdapter(creat_cmd3)
                reader3 = creat_cmd3.ExecuteReader()
                reader.Close()

                MsgBox("作成しました。")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        conn2.Close()
    End Sub
End Class
