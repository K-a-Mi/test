Imports Npgsql
Imports Microsoft.Office.Interop

Public Class Form1

    'DB接続パス
    Public conn As New NpgsqlConnection("Server=192.168.0.111; Port=5432; User Id=postgres; Password=brains; Database=brains")

    Public reader As NpgsqlDataReader

    Public mei As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'DB接続
        conn.Open()

        '実行するSQL
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("select * from tokumt;", conn)

        Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter(cmd)
        Dim dt As DataTable = New DataTable()

        da.Fill(dt)

        'SQL実行
        reader = cmd.ExecuteReader()

        'DataGridView1にデータ表示
        DataGridView1.DataSource = dt

        'Readでレコードデータを取得する(次のレコードデータを取得する)
        While (reader.Read())

            mei = reader("tokumei")

            'コンボボックスに"tokuryaku"フィールドの取得したデータをセット
            Me.ComboBox1.Items.Add(reader("tokuryak"))

        End While

        console()

        'ctl()
        conn.Close()

        'End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        MsgBox(mei)

    End Sub

    Private Sub ctl()

        Dim filename = "C:\Users\User\source\repos\sample1\sample1\test1.xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename)

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(1)
        Dim cs As Excel.Range = ws.Cells

        System.Runtime.InteropServices.Marshal.ReleaseComObject(cs)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ss)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wbs)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(ea)

    End Sub

    Private Sub console()

        MsgBox("OK")

    End Sub
End Class
