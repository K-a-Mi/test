Imports Npgsql
Imports Microsoft.Office.Interop
Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class Form1

    'DB接続パス
    Public conn As New NpgsqlConnection("Server=192.168.0.111; Port=5432; User Id=postgres; Password=brains; Database=brains")
    Public reader As NpgsqlDataReader
    Public form2 As Form2 = New Form2()
    Public form3 As Form3 = New Form3()
    Dim dt As DateTime = DateTime.Now
    Dim time_stp As String
    Dim T_cell As Integer
    Dim sheetname As String = ""

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            conn.Open()

            Dim sql As String = "select * from tokumt where tokumei = '" + TextBox1.Text + "' and tokucd = '" + TextBox2.Text + "';"
            Dim Chec_acc As NpgsqlCommand = New NpgsqlCommand(sql, conn)

            reader = Chec_acc.ExecuteReader()

            If reader.Read = True Then

                'ユーザーファイルの存在確認・作成
                Call file_check()

                '本日のシフトに入力あれば打刻フォームを開く(入力なければシフト入力フォームから)
                If time_stp <> "" Then

                    form3.U_name = TextBox1.Text
                    form3.get_cell = T_cell
                    form3.ws_name = sheetname
                    form3.ShowDialog()

                Else

                    form2.U_name = TextBox1.Text
                    form2.ShowDialog()

                End If

                TextBox1.Text = ""
                TextBox2.Text = ""

            Else

                MsgBox("登録されていません。")

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        conn.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()

    End Sub

    Private Sub file_check()

        Try
            'ユーザーファイルの存在を確認(新規ユーザーか)
            If File.Exists("C:\Users\User\" & reader("tokumei").ToString.Trim() & ".xlsx") = False Then

                Dim filename As String = "C:\Users\User\sample.xlsx"
                'ユーザーファイルが存在してなければ原本コピーして作成する
                File.Copy(filename, "C:\Users\User\" & reader("tokumei").ToString.Trim() & ".xlsx")

                'copyでプロセス残るので消す(これ以外の方法わからなかった)
                Dim ps As Process() = Process.GetProcessesByName("EXCEL")
                For Each p As Process In ps
                    p.Kill()
                Next

            End If

            Dim ea As Excel.Application = New Excel.Application
            Dim wbs As Excel.Workbooks = ea.Workbooks
            Dim wb As Excel.Workbook = wbs.Open("C:\Users\User\" & reader("tokumei").ToString.Trim() & ".xlsx", Password:="sample")
            Dim ss As Excel.Sheets = wb.Worksheets
            Dim sheetcount As Integer = ss.Count
            Dim ws As Excel.Worksheet = ss(1)
            Dim Ex_sheetname As String = ""

            'シートが複数ある場合シート名取得
            If sheetcount <> 1 Then

                ws = ss(2)
                Ex_sheetname = ws.Range("A1").Value

            End If

            Dim N_year As String = dt.AddYears(-1).ToString("yyyy")
            Dim N_month As String = dt.AddMonths(-1).ToString("MM")

            'シート名を決定する
            If dt.ToString("dd") >= 16 Then
                sheetname = dt.ToString("yyyy_MM")
            ElseIf dt.ToString("MM") = 1 Then
                sheetname = N_year & "_" & N_month
            Else
                sheetname = dt.ToString("yyyy") & "_" & N_month
            End If

            '今月分シートなければ作成・シート名変更
            If sheetcount = 1 Or sheetcount >= 2 And Ex_sheetname <> sheetname Then

                ws = ss("原紙")
                '原紙シートをコピー
                ws.Copy(After:=ws)
                'コピーしたシートを選択
                ws = ss("原紙 (2)")
                '名前の変更
                ws.Name = sheetname

                '日付をセットする
                ws.Range("K2").Value = Strings.Left(sheetname, 4)
                ws.Range("M2").Value = Strings.Right(sheetname, 2)

            End If

            '本日シフトの入力値取得(月をまたいでいる場合、前月の日数からセルを指定して取得しているが、年度またぎの12月の日数に変化ないため考慮していない)
            If Integer.Parse(dt.ToString("dd")) <= 15 Then

                Dim lastday As Integer = Date.DaysInMonth(Integer.Parse(dt.ToString("yyyy")), Integer.Parse(N_month))
                T_cell = lastday + Integer.Parse(dt.ToString("dd")) - 12
                time_stp = ws.Range("G" & T_cell).Value

            Else

                T_cell = Integer.Parse(dt.ToString("dd")) - 12
                time_stp = ws.Range("G" & T_cell).Value

            End If

            ea.DisplayAlerts = False
            '保存
            wb.Save()
            '終了
            ea.Quit()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            End
        End Try

    End Sub
End Class
