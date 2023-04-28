﻿Imports Npgsql
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
    Dim new_name As String
    Dim time_stp As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            conn.Open()

            Dim sql As String = "select * from tokumt where tokumei = '" + TextBox1.Text + "' and tokucd = '" + TextBox2.Text + "';"

            Dim Chec_acc As NpgsqlCommand = New NpgsqlCommand(sql, conn)

            'Chec_acc.ExecuteScalar()
            reader = Chec_acc.ExecuteReader()

            If reader.Read = True Then

                'ファイルの存在確認・作成
                Call file_check()

                '本日のシフトに入力あれば打刻フォームを開く(入力なければシフト入力フォームから)
                If time_stp <> "" Then

                    form3.U_name = TextBox1.Text
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

            '今月のシートを作成
            Call new_create_sheet()

        Else

            Dim ps As Process() = Process.GetProcessesByName("EXCEL")
            For Each p As Process In ps
                p.Kill()
            Next

            Try
                '月跨ぎスルー
                If Integer.Parse(dt.ToString("dd")) >= 16 Then

                    '今月シートがあるか確認
                    Dim ea As Excel.Application = New Excel.Application
                    Dim wbs As Excel.Workbooks = ea.Workbooks
                    Dim wb As Excel.Workbook = wbs.Open("C:\Users\User\" & reader("tokumei").ToString.Trim() & ".xlsx", Password:="sample")

                    Dim ss As Excel.Sheets = wb.Worksheets
                    Dim ws As Excel.Worksheet = ss(2)

                    'シート名の取得方法がわからなかった為シートに関数を入力してシート名を取得できるようにした
                    If ws.Range("A1").Value <> dt.ToString("yyyy_MM") Then

                        ws = ss("原紙")
                        '原紙シートをコピーして名前を年_月にする
                        ws.Copy(After:=ws)
                        ws = ss("原紙 (2)")
                        ws.Name = dt.ToString("yyyy_MM")

                        Dim cs As Excel.Range = ws.Cells

                        '1カ月の日付を入力するために指定セルに入力する
                        cs.Range("K2").Value = dt.ToString("yyyy")
                        cs.Range("M2").Value = dt.ToString("MM")

                    ElseIf ws.Range("A1").Value = dt.ToString("yyyy_MM") Then

                        '本日のシフトを入力しているかチェック
                        If Integer.Parse(dt.ToString("dd")) <= 15 Then

                            Dim lastday As Integer = Integer.Parse(Date.DaysInMonth(dt.ToString("yyyy"), dt.AddMonths(-1).ToString))
                            Dim get_cell As Integer = lastday + Integer.Parse(dt.ToString("dd")) - 12
                            time_stp = ws.Range("G" & get_cell).Value

                        Else

                            Dim get_day As Integer = Integer.Parse(dt.ToString("dd")) - 12
                            time_stp = ws.Range("G" & get_day).Value

                        End If

                    End If

                    ea.DisplayAlerts = False

                    '保存
                    wb.Save()
                    '終了
                    ea.Quit()

                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End If

    End Sub

    Sub new_create_sheet()

        '去年と先月を取得
        Dim dt_month As String = dt.AddMonths(-1).Month
        Dim dt_year As String = dt.AddYears(-1).Year

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open("C:\Users\User\" & reader("tokumei").ToString.Trim() & ".xlsx", Password:="sample")

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss("原紙")

        '原紙シートをコピー
        ws.Copy(After:=ws)
        'コピーしたシートを選択
        ws = ss("原紙 (2)")

        Dim cs As Excel.Range = ws.Cells

        '1カ月の日付を入力するために指定セルに入力する・シート名を変更
        If Integer.Parse(dt.ToString("dd")) >= 16 Then

            cs.Range("K2").Value = dt.ToString("yyyy")
            cs.Range("M2").Value = dt.ToString("MM")
            ws.Name = dt.ToString("yyyy_MM")

        ElseIf Integer.Parse(dt.ToString("MM")) = 1 Then

            cs.Range("K2").Value = dt_year
            cs.Range("M2").Value = dt_month
            ws.Name = dt_year & "_" & dt_month

        Else

            ws.Range("K2").Value = dt.ToString("yyyy")
            ws.Range("M2").Value = dt_month
            ws.Name = dt.ToString("yyyy_") & dt_month

        End If

        ea.DisplayAlerts = False

        '保存
        wb.Save()
        '終了
        ea.Quit()

    End Sub

End Class
