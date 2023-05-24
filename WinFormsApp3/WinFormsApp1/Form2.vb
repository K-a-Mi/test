Imports Microsoft.Office.Interop

Public Class Form2

    Public U_name As String
    '本日の日時取得
    Dim dt1 As DateTime = DateTime.Now
    Dim time_stp2 As String
    Dim S_time As String
    Dim N_year As String = dt1.AddYears(-1).ToString("yyyy")
    Dim N_month = dt1.AddMonths(-1).ToString("MM")


    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'ユーザー名表示
        Label1.Text = U_name

        'シフトパターンを取得するセル位置を設定
        Dim i As Integer = 4
        Dim filename = "C:\Users\User\" & U_name & ".xlsx"
        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")
        Dim ss As Excel.Sheets = wb.Worksheets

        'シフトパターンを読み込むシートを判断
        If Integer.Parse(dt1.ToString("dd")) <= 15 And dt1.ToString("MM") = 1 Then

            S_time = N_year.ToString & "_" & N_month

        ElseIf Integer.Parse(dt1.ToString("dd")) <= 15 Then

            S_time = dt1.ToString("yyyy") & "_" & N_month

        Else

            S_time = dt1.ToString("yyyy_MM")

        End If

        Dim ws As Excel.Worksheet = ss(S_time)
        Dim cs As Excel.Range = ws.Cells

        'シフトパターンを読み込んでコンボボックスにセット
        While ws.Range("B" & i).Value <> ""

            ComboBox1.Items.Add(ws.Range("B" & i).Value)
            i = i + 1

        End While

        ea.DisplayAlerts = False
        wb.Save()
        ea.Quit()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'コンボボックス入力チェック
        If ComboBox1.Text <> "" Then

            Dim dialog As DialogResult = MessageBox.Show(" シフトは　" & ComboBox1.Text & "　で間違いはないですか？", "確認メッセージ", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            'シフトに間違いない場合、本日のシフトを入力
            If dialog = DialogResult.Yes Then

                Dim filename = "C:\Users\User\" & U_name & ".xlsx"
                Dim ea As Excel.Application = New Excel.Application
                Dim wbs As Excel.Workbooks = ea.Workbooks
                Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")
                Dim ss As Excel.Sheets = wb.Worksheets
                Dim ws As Excel.Worksheet = ss(S_time)
                Dim cs As Excel.Range = ws.Cells
                Dim sif As Integer

                '入力セルを取得し入力
                If Integer.Parse(dt1.ToString("dd")) <= 15 Then

                    Dim year As Integer = Integer.Parse(dt1.ToString("yyyy"))
                    Dim lastday As Integer = Date.DaysInMonth(year, Integer.Parse(N_month))
                    sif = lastday + Integer.Parse(dt1.ToString("dd")) - 12
                    ws.Range("G" & sif).Value = ComboBox1.Text

                Else

                    sif = Integer.Parse(dt1.ToString("dd")) - 12
                    ws.Range("G" & sif).Value = ComboBox1.Text

                End If

                ea.DisplayAlerts = False
                wb.Save()
                ea.Quit()

                Dim form3 As Form3 = New Form3
                form3.U_name = U_name
                form3.ws_name = S_time
                form3.get_cell = sif
                form3.ShowDialog()

                ComboBox1.Items.Clear()
                ComboBox1.Text = ""

                Me.Close()

            End If

        Else

            MsgBox("シフトを選択してください。")

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()

    End Sub
End Class