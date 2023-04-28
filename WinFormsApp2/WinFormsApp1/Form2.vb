Imports Microsoft.Office.Interop

Public Class Form2

    Public U_name As String
    '本日の日時取得
    Dim dt1 As DateTime = DateTime.Now
    Dim time_stp2 As String


    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'ユーザー名表示
        Label1.Text = U_name

        Dim i As Integer = 4

        Dim filename = "C:\Users\User\" & U_name & ".xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(dt1.ToString("yyyy_MM"))
        Dim cs As Excel.Range = ws.Cells

        'シフトパターンを読み込む
        While ws.Range("B" & i).Value <> ""

            ComboBox1.Items.Add(ws.Range("B" & i).Value)
            i = i + 1

        End While

        ea.DisplayAlerts = False
        wb.Save()
        ea.Quit()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '去年と先月を取得
        Dim dt_month As String = dt1.AddMonths(-1).Month
        Dim dt_year As String = dt1.AddYears(-1).Year

        Dim ws_name As String

        'コンボボックス入力チェック
        If ComboBox1.Text <> "" Then

            Dim dialog As DialogResult = MessageBox.Show(" シフトは　" & ComboBox1.Text & "　で間違いはないですか？", "確認メッセージ", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If dialog = DialogResult.Yes Then

                Dim filename = "C:\Users\User\" & U_name & ".xlsx"

                Dim ea As Excel.Application = New Excel.Application
                Dim wbs As Excel.Workbooks = ea.Workbooks
                Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")

                Dim ss As Excel.Sheets = wb.Worksheets

                'シフトパターンを入力するシート名を変数に入れる
                If Integer.Parse(dt1.ToString("dd")) >= 16 Then

                    ws_name = dt1.ToString("yyyy_MM")

                ElseIf Integer.Parse(dt1.ToString("MM")) = 1 Then

                    ws_name = dt_year & "_" & dt_month

                Else

                    ws_name = dt1.ToString("yyyy_") & dt_month

                End If

                Dim ws As Excel.Worksheet = ss(ws_name)
                Dim cs As Excel.Range = ws.Cells

                If Integer.Parse(dt1.ToString("dd")) <= 15 Then

                    Dim lastday As Integer = Integer.Parse(Date.DaysInMonth(dt1.ToString("yyyy"), dt1.AddMonths(-1).ToString))
                    Dim sif As Integer = lastday + Integer.Parse(dt1.ToString("dd")) - 12
                    ws.Range("G" & sif).Value = ComboBox1.Text

                Else

                    Dim sif As Integer = Integer.Parse(dt1.ToString("dd")) - 12
                    ws.Range("G" & sif).Value = ComboBox1.Text

                End If

                ea.DisplayAlerts = False
                wb.Save()
                ea.Quit()

                Dim form3 As Form3 = New Form3
                form3.U_name = U_name
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