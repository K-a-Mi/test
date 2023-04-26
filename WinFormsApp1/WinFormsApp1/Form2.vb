Imports Microsoft.Office.Interop

Public Class Form2

    Public U_name As String
    '本日の日時取得
    Dim dt1 As DateTime = DateTime.Now


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

        'コンボボックス入力チェック
        If ComboBox1.Text <> "" Then

            Dim dialog As DialogResult = MessageBox.Show(" シフトは　" & ComboBox1.Text & "　で間違いはないですか？", "確認メッセージ", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If dialog = DialogResult.Yes Then

                Dim filename = "C:\Users\User\" & U_name & ".xlsx"

                Dim ea As Excel.Application = New Excel.Application
                Dim wbs As Excel.Workbooks = ea.Workbooks
                Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")

                Dim ss As Excel.Sheets = wb.Worksheets
                Dim ws As Excel.Worksheet = ss(dt1.ToString("yyyy_MM"))
                Dim cs As Excel.Range = ws.Cells

                Dim sif As Integer = Integer.Parse(dt1.ToString("dd")) + 3

                ws.Range("G" & sif).Value = ComboBox1.Text

                ea.DisplayAlerts = False
                wb.Save()
                ea.Quit()

                Dim form3 As Form3 = New Form3
                form3.U_name = U_name
                form3.ShowDialog()

                Me.Close()

            Else



            End If
        Else

            MsgBox("シフトを選択してください。")

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Close()

    End Sub
End Class