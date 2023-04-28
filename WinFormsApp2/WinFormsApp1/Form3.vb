Imports System.Net.Mail
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form3

    Public U_name As String
    Dim dt2 As DateTime = DateTime.Now
    Dim S_time As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Close()

    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label1.Text = U_name
        Label2.Text = dt2.ToString("yyyy/MM/dd")

        Dim filename = "C:\Users\User\" & U_name & ".xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(dt2.ToString("yyyy_MM"))
        'Dim cs As Excel.Range = ws.Cell

        '打刻入力セル設定
        Call set_cell()

        '本日出勤ボタン押していたら押せないようにする
        If ws.Range("G" & S_time).Value <> "" Then

            If ws.Range("H" & S_time).Value Is Nothing Then

            ElseIf ws.Range("H" & S_time).Value.ToString <> "" Then

                If ws.Range("I" & S_time).Value Is Nothing Then

                    Button2.Enabled = False

                ElseIf ws.Range("I" & S_time).Value.ToString <> "" Then

                    Button2.Enabled = False
                    Button3.Enabled = False

                End If

            End If

        End If

        ea.DisplayAlerts = False
        wb.Save()
        ea.Quit()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        dt2 = DateTime.Now

        Dim filename = "C:\Users\User\" & U_name & ".xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(dt2.ToString("yyyy_MM"))
        Dim cs As Excel.Range = ws.Cells

        'If Integer.Parse(dt2.ToString("dd")) <= 15 Then

        'Dim lastday As Integer = Integer.Parse(Date.DaysInMonth(dt2.ToString("yyyy"), dt2.AddMonths(-1).ToString))
        'S_time = lastday + Integer.Parse(dt2.ToString("dd")) - 12

        'Else

        'S_time = Integer.Parse(dt2.ToString("dd")) - 12

        'End If

        ws.Range("H" & S_time).Value = dt2.ToString("HH:mm")


        ea.DisplayAlerts = False
        wb.Save()
        ea.Quit()

        MsgBox("おはようございます。")

        Button2.Enabled = False

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        dt2 = DateTime.Now

        Dim filename = "C:\Users\User\" & U_name & ".xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(dt2.ToString("yyyy_MM"))
        Dim cs As Excel.Range = ws.Cells

        '入力セル設定
        'If Integer.Parse(dt2.ToString("dd")) <= 15 Then

        'Dim lastday As Integer = Integer.Parse(Date.DaysInMonth(dt2.ToString("yyyy"), dt2.AddMonths(-1).ToString))
        'S_time = lastday + Integer.Parse(dt2.ToString("dd")) - 12

        'Else

        'S_time = Integer.Parse(dt2.ToString("dd")) - 12

        'End If

        ws.Range("I" & S_time).Value = dt2.ToString("HH:mm")

        ea.DisplayAlerts = False
        wb.Save()
        ea.Quit()

        MsgBox("お疲れさまでした。")

        Button3.Enabled = False

    End Sub

    Sub set_cell()

        If Integer.Parse(dt2.ToString("dd")) <= 15 Then

            Dim lastday As Integer = Integer.Parse(Date.DaysInMonth(dt2.ToString("yyyy"), dt2.AddMonths(-1).ToString))
            S_time = lastday + Integer.Parse(dt2.ToString("dd")) - 12

        Else

            S_time = Integer.Parse(dt2.ToString("dd")) - 12

        End If

    End Sub
End Class