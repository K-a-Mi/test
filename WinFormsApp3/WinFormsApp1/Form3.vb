Imports System.Net.Mail
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form3

    Public U_name As String
    Public ws_name As String
    Public get_cell As Integer
    Dim dt2 As DateTime = DateTime.Now
    Dim S_time As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Close()

    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label1.Text = U_name
        Label2.Text = dt2.ToString("yyyy/MM/dd")
        Button2.Enabled = True
        Button3.Enabled = True

        Dim N_month As String = dt2.AddMonths(-1).ToString("MM")
        Dim N_year As String = dt2.AddYears(-1).ToString("yyyy")

        Dim filename = "C:\Users\User\" & U_name & ".xlsx"
        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")
        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(ws_name)

        '本日出勤ボタン押していたら押せないようにする
        If ws.Range("G" & get_cell).Value <> "" Then

            If ws.Range("H" & get_cell).Value Is Nothing Then

                Button3.Enabled = False

            ElseIf ws.Range("H" & get_cell).Value.ToString <> "" Then

                If ws.Range("I" & get_cell).Value Is Nothing Then

                    Button2.Enabled = False

                ElseIf ws.Range("I" & get_cell).Value.ToString <> "" Then

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

        '現在時刻を取得
        dt2 = DateTime.Now

        Dim filename = "C:\Users\User\" & U_name & ".xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")
        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(ws_name)
        Dim cs As Excel.Range = ws.Cells

        ws.Range("H" & get_cell).Value = dt2.ToString("HH:mm")

        ea.DisplayAlerts = False
        wb.Save()
        ea.Quit()

        MsgBox("おはようございます。")

        Button2.Enabled = False

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        '現在時刻を取得
        dt2 = DateTime.Now

        Dim filename = "C:\Users\User\" & U_name & ".xlsx"
        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")
        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(ws_name)
        Dim cs As Excel.Range = ws.Cells

        ws.Range("I" & get_cell).Value = dt2.ToString("HH:mm")

        ea.DisplayAlerts = False
        wb.Save()
        ea.Quit()

        MsgBox("お疲れさまでした。")

        Button3.Enabled = False

    End Sub

End Class