Imports System.Net.Mail
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form3

    Public U_name As String
    Dim dt2 As DateTime = DateTime.Now
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


        Dim S_time As Integer = Integer.Parse(dt2.ToString("dd")) + 3

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

        Dim filename = "C:\Users\User\" & U_name & ".xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(dt2.ToString("yyyy_MM"))
        Dim cs As Excel.Range = ws.Cells

        Dim S_time As Integer = Integer.Parse(dt2.ToString("dd")) + 3

        ws.Range("H" & S_time).Value = dt2.ToString("HH:mm")

        ea.DisplayAlerts = False
        wb.Save()
        ea.Quit()

        MsgBox("おはようございます。")

        Button2.Enabled = False

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim filename = "C:\Users\User\" & U_name & ".xlsx"

        Dim ea As Excel.Application = New Excel.Application
        Dim wbs As Excel.Workbooks = ea.Workbooks
        Dim wb As Excel.Workbook = wbs.Open(filename, Password:="sample")

        Dim ss As Excel.Sheets = wb.Worksheets
        Dim ws As Excel.Worksheet = ss(dt2.ToString("yyyy_MM"))
        Dim cs As Excel.Range = ws.Cells

        Dim S_time As Integer = Integer.Parse(dt2.ToString("dd")) + 3

        ws.Range("I" & S_time).Value = dt2.ToString("HH:mm")

        ea.DisplayAlerts = False
        wb.Save()
        ea.Quit()

        MsgBox("お疲れさまでした。")

        Button3.Enabled = False

    End Sub
End Class