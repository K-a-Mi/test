Imports Npgsql
Imports Microsoft.Office.Interop
Imports System.IO

Public Class WebForm5
    Inherits System.Web.UI.Page
    Public conn As New NpgsqlConnection("Server=192.168.0.111; Port=5432; User Id=postgres; Password=brains; Database=brains")
    Public reader As NpgsqlDataReader
    Dim dt As Date = DateTime.Now
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If DropDownList1.Items.Count < 1 And DropDownList3.Items.Count < 1 Then

            DropDownList1.Items.Add("--")
            DropDownList3.Items.Add("--")


            If DropDownList1.Items.Count = 1 And DropDownList3.Items.Count = 1 Then

                Dim i As Integer = 1
                While i <= 12

                    DropDownList1.Items.Add(i.ToString("00"))
                    DropDownList3.Items.Add(i.ToString("00"))
                    i = i + 1

                End While

            End If

            DropDownList1.SelectedIndex = 0
            DropDownList3.SelectedIndex = 0

        End If

    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs)

        Response.Redirect("WebForm4.aspx")

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs)

        Try


            If TextBox1.Text <> "" And (DropDownList1.SelectedValue <> "" And DropDownList1.SelectedValue <> "--") And (DropDownList2.SelectedValue <> "" And DropDownList2.SelectedValue <> "--") And TextBox2.Text <> "" And (DropDownList3.SelectedValue <> "" And DropDownList3.SelectedValue <> "--") And (DropDownList4.SelectedValue <> "" And DropDownList4.SelectedValue <> "--") Then

                If Integer.Parse(TextBox1.Text & DropDownList1.SelectedValue & DropDownList2.SelectedValue) < Integer.Parse(TextBox2.Text & DropDownList3.SelectedValue & DropDownList4.SelectedValue) Then

                    conn.Open()
                    Dim sheetname As String = ""

                    '有効ユーザー確認
                    Dim U_count As Integer = 0
                    Dim U_name() As String
                    Dim T_name() As String
                    Dim U_Chk As String = "select * from tokumt where mukouflg = '1' ;"
                    Dim Chk_U As NpgsqlCommand = New NpgsqlCommand(U_Chk, conn)
                    reader = Chk_U.ExecuteReader()

                    While reader.Read() = True
                        '有効ユーザーと店舗名を配列に格納する
                        ReDim Preserve U_name(U_count)
                        ReDim Preserve T_name(U_count)
                        U_name(U_count) = Trim(reader("tokumei"))
                        T_name(U_count) = Trim(reader("tenpo"))
                        U_count = U_count + 1

                    End While

                    reader.Close()
                    Dim Create_fu As Integer = 0

                    While Create_fu < U_count

                        If File.Exists("C:\Users\User\" & U_name(Create_fu).ToString.Trim() & ".xlsx") = False Then

                            Dim filename As String = "C:\Users\User\sample.xlsx"
                            'ユーザーファイルが存在してなければ原本コピーして作成する
                            File.Copy(filename, "C:\Users\User\" & U_name(Create_fu).ToString.Trim() & ".xlsx")

                            'copyでプロセス残るので消す(これ以外の方法わからなかった)
                            Dim ps As Process() = Process.GetProcessesByName("EXCEL")
                            For Each p As Process In ps
                                p.Kill()
                            Next

                        End If

                        Dim ea As Excel.Application = New Excel.Application
                        Dim wbs As Excel.Workbooks = ea.Workbooks
                        Dim wb As Excel.Workbook = wbs.Open("C:\Users\User\" & U_name(Create_fu).ToString.Trim() & ".xlsx", Password:="sample")
                        Dim ss As Excel.Sheets = wb.Worksheets
                        Dim sheetcount As Integer = ss.Count
                        Dim ws As Excel.Worksheet = ss("原紙")
                        Dim Ex_sheetname As String = ""

                        '原紙シートをコピー
                        ws.Copy(After:=ws)
                        'コピーしたシートを選択
                        ws = ss("原紙 (2)")
                        '名前の変更
                        ws.Name = "【" & sheetcount.ToString() & "】 作成日 (" & dt.ToString("yyyy_MM_dd") & "_" & dt.ToString("HH_mm") & ")"

                        'シフト表を取得
                        Dim s As Integer = 4
                        Dim S_count As Integer = 0
                        Dim ST_TS As String = "=IF(G4=$B$4,IF(H4>$C$4,CEILING(H4,""0:15""),$C$4),"""")"
                        Dim KK_TS As String = "=IF(G4=$B$4,$E$4,"""")"
                        Dim T_shift As String = "select * from " & T_name(Create_fu) & ";"
                        Dim TS_set As NpgsqlCommand = New NpgsqlCommand(T_shift, conn)
                        reader = TS_set.ExecuteReader()
                        'シフトを入力
                        While reader.Read() = True

                            S_count = S_count + 1
                            ws.Range("B" & s).Value = reader("pattern").ToString
                            ws.Range("C" & s).Value = reader("s_kin").ToString
                            ws.Range("D" & s).Value = reader("t_kin").ToString
                            ws.Range("E" & s).Value = reader("k_kei").ToString
                            s = s + 1

                        End While

                        s = 4
                        Dim while_count As Integer = 0
                        'シフトパターン数によって式を変更する
                        If S_count > 1 Then

                            ST_TS = "=IF(G4=$B$4,IF(H4>$C$4,CEILING(H4,""0:15""),$C$4),"
                            KK_TS = "=IF(G4=$B$4,$E$4,"

                            While S_count > 1

                                s = s + 1
                                while_count = while_count + 1
                                ST_TS = ST_TS & "IF(G4=$B$" & s & ",IF(H4>$C$" & s & ",CEILING(H4,""0:15""),$C$" & s & "),"
                                KK_TS = KK_TS & "IF(G4=$B$" & s & ",$E$" & s & ","
                                S_count = S_count - 1

                            End While

                            ST_TS = ST_TS & """"""
                            ST_TS = ST_TS & ")"
                            KK_TS = KK_TS & """"""
                            KK_TS = KK_TS & ")"

                            While while_count > 0

                                ST_TS = ST_TS & ")"
                                KK_TS = KK_TS & ")"
                                while_count = while_count - 1

                            End While

                        End If
                        '式を入力する
                        ws.Range("K4").Value = ST_TS
                        ws.Range("M4").Value = KK_TS
                        reader.Close()

                        '日付をセットする
                        ws.Range("K2").Value = TextBox1.Text & "/" & DropDownList1.SelectedValue & "/" & DropDownList2.SelectedValue
                        ws.Range("M2").Value = TextBox2.Text & "/" & DropDownList3.SelectedValue & "/" & DropDownList4.SelectedValue

                        Dim L_year As String = TextBox1.Text
                        Dim L_month As String = DropDownList1.SelectedValue
                        Dim L_day As String = DropDownList2.SelectedValue
                        Dim R_year As String = TextBox2.Text
                        Dim R_month As String = DropDownList3.SelectedValue
                        Dim R_day As String = DropDownList4.SelectedValue
                        Dim E_count As Integer = 5
                        Dim G_D As Integer = Integer.Parse(DropDownList2.SelectedValue) + 1

                        ws.Range("J4").NumberFormat = "mm/dd"
                        ws.Range("J4").Value = DropDownList1.SelectedValue & "/" & DropDownList2.SelectedValue

                        If Integer.Parse(L_year & L_month) <> Integer.Parse(R_year & R_month) Then

                            Dim M_ld As Integer = Date.DaysInMonth(Integer.Parse(TextBox1.Text), Integer.Parse(DropDownList1.SelectedValue))

                            While M_ld >= G_D

                                ws.Range("J" & E_count).NumberFormat = "mm/dd"
                                ws.Range("J" & E_count).Value = (L_month & "/" & G_D).ToString
                                E_count = E_count + 1
                                G_D = G_D + 1

                            End While

                            If L_month = "12" Then

                                L_month = 1.ToString("00")
                                L_year = (Integer.Parse(L_year) + 1).ToString("0000")

                            Else

                                L_month = (Integer.Parse(L_month) + 1).ToString("00")

                            End If

                            While Integer.Parse(L_year & L_month) < Integer.Parse(R_year & R_month)

                                Dim D_count As Integer = 1
                                M_ld = 0
                                M_ld = Date.DaysInMonth(Integer.Parse(L_year), Integer.Parse(L_month))

                                While D_count <= M_ld

                                    ws.Range("J" & E_count).NumberFormat = "mm/dd"
                                    ws.Range("J" & E_count).Value = (L_month & "/" & D_count).ToString
                                    E_count = E_count + 1
                                    D_count = D_count + 1

                                End While

                                If L_month <> "12" Then

                                    L_month = (Integer.Parse(L_month) + 1).ToString("00")

                                Else

                                    L_month = 1.ToString("00")
                                    L_year = (Integer.Parse(L_year) + 1).ToString("0000")

                                End If

                            End While

                            If Integer.Parse(L_year & L_month) = Integer.Parse(R_year & R_month) Then

                                Dim D_count As Integer = 1
                                While D_count <= Integer.Parse(R_day)

                                    ws.Range("J" & E_count).NumberFormat = "mm/dd"
                                    ws.Range("J" & E_count).Value = (R_month & "/" & D_count).ToString
                                    E_count = E_count + 1
                                    D_count = D_count + 1

                                End While

                            End If

                        Else

                            While Integer.Parse(R_day) >= G_D

                                ws.Range("J" & E_count).NumberFormat = "mm/dd"
                                ws.Range("J" & E_count).Value = (L_month & "/" & G_D).ToString
                                E_count = E_count + 1
                                G_D = G_D + 1

                            End While

                        End If

                        Dim TC_Ttl As String = "Select * from time_stamp where tokumei = '" & Trim(U_name(Create_fu)) & "' and s_dt between '" & TextBox1.Text & "/" & DropDownList1.SelectedValue & "/" & DropDownList2.SelectedValue & "' and '" & TextBox2.Text & "/" & DropDownList3.SelectedValue & "/" & DropDownList4.SelectedValue & "' order by s_dt;"

                        Dim TS_reader As NpgsqlDataReader
                        Dim Ttl_TC As NpgsqlCommand = New NpgsqlCommand(TC_Ttl, conn)
                        Dim Ex_select As Integer = 4

                        TS_reader = Ttl_TC.ExecuteReader()
                        '入力された日付に合わせてデータを入力していく
                        While TS_reader.Read() = True
                            While Strings.Right(ws.Range("j" & Ex_select).Value, 5) <> Strings.Right(TS_reader("s_dt"), 5)

                                ws.Range("G" & Ex_select).Value = ""
                                ws.Range("H" & Ex_select).Value = ""
                                ws.Range("I" & Ex_select).Value = ""

                                Ex_select = Ex_select + 1

                            End While

                            ws.Range("G" & Ex_select).Value = TS_reader("shift")
                            ws.Range("H" & Ex_select).Value = TS_reader("s_time")
                            ws.Range("I" & Ex_select).Value = TS_reader("t_time")

                            ws.Range("K" & Ex_select).NumberFormat = "hh:mm"
                            ws.Range("L" & Ex_select).NumberFormat = "hh:mm"
                            ws.Range("M" & Ex_select).NumberFormat = "hh:mm"
                            ws.Range("N" & Ex_select).NumberFormat = "hh:mm"

                            Ex_select = Ex_select + 1

                        End While

                        If Ex_select > 4 Then

                            Ex_select = Ex_select - 1

                            'オートフィル設定
                            ws.Range("K4").AutoFill(ws.Range("K4", "K" & Ex_select.ToString), Excel.XlAutoFillType.xlFillValues)
                            ws.Range("L4").AutoFill(ws.Range("L4", "L" & Ex_select.ToString), Excel.XlAutoFillType.xlFillValues)
                            ws.Range("M4").AutoFill(ws.Range("M4", "M" & Ex_select.ToString), Excel.XlAutoFillType.xlFillValues)
                            ws.Range("N4").AutoFill(ws.Range("N4", "N" & Ex_select.ToString), Excel.XlAutoFillType.xlFillValues)

                            ws.Range("Q3").Value = "=COUNT(N4:N" & Ex_select.ToString
                            ws.Range("Q4").Value = "=SUM(N4:N" & Ex_select.ToString & ")"

                        Else

                            ws.Range("Q3").Value = "=COUNT(N4)"
                            ws.Range("Q4").Value = "=SUM(N4)"

                        End If
                        TS_reader.Close()

                        ea.DisplayAlerts = False
                        '保存
                        wb.Save()
                        '終了
                        ea.Quit()

                        Create_fu = Create_fu + 1

                    End While

                    Dim SC_msg As String = "alert('完了しました。')"
                    ClientScript.RegisterStartupScript(Me.GetType(), "SC_msg", SC_msg, True)

                Else

                    Dim D_erro As String = "alert('正しい期間を入力してください。')"
                    ClientScript.RegisterStartupScript(Me.GetType(), "D_erro", D_erro, True)

                End If

            Else

                Dim N_erro As String = "alert('期間を入力してください。')"
                ClientScript.RegisterStartupScript(Me.GetType(), "N_erro", N_erro, True)

            End If

        Catch ex As Exception

        End Try
        conn.Close()
    End Sub

    Protected Sub DropDownList3_SelectedIndexChanged(sender As Object, e As EventArgs)

        Try
            '月が選択されたら日数を取得してDropDownListに追加していく
            If DropDownList3.SelectedIndex <> 0 Then

                DropDownList4.Enabled = True

                Dim M_day As Integer = Date.DaysInMonth(Integer.Parse(TextBox2.Text), Integer.Parse(DropDownList3.SelectedValue))
                Dim i As Integer = 1
                While i <= M_day

                    DropDownList4.Items.Add(i.ToString("00"))
                    i = i + 1

                End While

            Else

                DropDownList4.Enabled = False

            End If


        Catch ex As Exception

        End Try

    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged1(sender As Object, e As EventArgs)

        Try
            '月が選択されたら日数を取得してDropDownListに追加していく
            If DropDownList1.SelectedIndex <> 0 Then

                DropDownList2.Enabled = True

                Dim M_day As Integer = Date.DaysInMonth(Integer.Parse(TextBox1.Text), Integer.Parse(DropDownList1.SelectedValue))
                Dim i As Integer = 1
                While i <= M_day

                    DropDownList2.Items.Add(i.ToString("00"))
                    i = i + 1

                End While

            Else

                DropDownList2.Enabled = False

            End If


        Catch ex As Exception

        End Try

    End Sub

    Protected Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

        Try

            If TextBox1.Text <> "" Then

                DropDownList1.Enabled = True
                DropDownList1.SelectedIndex = 0
                DropDownList2.Items.Clear()
                DropDownList2.Enabled = False

            Else

                DropDownList1.Enabled = False
                DropDownList2.Enabled = False
                DropDownList1.SelectedIndex = 0
                DropDownList2.Items.Clear()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub TextBox2_TextChanged(sender As Object, e As EventArgs)

        Try

            If TextBox2.Text <> "" Then

                DropDownList3.Enabled = True
                DropDownList3.SelectedIndex = 0
                DropDownList4.Items.Clear()
                DropDownList4.Enabled = False

            Else

                DropDownList3.Enabled = False
                DropDownList4.Enabled = False
                DropDownList3.SelectedIndex = 0
                DropDownList4.Items.Clear()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Sub Button3_Click(sender As Object, e As EventArgs)

        Try
            conn.Open()

            If TextBox1.Text <> "" And (DropDownList1.SelectedValue <> "" And DropDownList1.SelectedValue <> "--") And (DropDownList2.SelectedValue <> "" And DropDownList2.SelectedValue <> "--") And TextBox2.Text <> "" And (DropDownList3.SelectedValue <> "" And DropDownList3.SelectedValue <> "--") And (DropDownList4.SelectedValue <> "" And DropDownList4.SelectedValue <> "--") And TextBox3.Text <> "" Then
                If Integer.Parse(TextBox1.Text & DropDownList1.SelectedValue & DropDownList2.SelectedValue) < Integer.Parse(TextBox2.Text & DropDownList3.SelectedValue & DropDownList4.SelectedValue) Then

                    Dim Chk As String = "select * from tokumt where tokumei = '" & TextBox3.Text & "';"
                    Dim ST_Chk As NpgsqlCommand = New NpgsqlCommand(Chk, conn)
                    reader = ST_Chk.ExecuteReader()

                    If reader.Read() = True Then
                        If reader("mukouflg") <> "0" Then

                            Dim sheetname As String = ""
                            reader.Close()
                            '有効ユーザー確認
                            Dim U_count As Integer = 0
                            Dim U_name() As String
                            Dim T_name() As String
                            Dim U_Chk As String = "select * from tokumt where tokumei = '" & TextBox3.Text & "' and mukouflg = '1' ;"
                            Dim Chk_U As NpgsqlCommand = New NpgsqlCommand(U_Chk, conn)
                            reader = Chk_U.ExecuteReader()

                            While reader.Read() = True
                                '有効ユーザーと店舗名を配列に格納する
                                ReDim Preserve U_name(U_count)
                                ReDim Preserve T_name(U_count)
                                U_name(U_count) = Trim(reader("tokumei"))
                                T_name(U_count) = Trim(reader("tenpo"))
                                U_count = U_count + 1

                            End While

                            reader.Close()
                            Dim Create_fu As Integer = 0

                            While Create_fu < U_count

                                If File.Exists("C:\Users\User\" & U_name(Create_fu).ToString.Trim() & ".xlsx") = False Then

                                    Dim filename As String = "C:\Users\User\sample.xlsx"
                                    'ユーザーファイルが存在してなければ原本コピーして作成する
                                    File.Copy(filename, "C:\Users\User\" & U_name(Create_fu).ToString.Trim() & ".xlsx")

                                    'copyでプロセス残るので消す(これ以外の方法わからなかった)
                                    Dim ps As Process() = Process.GetProcessesByName("EXCEL")
                                    For Each p As Process In ps
                                        p.Kill()
                                    Next

                                End If

                                Dim ea As Excel.Application = New Excel.Application
                                Dim wbs As Excel.Workbooks = ea.Workbooks
                                Dim wb As Excel.Workbook = wbs.Open("C:\Users\User\" & U_name(Create_fu).ToString.Trim() & ".xlsx", Password:="sample")
                                Dim ss As Excel.Sheets = wb.Worksheets
                                Dim sheetcount As Integer = ss.Count
                                Dim ws As Excel.Worksheet = ss("原紙")
                                Dim Ex_sheetname As String = ""

                                '原紙シートをコピー
                                ws.Copy(After:=ws)
                                'コピーしたシートを選択
                                ws = ss("原紙 (2)")
                                '名前の変更
                                ws.Name = "【" & sheetcount.ToString() & "】 作成日 (" & dt.ToString("yyyy_MM_dd") & "_" & dt.ToString("HH_mm") & ")"

                                'シフト表を取得
                                Dim s As Integer = 4
                                Dim S_count As Integer = 0
                                Dim ST_TS As String = "=IF(G4=$B$4,IF(H4>$C$4,CEILING(H4,""0:15""),$C$4),"""")"
                                Dim KK_TS As String = "=IF(G4=$B$4,$E$4,"""")"
                                Dim T_shift As String = "select * from " & T_name(Create_fu) & ";"
                                Dim TS_set As NpgsqlCommand = New NpgsqlCommand(T_shift, conn)
                                reader = TS_set.ExecuteReader()
                                'シフトを入力
                                While reader.Read() = True

                                    S_count = S_count + 1
                                    ws.Range("B" & s).Value = reader("pattern").ToString
                                    ws.Range("C" & s).Value = reader("s_kin").ToString
                                    ws.Range("D" & s).Value = reader("t_kin").ToString
                                    ws.Range("E" & s).Value = reader("k_kei").ToString
                                    s = s + 1

                                End While

                                s = 4
                                Dim while_count As Integer = 0
                                'シフトパターン数によって式を変更する
                                If S_count > 1 Then

                                    ST_TS = "=IF(G4=$B$4,IF(H4>$C$4,CEILING(H4,""0:15""),$C$4),"
                                    KK_TS = "=IF(G4=$B$4,$E$4,"

                                    While S_count > 1

                                        s = s + 1
                                        while_count = while_count + 1
                                        ST_TS = ST_TS & "IF(G4=$B$" & s & ",IF(H4>$C$" & s & ",CEILING(H4,""0:15""),$C$" & s & "),"
                                        KK_TS = KK_TS & "IF(G4=$B$" & s & ",$E$" & s & ","
                                        S_count = S_count - 1

                                    End While

                                    ST_TS = ST_TS & """"""
                                    ST_TS = ST_TS & ")"
                                    KK_TS = KK_TS & """"""
                                    KK_TS = KK_TS & ")"

                                    While while_count > 0

                                        ST_TS = ST_TS & ")"
                                        KK_TS = KK_TS & ")"
                                        while_count = while_count - 1

                                    End While

                                End If
                                '式を入力する
                                ws.Range("K4").Value = ST_TS
                                ws.Range("M4").Value = KK_TS
                                reader.Close()

                                '日付をセットする
                                ws.Range("K2").Value = TextBox1.Text & "/" & DropDownList1.SelectedValue & "/" & DropDownList2.SelectedValue
                                ws.Range("M2").Value = TextBox2.Text & "/" & DropDownList3.SelectedValue & "/" & DropDownList4.SelectedValue

                                Dim L_year As String = TextBox1.Text
                                Dim L_month As String = DropDownList1.SelectedValue
                                Dim L_day As String = DropDownList2.SelectedValue
                                Dim R_year As String = TextBox2.Text
                                Dim R_month As String = DropDownList3.SelectedValue
                                Dim R_day As String = DropDownList4.SelectedValue
                                Dim E_count As Integer = 5
                                Dim G_D As Integer = Integer.Parse(DropDownList2.SelectedValue) + 1

                                ws.Range("J4").NumberFormat = "mm/dd"
                                ws.Range("J4").Value = DropDownList1.SelectedValue & "/" & DropDownList2.SelectedValue

                                If Integer.Parse(L_year & L_month) <> Integer.Parse(R_year & R_month) Then

                                    Dim M_ld As Integer = Date.DaysInMonth(Integer.Parse(TextBox1.Text), Integer.Parse(DropDownList1.SelectedValue))

                                    While M_ld >= G_D

                                        ws.Range("J" & E_count).NumberFormat = "mm/dd"
                                        ws.Range("J" & E_count).Value = (L_month & "/" & G_D).ToString
                                        E_count = E_count + 1
                                        G_D = G_D + 1

                                    End While

                                    If L_month = "12" Then

                                        L_month = 1.ToString("00")
                                        L_year = (Integer.Parse(L_year) + 1).ToString("0000")

                                    Else

                                        L_month = (Integer.Parse(L_month) + 1).ToString("00")

                                    End If

                                    While Integer.Parse(L_year & L_month) < Integer.Parse(R_year & R_month)

                                        Dim D_count As Integer = 1
                                        M_ld = 0
                                        M_ld = Date.DaysInMonth(Integer.Parse(L_year), Integer.Parse(L_month))

                                        While D_count <= M_ld

                                            ws.Range("J" & E_count).NumberFormat = "mm/dd"
                                            ws.Range("J" & E_count).Value = (L_month & "/" & D_count).ToString
                                            E_count = E_count + 1
                                            D_count = D_count + 1

                                        End While

                                        If L_month <> "12" Then

                                            L_month = (Integer.Parse(L_month) + 1).ToString("00")

                                        Else

                                            L_month = 1.ToString("00")
                                            L_year = (Integer.Parse(L_year) + 1).ToString("0000")

                                        End If

                                    End While

                                    If Integer.Parse(L_year & L_month) = Integer.Parse(R_year & R_month) Then

                                        Dim D_count As Integer = 1
                                        While D_count <= Integer.Parse(R_day)

                                            ws.Range("J" & E_count).NumberFormat = "mm/dd"
                                            ws.Range("J" & E_count).Value = (R_month & "/" & D_count).ToString
                                            E_count = E_count + 1
                                            D_count = D_count + 1

                                        End While

                                    End If

                                Else

                                    While Integer.Parse(R_day) >= G_D

                                        ws.Range("J" & E_count).NumberFormat = "mm/dd"
                                        ws.Range("J" & E_count).Value = (L_month & "/" & G_D).ToString
                                        E_count = E_count + 1
                                        G_D = G_D + 1

                                    End While

                                End If

                                Dim TC_Ttl As String = "Select * from time_stamp where tokumei = '" & Trim(U_name(Create_fu)) & "' and s_dt between '" & TextBox1.Text & "/" & DropDownList1.SelectedValue & "/" & DropDownList2.SelectedValue & "' and '" & TextBox2.Text & "/" & DropDownList3.SelectedValue & "/" & DropDownList4.SelectedValue & "' order by s_dt;"

                                Dim TS_reader As NpgsqlDataReader
                                Dim Ttl_TC As NpgsqlCommand = New NpgsqlCommand(TC_Ttl, conn)
                                Dim Ex_select As Integer = 4

                                TS_reader = Ttl_TC.ExecuteReader()
                                '入力された日付に合わせてデータを入力していく
                                While TS_reader.Read() = True
                                    While Strings.Right(ws.Range("j" & Ex_select).Value, 5) <> Strings.Right(TS_reader("s_dt"), 5)

                                        ws.Range("G" & Ex_select).Value = ""
                                        ws.Range("H" & Ex_select).Value = ""
                                        ws.Range("I" & Ex_select).Value = ""

                                        Ex_select = Ex_select + 1

                                    End While

                                    ws.Range("G" & Ex_select).Value = TS_reader("shift")
                                    ws.Range("H" & Ex_select).Value = TS_reader("s_time")
                                    ws.Range("I" & Ex_select).Value = TS_reader("t_time")

                                    ws.Range("K" & Ex_select).NumberFormat = "hh:mm"
                                    ws.Range("L" & Ex_select).NumberFormat = "hh:mm"
                                    ws.Range("M" & Ex_select).NumberFormat = "hh:mm"
                                    ws.Range("N" & Ex_select).NumberFormat = "hh:mm"

                                    Ex_select = Ex_select + 1

                                End While

                                If Ex_select > 4 Then

                                    Ex_select = Ex_select - 1

                                    'オートフィル設定
                                    ws.Range("K4").AutoFill(ws.Range("K4", "K" & Ex_select.ToString), Excel.XlAutoFillType.xlFillValues)
                                    ws.Range("L4").AutoFill(ws.Range("L4", "L" & Ex_select.ToString), Excel.XlAutoFillType.xlFillValues)
                                    ws.Range("M4").AutoFill(ws.Range("M4", "M" & Ex_select.ToString), Excel.XlAutoFillType.xlFillValues)
                                    ws.Range("N4").AutoFill(ws.Range("N4", "N" & Ex_select.ToString), Excel.XlAutoFillType.xlFillValues)

                                    ws.Range("Q3").Value = "=COUNT(N4:N" & Ex_select.ToString
                                    ws.Range("Q4").Value = "=SUM(N4:N" & Ex_select.ToString & ")"

                                Else

                                    ws.Range("Q3").Value = "=COUNT(N4)"
                                    ws.Range("Q4").Value = "=SUM(N4)"

                                End If
                                TS_reader.Close()

                                ea.DisplayAlerts = False
                                '保存
                                wb.Save()
                                '終了
                                ea.Quit()

                                Create_fu = Create_fu + 1

                            End While

                            Dim SC_msg As String = "alert('完了しました。')"
                            ClientScript.RegisterStartupScript(Me.GetType(), "SC_msg", SC_msg, True)

                        Else

                            Dim N_U = "alert('無効なユーザーです。')"
                            ClientScript.RegisterStartupScript(Me.GetType, "N_U", N_U, True)

                        End If

                    Else

                        Dim N_U = "alert('登録されていません。')"
                        ClientScript.RegisterStartupScript(Me.GetType, "N_U", N_U, True)

                    End If

                Else

                    Dim D_erro As String = "alert('正しい期間を入力してください。')"
                    ClientScript.RegisterStartupScript(Me.GetType(), "D_erro", D_erro, True)

                End If

            Else

                Dim N_erro As String = "alert('期間・ユーザーを入力してください。')"
                ClientScript.RegisterStartupScript(Me.GetType(), "N_erro", N_erro, True)


            End If

        Catch ex As Exception

        End Try
        conn.Close()
    End Sub
End Class