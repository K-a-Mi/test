Imports Npgsql

Public Class WebForm6
    Inherits System.Web.UI.Page
    Public conn As New NpgsqlConnection("Server=192.168.0.111; Port=5432; User Id=postgres; Password=brains; Database=brains")
    Public reader As NpgsqlDataReader
    Public index As Integer = 0
    Public pass As String = ""
    Public tokumei As String = ""
    Public tenpo As String = ""
    Public flg As String = ""
    Public tenryak As String = ""
    Dim Upd_t1 As String = ""
    Dim Upd_t2 As String = ""
    Protected column_1 As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs)
        '########################################################################
        '############################             登録用           ###############################
        '########################################################################
        Try

            If TextBox1.Text.Length <= 20 And TextBox1.Text.Length > 0 Then
                If TextBox2.Text.Length = 5 Then
                    If TextBox3.Text.Length <= 10 And TextBox3.Text.Length > 0 Then
                        conn.Open()

                        Dim Tname_Chk As String = "select * from tenpomt where tenmei = '" & TextBox3.Text & "';"
                        Dim T_Chk As NpgsqlCommand = New NpgsqlCommand(Tname_Chk, conn)
                        reader = T_Chk.ExecuteReader()

                        If reader.Read() = True Then

                            Dim Insert_U As String = "insert into tokumt values('" & TextBox2.Text & "','" & TextBox1.Text & "','" & TextBox3.Text & "', '1' , now() ,'" & reader("tenryaku") & "');"
                            reader.Close()
                            Dim I_User As NpgsqlCommand = New NpgsqlCommand(Insert_U, conn)
                            reader = I_User.ExecuteReader()

                            Dim Scsses As String = "alert('登録しました。')"
                            ClientScript.RegisterStartupScript(Me.GetType(), "Scsses", Scsses, True)

                            TextBox1.Text = ""
                            TextBox2.Text = ""
                            TextBox3.Text = ""

                        Else

                            Dim Tname_e As String = "alert('店舗が存在していません。')"
                            ClientScript.RegisterStartupScript(Me.GetType(), "Tnamer_e", Tname_e, True)

                        End If

                    Else

                        Dim Tname_e As String = "alert('店舗名は10文字以下で入力してください。')"
                        ClientScript.RegisterStartupScript(Me.GetType(), "Tname_e", Tname_e, True)

                    End If

                Else

                    Dim Pass_e As String = "alert('パスワードは5文字で入力してください。')"
                    ClientScript.RegisterStartupScript(Me.GetType(), "Pass_e", Pass_e, True)

                End If

            Else

                Dim Uname_e As String = "alert('ユーザー名は20文字以下で入力してください。')"
                ClientScript.RegisterStartupScript(Me.GetType(), "Uname_e", Uname_e, True)

            End If
        Catch ex As Exception

        End Try
        conn.Close()
    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs)

        Response.Redirect("WebForm4.aspx")

    End Sub

    Protected Sub TextBox2_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Protected Sub Button3_Click(sender As Object, e As EventArgs)
        '########################################################################
        '#########################             ユーザー名で検索           ############################
        '########################################################################
        Try
            If TextBox1.Text.Length <= 20 And TextBox1.Text.Length > 0 Then

                conn.Open()

                Dim command As String = "select tokucd as パスワード,tokumei as ユーザー名,tokuryak as 店舗名,mukouflg as 無効FLG,tenpo as 店舗略称 from tokumt where tokumei = '" & TextBox1.Text & "';"
                Dim cmd As NpgsqlCommand = New NpgsqlCommand(command, conn)
                Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter(cmd)
                Dim dt As DataTable = New DataTable()
                da.Fill(dt)
                reader = cmd.ExecuteReader()

                If reader.Read() = True Then

                    GridView1.DataSource = dt
                    '編集ボタン付ける
                    GridView1.AutoGenerateEditButton = True
                    GridView1.DataBind()
                    column_1 = GridView1.HeaderRow.Cells(1).Text


                Else

                    Dim Uname_e As String = "alert('登録されていません。')"
                    ClientScript.RegisterStartupScript(Me.GetType(), "Uname_e", Uname_e, True)

                End If

            Else

                Dim Uname_e As String = "alert('ユーザー名は20文字以下で入力してください。')"
                ClientScript.RegisterStartupScript(Me.GetType(), "Uname_e", Uname_e, True)

            End If

        Catch ex As Exception

        End Try
        conn.Close()

    End Sub

    Protected Sub Button4_Click(sender As Object, e As EventArgs)
        '########################################################################
        '##########################             店舗名で検索           #############################
        '########################################################################
        Try
            If TextBox3.Text.Length <= 10 And TextBox3.Text.Length > 0 Then

                conn.Open()

                Dim command As String = "select tokucd as パスワード,tokumei as ユーザー名,tokuryak as 店舗名,mukouflg as 無効FLG,tenpo as 店舗略称 from tokumt where tokuryak = '" & TextBox3.Text & "';"
                Dim cmd As NpgsqlCommand = New NpgsqlCommand(command, conn)
                Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter(cmd)
                Dim dt As DataTable = New DataTable()
                da.Fill(dt)
                reader = cmd.ExecuteReader()

                If reader.Read() = True Then

                    GridView1.DataSource = dt
                    GridView1.AutoGenerateEditButton = True
                    GridView1.EditIndex = -1
                    GridView1.DataBind()

                Else

                    Dim Uname_e As String = "alert('存在しない店舗名です。')"
                    ClientScript.RegisterStartupScript(Me.GetType(), "Uname_e", Uname_e, True)

                End If

            Else

                Dim Uname_e As String = "alert('店舗名は10文字以下で入力してください。')"
                ClientScript.RegisterStartupScript(Me.GetType(), "Uname_e", Uname_e, True)

            End If

        Catch ex As Exception

        End Try
        reader.Close()
        conn.Close()
    End Sub

    Protected Sub GridView1_RowEditing(sender As Object, e As GridViewEditEventArgs)

        'GridView1.EditIndex = -1
        Dim column_c As Integer = GridView1.HeaderRow.Cells.Count
        'Dim bound As BoundField = GridView1.HeaderRow.Cells(1)
        'bound.ReadOnly = True
        GridView1.Rows(index).Cells(1).Text = pass
        GridView1.Rows(index).Cells(2).Text = tokumei
        'GridView1.Rows(index).Cells(3).Text = tenpo
        'GridView1.Rows(index).Cells(4).Text = flg
        GridView1.Rows(index).Cells(5).Text = tenryak

        'GridView1.EditIndex = index

    End Sub

    Protected Sub GridView1_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs)


    End Sub

    Protected Sub GridView1_RowCommand(sender As Object, e As GridViewCommandEventArgs)

        If e.CommandName = "Edit" Then

            index = Convert.ToInt32(e.CommandArgument)
            pass = GridView1.Rows(index).Cells(1).Text
            tokumei = GridView1.Rows(index).Cells(2).Text
            tenpo = GridView1.Rows(index).Cells(3).Text
            flg = GridView1.Rows(index).Cells(4).Text
            tenryak = GridView1.Rows(index).Cells(5).Text

            Dim column_1 As String = GridView1.HeaderRow.Cells(1).Text
            Dim dolumn_c As Integer = GridView1.HeaderRow.Cells.Count

        ElseIf e.CommandName = "Update" Then

            Dim editindex As Integer = CType(sender, GridView).EditIndex
            Dim row1 As GridViewRow = CType(sender, GridView).Rows(editindex)
            Dim id As String = row1.Cells(1).Text

            index = Convert.ToInt32(e.CommandArgument)
            tokumei = GridView1.Rows(index).Cells(2).Text

            For Each control As Control In row1.Cells(3).Controls
                If (TypeOf control Is TextBox) Then

                    Upd_t1 = CType(control, TextBox).Text

                End If
            Next

            For Each control As Control In row1.Cells(4).Controls
                If (TypeOf control Is TextBox) Then

                    Upd_t2 = CType(control, TextBox).Text

                End If
            Next


        End If

    End Sub

    Protected Sub GridView1_RowUpdating(sender As Object, e As GridViewUpdateEventArgs)

        Try
            conn.Open()

            Dim command As String = "select tokucd as パスワード,tokumei as ユーザー名,tokuryak as 店舗名,mukouflg as 無効FLG,tenpo as 店舗略称 from tokumt where tokuryak = '" & TextBox3.Text & "';"

            Dim Tname_Chk As String = "select * from tenpomt where tenmei = '" & Trim(Upd_t1) & "';"
            Dim T_Chk As NpgsqlCommand = New NpgsqlCommand(Tname_Chk, conn)
            reader = T_Chk.ExecuteReader()

            If reader.Read() = True Then

                Dim update_U As String = "update tokumt set tenpo = '" & reader("tenryaku") & "', mukouflg = '" & Upd_t2 & "', tokuryak = '" & Upd_t1 & "' where tokumei = '" & Trim(GridView1.Rows(index).Cells(2).Text) & "';"
                reader.Close()
                Dim U_update As NpgsqlCommand = New NpgsqlCommand(update_U, conn)
                reader = U_update.ExecuteReader()
                reader.Close()

                Dim update_S As String = "alert('更新しました。')"
                ClientScript.RegisterStartupScript(Me.GetType(), "update_S", update_S, True)

                If TextBox3.Text.Length <= 10 And TextBox3.Text.Length > 0 Then

                    Dim cmd As NpgsqlCommand = New NpgsqlCommand(command, conn)
                    Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter(cmd)
                    Dim dt As DataTable = New DataTable()
                    da.Fill(dt)

                    reader = cmd.ExecuteReader()

                    If reader.Read() = True Then

                        reader.Close()
                        GridView1.DataSource = dt
                        GridView1.AutoGenerateEditButton = True
                        GridView1.EditIndex = -1
                        GridView1.DataBind()

                    Else

                        Dim Uname_e As String = "alert('存在しない店舗名です。')"
                        ClientScript.RegisterStartupScript(Me.GetType(), "Uname_e", Uname_e, True)

                    End If

                Else

                    Dim Uname_e As String = "alert('店舗名は10文字以下で入力してください。')"
                    ClientScript.RegisterStartupScript(Me.GetType(), "Uname_e", Uname_e, True)

                End If

            Else

                reader.Close()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand(command, conn)
                Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter(cmd)
                Dim dt As DataTable = New DataTable()
                da.Fill(dt)

                GridView1.DataSource = dt
                GridView1.AutoGenerateEditButton = True
                GridView1.EditIndex = -1
                GridView1.DataBind()

                Dim Tname_e As String = "alert('店舗が存在していません。')"
                ClientScript.RegisterStartupScript(Me.GetType(), "Tnamer_e", Tname_e, True)

            End If

        Catch ex As Exception

        End Try
        reader.Close()
        conn.Close()
    End Sub
End Class