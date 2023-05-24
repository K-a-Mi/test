Public Class WebForm4
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs)

        Response.Redirect("WebForm5.aspx")

    End Sub

    Protected Sub Button5_Click(sender As Object, e As EventArgs)

        Response.Redirect("WebForm1.aspx")

    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs)

        Response.Redirect("WebForm6.aspx")

    End Sub
End Class