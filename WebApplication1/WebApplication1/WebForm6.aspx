<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm6.aspx.vb" Inherits="WebApplication1.WebForm6" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div align="center">
            <asp:Label ID="Label1" runat="server" Text="ユーザー設定"></asp:Label>
            <br />
            <br />
            <asp:Label ID="Label2" runat="server" Text="ユーザー名：" MaxLength="20"></asp:Label>
            <asp:TextBox ID="TextBox1" runat="server" MaxLength="20"></asp:TextBox>
             &nbsp;
             &nbsp;
            <asp:Label ID="Label3" runat="server" Text="パスワード："></asp:Label>
            <asp:TextBox ID="TextBox2" runat="server" MaxLength="5" OnTextChanged="TextBox2_TextChanged"></asp:TextBox>
            &nbsp;
            &nbsp;
            <asp:Label ID="Label4" runat="server" Text="店舗名："></asp:Label>
            <asp:TextBox ID="TextBox3" runat="server" MaxLength="10"></asp:TextBox>
            <br />
            <br />
            <asp:Button ID="Button1" runat="server" Text="登録" OnClick="Button1_Click" />
            <br />
            <br />
            <asp:Button ID="Button3" runat="server" Text="ユーザー名で検索" OnClick="Button3_Click" />
            &nbsp;
            &nbsp;
            <asp:Button ID="Button4" runat="server" Text="店舗名で検索" OnClick="Button4_Click" />
            <br />
            <br />
            <asp:GridView ID="GridView1" runat="server" AutoGenerateEditButton="true" OnRowCommand="GridView1_RowCommand" OnRowEditing="GridView1_RowEditing" OnRowCancelingEdit="GridView1_RowCancelingEdit" OnRowUpdating="GridView1_RowUpdating">

            </asp:GridView>
            <br />
            <asp:Button ID="Button2" runat="server" Text="戻る" OnClick="Button2_Click" />
        </div>
    </form>
</body>
</html>
