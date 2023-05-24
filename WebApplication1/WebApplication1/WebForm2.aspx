<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm2.aspx.vb" Inherits="WebApplication1.WebForm2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server" defaultfocus="DropDownList1">
        <div align="center">
            <asp:Label ID="Label1" runat="server" Text="ログインユーザー："></asp:Label>
            <asp:Label ID="Label2" runat="server" Text=""></asp:Label>
            <br />
            <br />
            <asp:DropDownList ID="DropDownList1" runat="server" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged">
            </asp:DropDownList>
            &nbsp;
            &nbsp;
            
            <input type="button" ID="btn2" class="btn2" value="進む" />
            <script>
                let btn3 = document.getElementById('btn2');
                let lab2 = document.getElementById('Label2')
                let D_list = document.getElementById('DropDownList1');
                //asp.net(vb)から変数受け取り
                var data = '<%=data %>';

                btn3.onclick = function () {
                    if (D_list.value != '<未選択>') {
                        let result = window.confirm('シフトは ' + D_list.value + ' で間違いないですか？');
                        if (result == true) { window.location.href = 'WebForm3.aspx?name=' + lab2.innerText + '&tenpo=' + data + '&shift=' + D_list.value; }
                    } else {
                        let check = window.alert('シフトを選択してください。');
                    }
                }
            </script>
            <br />
            <br />
            <asp:Button ID="Button1" runat="server" Text="戻る" OnClick="Button1_Click" />
        </div>
    </form>
</body>
</html>
