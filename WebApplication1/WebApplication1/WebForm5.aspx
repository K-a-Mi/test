<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm5.aspx.vb" Inherits="WebApplication1.WebForm5" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div align="center">
            <asp:Label ID="Label1" runat="server" Text="集計"></asp:Label>
            <br />
            <br />
            <asp:TextBox ID="TextBox1" runat="server" Style="width: 10%;" MaxLength="4" oninput="value = value.replace(/[０-９]/g,s => String.fromCharCode(s.charCodeAt(0) - 65248)).replace(/\D/g,'');" OnTextChanged="TextBox1_TextChanged" AutoPostBack="True"></asp:TextBox>
            <asp:Label ID="Label2" runat="server" Text="年"></asp:Label>
            <asp:DropDownList ID="DropDownList1" runat="server" Style="width: 10%;" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged1"  AutoPostBack="True" Enabled="False"></asp:DropDownList>
            <script>
//onchange="changeB()"
                function changeB() {
                    let ddl1 = document.getElementById('DropDownList1');
                    let ddl2 = document.getElementById('DropDownList2');
                    if (ddl1.value == '--') {
                        ddl2.disabled = true
                    } else {
                        ddl2.disabled = false
                    }
                }
            </script>
            <asp:Label ID="Label3" runat="server" Text="月"></asp:Label>
            <asp:DropDownList ID="DropDownList2" runat="server" Style="width: 10%;" Enabled="False"></asp:DropDownList>
            <asp:Label ID="Label4" runat="server" Text="日"></asp:Label>
            <asp:Label ID="Label5" runat="server" Text="～"></asp:Label>
            <asp:TextBox ID="TextBox2" runat="server" Style="width: 10%;" MaxLength="4" oninput="value = value.replace(/[０-９]/g,s => String.fromCharCode(s.charCodeAt(0) - 65248)).replace(/\D/g,'');" AutoPostBack="True" OnTextChanged="TextBox2_TextChanged"></asp:TextBox>
            <asp:Label ID="Label6" runat="server" Text="年"></asp:Label>
            <asp:DropDownList ID="DropDownList3" runat="server" Style="width: 10%;"  OnSelectedIndexChanged="DropDownList3_SelectedIndexChanged" AutoPostBack="True" Enabled="False"></asp:DropDownList>
            <script>
//onchange="changeA()"
                function changeA() {
                    let ddl3 = document.getElementById('DropDownList3');
                    let ddl4 = document.getElementById('DropDownList4');
                    if (ddl3.value == '--') {
                        ddl4.disabled = true
                    } else {
                        ddl4.disabled = false
                    }
                }
            </script>
            <asp:Label ID="Label7" runat="server" Text="月"></asp:Label>
            <asp:DropDownList ID="DropDownList4" runat="server" Style="width: 10%;" Enabled="False"></asp:DropDownList>
            <asp:Label ID="Label8" runat="server" Text="日"></asp:Label>
            <br />
            <br />
            <asp:Label ID="Label9" runat="server" Text="ユーザーネーム："></asp:Label>
            <asp:TextBox ID="TextBox3" runat="server"></asp:TextBox>
            <asp:Button ID="Button3" runat="server" Text="指定エクスポート" style="display: inline" OnClick="Button3_Click" onclientclick="EN_C2()"/>
            <asp:Button ID="Button5" runat="server" Text="指定エクスポート"  style="display: none" Enabled="false"/>
            <script>
//実際処理が走るボタンを非表示にして、ダミーのボタンを非活性で表示させる
                let ex_btn3 = document.getElementById("Button3");
                let btn5 = document.getElementById("Button5");
                function EN_C2() {
                    ex_btn3.style.display = 'none';
                    btn5.style.display = 'inline';
                }
            </script>
            <br />
            <br />
            <asp:Button ID="Button1" runat="server" Text="一括エクスポート" style="display: inline" OnClick="Button1_Click" OnClientClick="EN_C()"/>
            <asp:Button ID="Button4" runat="server" Text="一括エクスポート" style="display: none" Enabled="false"/>
            <script>
//OnClientClick="EN_C()"
//実際処理が走るボタンを非表示にして、ダミーのボタンを非活性で表示させる
                let btn1 = document.getElementById("Button1");
                let btn0 = document.getElementById("Button4");
                function EN_C() {
                    btn1.style.display = 'none';
                    btn0.style.display = 'inline';
                }
            </script>
            <br />
            <br />
            <asp:Button ID="Button2" runat="server" Text="戻る" OnClick="Button2_Click" />
        </div>
    </form>
</body>
</html>
