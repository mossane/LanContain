<%@ Page Language="C#" AutoEventWireup="true"  EnableEventValidation="false" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>LanContain</title>
</head>
<body>
    <form id="form2" runat="server">
    <div>
        <center>
        <asp:Label ID="Label1" runat="server" Text="IP地址段：" Font-Names="宋体" Font-Size="Small"></asp:Label><asp:TextBox ID="TextBox1" runat="server" Font-Names="宋体" Font-Size="Small" Width="108px"></asp:TextBox><asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="扫描" Font-Names="宋体" Font-Size="Small" />&nbsp;</center>
        <center>
            <asp:Label ID="Label3" runat="server" Font-Names="宋体" Font-Size="Small" Text="用户："></asp:Label><asp:TextBox ID="TextBox2" runat="server" Font-Names="宋体" Font-Size="Small" Width="130px"></asp:TextBox>&nbsp;</center>
        <center>
            <asp:Label ID="Label4" runat="server" Font-Names="宋体" Font-Size="Small" Text="密码："></asp:Label><asp:TextBox ID="TextBox3" runat="server" Font-Names="宋体" Font-Size="Small" AutoCompleteType="Pager" TextMode="Password" Width="130px"></asp:TextBox>&nbsp;</center>
        <center>
        <asp:CheckBox ID="CheckBox1" runat="server" Font-Names="宋体" Font-Size="Small" Text="DNS扫描" />
        <asp:CheckBox ID="CheckBox2" runat="server" Font-Names="宋体" Font-Size="Small" Text="WMI扫描" /></center>
        <center>
            &nbsp;</center>
        <center>
        <asp:Label ID="Label2" runat="server" Font-Names="宋体" Font-Size="Small" ForeColor="Red"></asp:Label><br />
        </center>
        <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="删除所选" Width="70px" Font-Names="宋体" Font-Size="Small" />
        <asp:Button ID="Button3" runat="server" OnClick="Button3_Click" Text="导出Excel" 
            Font-Names="宋体" Font-Size="Small" Width="80px" /><br />
        <center>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" 
                Font-Names="宋体" Font-Size="Small" DataKeyNames="序号" AllowSorting="True" 
                OnSorting="GridView1_Sorting" OnRowDeleting="GridView1_RowDeleting" 
                OnRowDataBound="GridView1_RowDataBound">
            <Columns>
                <asp:BoundField DataField="序号" HeaderText="序号" SortExpression="序号" >
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="IP地址" HeaderText="IP地址" SortExpression="IP地址" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="计算机名" HeaderText="计算机名" SortExpression="计算机名" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="工作组(域)" HeaderText="工作组(域)" SortExpression="工作组(域)" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="登录名" HeaderText="登录名" SortExpression="登录名" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="MAC地址" HeaderText="MAC地址" SortExpression="MAC地址" />
                <asp:BoundField DataField="TTL" HeaderText="TTL" SortExpression="TTL" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="回传" HeaderText="回传" SortExpression="回传" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="子网掩码" HeaderText="子网掩码" SortExpression="子网掩码" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="默认网关" HeaderText="默认网关" SortExpression="默认网关" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="DNS服务器1" HeaderText="DNS服务器1" SortExpression="DNS服务器1" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="DNS服务器2" HeaderText="DNS服务器2" SortExpression="DNS服务器2" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="DNS服务器3" HeaderText="DNS服务器3" SortExpression="DNS服务器3" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="WINS服务器1" HeaderText="WINS服务器1" SortExpression="WINS服务器1" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="WINS服务器2" HeaderText="WINS服务器2" SortExpression="WINS服务器2" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="日期时间" HeaderText="日期时间" SortExpression="日期时间" >
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:TemplateField HeaderText="选择" ShowHeader="False">
                    <headertemplate>
                        <asp:CheckBox id="chkSelectAll" runat="server" OnCheckedChanged="chkSelectAll_CheckedChanged" AutoPostBack="True"></asp:CheckBox>
                    </headertemplate>
                    <itemtemplate>
                        <asp:CheckBox id="chkSelect" runat="server" OnCheckedChanged="chkSelect_CheckedChanged" AutoPostBack="True"></asp:CheckBox>
                    </itemtemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="删除" ShowHeader="False">
                    <headertemplate>
                    </headertemplate>
                    <ItemTemplate>
                    <asp:LinkButton ID="BtnDelete" runat="server" CausesValidation="False" CommandName="Delete" 
                Text="删除" OnClientClick="return confirm('确认要删除此行信息吗？')"></asp:LinkButton>
                    </ItemTemplate> 
                </asp:TemplateField>
            </Columns>
            <RowStyle HorizontalAlign="Justify" />
        </asp:GridView>
            &nbsp;
        </center>
        </div>
    </form>
</body>
</html>