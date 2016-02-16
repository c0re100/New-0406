<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="jhb.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "請進入聊天室再進入股票！"
window.close()
</script>
<%Response.End
end if

Jname=info(0)
sql="delete * from 交易 where 時間<date()"
conn.execute sql
sql="select * from 交易 where 帳號='" & Jname & "' "
set rs=conn.execute(sql)

%>
<body bgcolor="#990000"><table border=1 bgcolor="#bee0fc" align=center width=550 cellpadding="10" cellspacing="13"> 
<option><p align=center><font size=5 color=#00ff00></font></p>
	<tr><td bgcolor=#cce8d6>
		<table border=0 cellpadding=1 cellspacing=1 bgcolor="#51a8ff" width="100%" >		</table>
		
		<P align=center><front size=3pt>經紀人今天為您代理的交易如下，扣除1%的傭金就是您在股市的收益</front></p>
		<table border=0 cellpadding=1 cellspacing=1 bgcolor="#51a8ff" width="100%" 
     >
	
		
		
        <P align=center></P>
        <TBODY>
			<tr bgcolor=#c4deff>
				<td>股票名稱</td><td>買賣操作</td><td>交易數量</td><td>時間</td><td>收益</td>
			</tr>
<%DO UNTIL RS.EOF%>	
			<tr bgcolor=#f7f7f7>
				<td><%=rs("企業")%>
				<td><%=rs("操作")%></td>
				<td><%=rs("交易量")%></td>
				<td><%=rs("時間")%></td>
				<td><%=formatnumber(rs("收獲"),0)%></td>
				
		    </tr>
<%
rs.movenext
loop
set rs=nothing
conn.close
set conn=nothing
%></p>

		</table>