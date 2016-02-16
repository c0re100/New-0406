<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="data.asp"-->
<%
info=Session("info")
if info(0)="" then response.redirect"../../error.asp?id=210"
sql="SELECT 識別,行動點 FROM 寵物狀態 where user='"&info(0)&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('您的寵物已經死了或您不是這隻寵物的主人！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("識別")>=rs("行動點") then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的寵物累了該歇歇了！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
name=request("name")
sql="select 攻擊,防御,體力 from 道具列表 where id="&name&" and 擁有者='"&info(0)&"'"
set rs=conn.execute(sql)
at=rs("攻擊")
guard=rs("防御")
power=rs("體力")

sql="update 寵物狀態 set 識別=識別+1,體力=體力+"&power&",攻擊=攻擊+"&at&",防御=防御+"&guard&" where user='"&info(0)&"'"
conn.Execute(sql)
conn.execute("delete from 道具列表 where id="&name&" and 擁有者='"&info(0)&"'")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>道具使用</title>
<link rel="stylesheet" href="setup.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#3a4b91" text="#000000" background="../../linjianww/0064.jpg">
<div align="center">
<center>
<br>
<br>
</center>
</div>
<div align="center">
<center>
<table border="1" width="380" cellspacing="1" cellpadding="0" height="173" bordercolor="#000000">
<tr>
<td class="p2" width="100%">
<p align="center"><font size="2">您已經成功使用了道具 ，刷新即可看到效果！<br>
<br>
攻擊增加<%=at%> 防御增加<%=guard%> 體力增加<%=power%><br>
<font color="#000000"><br>
</font></font>
<table border="0" width="320" cellspacing="0" cellpadding="0" align="center">
<td class="p3" width="100%" height="19">
<p align="center"><font color="#000000"><a href="feedsheep.asp">&gt;&gt;返回</a>
</font>
</td>
</table>

</td>
</tr>
</table>
</center>
</div>
</body>

</html>